package com.example.outlook

import com.example.outlook.data.GraphApi
import com.example.outlook.data.GraphMailRepository
import com.example.outlook.domain.OutlookError
import com.example.outlook.domain.OutlookException
import kotlinx.coroutines.test.runTest
import kotlinx.serialization.SerializationException
import org.junit.Assert.assertEquals
import org.junit.Assert.assertFalse
import org.junit.Assert.assertNull
import org.junit.Assert.assertTrue
import org.junit.Assert.fail
import org.junit.Test
import java.io.IOException
import java.time.Instant
import kotlin.time.Duration
import kotlin.time.Duration.Companion.seconds

class GraphMailRepositoryTest {

    private val api = FakeGraphApi()
    private val auth = FakeAuth()
    private val sleeps = mutableListOf<Duration>()
    private val now = Instant.parse("2026-10-24T12:00:00Z")
    private val repo = GraphMailRepository(api, auth, 10.seconds, clock = { now }, sleep = { sleeps += it })

    private suspend fun expectError(expected: OutlookError, block: suspend () -> Unit) {
        try {
            block()
            fail("expected $expected")
        } catch (e: OutlookException) {
            assertEquals(expected, e.error)
        }
    }

    @Test
    fun `parses first page with spec query and bearer token`() = runTest {
        api.queueList(fixture("inbox_page1.json"))

        val page = repo.listInbox(pageSize = 25, pageToken = null)

        assertEquals(
            "list top=25 select=${GraphApi.LIST_SELECT} orderBy=${GraphApi.LIST_ORDER_BY}",
            api.calls.single(),
        )
        assertEquals("Bearer token-1", api.authorizations.single())
        assertEquals(2, page.messages.size)
        val first = page.messages[0]
        assertEquals("AAMkAGI2THVSAAA=", first.id)
        assertEquals("Quarterly planning", first.subject)
        assertEquals("Alex Wilber", first.sender?.displayName)
        assertEquals(Instant.parse("2026-10-20T14:05:11Z"), first.receivedAt)
        assertFalse(first.isRead)
        val second = page.messages[1]
        assertEquals("", second.subject)
        assertEquals("noreply@contoso.com", second.sender?.displayName)
        assertTrue(page.nextPageToken!!.startsWith("https://graph.microsoft.com/v1.0/"))
    }

    @Test
    fun `follows nextLink and ends on last page`() = runTest {
        api.queueList(fixture("inbox_page1.json"))
        api.queueList(fixture("inbox_page2.json"))

        val page1 = repo.listInbox(25, null)
        val page2 = repo.listInbox(25, page1.nextPageToken)

        assertEquals("next ${page1.nextPageToken}", api.calls[1])
        assertEquals("Last page", page2.messages.single().subject)
        assertNull(page2.nextPageToken)
    }

    @Test
    fun `clamps page size`() = runTest {
        api.queueList(fixture("inbox_page2.json"))
        repo.listInbox(pageSize = 5000, pageToken = null)
        assertTrue(api.calls.single().startsWith("list top=100 "))
    }

    @Test(expected = IllegalArgumentException::class)
    fun `rejects page token for another host`() = runTest {
        repo.listInbox(25, "https://evil.example.com/steal")
    }

    @Test
    fun `message detail has html and plain text bodies`() = runTest {
        api.queueDetail(fixture("message_detail.json"))

        val message = repo.getMessage("AAMkAGI2THVSAAA=")

        assertTrue(api.calls.single().contains("prefer=${GraphApi.PREFER_HTML_BODY}"))
        assertEquals(listOf("Adele Vance", "team@contoso.com"), message.toRecipients.map { it.displayName })
        assertTrue(message.body.html!!.contains("<ul>"))
        assertEquals("Hi team,\n\nAgenda:\n\n• Budget & hiring\n• Roadmap\n\nThanks\nAlex", message.body.text)
    }

    @Test
    fun `plain text body has no html`() = runTest {
        api.queueDetail("""{"id":"x","body":{"contentType":"text","content":"hello"}}""")
        val body = repo.getMessage("x").body
        assertNull(body.html)
        assertEquals("hello", body.text)
    }

    @Test
    fun `401 refreshes token once then succeeds`() = runTest {
        api.listResponses.addLast(httpError(401))
        api.queueList(fixture("inbox_page2.json"))

        repo.listInbox(25, null)

        assertEquals(listOf(false, true), auth.tokenRequests)
        assertEquals(listOf("Bearer token-1", "Bearer token-2"), api.authorizations)
    }

    @Test
    fun `second 401 surfaces AuthRequired`() = runTest {
        api.listResponses.addLast(httpError(401))
        api.listResponses.addLast(httpError(401))
        expectError(OutlookError.AuthRequired) { repo.listInbox(25, null) }
        assertEquals(2, api.calls.size)
    }

    @Test
    fun `no account surfaces AuthRequired without calling Graph`() = runTest {
        auth.account = null
        expectError(OutlookError.AuthRequired) { repo.listInbox(25, null) }
        assertTrue(api.calls.isEmpty())
    }

    @Test
    fun `consent failure from token acquisition is passed through`() = runTest {
        auth.tokenError = OutlookError.ConsentRequired
        expectError(OutlookError.ConsentRequired) { repo.getMessage("x") }
    }

    @Test
    fun `short 429 waits Retry-After and retries once`() = runTest {
        api.listResponses.addLast(httpError(429, mapOf("Retry-After" to "3")))
        api.queueList(fixture("inbox_page2.json"))

        repo.listInbox(25, null)

        assertEquals(listOf(3.seconds), sleeps)
        assertEquals(2, api.calls.size)
    }

    @Test
    fun `long 429 surfaces Throttled with Retry-After`() = runTest {
        api.listResponses.addLast(httpError(429, mapOf("Retry-After" to "120")))
        expectError(OutlookError.Throttled(120.seconds)) { repo.listInbox(25, null) }
        assertTrue(sleeps.isEmpty())
    }

    @Test
    fun `repeated 429 surfaces Throttled after one retry`() = runTest {
        api.listResponses.addLast(httpError(429, mapOf("Retry-After" to "2")))
        api.listResponses.addLast(httpError(429, mapOf("Retry-After" to "2")))
        expectError(OutlookError.Throttled(2.seconds)) { repo.listInbox(25, null) }
        assertEquals(listOf(2.seconds), sleeps)
    }

    @Test
    fun `Retry-After parsing handles seconds, dates and missing values`() {
        assertEquals(7.seconds, repo.retryAfter("7"))
        assertEquals(30.seconds, repo.retryAfter("Sat, 24 Oct 2026 12:00:30 GMT"))
        assertEquals(GraphMailRepository.DEFAULT_RETRY_AFTER, repo.retryAfter(null))
        assertEquals(GraphMailRepository.DEFAULT_RETRY_AFTER, repo.retryAfter("soon"))
    }

    @Test
    fun `network failure maps to Network`() = runTest {
        api.listResponses.addLast(IOException("offline"))
        expectError(OutlookError.Network) { repo.listInbox(25, null) }
    }

    @Test
    fun `403 maps to ConsentRequired`() = runTest {
        api.detailResponses.addLast(httpError(403))
        expectError(OutlookError.ConsentRequired) { repo.getMessage("x") }
    }

    @Test
    fun `other HTTP errors and bad payloads map to Unknown`() = runTest {
        api.listResponses.addLast(httpError(500))
        expectError(OutlookError.Unknown("Graph HTTP 500")) { repo.listInbox(25, null) }

        api.listResponses.addLast(SerializationException("bad"))
        expectError(OutlookError.Unknown("Unexpected Graph response format")) { repo.listInbox(25, null) }
    }
}
