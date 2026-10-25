package com.example.outlook

import com.example.outlook.data.GraphMailRepository
import com.example.outlook.domain.OutlookAccount
import com.example.outlook.domain.OutlookError
import com.example.outlook.domain.OutlookResult
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.test.runTest
import org.junit.Assert.assertEquals
import org.junit.Assert.assertNull
import org.junit.Assert.assertTrue
import org.junit.Test

class OutlookMailClientTest {

    private val api = FakeGraphApi()
    private val auth = FakeAuth()
    private val accountState = MutableStateFlow<OutlookAccount?>(null)
    private val client = OutlookMailClient(auth, GraphMailRepository(api, auth, sleep = {}), accountState)

    @Test
    fun `refreshAccount publishes current account`() = runTest {
        assertEquals("user@contoso.com", client.refreshAccount()?.username)
        assertEquals("user@contoso.com", accountState.value?.username)
    }

    @Test
    fun `listInbox wraps success`() = runTest {
        api.queueList(fixture("inbox_page1.json"))
        val result = client.listInbox(25, null)
        assertEquals(2, (result as OutlookResult.Success).value.messages.size)
    }

    @Test
    fun `AuthRequired clears account state`() = runTest {
        client.refreshAccount()
        api.listResponses.addLast(httpError(401))
        api.listResponses.addLast(httpError(401))

        val result = client.listInbox(25, null)

        assertEquals(OutlookError.AuthRequired, result.errorOrNull())
        assertNull(accountState.value)
    }

    @Test
    fun `invalid page token becomes Unknown failure, not a crash`() = runTest {
        val result = client.listInbox(25, "https://example.com/next")
        assertTrue(result.errorOrNull() is OutlookError.Unknown)
    }

    @Test
    fun `signOut clears MSAL account and account state`() = runTest {
        client.refreshAccount()
        val result = client.signOut()
        assertTrue(result is OutlookResult.Success)
        assertEquals(1, auth.signOutCalls)
        assertNull(accountState.value)
        assertNull(client.refreshAccount())
    }

    @Test
    fun `getMessage maps throttling`() = runTest {
        api.detailResponses.addLast(httpError(429, mapOf("Retry-After" to "60")))
        assertEquals(OutlookError.Throttled(kotlin.time.Duration.parse("60s")), client.getMessage("x").errorOrNull())
    }
}
