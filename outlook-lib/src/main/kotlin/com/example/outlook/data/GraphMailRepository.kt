package com.example.outlook.data

import com.example.outlook.auth.AuthProvider
import com.example.outlook.domain.EmailAddress
import com.example.outlook.domain.Message
import com.example.outlook.domain.MessageBody
import com.example.outlook.domain.MessagePage
import com.example.outlook.domain.MessageSummary
import com.example.outlook.domain.OutlookError
import com.example.outlook.domain.OutlookException
import kotlinx.coroutines.delay
import kotlinx.serialization.SerializationException
import retrofit2.HttpException
import java.io.IOException
import java.time.Instant
import java.time.ZonedDateTime
import java.time.format.DateTimeFormatter
import java.time.format.DateTimeParseException
import kotlin.time.Duration
import kotlin.time.Duration.Companion.seconds

/**
 * Reads Inbox mail through Microsoft Graph. Handles token attachment, one forced token refresh on HTTP 401,
 * Retry-After on HTTP 429, and maps every failure to a typed [OutlookError].
 */
internal class GraphMailRepository(
    private val api: GraphApi,
    private val auth: AuthProvider,
    private val maxAutoRetryWait: Duration = 10.seconds,
    private val clock: () -> Instant = Instant::now,
    private val sleep: suspend (Duration) -> Unit = { delay(it) },
) {

    suspend fun listInbox(pageSize: Int, pageToken: String?): MessagePage {
        val dto = if (pageToken == null) {
            val top = pageSize.coerceIn(1, MAX_PAGE_SIZE)
            withGraph { authz -> api.listInbox(authorization = authz, top = top) }
        } else {
            require(pageToken.startsWith(GRAPH_ORIGIN)) { "pageToken was not issued by this library" }
            withGraph { authz -> api.listInboxNextPage(authorization = authz, nextLink = pageToken) }
        }
        return MessagePage(
            messages = dto.value.map { it.toSummary() },
            nextPageToken = dto.nextLink?.takeIf { it.startsWith(GRAPH_ORIGIN) },
        )
    }

    suspend fun getMessage(id: String): Message {
        require(id.isNotBlank()) { "id is required" }
        val dto = withGraph { authz -> api.getMessage(authorization = authz, id = id) }
        return Message(
            summary = dto.toSummary().copy(id = dto.id.ifEmpty { id }),
            toRecipients = dto.toRecipients.mapNotNull { it.toEmailAddress() },
            body = dto.body.toMessageBody(),
        )
    }

    /** Runs a Graph call with a bearer token, retrying once on 401 (fresh token) and once on a short 429. */
    private suspend fun <T> withGraph(call: suspend (authorization: String) -> T): T {
        var token = auth.acquireToken(forceRefresh = false)
        var refreshedToken = false
        var retriedThrottle = false
        while (true) {
            try {
                return call("Bearer $token")
            } catch (e: HttpException) {
                when (e.code()) {
                    401 -> {
                        if (refreshedToken) throw OutlookException(OutlookError.AuthRequired, e)
                        refreshedToken = true
                        token = auth.acquireToken(forceRefresh = true)
                    }
                    429 -> {
                        val wait = retryAfter(e.response()?.headers()?.get("Retry-After"))
                        if (retriedThrottle || wait > maxAutoRetryWait) {
                            throw OutlookException(OutlookError.Throttled(wait), e)
                        }
                        retriedThrottle = true
                        sleep(wait)
                    }
                    403 -> throw OutlookException(OutlookError.ConsentRequired, e)
                    else -> throw OutlookException(OutlookError.Unknown("Graph HTTP ${e.code()}"), e)
                }
            } catch (e: IOException) {
                throw OutlookException(OutlookError.Network, e)
            } catch (e: SerializationException) {
                throw OutlookException(OutlookError.Unknown("Unexpected Graph response format"), e)
            }
        }
    }

    /** Retry-After is either delta-seconds or an HTTP date. */
    internal fun retryAfter(header: String?): Duration {
        val value = header?.trim().orEmpty()
        value.toLongOrNull()?.let { return it.coerceAtLeast(0).seconds }
        if (value.isNotEmpty()) {
            try {
                val at = ZonedDateTime.parse(value, DateTimeFormatter.RFC_1123_DATE_TIME).toInstant()
                return (at.epochSecond - clock().epochSecond).coerceAtLeast(0).seconds
            } catch (_: DateTimeParseException) {
                // Fall through to the default.
            }
        }
        return DEFAULT_RETRY_AFTER
    }

    private fun MessageDto.toSummary() = MessageSummary(
        id = id,
        subject = subject.orEmpty(),
        sender = from?.toEmailAddress(),
        receivedAt = receivedDateTime?.let { runCatching { Instant.parse(it) }.getOrNull() },
        preview = bodyPreview.orEmpty(),
        isRead = isRead,
    )

    private fun RecipientDto.toEmailAddress(): EmailAddress? =
        emailAddress?.let { EmailAddress(name = it.name, address = it.address) }

    private fun ItemBodyDto?.toMessageBody(): MessageBody {
        val content = this?.content.orEmpty()
        return if (this?.contentType.equals("html", ignoreCase = true)) {
            MessageBody(html = content, text = HtmlText.toPlainText(content))
        } else {
            MessageBody(html = null, text = content)
        }
    }

    companion object {
        const val GRAPH_ORIGIN = "https://graph.microsoft.com/"
        const val MAX_PAGE_SIZE = 100
        val DEFAULT_RETRY_AFTER: Duration = 5.seconds
    }
}
