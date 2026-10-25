package com.example.outlook.domain

import java.time.Instant

/** The signed-in Microsoft account. Contains no tokens. */
public data class OutlookAccount(
    /** Stable account identifier from the identity platform. */
    val id: String,
    /** The account's sign-in name, usually an email address. */
    val username: String,
)

/** A message sender or recipient. */
public data class EmailAddress(
    val name: String?,
    val address: String?,
) {
    /** Name when present, otherwise the address, otherwise an empty string. */
    public val displayName: String get() = name?.takeIf { it.isNotBlank() } ?: address.orEmpty()
}

/** One row of the Inbox list. */
public data class MessageSummary(
    val id: String,
    val subject: String,
    val sender: EmailAddress?,
    val receivedAt: Instant?,
    val preview: String,
    val isRead: Boolean,
)

/** A message body. [html] is null when the server returned plain text only. */
public data class MessageBody(
    val html: String?,
    val text: String,
)

/** A full message with recipients and body. */
public data class Message(
    val summary: MessageSummary,
    val toRecipients: List<EmailAddress>,
    val body: MessageBody,
)

/**
 * A page of Inbox messages.
 *
 * [nextPageToken] is opaque. Pass it back to `listInbox` to load the next page. It is null on the last page.
 */
public data class MessagePage(
    val messages: List<MessageSummary>,
    val nextPageToken: String?,
)
