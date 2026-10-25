package com.example.outlook.data

import kotlinx.serialization.SerialName
import kotlinx.serialization.Serializable
import kotlinx.serialization.json.Json

@Serializable
internal data class MessageListDto(
    val value: List<MessageDto> = emptyList(),
    @SerialName("@odata.nextLink") val nextLink: String? = null,
)

@Serializable
internal data class MessageDto(
    val id: String,
    val subject: String? = null,
    val from: RecipientDto? = null,
    val toRecipients: List<RecipientDto> = emptyList(),
    val receivedDateTime: String? = null,
    val bodyPreview: String? = null,
    val isRead: Boolean = false,
    val body: ItemBodyDto? = null,
)

@Serializable
internal data class RecipientDto(
    val emailAddress: EmailAddressDto? = null,
)

@Serializable
internal data class EmailAddressDto(
    val name: String? = null,
    val address: String? = null,
)

@Serializable
internal data class ItemBodyDto(
    val contentType: String? = null,
    val content: String? = null,
)

internal val GraphJson: Json = Json {
    ignoreUnknownKeys = true
    explicitNulls = false
    coerceInputValues = true
}
