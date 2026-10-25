package com.example.outlook.data

import kotlinx.serialization.ExperimentalSerializationApi
import okhttp3.ConnectionSpec
import okhttp3.MediaType.Companion.toMediaType
import okhttp3.OkHttpClient
import retrofit2.Retrofit
import retrofit2.converter.kotlinx.serialization.asConverterFactory
import retrofit2.http.GET
import retrofit2.http.Header
import retrofit2.http.Path
import retrofit2.http.Query
import retrofit2.http.Url
import java.util.concurrent.TimeUnit

/** Microsoft Graph v1.0 mail endpoints used by the library. Query values are pre-encoded. */
internal interface GraphApi {

    @GET("me/mailFolders/inbox/messages")
    suspend fun listInbox(
        @Header("Authorization") authorization: String,
        @Query(value = "\$select", encoded = true) select: String = LIST_SELECT,
        @Query(value = "\$orderby", encoded = true) orderBy: String = LIST_ORDER_BY,
        @Query(value = "\$top", encoded = true) top: Int,
    ): MessageListDto

    /** Follows an `@odata.nextLink` exactly as Graph returned it. */
    @GET
    suspend fun listInboxNextPage(
        @Header("Authorization") authorization: String,
        @Url nextLink: String,
    ): MessageListDto

    @GET("me/messages/{id}")
    suspend fun getMessage(
        @Header("Authorization") authorization: String,
        @Path("id") id: String,
        @Query(value = "\$select", encoded = true) select: String = DETAIL_SELECT,
        @Header("Prefer") prefer: String = PREFER_HTML_BODY,
    ): MessageDto

    companion object {
        const val BASE_URL = "https://graph.microsoft.com/v1.0/"
        const val LIST_SELECT = "id,subject,from,receivedDateTime,bodyPreview,isRead"
        const val LIST_ORDER_BY = "receivedDateTime%20desc"
        const val DETAIL_SELECT = "subject,from,toRecipients,receivedDateTime,body"
        const val PREFER_HTML_BODY = "outlook.body-content-type=\"html\""

        @OptIn(ExperimentalSerializationApi::class)
        fun create(): GraphApi {
            val client = OkHttpClient.Builder()
                // TLS only. Cleartext is not in the list, so http:// requests fail.
                .connectionSpecs(listOf(ConnectionSpec.MODERN_TLS))
                .connectTimeout(20, TimeUnit.SECONDS)
                .readTimeout(30, TimeUnit.SECONDS)
                .build()
            return Retrofit.Builder()
                .baseUrl(BASE_URL)
                .client(client)
                .addConverterFactory(GraphJson.asConverterFactory("application/json".toMediaType()))
                .build()
                .create(GraphApi::class.java)
        }
    }
}
