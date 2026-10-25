package com.example.outlook

import android.app.Activity
import com.example.outlook.auth.AuthProvider
import com.example.outlook.data.GraphApi
import com.example.outlook.data.GraphJson
import com.example.outlook.data.MessageDto
import com.example.outlook.data.MessageListDto
import com.example.outlook.domain.OutlookAccount
import com.example.outlook.domain.OutlookError
import com.example.outlook.domain.OutlookException
import okhttp3.Protocol
import okhttp3.Request
import okhttp3.ResponseBody.Companion.toResponseBody
import retrofit2.HttpException
import retrofit2.Response

internal fun fixture(name: String): String =
    checkNotNull(object {}.javaClass.classLoader!!.getResource(name)) { "missing fixture $name" }.readText()

internal fun httpError(code: Int, headers: Map<String, String> = emptyMap()): HttpException {
    val raw = okhttp3.Response.Builder()
        .request(Request.Builder().url("https://graph.microsoft.com/v1.0/me/messages").build())
        .protocol(Protocol.HTTP_1_1)
        .code(code)
        .message("HTTP $code")
        .apply { headers.forEach { (k, v) -> header(k, v) } }
        .build()
    return HttpException(Response.error<Any>("{}".toResponseBody(), raw))
}

internal class FakeAuth(
    var account: OutlookAccount? = OutlookAccount("id-1", "user@contoso.com"),
    var tokenError: OutlookError? = null,
) : AuthProvider {
    val tokenRequests = mutableListOf<Boolean>()
    var signOutCalls = 0
    private var tokenCounter = 0

    override suspend fun currentAccount(): OutlookAccount? = account

    override suspend fun signIn(activity: Activity): OutlookAccount =
        account ?: throw OutlookException(OutlookError.AuthRequired)

    override suspend fun signOut() {
        signOutCalls++
        account = null
    }

    override suspend fun acquireToken(forceRefresh: Boolean): String {
        tokenRequests += forceRefresh
        tokenError?.let { throw OutlookException(it) }
        if (account == null) throw OutlookException(OutlookError.AuthRequired)
        return "token-${++tokenCounter}"
    }
}

/** Scripted Graph API. Each queued entry is either a DTO to return or a Throwable to throw. */
internal class FakeGraphApi : GraphApi {
    val listResponses = ArrayDeque<Any>()
    val detailResponses = ArrayDeque<Any>()
    val calls = mutableListOf<String>()
    val authorizations = mutableListOf<String>()

    fun queueList(json: String) = listResponses.addLast(GraphJson.decodeFromString<MessageListDto>(json))
    fun queueDetail(json: String) = detailResponses.addLast(GraphJson.decodeFromString<MessageDto>(json))

    override suspend fun listInbox(authorization: String, select: String, orderBy: String, top: Int): MessageListDto {
        calls += "list top=$top select=$select orderBy=$orderBy"
        authorizations += authorization
        return next(listResponses)
    }

    override suspend fun listInboxNextPage(authorization: String, nextLink: String): MessageListDto {
        calls += "next $nextLink"
        authorizations += authorization
        return next(listResponses)
    }

    override suspend fun getMessage(authorization: String, id: String, select: String, prefer: String): MessageDto {
        calls += "get $id select=$select prefer=$prefer"
        authorizations += authorization
        return next(detailResponses)
    }

    @Suppress("UNCHECKED_CAST")
    private fun <T> next(queue: ArrayDeque<Any>): T {
        val item = queue.removeFirstOrNull() ?: error("no scripted response")
        if (item is Throwable) throw item
        return item as T
    }
}
