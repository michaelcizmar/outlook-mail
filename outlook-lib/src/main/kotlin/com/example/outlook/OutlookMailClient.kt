package com.example.outlook

import android.app.Activity
import com.example.outlook.auth.AuthProvider
import com.example.outlook.data.GraphMailRepository
import com.example.outlook.domain.Message
import com.example.outlook.domain.MessagePage
import com.example.outlook.domain.OutlookAccount
import com.example.outlook.domain.OutlookError
import com.example.outlook.domain.OutlookException
import com.example.outlook.domain.OutlookResult
import kotlinx.coroutines.flow.MutableStateFlow
import kotlin.coroutines.cancellation.CancellationException

/** Testable core behind [OutlookMail]. Converts every failure into an [OutlookResult.Failure]. */
internal class OutlookMailClient(
    private val auth: AuthProvider,
    private val repository: GraphMailRepository,
    private val accountState: MutableStateFlow<OutlookAccount?>,
) {

    suspend fun refreshAccount(): OutlookAccount? {
        val account = runCatchingOutlook { auth.currentAccount() }.getOrNull()
        accountState.value = account
        return account
    }

    suspend fun signIn(activity: Activity): OutlookResult<String> = runCatchingOutlook {
        val account = auth.signIn(activity)
        accountState.value = account
        account.username
    }

    suspend fun signOut(): OutlookResult<Unit> {
        val result = runCatchingOutlook { auth.signOut() }
        // Clear cached account state even if MSAL reported a problem, so the UI never shows a stale account.
        accountState.value = null
        return result
    }

    suspend fun listInbox(pageSize: Int, pageToken: String?): OutlookResult<MessagePage> =
        runCatchingOutlook { repository.listInbox(pageSize, pageToken) }.also(::onAuthFailure)

    suspend fun getMessage(id: String): OutlookResult<Message> =
        runCatchingOutlook { repository.getMessage(id) }.also(::onAuthFailure)

    /** After AuthRequired the cached session is unusable; reflect that in the account flow. */
    private fun onAuthFailure(result: OutlookResult<*>) {
        if (result.errorOrNull() == OutlookError.AuthRequired) accountState.value = null
    }

    private suspend fun <T> runCatchingOutlook(block: suspend () -> T): OutlookResult<T> = try {
        OutlookResult.Success(block())
    } catch (e: CancellationException) {
        throw e
    } catch (e: OutlookException) {
        OutlookResult.Failure(e.error)
    } catch (e: IllegalArgumentException) {
        OutlookResult.Failure(OutlookError.Unknown(e.message ?: "Invalid argument"))
    } catch (e: Exception) {
        OutlookResult.Failure(OutlookError.Unknown(e::class.java.simpleName))
    }
}
