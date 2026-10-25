package com.example.outlook

import android.app.Activity
import android.content.Context
import com.example.outlook.auth.MsalAuthProvider
import com.example.outlook.data.GraphApi
import com.example.outlook.data.GraphMailRepository
import com.example.outlook.domain.Message
import com.example.outlook.domain.MessagePage
import com.example.outlook.domain.OutlookAccount
import com.example.outlook.domain.OutlookError
import com.example.outlook.domain.OutlookException
import com.example.outlook.domain.OutlookResult
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.asStateFlow
import kotlinx.coroutines.sync.Mutex
import kotlinx.coroutines.sync.withLock
import kotlin.coroutines.cancellation.CancellationException

/**
 * Public entry point. Signs a user into a Microsoft account and reads Outlook Inbox mail through Microsoft Graph.
 *
 * Callers never see MSAL or Graph types, and no call returns a raw token.
 *
 * ```
 * OutlookMail.initialize(context, OutlookConfig(clientId, redirectUri))
 * OutlookMail.signIn(activity)
 * when (val page = OutlookMail.listInbox()) {
 *     is OutlookResult.Success -> show(page.value.messages)
 *     is OutlookResult.Failure -> handle(page.error)
 * }
 * ```
 */
public object OutlookMail {

    private val initLock = Mutex()
    private val accountState = MutableStateFlow<OutlookAccount?>(null)

    @Volatile
    private var client: OutlookMailClient? = null

    @Volatile
    private var activeConfig: OutlookConfig? = null

    /** The signed-in account, or null. Updates on sign-in, sign-out, and when a session becomes unusable. */
    public val account: StateFlow<OutlookAccount?> = accountState.asStateFlow()

    /** Sets up the library. Safe to call more than once with the same config. */
    public suspend fun initialize(context: Context, config: OutlookConfig): OutlookResult<Unit> = initLock.withLock {
        if (client != null && activeConfig == config) return@withLock OutlookResult.Success(Unit)
        try {
            val auth = MsalAuthProvider.create(context, config)
            val repository = GraphMailRepository(GraphApi.create(), auth, config.maxAutoRetryWait)
            client = OutlookMailClient(auth, repository, accountState).also { it.refreshAccount() }
            activeConfig = config
            OutlookResult.Success(Unit)
        } catch (e: CancellationException) {
            throw e
        } catch (e: OutlookException) {
            OutlookResult.Failure(e.error)
        } catch (e: Exception) {
            OutlookResult.Failure(OutlookError.Unknown("Initialization failed: ${e::class.java.simpleName}"))
        }
    }

    /** Runs interactive sign-in and returns the account name. Returns the existing account if already signed in. */
    public suspend fun signIn(activity: Activity): OutlookResult<String> = requireClient().signIn(activity)

    /** Clears the account and MSAL's cached tokens. */
    public suspend fun signOut(): OutlookResult<Unit> = requireClient().signOut()

    /** Returns the signed-in account, or null. */
    public suspend fun currentAccount(): OutlookAccount? = requireClient().refreshAccount()

    /**
     * Returns a page of Inbox message summaries, newest first.
     *
     * @param pageSize messages per page, 1 to 100. Applies to the first page; later pages keep the same size.
     * @param pageToken null for the first page, otherwise [MessagePage.nextPageToken] from the previous page.
     */
    public suspend fun listInbox(pageSize: Int = 25, pageToken: String? = null): OutlookResult<MessagePage> =
        requireClient().listInbox(pageSize, pageToken)

    /** Returns one message with its HTML body (when available) and a plain-text body. */
    public suspend fun getMessage(id: String): OutlookResult<Message> = requireClient().getMessage(id)

    private fun requireClient(): OutlookMailClient =
        checkNotNull(client) { "OutlookMail.initialize must complete before other calls" }
}
