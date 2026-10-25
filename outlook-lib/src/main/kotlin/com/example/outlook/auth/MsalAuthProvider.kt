package com.example.outlook.auth

import android.app.Activity
import android.content.Context
import com.example.outlook.OutlookConfig
import com.example.outlook.domain.OutlookAccount
import com.example.outlook.domain.OutlookError
import com.example.outlook.domain.OutlookException
import com.microsoft.identity.client.AcquireTokenSilentParameters
import com.microsoft.identity.client.AuthenticationCallback
import com.microsoft.identity.client.IAccount
import com.microsoft.identity.client.IAuthenticationResult
import com.microsoft.identity.client.IPublicClientApplication
import com.microsoft.identity.client.ISingleAccountPublicClientApplication
import com.microsoft.identity.client.Logger
import com.microsoft.identity.client.PublicClientApplication
import com.microsoft.identity.client.SignInParameters
import com.microsoft.identity.client.exception.MsalClientException
import com.microsoft.identity.client.exception.MsalDeclinedScopeException
import com.microsoft.identity.client.exception.MsalException
import com.microsoft.identity.client.exception.MsalServiceException
import com.microsoft.identity.client.exception.MsalUiRequiredException
import com.microsoft.identity.client.exception.MsalUserCancelException
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.suspendCancellableCoroutine
import kotlinx.coroutines.withContext
import org.json.JSONArray
import org.json.JSONObject
import java.io.File
import kotlin.coroutines.resume
import kotlin.coroutines.resumeWithException

/**
 * MSAL single-account public client. Tokens stay in MSAL's own encrypted cache and are only ever
 * returned to library-internal callers.
 */
internal class MsalAuthProvider private constructor(
    private val app: ISingleAccountPublicClientApplication,
) : AuthProvider {

    override suspend fun currentAccount(): OutlookAccount? = msalAccount()?.toOutlookAccount()

    override suspend fun signIn(activity: Activity): OutlookAccount {
        // MSAL rejects signIn while an account is present; return the existing one instead.
        msalAccount()?.let { return it.toOutlookAccount() }

        val result = suspendCancellableCoroutine { cont ->
            val params = SignInParameters.builder()
                .withActivity(activity)
                .withScopes(AuthProvider.SCOPES)
                .withCallback(object : AuthenticationCallback {
                    override fun onSuccess(authenticationResult: IAuthenticationResult) {
                        cont.resume(authenticationResult)
                    }

                    override fun onError(exception: MsalException) {
                        cont.resumeWithException(exception.toOutlookException())
                    }

                    override fun onCancel() {
                        cont.resumeWithException(OutlookException(OutlookError.AuthRequired))
                    }
                })
                .build()
            app.signIn(params)
        }
        return result.account.toOutlookAccount()
    }

    override suspend fun signOut() {
        withContext(Dispatchers.IO) {
            try {
                app.signOut()
            } catch (e: MsalException) {
                // No account to remove is not an error for sign-out.
                if (e.errorCode != "no_current_account") throw e.toOutlookException()
            }
        }
    }

    override suspend fun acquireToken(forceRefresh: Boolean): String = withContext(Dispatchers.IO) {
        val account = msalAccount() ?: throw OutlookException(OutlookError.AuthRequired)
        val params = AcquireTokenSilentParameters.Builder()
            .forAccount(account)
            .fromAuthority(account.authority)
            .withScopes(AuthProvider.SCOPES)
            .forceRefresh(forceRefresh)
            .build()
        try {
            app.acquireTokenSilent(params).accessToken
        } catch (e: MsalException) {
            throw e.toOutlookException()
        }
    }

    private suspend fun msalAccount(): IAccount? = withContext(Dispatchers.IO) {
        try {
            app.currentAccount?.currentAccount
        } catch (e: MsalException) {
            throw e.toOutlookException()
        }
    }

    companion object {
        /** Creates the MSAL application from host-supplied configuration. */
        suspend fun create(context: Context, config: OutlookConfig): MsalAuthProvider {
            configureMsalLogging()
            val appContext = context.applicationContext
            val configFile = withContext(Dispatchers.IO) { writeConfigFile(appContext, config) }
            try {
                val app = suspendCancellableCoroutine { cont ->
                    PublicClientApplication.createSingleAccountPublicClientApplication(
                        appContext,
                        configFile,
                        object : IPublicClientApplication.ISingleAccountApplicationCreatedListener {
                            override fun onCreated(application: ISingleAccountPublicClientApplication) {
                                cont.resume(application)
                            }

                            override fun onError(exception: MsalException) {
                                cont.resumeWithException(exception.toOutlookException())
                            }
                        },
                    )
                }
                return MsalAuthProvider(app)
            } finally {
                withContext(Dispatchers.IO) { configFile.delete() }
            }
        }

        /** MSAL takes its configuration as a file. Build it from runtime values; nothing is hard-coded. */
        private fun writeConfigFile(context: Context, config: OutlookConfig): File {
            val json = JSONObject()
                .put("client_id", config.clientId)
                .put("redirect_uri", config.redirectUri)
                .put("account_mode", "SINGLE")
                .put("broker_redirect_uri_registered", true)
                // System browser or Custom Tabs only. Never an embedded WebView credential form.
                .put("authorization_user_agent", "BROWSER")
                .put(
                    "authorities",
                    JSONArray().put(
                        JSONObject()
                            .put("type", "AAD")
                            .put("authority_url", config.authority)
                            .put("default", true),
                    ),
                )
            val file = File(context.noBackupFilesDir, "outlook_msal_config.json")
            file.writeText(json.toString())
            return file
        }

        /** Keep MSAL quiet: errors only, never PII, never to logcat. */
        private fun configureMsalLogging() {
            Logger.getInstance().apply {
                setEnablePII(false)
                setEnableLogcatLog(false)
                setLogLevel(Logger.LogLevel.ERROR)
            }
        }

        private fun IAccount.toOutlookAccount() = OutlookAccount(id = id, username = username)

        private fun MsalException.toOutlookException(): OutlookException {
            val kind = when (this) {
                is MsalUserCancelException -> AuthFailureKind.USER_CANCEL
                is MsalDeclinedScopeException -> AuthFailureKind.DECLINED_SCOPE
                is MsalUiRequiredException -> AuthFailureKind.UI_REQUIRED
                is MsalServiceException -> AuthFailureKind.SERVICE
                is MsalClientException -> AuthFailureKind.CLIENT
                else -> AuthFailureKind.OTHER
            }
            return OutlookException(AuthErrorMapper.map(kind, errorCode, message), this)
        }
    }
}
