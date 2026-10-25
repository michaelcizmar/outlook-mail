package com.example.outlook.auth

import android.app.Activity
import com.example.outlook.domain.OutlookAccount

/**
 * Internal seam between the library and MSAL. Every failure is thrown as
 * [com.example.outlook.domain.OutlookException]. Tokens never leave the library.
 */
internal interface AuthProvider {
    /** Returns the signed-in account, or null. */
    suspend fun currentAccount(): OutlookAccount?

    /** Runs interactive sign-in in the system browser or Custom Tabs. */
    suspend fun signIn(activity: Activity): OutlookAccount

    /** Removes the account and MSAL's cached tokens. */
    suspend fun signOut()

    /**
     * Returns an access token for Graph, acquired silently from MSAL's cache.
     * Throws AuthRequired or ConsentRequired when interactive sign-in is needed.
     */
    suspend fun acquireToken(forceRefresh: Boolean = false): String

    companion object {
        /** Least-privilege delegated scopes. */
        val SCOPES: List<String> = listOf("User.Read", "Mail.Read")
    }
}
