package com.example.outlook

import kotlin.time.Duration
import kotlin.time.Duration.Companion.seconds

/**
 * Configuration supplied by the host app. The library hard-codes none of these values.
 *
 * @param clientId Application (client) ID from the Entra app registration.
 * @param redirectUri Android redirect URI from the app registration, in the form
 *   `msauth://<package name>/<URL-encoded signature hash>`.
 * @param authority Sign-in authority. Defaults to `common`, which accepts personal and work accounts.
 * @param maxAutoRetryWait When Graph returns HTTP 429 with a Retry-After at or below this value, the library waits
 *   and retries once. Longer waits are returned to the caller as [com.example.outlook.domain.OutlookError.Throttled].
 */
public data class OutlookConfig(
    val clientId: String,
    val redirectUri: String,
    val authority: String = DEFAULT_AUTHORITY,
    val maxAutoRetryWait: Duration = 10.seconds,
) {
    init {
        require(clientId.isNotBlank()) { "clientId is required" }
        require(redirectUri.startsWith("msauth://")) { "redirectUri must use the msauth:// scheme" }
        require(authority.startsWith("https://")) { "authority must be an https URL" }
    }

    public companion object {
        public const val DEFAULT_AUTHORITY: String = "https://login.microsoftonline.com/common"
    }
}
