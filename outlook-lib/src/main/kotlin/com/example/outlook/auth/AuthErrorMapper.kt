package com.example.outlook.auth

import com.example.outlook.domain.OutlookError

/** The MSAL exception family a failure came from. Kept separate from MSAL types so the mapping is unit-testable. */
internal enum class AuthFailureKind { USER_CANCEL, UI_REQUIRED, DECLINED_SCOPE, SERVICE, CLIENT, OTHER }

internal object AuthErrorMapper {

    /** AADSTS codes that mean consent is missing: user consent, admin consent required, or admin approval pending. */
    private val CONSENT_CODES = listOf("AADSTS65001", "AADSTS65004", "AADSTS90094", "AADSTS90008", "AADSTS900941")

    private val CONSENT_ERROR_CODES = setOf("consent_required", "access_denied", "interaction_required_consent")

    private val NETWORK_ERROR_CODES = setOf("io_error", "device_network_not_available", "network_not_available")

    private val NO_ACCOUNT_ERROR_CODES = setOf("no_current_account", "no_account_found", "no_tokens_found")

    fun map(kind: AuthFailureKind, errorCode: String?, message: String?): OutlookError {
        val code = errorCode.orEmpty()
        val consentHint = CONSENT_CODES.any { message?.contains(it) == true } || code in CONSENT_ERROR_CODES
        return when (kind) {
            AuthFailureKind.USER_CANCEL -> OutlookError.AuthRequired
            AuthFailureKind.DECLINED_SCOPE -> OutlookError.ConsentRequired
            AuthFailureKind.UI_REQUIRED -> if (consentHint) OutlookError.ConsentRequired else OutlookError.AuthRequired
            AuthFailureKind.SERVICE -> when {
                consentHint -> OutlookError.ConsentRequired
                code == "invalid_grant" -> OutlookError.AuthRequired
                else -> OutlookError.Unknown("auth service error: ${code.ifEmpty { "unspecified" }}")
            }
            AuthFailureKind.CLIENT -> when (code) {
                in NETWORK_ERROR_CODES -> OutlookError.Network
                in NO_ACCOUNT_ERROR_CODES -> OutlookError.AuthRequired
                else -> OutlookError.Unknown("auth client error: ${code.ifEmpty { "unspecified" }}")
            }
            AuthFailureKind.OTHER -> OutlookError.Unknown("auth error: ${code.ifEmpty { "unspecified" }}")
        }
    }
}
