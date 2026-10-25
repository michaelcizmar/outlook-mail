package com.example.outlook

import com.example.outlook.auth.AuthErrorMapper
import com.example.outlook.auth.AuthFailureKind
import com.example.outlook.domain.OutlookError
import org.junit.Assert.assertEquals
import org.junit.Assert.assertTrue
import org.junit.Test

class AuthErrorMapperTest {

    private fun map(kind: AuthFailureKind, code: String? = null, message: String? = null) =
        AuthErrorMapper.map(kind, code, message)

    @Test
    fun `user cancel needs sign-in`() {
        assertEquals(OutlookError.AuthRequired, map(AuthFailureKind.USER_CANCEL, "user_cancelled"))
    }

    @Test
    fun `ui required without consent hint needs sign-in`() {
        assertEquals(OutlookError.AuthRequired, map(AuthFailureKind.UI_REQUIRED, "invalid_grant", "token expired"))
    }

    @Test
    fun `consent AADSTS codes map to ConsentRequired`() {
        assertEquals(
            OutlookError.ConsentRequired,
            map(AuthFailureKind.UI_REQUIRED, "invalid_grant", "AADSTS65001: The user or administrator has not consented"),
        )
        assertEquals(
            OutlookError.ConsentRequired,
            map(AuthFailureKind.SERVICE, "access_denied", "AADSTS90094: Admin approval required"),
        )
        assertEquals(OutlookError.ConsentRequired, map(AuthFailureKind.DECLINED_SCOPE, "declined_scope"))
    }

    @Test
    fun `client network codes map to Network`() {
        assertEquals(OutlookError.Network, map(AuthFailureKind.CLIENT, "device_network_not_available"))
        assertEquals(OutlookError.Network, map(AuthFailureKind.CLIENT, "io_error"))
    }

    @Test
    fun `missing account maps to AuthRequired`() {
        assertEquals(OutlookError.AuthRequired, map(AuthFailureKind.CLIENT, "no_current_account"))
    }

    @Test
    fun `unknown errors carry only the code, never the message`() {
        val error = map(AuthFailureKind.SERVICE, "unknown_error", "contains user@contoso.com")
        assertTrue(error is OutlookError.Unknown)
        assertTrue(!(error as OutlookError.Unknown).message.contains("@"))
    }
}
