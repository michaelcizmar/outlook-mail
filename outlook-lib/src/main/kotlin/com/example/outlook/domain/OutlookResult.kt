package com.example.outlook.domain

import kotlin.time.Duration

/** Typed failure returned by every library call. */
public sealed class OutlookError {
    /** No usable session. The caller should run interactive sign-in. Also returned when the user cancels sign-in. */
    public data object AuthRequired : OutlookError()

    /** The user or the tenant admin has not granted the requested permissions. */
    public data object ConsentRequired : OutlookError()

    /** Microsoft Graph throttled the request. Retry after [retryAfter]. */
    public data class Throttled(val retryAfter: Duration) : OutlookError()

    /** The device is offline or the connection failed. */
    public data object Network : OutlookError()

    /** Anything else. [message] never contains tokens or mail content. */
    public data class Unknown(val message: String) : OutlookError()
}

/** Result of a library call. */
public sealed class OutlookResult<out T> {
    public data class Success<out T>(val value: T) : OutlookResult<T>()
    public data class Failure(val error: OutlookError) : OutlookResult<Nothing>()

    public fun getOrNull(): T? = (this as? Success)?.value
    public fun errorOrNull(): OutlookError? = (this as? Failure)?.error
}

/** Internal carrier for an [OutlookError] across suspend boundaries. */
internal class OutlookException(val error: OutlookError, cause: Throwable? = null) :
    Exception(error.toString(), cause)
