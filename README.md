# outlook-mail

An Android library (`outlook-lib`) that signs a user into a Microsoft account with MSAL and reads Outlook Inbox mail through Microsoft Graph, plus a minimal harness app (`harness-app`) that exercises it on a device.

The library is the product. The harness exists only to prove the library works.

| Module | Type | Purpose |
| --- | --- | --- |
| `outlook-lib` | Android library (AAR) | MSAL sign-in, Graph mail reads, typed errors. No UI code. |
| `harness-app` | Android application | One activity: sign in, sign out, Inbox list with paging, plain-text message view. |

Stack: Kotlin, coroutines and Flow, MSAL for Android 8.6.0, Retrofit 3 + OkHttp 5 + kotlinx.serialization. Min SDK 26, compile/target SDK 37, AGP 9.4.1, Gradle 9.8.2 (wrapper included).

---

## 1. Register the app in Microsoft Entra

The app registration does not exist yet. Create it once.

1. Sign in to the [Microsoft Entra admin center](https://entra.microsoft.com) and go to **Identity > Applications > App registrations > New registration**.
2. **Name:** anything, for example `Outlook Mail Harness`.
3. **Supported account types:** *Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)*. This matches the `common` authority.
4. Leave **Redirect URI** empty and select **Register**.
5. Copy the **Application (client) ID** from the Overview page.
6. Get the signature hash of the keystore that signs the harness app. For the debug keystore:

   ```bash
   keytool -exportcert -alias androiddebugkey -keystore ~/.android/debug.keystore -storepass android \
     | openssl sha1 -binary | openssl base64
   ```

   On Windows the keystore is at `%USERPROFILE%\.android\debug.keystore`. The output is a 28-character Base64 string such as `Xo8WBi6jzSxKDVR4drqm84yr9iU=`. Each developer machine has its own debug keystore, so each needs its own platform entry (or a shared debug keystore).
7. Go to **Authentication > Add a platform > Android**.
   - **Package name:** `com.example.outlook.harness`
   - **Signature hash:** the value from step 6.
   - Select **Configure**. Entra shows the generated **Redirect URI** (`msauth://com.example.outlook.harness/<URL-encoded hash>`). Copy it exactly.
8. Under **Authentication > Advanced settings**, leave **Allow public client flows** at its default. MSAL's interactive flow on Android does not need it.
9. Go to **API permissions > Add a permission > Microsoft Graph > Delegated permissions** and add `User.Read` and `Mail.Read`. `User.Read` is usually present already. Do not add anything else.
10. **Admin consent for work tenants:** many corporate tenants block user consent for `Mail.Read`. In that case a tenant admin must select **Grant admin consent for <tenant>** on the API permissions page (in their tenant), or approve the request through the tenant's admin consent workflow. Until then, sign-in with a work account returns `ConsentRequired`. Personal accounts consent at first sign-in.

## 2. Fill in `msal_config.json`

Edit `harness-app/src/main/res/raw/msal_config.json`:

```json
{
  "client_id": "<Application (client) ID from step 5>",
  "redirect_uri": "<Redirect URI from step 7, exactly as Entra shows it>",
  "authority": "https://login.microsoftonline.com/common"
}
```

Then add the raw (not URL-encoded) signature hash to `local.properties` in the project root so the manifest's redirect intent filter matches:

```properties
msal.signatureHash=Xo8WBi6jzSxKDVR4drqm84yr9iU=
```

`local.properties` is git-ignored. The harness shows "Fill in res/raw/msal_config.json" until the placeholders are replaced.

The authority can be changed to a single tenant (`https://login.microsoftonline.com/<tenant-id>`), `organizations`, or `consumers`. The library hard-codes none of these values; the harness reads them from this file and passes them to `OutlookMail.initialize`.

## 3. Build and run the harness

```bash
./gradlew assembleDebug                         # builds the AAR and the harness APK
./gradlew :outlook-lib:testDebugUnitTest        # library unit tests
./gradlew :outlook-lib:assembleRelease          # release AAR at outlook-lib/build/outputs/aar/
./gradlew :harness-app:installDebug             # install on a connected device
```

Or open the project in Android Studio and run the `harness-app` configuration. Use a device or emulator with a browser installed (Chrome or another Custom Tabs browser); sign-in runs in the system browser, never in an embedded WebView.

## 4. Library public API

The entry point is the `OutlookMail` object. Callers never see MSAL or Graph types, and no call returns a token.

| Call | Returns |
| --- | --- |
| `suspend initialize(context, config: OutlookConfig)` | `OutlookResult<Unit>` |
| `suspend signIn(activity)` | `OutlookResult<String>`: the account name |
| `suspend signOut()` | `OutlookResult<Unit>`: clears the account and MSAL's token cache |
| `suspend currentAccount()` | `OutlookAccount?` |
| `account: StateFlow<OutlookAccount?>` | Emits on sign-in, sign-out, and when a session becomes unusable |
| `suspend listInbox(pageSize = 25, pageToken = null)` | `OutlookResult<MessagePage>`: summaries newest first, plus `nextPageToken` |
| `suspend getMessage(id)` | `OutlookResult<Message>`: summary, recipients, `body.html` (null for text-only mail) and `body.text` |

Errors are `OutlookError`: `AuthRequired`, `ConsentRequired`, `Throttled(retryAfter)`, `Network`, `Unknown(message)`. `Unknown.message` never contains tokens, addresses or mail content.

```kotlin
val config = OutlookConfig(
    clientId = "<client id>",
    redirectUri = "msauth://com.example.yourapp/<url-encoded hash>",
    // authority defaults to https://login.microsoftonline.com/common
)

lifecycleScope.launch {
    OutlookMail.initialize(applicationContext, config)
    if (OutlookMail.currentAccount() == null) {
        OutlookMail.signIn(this@MyActivity)
    }

    when (val result = OutlookMail.listInbox(pageSize = 25)) {
        is OutlookResult.Success -> {
            render(result.value.messages)
            val next = result.value.nextPageToken  // pass to listInbox for the next page; null on the last page
        }
        is OutlookResult.Failure -> when (val error = result.error) {
            OutlookError.AuthRequired -> OutlookMail.signIn(this@MyActivity)
            OutlookError.ConsentRequired -> showAdminConsentHelp()
            is OutlookError.Throttled -> retryLater(error.retryAfter)
            OutlookError.Network -> showOfflineWithRetry()
            is OutlookError.Unknown -> showGenericError()
        }
    }
}
```

### Behaviour

- **Auth.** MSAL single-account public client, scopes `User.Read` and `Mail.Read` only. Every Graph call acquires a token silently from MSAL's cache. When silent acquisition fails the call returns `AuthRequired` (or `ConsentRequired`), and the caller runs `signIn(activity)`, which is the interactive fallback. The library cannot open a browser without an activity, so it never starts interactive sign-in on its own.
- **HTTP 401.** The library forces one token refresh and retries once. A second 401 returns `AuthRequired`.
- **HTTP 429.** The library reads `Retry-After` (seconds or HTTP date). If the wait is at or below `OutlookConfig.maxAutoRetryWait` (default 10 s) it waits and retries once; otherwise, or if the retry is also throttled, it returns `Throttled(retryAfter)`.
- **HTTP 403** returns `ConsentRequired` (the token lacks `Mail.Read`). Other HTTP errors return `Unknown("Graph HTTP <code>")`. I/O failures return `Network`.
- **Paging.** `nextPageToken` is Graph's `@odata.nextLink`, followed as returned. Tokens that do not point at `https://graph.microsoft.com/` are rejected so a bearer token is never sent elsewhere.
- **Bodies.** `getMessage` asks Graph for the HTML body. The library also derives `body.text` with a dependency-free HTML-to-text conversion (scripts, styles and comments dropped). The library never renders or fetches remote content.

### Consuming the AAR in another app

`outlook-lib` declares its dependencies (MSAL, Retrofit, OkHttp, coroutines, kotlinx.serialization) as `implementation`. When consuming the AAR as a file rather than as a Gradle module or a published Maven artifact, add those dependencies to the host app too, and add Microsoft's display-mask repository that MSAL needs:

```kotlin
maven { url = uri("https://pkgs.dev.azure.com/MicrosoftDeviceSDK/DuoSDK-Public/_packaging/Duo-SDK-Feed/maven/v1") }
```

The host app must also declare MSAL's `BrowserTabActivity` with its own package name and signature hash (copy the block from `harness-app/src/main/AndroidManifest.xml`), set `android:allowBackup="false"`, and forbid cleartext traffic.

## 5. Security and privacy

- Least-privilege delegated scopes: `User.Read`, `Mail.Read`.
- Tokens live only in MSAL's cache. No library or app code stores them, and the public API never returns one.
- MSAL logging is set to errors only, PII disabled, logcat disabled. The library has no HTTP logging, so no Graph response content reaches logcat in any build type.
- OkHttp uses TLS connection specs only; the harness also sets `usesCleartextTraffic="false"` and a network security config that forbids cleartext.
- The harness sets `android:allowBackup="false"` and excludes all data from cloud backup and device transfer.
- The harness shows bodies as plain text only. If HTML rendering is added later, use a WebView with JavaScript disabled, file and content access blocked, and mixed content blocked.
- On sign-out the harness clears its in-memory message list and the library clears MSAL's account and token cache and its account state.

## 6. Project structure

```
outlook-lib/
  src/main/kotlin/com/example/outlook/
    OutlookMail.kt          public facade
    OutlookConfig.kt        host-supplied configuration
    OutlookMailClient.kt    testable core behind the facade
    auth/                   MSAL wrapper and error mapping
    data/                   Graph Retrofit interface, DTOs, repository, HTML-to-text
    domain/                 public models, OutlookError, OutlookResult
  src/test/                 unit tests with fake auth and fake Graph API, JSON fixtures
harness-app/
  src/main/kotlin/com/example/outlook/harness/MainActivity.kt
  src/main/res/raw/msal_config.json
```

`outlook-lib` does not depend on `harness-app` and contains no UI code.

## 7. Tests

Unit tests (`./gradlew :outlook-lib:testDebugUnitTest`) use a fake auth provider and a scripted fake Graph API with JSON fixtures (`inbox_page1.json`, `inbox_page2.json`, `message_detail.json`). They cover parsing, paging, page-token validation, the 401 refresh-once rule, 429 Retry-After handling, 403/5xx/network/format error mapping, MSAL error classification, HTML-to-text conversion, and sign-out state clearing.

### Manual test checklist

Run on a real device after completing sections 1 and 2.

- [ ] **Personal account, first sign-in.** Tap Sign in, complete sign-in and consent in the browser. The 25 most recent Inbox messages appear.
- [ ] **Work account, first sign-in.** Sign out, then sign in with a work account in a tenant that allows user consent. Inbox loads.
- [ ] **Paging.** Scroll to the end of the list (or tap Load more). The next 25 messages append. On the last page the button disappears.
- [ ] **Message view.** Tap a message. Subject, sender, recipients and plain-text body appear. Back returns to the list.
- [ ] **Silent sign-in after restart.** Force-stop the app and reopen it. It shows "Signed in as ..." and loads the Inbox without a browser prompt.
- [ ] **Sign-out.** Tap Sign out. The list clears. Force-stop and reopen: the app shows "Signed out" and does not load mail.
- [ ] **Token expiry.** Leave the app signed in for more than an hour (access tokens last about 60 to 90 minutes), then pull a new page. It loads without a prompt (silent refresh). Revoking sessions for the account (account security page or Entra) and retrying shows "sign-in required".
- [ ] **Airplane mode.** Enable airplane mode and load or open a message. The status shows "network unavailable" with a Retry button. Disable airplane mode and tap Retry.
- [ ] **Admin consent tenant.** Sign in with a work account in a tenant that blocks user consent for `Mail.Read`. The status shows "consent required". After an admin grants consent, sign-in succeeds.
- [ ] **Logcat.** During all of the above, `adb logcat` shows no tokens, email addresses, subjects or bodies.

## 8. Known limitations

- Read-only Inbox: no send, reply, forward, delete, move, attachments, calendar or contacts.
- One account at a time. No background sync, push notifications or persistent message cache.
- Interactive sign-in requires an `Activity`; the library returns `AuthRequired` instead of prompting by itself.
- `HtmlText` is a pragmatic converter for display, not a full HTML parser. Complex layouts (nested tables, CSS-driven content) may lose structure. `body.html` is always available for a proper renderer.
- `403` is mapped to `ConsentRequired`. Rare 403 causes unrelated to consent (for example, an account without an Exchange Online mailbox) are reported the same way.
- The package names `com.example.outlook` and `com.example.outlook.harness` are placeholders. Changing the harness `applicationId` requires a new Android platform entry and redirect URI in Entra.
- The harness is unstyled on purpose.
