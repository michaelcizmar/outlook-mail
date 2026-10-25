package com.example.outlook.harness

import android.os.Bundle
import android.view.View
import android.widget.AbsListView
import android.widget.ArrayAdapter
import android.widget.Button
import android.widget.LinearLayout
import android.widget.ListView
import android.widget.ScrollView
import android.widget.TextView
import androidx.activity.ComponentActivity
import androidx.activity.OnBackPressedCallback
import androidx.lifecycle.lifecycleScope
import com.example.outlook.OutlookConfig
import com.example.outlook.OutlookMail
import com.example.outlook.domain.Message
import com.example.outlook.domain.MessageSummary
import com.example.outlook.domain.OutlookError
import com.example.outlook.domain.OutlookResult
import kotlinx.coroutines.Job
import kotlinx.coroutines.launch
import org.json.JSONObject

/**
 * Test harness for outlook-lib. Deliberately plain: one activity, programmatic views, no styling or navigation.
 * Calls only the library's public API.
 */
class MainActivity : ComponentActivity() {

    private lateinit var status: TextView
    private lateinit var signInButton: Button
    private lateinit var signOutButton: Button
    private lateinit var retryButton: Button
    private lateinit var listView: ListView
    private lateinit var loadMoreButton: Button
    private lateinit var detailScroll: ScrollView
    private lateinit var detailText: TextView
    private lateinit var adapter: ArrayAdapter<String>

    // In-memory only. Cleared on sign-out.
    private val messages = mutableListOf<MessageSummary>()
    private var nextPageToken: String? = null
    private var loading = false
    private var retryAction: (() -> Unit)? = null

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        buildViews()
        onBackPressedDispatcher.addCallback(this, object : OnBackPressedCallback(true) {
            override fun handleOnBackPressed() {
                if (detailScroll.visibility == View.VISIBLE) showList() else finish()
            }
        })
        lifecycleScope.launch { initialize() }
    }

    private suspend fun initialize() {
        val config = readConfig() ?: run {
            setStatus("Fill in res/raw/msal_config.json with valid values (see README), then rebuild.")
            return
        }
        when (val result = OutlookMail.initialize(applicationContext, config)) {
            is OutlookResult.Failure -> showError("Initialization failed", result.error) { lifecycleScope.launch { initialize() } }
            is OutlookResult.Success -> {
                val account = OutlookMail.currentAccount()
                updateButtons(signedIn = account != null)
                if (account != null) {
                    setStatus("Signed in as ${account.username}")
                    loadFirstPage()
                } else {
                    setStatus("Signed out")
                }
            }
        }
    }

    private fun readConfig(): OutlookConfig? {
        val json = resources.openRawResource(R.raw.msal_config).bufferedReader().use { JSONObject(it.readText()) }
        val clientId = json.optString("client_id")
        val redirectUri = json.optString("redirect_uri")
        if (clientId.startsWith("REPLACE_") || redirectUri.contains("REPLACE_")) return null
        val authority = json.optString("authority").ifEmpty { OutlookConfig.DEFAULT_AUTHORITY }
        // OutlookConfig validates its values; treat an invalid file like an unfilled one.
        return runCatching { OutlookConfig(clientId = clientId, redirectUri = redirectUri, authority = authority) }.getOrNull()
    }

    private fun signIn(): Job = lifecycleScope.launch {
        setStatus("Signing in...")
        when (val result = OutlookMail.signIn(this@MainActivity)) {
            is OutlookResult.Success -> {
                setStatus("Signed in as ${result.value}")
                updateButtons(signedIn = true)
                loadFirstPage()
            }
            is OutlookResult.Failure -> showError("Sign-in failed", result.error) { signIn() }
        }
    }

    private fun signOut(): Job = lifecycleScope.launch {
        OutlookMail.signOut()
        messages.clear()
        nextPageToken = null
        adapter.clear()
        detailText.text = ""
        showList()
        updateButtons(signedIn = false)
        loadMoreButton.visibility = View.GONE
        setStatus("Signed out")
    }

    private fun loadFirstPage() {
        messages.clear()
        nextPageToken = null
        adapter.clear()
        loadPage(null)
    }

    private fun loadPage(pageToken: String?) {
        if (loading) return
        loading = true
        setStatus(if (pageToken == null) "Loading Inbox..." else "Loading more...")
        lifecycleScope.launch {
            val result = OutlookMail.listInbox(pageSize = 25, pageToken = pageToken)
            loading = false
            when (result) {
                is OutlookResult.Success -> {
                    messages += result.value.messages
                    nextPageToken = result.value.nextPageToken
                    adapter.addAll(result.value.messages.map(::rowText))
                    loadMoreButton.visibility = if (nextPageToken != null) View.VISIBLE else View.GONE
                    setStatus("${messages.size} messages")
                }
                is OutlookResult.Failure -> showError("Could not load Inbox", result.error) { loadPage(pageToken) }
            }
        }
    }

    private fun openMessage(summary: MessageSummary): Job = lifecycleScope.launch {
        setStatus("Loading message...")
        when (val result = OutlookMail.getMessage(summary.id)) {
            is OutlookResult.Success -> {
                detailText.text = detailText(result.value)
                showDetail()
                setStatus("")
            }
            is OutlookResult.Failure -> showError("Could not load message", result.error) { openMessage(summary) }
        }
    }

    private fun showError(prefix: String, error: OutlookError, retry: () -> Unit) {
        val reason = when (error) {
            OutlookError.AuthRequired -> "sign-in required"
            OutlookError.ConsentRequired -> "consent required (a work tenant may need admin consent for Mail.Read)"
            is OutlookError.Throttled -> "throttled, retry after ${error.retryAfter.inWholeSeconds}s"
            OutlookError.Network -> "network unavailable"
            is OutlookError.Unknown -> error.message
        }
        setStatus("$prefix: $reason")
        if (error == OutlookError.AuthRequired) updateButtons(signedIn = false)
        retryAction = retry
        retryButton.visibility = View.VISIBLE
    }

    private fun rowText(m: MessageSummary): String {
        val unread = if (m.isRead) "" else "[unread] "
        return "$unread${m.subject.ifBlank { "(no subject)" }}\n${m.sender?.displayName.orEmpty()}"
    }

    private fun detailText(m: Message): String = buildString {
        appendLine("Subject: ${m.summary.subject.ifBlank { "(no subject)" }}")
        appendLine("From: ${m.summary.sender?.displayName.orEmpty()}")
        appendLine("To: ${m.toRecipients.joinToString { it.displayName }}")
        appendLine("Received: ${m.summary.receivedAt ?: ""}")
        appendLine()
        // Plain text only. The harness never renders HTML.
        append(m.body.text)
    }

    private fun setStatus(text: String) {
        status.text = text
        retryButton.visibility = View.GONE
    }

    private fun updateButtons(signedIn: Boolean) {
        signInButton.isEnabled = !signedIn
        signOutButton.isEnabled = signedIn
    }

    private fun showList() {
        detailScroll.visibility = View.GONE
        listView.visibility = View.VISIBLE
    }

    private fun showDetail() {
        listView.visibility = View.GONE
        detailScroll.visibility = View.VISIBLE
        detailScroll.scrollTo(0, 0)
    }

    private fun buildViews() {
        val pad = (16 * resources.displayMetrics.density).toInt()
        status = TextView(this).apply { setPadding(pad, pad, pad, pad / 2) }
        signInButton = Button(this).apply { text = "Sign in"; isEnabled = false; setOnClickListener { signIn() } }
        signOutButton = Button(this).apply { text = "Sign out"; isEnabled = false; setOnClickListener { signOut() } }
        retryButton = Button(this).apply {
            text = "Retry"; visibility = View.GONE
            setOnClickListener { retryAction?.invoke() }
        }
        val buttons = LinearLayout(this).apply {
            orientation = LinearLayout.HORIZONTAL
            setPadding(pad, 0, pad, 0)
            addView(signInButton); addView(signOutButton); addView(retryButton)
        }

        adapter = ArrayAdapter(this, android.R.layout.simple_list_item_1, mutableListOf())
        loadMoreButton = Button(this).apply {
            text = "Load more"; visibility = View.GONE
            setOnClickListener { nextPageToken?.let(::loadPage) }
        }
        listView = ListView(this).apply {
            addFooterView(loadMoreButton)
            this.adapter = this@MainActivity.adapter
            setOnItemClickListener { _, _, position, _ -> messages.getOrNull(position)?.let { openMessage(it) } }
            setOnScrollListener(object : AbsListView.OnScrollListener {
                override fun onScrollStateChanged(view: AbsListView, scrollState: Int) = Unit
                override fun onScroll(view: AbsListView, first: Int, visible: Int, total: Int) {
                    if (total > 1 && first + visible >= total) nextPageToken?.let(::loadPage)
                }
            })
        }

        detailText = TextView(this).apply { setPadding(pad, pad, pad, pad); setTextIsSelectable(true) }
        detailScroll = ScrollView(this).apply { visibility = View.GONE; addView(detailText) }

        val content = LinearLayout(this).apply {
            orientation = LinearLayout.VERTICAL
            addView(status)
            addView(buttons)
            addView(listView, LinearLayout.LayoutParams(LinearLayout.LayoutParams.MATCH_PARENT, 0, 1f))
            addView(detailScroll, LinearLayout.LayoutParams(LinearLayout.LayoutParams.MATCH_PARENT, 0, 1f))
            fitsSystemWindows = true
        }
        setContentView(content)
    }
}
