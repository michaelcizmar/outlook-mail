package com.example.outlook.data

/**
 * Converts an HTML mail body to readable plain text without any Android or rendering dependency.
 * Scripts, styles and comments are dropped; nothing is fetched.
 */
internal object HtmlText {

    private val DROP_BLOCKS = Regex(
        "<!--.*?-->|<(script|style|head|title)\\b[^>]*>.*?</\\1\\s*>",
        setOf(RegexOption.IGNORE_CASE, RegexOption.DOT_MATCHES_ALL),
    )
    private val LINE_BREAK = Regex("<br\\s*/?>", RegexOption.IGNORE_CASE)
    private val PARAGRAPH_END = Regex("</(p|h[1-6]|blockquote|table|ul|ol|pre)\\s*>", RegexOption.IGNORE_CASE)
    private val LINE_END = Regex("</(div|tr)\\s*>", RegexOption.IGNORE_CASE)
    private val LIST_ITEM = Regex("<li\\b[^>]*>", RegexOption.IGNORE_CASE)
    private val CELL_END = Regex("</t[dh]\\s*>", RegexOption.IGNORE_CASE)
    private val ANY_TAG = Regex("<[^>]+>")
    private val ENTITY = Regex("&(#[xX][0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);")
    private val HORIZONTAL_SPACE = Regex("[ \\t\\u00A0]+")
    private val EXTRA_BLANK_LINES = Regex("\\n{3,}")

    private val NAMED_ENTITIES = mapOf(
        "nbsp" to " ", "amp" to "&", "lt" to "<", "gt" to ">", "quot" to "\"", "apos" to "'",
        "ndash" to "–", "mdash" to "—", "hellip" to "…", "copy" to "©",
        "reg" to "®", "trade" to "™", "rsquo" to "’", "lsquo" to "‘",
        "rdquo" to "”", "ldquo" to "“", "bull" to "•", "zwnj" to "",
    )

    fun toPlainText(html: String): String {
        val withBreaks = html
            .replace(DROP_BLOCKS, "")
            .replace("\r", "")
            .replace("\n", " ")
            .replace(LINE_BREAK, "\n")
            .replace(LIST_ITEM, "\n• ")
            .replace(CELL_END, "\t")
            .replace(PARAGRAPH_END, "\n\n")
            .replace(LINE_END, "\n")
            .replace(ANY_TAG, "")
        return decodeEntities(withBreaks)
            .lines()
            .joinToString("\n") { it.replace(HORIZONTAL_SPACE, " ").trim() }
            .replace(EXTRA_BLANK_LINES, "\n\n")
            .trim()
    }

    private fun decodeEntities(text: String): String = text.replace(ENTITY) { match ->
        val body = match.groupValues[1]
        when {
            body.startsWith("#x") || body.startsWith("#X") ->
                body.substring(2).toIntOrNull(16)?.let(::codePointToString)
            body.startsWith("#") -> body.substring(1).toIntOrNull()?.let(::codePointToString)
            else -> NAMED_ENTITIES[body.lowercase()]
        } ?: match.value
    }

    private fun codePointToString(codePoint: Int): String? =
        if (Character.isValidCodePoint(codePoint)) String(Character.toChars(codePoint)) else null
}
