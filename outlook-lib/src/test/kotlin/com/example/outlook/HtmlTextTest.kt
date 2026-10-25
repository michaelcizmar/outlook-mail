package com.example.outlook

import com.example.outlook.data.HtmlText
import org.junit.Assert.assertEquals
import org.junit.Test

class HtmlTextTest {

    @Test
    fun `drops scripts styles and comments`() {
        val html = "<style>.a{}</style><!-- hidden --><p>Visible</p><SCRIPT type='x'>evil()</SCRIPT>"
        assertEquals("Visible", HtmlText.toPlainText(html))
    }

    @Test
    fun `decodes named and numeric entities once`() {
        assertEquals("a & b < c © — &lt;", HtmlText.toPlainText("a &amp; b &lt; c &#169; &#x2014; &amp;lt;"))
    }

    @Test
    fun `keeps line structure and collapses whitespace`() {
        val html = "<div>Line   one</div>\n\n<div>Line two<br/>Line three</div><p></p><p></p><p>End</p>"
        assertEquals("Line one\nLine two\nLine three\n\nEnd", HtmlText.toPlainText(html))
    }

    @Test
    fun `leaves unknown entities untouched`() {
        assertEquals("&bogus;", HtmlText.toPlainText("&bogus;"))
    }
}
