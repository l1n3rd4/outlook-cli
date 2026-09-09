"""Unit tests for modular HTML email rendering and cleaning."""

from __future__ import annotations

import pytest
from bs4 import BeautifulSoup, Tag
from rich.console import Console

from outlook_cli.formatter.html import (
    clean_email_html,
    format_inline,
    html_to_clean_text,
    render_code_block,
    render_email_body,
    render_html_email,
    render_list,
    render_table,
)
from outlook_cli.formatter.html.code import extract_language_from_pre
from outlook_cli.formatter.html.tables import is_layout_table


@pytest.fixture
def test_console():
    return Console(record=True, width=100)


# ============================================================================
# Cleaner Tests
# ============================================================================


def test_cleaner_removes_noise_tags():
    html = """
    <html>
        <head><meta charset="utf-8"><title>Test</title><link rel="stylesheet" href="style.css"></head>
        <body>
            <script>alert(1);</script>
            <style>body { color: red; }</style>
            <noscript>No script</noscript>
            <iframe src="frame.html"></iframe>
            <svg><circle r="10"/></svg>
            <video src="video.mp4"></video>
            <audio src="audio.mp3"></audio>
            <canvas></canvas>
            <form action="/"><input type="text"></form>
            <template><p>Template content</p></template>
            <p>Real content</p>
        </body>
    </html>
    """
    soup = BeautifulSoup(html, "html.parser")
    clean_email_html(soup)

    assert soup.find("script") is None
    assert soup.find("style") is None
    assert soup.find("meta") is None
    assert soup.find("link") is None
    assert soup.find("iframe") is None
    assert soup.find("svg") is None
    assert soup.find("form") is None
    assert "Real content" in soup.get_text()


def test_cleaner_removes_comments_and_mso_conditionals():
    html = """
    <div>
        <!--[if mso]>
        <p>Outlook only</p>
        <![endif]-->
        <!-- Normal comment -->
        <p>Visible content</p>
    </div>
    """
    soup = BeautifulSoup(html, "html.parser")
    clean_email_html(soup)

    text = soup.get_text()
    assert "Outlook only" not in text
    assert "Normal comment" not in text
    assert "Visible content" in text


def test_cleaner_removes_hidden_elements_and_tracking_pixels():
    html = """
    <div>
        <span hidden>Hidden by attr</span>
        <span aria-hidden="true">Aria hidden</span>
        <div style="display: none">Display none</div>
        <div style="display:none; color: blue">Display none compact</div>
        <div style="visibility: hidden">Visibility hidden</div>
        <div style="visibility:hidden">Visibility hidden compact</div>
        <img src="tracker.gif" width="1" height="1" alt="pixel" />
        <img src="tracker0.gif" width="0" height="0" />
        <img src="tracker_style.gif" style="width: 1px; height: 1px" />
        <img src="photo.jpg" width="300" height="200" alt="real photo" />
        <p>Visible body</p>
    </div>
    """
    soup = BeautifulSoup(html, "html.parser")
    clean_email_html(soup)

    text = soup.get_text()
    assert "Hidden by attr" not in text
    assert "Aria hidden" not in text
    assert "Display none" not in text
    assert "Visibility hidden" not in text
    assert soup.find("img", {"src": "tracker.gif"}) is None
    assert soup.find("img", {"src": "tracker0.gif"}) is None
    assert soup.find("img", {"src": "tracker_style.gif"}) is None
    assert soup.find("img", {"src": "photo.jpg"}) is not None
    assert "Visible body" in text


# ============================================================================
# Inline Formatting Tests
# ============================================================================


def test_format_inline_tags_and_whitespace():
    html = """
    <p>
        <strong>Bold</strong> and <b>also bold</b>.
        <em>Italic</em> and <i>also italic</i>.
        <u>Underlined</u>, <del>deleted</del>, <s>struck</s>, <strike>also struck</strike>.
        <code>inline code</code>, <kbd>Ctrl+C</kbd>, <samp>output</samp>.
        <mark>highlighted</mark>.
        <!-- inline comment -->
        <br>
        Line after break. &amp; &lt;escaped&gt;
    </p>
    """
    soup = BeautifulSoup(html, "html.parser")
    p_tag = soup.find("p")
    assert isinstance(p_tag, Tag)
    text = format_inline(p_tag)

    plain = text.plain
    assert "Bold and also bold." in plain
    assert "Italic and also italic." in plain
    assert "Underlined, deleted, struck, also struck." in plain
    assert "inline code, Ctrl+C, output." in plain
    assert "highlighted." in plain
    assert "\nLine after break. & <escaped>" in plain


def test_format_inline_links_and_mailto():
    html = """
    <p>
        Visit <a href="https://example.com">Example</a> or
        email <a href="mailto:support@example.com">Support</a>.
        <a href="/relative/path">Relative</a>
        <a>Anchor without href</a>
    </p>
    """
    soup = BeautifulSoup(html, "html.parser")
    p_tag = soup.find("p")
    assert isinstance(p_tag, Tag)

    text = format_inline(p_tag, base_url="https://base.org")
    plain = text.plain

    assert "Visit Example" in plain
    assert "email Support" in plain
    assert "Relative" in plain
    assert "Anchor without href" in plain


def test_format_inline_images_and_subblocks():
    html = """
    <p>
        Text before
        <img src="test.jpg" alt="A nice chart" />
        <img src="test2.jpg" title="Photo title" />
        <img src="spacer.gif" />
        <ul><li>Ignored inside inline</li></ul>
        Text after
    </p>
    """
    soup = BeautifulSoup(html, "html.parser")
    p_tag = soup.find("p")
    assert isinstance(p_tag, Tag)

    text = format_inline(p_tag)
    plain = text.plain

    assert "Text before [A nice chart] [Photo title] Text after" in plain
    assert "Ignored inside inline" not in plain


def test_format_inline_leading_whitespace_trim():
    soup = BeautifulSoup("<p>   Leading spaces   </p>", "html.parser")
    p_tag = soup.find("p")
    assert isinstance(p_tag, Tag)
    text = format_inline(p_tag)
    assert text.plain == "Leading spaces"


# ============================================================================
# Lists Tests
# ============================================================================


def test_render_list_nested_and_ordered(test_console):
    html = """
    <ol>
        <li>First
            <ul>
                <li>Nested subitem 1
                    <ul>
                        <li>Deep subitem
                            <ul>
                                <li>Deepest subitem</li>
                            </ul>
                        </li>
                    </ul>
                </li>
                <li>Nested subitem 2</li>
            </ul>
        </li>
        <li>Second</li>
        <li></li>
    </ol>
    """
    soup = BeautifulSoup(html, "html.parser")
    ol_tag = soup.find("ol")
    assert isinstance(ol_tag, Tag)

    render_list(ol_tag, test_console)
    output = test_console.export_text()

    assert "1. First" in output
    assert "2. Second" in output
    assert "◦ Nested subitem 1" in output
    assert "▪ Deep subitem" in output
    assert "▪ Deepest subitem" in output


# ============================================================================
# Code Block Tests
# ============================================================================


def test_extract_language():
    soup1 = BeautifulSoup('<pre class="language-python"><code>x = 1</code></pre>', "html.parser")
    assert extract_language_from_pre(soup1.find("pre")) == "python"

    soup2 = BeautifulSoup('<pre><code class="highlight-javascript">console.log(1);</code></pre>', "html.parser")
    assert extract_language_from_pre(soup2.find("pre")) == "javascript"

    soup3 = BeautifulSoup('<pre class="json"><code>{}</code></pre>', "html.parser")
    assert extract_language_from_pre(soup3.find("pre")) == "json"

    soup4 = BeautifulSoup("<pre><code>plain</code></pre>", "html.parser")
    assert extract_language_from_pre(soup4.find("pre")) == ""


def test_render_code_block_with_and_without_syntax(test_console, monkeypatch):
    # Empty pre
    empty_soup = BeautifulSoup("<pre></pre>", "html.parser")
    render_code_block(empty_soup.find("pre"), test_console)

    # Valid syntax
    py_soup = BeautifulSoup('<pre class="language-python"><code>print("hello")</code></pre>', "html.parser")
    render_code_block(py_soup.find("pre"), test_console)

    # Unknown language
    plain_soup = BeautifulSoup("<pre><code>echo hello</code></pre>", "html.parser")
    render_code_block(plain_soup.find("pre"), test_console)

    # Syntax exception fallback
    def raise_syntax(*args, **kwargs):
        raise RuntimeError("Syntax error")

    monkeypatch.setattr("outlook_cli.formatter.html.code.Syntax", raise_syntax)
    render_code_block(py_soup.find("pre"), test_console)

    output = test_console.export_text()
    assert "PYTHON" in output
    assert "print(\"hello\")" in output
    assert "Code" in output


# ============================================================================
# Tables Tests
# ============================================================================


def test_is_layout_table():
    # role=presentation
    t1 = BeautifulSoup('<table role="presentation"><tr><td>A</td><td>B</td></tr></table>', "html.parser").find("table")
    assert is_layout_table(t1) is True

    # Single cell
    t2 = BeautifulSoup("<table><tr><td>Single cell</td></tr></table>", "html.parser").find("table")
    assert is_layout_table(t2) is True

    # Cell containing block elements (p, div, etc.)
    t3 = BeautifulSoup("<table><tr><td><p>Paragraph inside</p></td><td>Col 2</td></tr></table>", "html.parser").find("table")
    assert is_layout_table(t3) is True

    # Real data table
    t4 = BeautifulSoup("<table><tr><th>Col 1</th><th>Col 2</th></tr><tr><td>Val 1</td><td>Val 2</td></tr></table>", "html.parser").find("table")
    assert is_layout_table(t4) is False


def test_render_table_data_and_layout(test_console):
    # Empty table
    empty_t = BeautifulSoup("<table></table>", "html.parser").find("table")
    render_table(empty_t, test_console)

    # Layout table with walk callback
    layout_t = BeautifulSoup('<table role="presentation"><tr><td><p>Layout content</p></td></tr></table>', "html.parser").find("table")
    called = []
    render_table(layout_t, test_console, walk_children=lambda node: called.append(True))
    assert len(called) == 1

    # Data table with headers
    data_t = BeautifulSoup("""
    <table>
        <tr><th>Name</th><th>Role</th></tr>
        <tr><td>Alice</td><td>Developer</td></tr>
        <tr><td>Mismatched</td></tr>
    </table>
    """, "html.parser").find("table")
    render_table(data_t, test_console)

    # Key-value table without headers (2-8 columns)
    kv_t = BeautifulSoup("""
    <table>
        <tr><td>Key 1</td><td>Value 1</td></tr>
        <tr><td>Key 2</td></tr>
    </table>
    """, "html.parser").find("table")
    render_table(kv_t, test_console)

    # Table without tr elements
    no_tr_t = BeautifulSoup("<table><td>1</td><td>2</td></table>", "html.parser").find("table")
    render_table(no_tr_t, test_console)

    output = test_console.export_text()
    assert "Name" in output
    assert "Alice" in output
    assert "Developer" in output
    assert "Key 1" in output
    assert "Value 1" in output


# ============================================================================
# Body & HTML Email Reader Tests
# ============================================================================


def test_render_email_body_full_elements(test_console):
    html = """
    <div>
        Plain text at root
        <!-- container comment -->
        <h1>Main Heading 1</h1>
        <h2>Section Heading 2</h2>
        <h3>Subsection Heading 3</h3>
        <h4>Minor Heading 4</h4>
        <h5>Minor Heading 5</h5>
        <h6>Minor Heading 6</h6>
        <hr />
        <p>Paragraph with <strong>bold</strong> text.</p>
        <pre class="language-bash"><code>git status</code></pre>
        <ul><li>Bullet item</li></ul>
        <ol><li>Ordered item</li></ol>
        <table>
            <tr><th>Header 1</th></tr>
            <tr><td>Cell 1</td></tr>
        </table>
        <blockquote>
            Quoted email reply message.
        </blockquote>
        <dl>
            <dt>Definition Term</dt>
            <dd>Definition Description</dd>
        </dl>
        <section>
            <div>
                <p>Nested container paragraph</p>
            </div>
        </section>
        <section>
            <div>
                <span>Inner text</span>
            </div>
        </section>
        <div class="leaf-div">
            Direct div text without leaves
        </div>
    </div>
    """
    soup = BeautifulSoup(html, "html.parser")
    clean_email_html(soup)
    render_email_body(soup.body or soup, test_console)

    output = test_console.export_text()
    assert "Plain text at root" in output
    assert "Main Heading 1" in output
    assert "Section Heading 2" in output
    assert "Subsection Heading 3" in output
    assert "Minor Heading 4" in output
    assert "Minor Heading 5" in output
    assert "Minor Heading 6" in output
    assert "Paragraph with bold text." in output
    assert "git status" in output
    assert "Quoted email reply message." in output
    assert "Definition Term" in output
    assert "Definition Description" in output
    assert "Nested container paragraph" in output
    assert "Direct div text without leaves" in output


def test_render_html_email_empty_and_normal(test_console):
    # Empty
    render_html_email("", console=test_console)
    render_html_email("   ", console=test_console)
    assert test_console.export_text() == ""

    # Normal email
    html = "<h2>Meeting Notes</h2><p>Discussion on <b>Q3 goals</b>.</p>"
    render_html_email(html, console=test_console)
    output = test_console.export_text()
    assert "Meeting Notes" in output
    assert "Discussion on Q3 goals." in output


def test_html_to_clean_text():
    assert html_to_clean_text("") == ""

    html = """
    <html>
        <head><style>p { color: red; }</style></head>
        <body>
            <h3>Report</h3>
            <p>First paragraph.</p>
            <p>Second paragraph.</p>
        </body>
    </html>
    """
    clean = html_to_clean_text(html)
    assert "Report" in clean
    assert "First paragraph." in clean
    assert "Second paragraph." in clean
    assert "<style>" not in clean


def test_html_to_clean_text_fallback_on_exception(monkeypatch):
    monkeypatch.setattr("outlook_cli.formatter.html.BeautifulSoup", lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("BS4 failure")))
    result = html_to_clean_text("<p>Fallback <b>plain</b> text</p>")
    assert "Fallback plain text" in result
