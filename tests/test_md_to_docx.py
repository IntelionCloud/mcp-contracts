"""Tests for MD → DOCX inline formatting (bold, italic).

The original parser only handled tracked-changes ({++…++}, {--…--}) and
footnote refs ([^N]); markdown bold (**…**) and HTML <strong>/<b>/<em>/<i>
inside table cells leaked through verbatim into the produced DOCX.

This module covers the segmenter and the end-to-end DOCX build.
"""
from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from core._md_to_docx import build_docx, parse_segments


# ---------------------------------------------------------------------------
# parse_segments — unit tests
# ---------------------------------------------------------------------------


def test_md_bold():
    assert parse_segments("hello **world** end", accept=True) == [
        ("hello ", "normal"),
        ("world", "bold"),
        (" end", "normal"),
    ]


def test_html_strong():
    assert parse_segments("a <strong>b</strong> c", accept=True) == [
        ("a ", "normal"),
        ("b", "bold"),
        (" c", "normal"),
    ]


def test_html_b_short_tag():
    assert parse_segments("<b>x</b>", accept=True) == [("x", "bold")]


def test_html_em_and_i_italic():
    assert parse_segments("<em>x</em> <i>y</i>", accept=True) == [
        ("x", "italic"),
        (" ", "normal"),
        ("y", "italic"),
    ]


def test_md_bold_does_not_consume_single_asterisks():
    # "5 * 3 = 15" must remain plain text — single * is not bold.
    assert parse_segments("5 * 3 = 15", accept=True) == [("5 * 3 = 15", "normal")]


def test_md_bold_with_escaped_asterisks():
    # Bold span containing \* (escaped asterisk) — the real bug case:
    # "+7(913)\*\*\*-\*\*-16" inside bold markers must render as bold text
    # with literal asterisks, not be rejected/split by the bold regex.
    out = parse_segments(r"**+7(913)\*\*\*-\*\*-16** (СБП)", accept=True)
    assert out == [
        ("+7(913)***-**-16", "bold"),
        (" (СБП)", "normal"),
    ]


def test_md_bold_combined_with_tracked_insertion():
    # Order is preserved; both markers are detected independently.
    out = parse_segments("see **note**, also {++added++} fact", accept=False)
    assert out == [
        ("see ", "normal"),
        ("note", "bold"),
        (", also ", "normal"),
        ("added", "insert"),
        (" fact", "normal"),
    ]


def test_footnote_ref_still_works():
    assert parse_segments("see [^1] there", accept=False) == [
        ("see ", "normal"),
        ("[^1]", "footnote"),
        (" there", "normal"),
    ]


def test_no_markers_returns_single_normal_segment():
    assert parse_segments("plain text", accept=True) == [("plain text", "normal")]


# ---------------------------------------------------------------------------
# end-to-end — build a DOCX and assert real bold/italic runs land in XML
# ---------------------------------------------------------------------------


def _document_xml(docx_path: Path) -> str:
    with zipfile.ZipFile(docx_path) as z:
        return z.read("word/document.xml").decode("utf-8")


@pytest.fixture
def md_with_bold_in_table(tmp_path: Path) -> Path:
    md = tmp_path / "in.md"
    md.write_text(
        "Title\n\n"
        "| left | right |\n"
        "| --- | --- |\n"
        "| **Agrobiotech LLC**, represented by **Zurab Munjishvili**. | "
        "<strong>Intelion Cloud LLC</strong>, <em>note</em>. |\n",
        encoding="utf-8",
    )
    return md


def test_docx_contains_bold_runs_inside_table_cells(md_with_bold_in_table, tmp_path):
    out = tmp_path / "out.docx"
    build_docx(str(md_with_bold_in_table), str(out), accept=False)

    xml = _document_xml(out)

    # The literal markers must NOT appear in the rendered text — they were
    # leaking through before the fix.
    assert "**" not in xml, "literal '**' leaked into docx"
    assert "<strong>" not in xml.replace("&lt;", "<"), "<strong> leaked into docx"
    assert "&lt;strong&gt;" not in xml, "<strong> escaped into text"

    # The actual bold-bearing names must be present as plain text.
    assert "Agrobiotech LLC" in xml
    assert "Zurab Munjishvili" in xml
    assert "Intelion Cloud LLC" in xml
    assert "note" in xml

    # And there must be at least one <w:b/> formatting run for each
    # bolded fragment (3: Agrobiotech LLC, Zurab Munjishvili, Intelion Cloud LLC).
    bold_runs = xml.count("<w:b/>")
    assert bold_runs >= 3, f"expected ≥3 bold runs, got {bold_runs}"

    # And at least one italic run for <em>note</em>.
    italic_runs = xml.count("<w:i/>")
    assert italic_runs >= 1, f"expected ≥1 italic run, got {italic_runs}"


# ---------------------------------------------------------------------------
# Markdown ATX headings (# / ## / ### / ####) — must render as bold runs,
# not as literal "## " text in the body.
# ---------------------------------------------------------------------------


def test_markdown_headings_are_not_literal(tmp_path: Path):
    md = tmp_path / "h.md"
    md.write_text(
        "Body lead-in.\n\n"
        "## Section heading\n\n"
        "Some body.\n\n"
        "### Sub-section heading\n\n"
        "More body.\n\n"
        "#### 3.1. Deep heading\n\n"
        "Tail.\n",
        encoding="utf-8",
    )
    out = tmp_path / "h.docx"
    build_docx(str(md), str(out), accept=False)
    xml = _document_xml(out)

    # Hashes must NOT leak as literal text into the rendered body.
    # We check the explicit "## ", "### ", "#### " prefixes rather than
    # bare "#" because the latter could legitimately appear elsewhere.
    assert "## " not in xml, "literal '## ' leaked into docx"
    assert "### " not in xml, "literal '### ' leaked into docx"
    assert "#### " not in xml, "literal '#### ' leaked into docx"

    assert "Section heading" in xml
    assert "Sub-section heading" in xml
    assert "3.1. Deep heading" in xml

    # Implicit title (Body lead-in) + 3 explicit headings = ≥4 bold runs.
    bold_runs = xml.count("<w:b/>")
    assert bold_runs >= 4, f"expected ≥4 bold runs, got {bold_runs}"


def test_explicit_h1_first_line_does_not_double_render(tmp_path: Path):
    """When the first line is `# Title`, it should be rendered as a single
    h1 (centered, bold) — not as both an implicit title AND a heading.
    """
    md = tmp_path / "h1.md"
    md.write_text("# Top Title\n\nBody.\n", encoding="utf-8")
    out = tmp_path / "h1.docx"
    build_docx(str(md), str(out), accept=False)
    xml = _document_xml(out)
    assert "# Top Title" not in xml, "literal '# Top Title' leaked"
    assert xml.count("Top Title") == 1, "title text should appear exactly once"


def _h1_paragraphs_with_page_break_flag(xml: str):
    """Return list of bool flags — True if the paragraph carrying the given
    h1 text has <w:pageBreakBefore/> in its <w:pPr>.

    h1 paragraphs in our converter are produced with bold + 13pt font, so
    we look for paragraphs containing that combination plus the heading
    text.
    """
    import re as _re
    flags = []
    for p_match in _re.finditer(r"<w:p\b[^>]*>(.*?)</w:p>", xml, _re.DOTALL):
        body = p_match.group(1)
        if "Top Title" in body or "Annex" in body:
            flags.append("<w:pageBreakBefore/>" in body)
    return flags


def test_h1_after_other_content_gets_page_break(tmp_path: Path):
    md = tmp_path / "pb.md"
    md.write_text(
        "# Top Title\n\n"
        "Body of the first section.\n\n"
        "# Annex\n\n"
        "Body of the annex.\n",
        encoding="utf-8",
    )
    out = tmp_path / "pb.docx"
    build_docx(str(md), str(out), accept=False)
    xml = _document_xml(out)

    flags = _h1_paragraphs_with_page_break_flag(xml)
    assert len(flags) == 2, f"expected 2 h1 paragraphs, got {len(flags)}: {flags}"
    # First h1 must NOT have pageBreakBefore (would force a blank cover page).
    assert flags[0] is False, "first h1 should not carry pageBreakBefore"
    # Second h1 must have pageBreakBefore — it starts a new top-level section.
    assert flags[1] is True, "second h1 should carry pageBreakBefore"


def test_h1_only_no_page_break(tmp_path: Path):
    md = tmp_path / "pb1.md"
    md.write_text("# Top Title\n\nOnly body.\n", encoding="utf-8")
    out = tmp_path / "pb1.docx"
    build_docx(str(md), str(out), accept=False)
    xml = _document_xml(out)
    assert "<w:pageBreakBefore/>" not in xml, \
        "single h1 must not carry pageBreakBefore"


def test_h2_does_not_get_page_break(tmp_path: Path):
    md = tmp_path / "pb2.md"
    md.write_text(
        "# Top Title\n\nBody.\n\n## Section\n\nMore.\n",
        encoding="utf-8",
    )
    out = tmp_path / "pb2.docx"
    build_docx(str(md), str(out), accept=False)
    xml = _document_xml(out)
    # No h1 after content here, so no pageBreakBefore should appear.
    assert "<w:pageBreakBefore/>" not in xml, \
        "h2 must not trigger pageBreakBefore"


def test_table_first_line_is_not_consumed_as_title(tmp_path: Path):
    """When the document begins with a table, the first row must be
    rendered inside the table — not as a centered bold paragraph with
    raw '|' separators.
    """
    md = tmp_path / "tbl.md"
    md.write_text(
        "| HEADER LEFT | HEADER RIGHT |\n"
        "| --- | --- |\n"
        "| body left | body right |\n",
        encoding="utf-8",
    )
    out = tmp_path / "tbl.docx"
    build_docx(str(md), str(out), accept=False)
    xml = _document_xml(out)
    # Pipe characters from the markdown table syntax must not appear in the
    # rendered DOCX body text.
    assert "| HEADER" not in xml, "literal '|' leaked from table-row title"
    assert "HEADER LEFT" in xml
    assert "HEADER RIGHT" in xml
    assert "body left" in xml
