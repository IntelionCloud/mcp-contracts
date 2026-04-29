"""End-to-end tests for the redline pipeline using real DOCX fixtures.

Covers the 4 canonical reviewer-comment operations:
  * find    — list_comments / extract_comments returns existing comments
  * add     — apply_changes(modify, comment=...) produces a new comment
  * reply   — apply_changes(reply, target_id=...) nests under a parent
  * delete  — delete_comment strips definition + anchors

Fixtures live next to this file; see `_make_fixtures.py` to regenerate.
"""
from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from core import redline
from core.docx_parser import extract_comments


FIXTURES = Path(__file__).parent / "fixtures"
SAMPLE_PLAIN = FIXTURES / "sample_plain.docx"
SAMPLE_WITH_COMMENT = FIXTURES / "sample_with_comment.docx"


@pytest.fixture
def plain_docx(tmp_path: Path) -> Path:
    """Fresh copy of the plain sample (writable; tests may produce derived files
    next to it)."""
    dst = tmp_path / "plain.docx"
    shutil.copy(SAMPLE_PLAIN, dst)
    return dst


@pytest.fixture
def with_comment_docx(tmp_path: Path) -> Path:
    """Fresh copy of the seeded one-comment sample."""
    dst = tmp_path / "with_comment.docx"
    shutil.copy(SAMPLE_WITH_COMMENT, dst)
    return dst


# ---------------------------------------------------------------------------
# 1. find — read comments back from a DOCX
# ---------------------------------------------------------------------------


def test_find_comment_in_seeded_fixture(with_comment_docx: Path):
    """The fixture was seeded with one comment by 'Original Reviewer';
    `extract_comments` must surface it verbatim."""
    comments = extract_comments(str(with_comment_docx))

    assert len(comments) == 1, f"expected 1 comment in fixture, got {comments!r}"
    c = comments[0]
    assert c["author"] == "Original Reviewer"
    assert "rate basis" in c["text"]
    # ID assigned by adeu starts at 1.
    assert c["id"] == "1"


def test_find_in_plain_docx_returns_empty(plain_docx: Path):
    """No comments part → no errors, just empty list."""
    assert extract_comments(str(plain_docx)) == []


# ---------------------------------------------------------------------------
# 2. add — modify-with-comment creates a new comment
# ---------------------------------------------------------------------------


def test_add_comment_via_modify_edit(plain_docx: Path, tmp_path: Path):
    """Using `comment` on a ModifyText creates a new reviewer comment
    anchored to the edit."""
    out = tmp_path / "with_added.docx"
    res = redline.apply_changes(
        str(plain_docx),
        edits=[{
            "type": "modify",
            "target_text": "Akash hourly rates",
            "new_text": "Akash hourly rates (TBD)",
            "comment": "Need to clarify the rate basis.",
        }],
        author="Original Reviewer",
        output_path=str(out),
    )
    assert res["applied"] == 1
    assert res["skipped"] == 0

    comments = extract_comments(str(out))
    assert len(comments) == 1
    assert comments[0]["author"] == "Original Reviewer"
    assert "rate basis" in comments[0]["text"]


# ---------------------------------------------------------------------------
# 3. reply — answer to an existing comment by ID
# ---------------------------------------------------------------------------


def test_reply_to_existing_comment(with_comment_docx: Path, tmp_path: Path):
    """Reply nests as a second comment in the thread; both authors visible."""
    out = tmp_path / "with_reply.docx"
    res = redline.apply_changes(
        str(with_comment_docx),
        edits=[{
            "type": "reply",
            "target_id": "Com:1",
            "text": "Will provide the formula by Friday.",
        }],
        author="Max Vyaznikov",
        output_path=str(out),
    )
    assert res["applied"] == 1
    assert res["skipped"] == 0

    comments = extract_comments(str(out))
    assert len(comments) == 2

    by_author = {c["author"]: c for c in comments}
    assert set(by_author) == {"Original Reviewer", "Max Vyaznikov"}
    assert "rate basis" in by_author["Original Reviewer"]["text"]
    assert "formula by Friday" in by_author["Max Vyaznikov"]["text"]


# ---------------------------------------------------------------------------
# 4. delete — strip a comment by ID (definition + inline anchors)
# ---------------------------------------------------------------------------


def test_delete_comment_removes_it(with_comment_docx: Path, tmp_path: Path):
    """`delete_comment("Com:1")` must leave zero comments behind."""
    out = tmp_path / "deleted.docx"
    res = redline.delete_comment(
        str(with_comment_docx),
        "Com:1",
        output_path=str(out),
    )
    # 1 definition removed + N inline anchors (start/end/reference) — exact
    # count varies by anchor topology, but it must be ≥ 1.
    assert res["removed"] >= 1
    assert Path(res["output_path"]).exists()
    assert extract_comments(str(out)) == []


def test_delete_comment_accepts_bare_numeric_id(with_comment_docx: Path, tmp_path: Path):
    """`delete_comment("1")` (without 'Com:' prefix) works too."""
    out = tmp_path / "deleted.docx"
    res = redline.delete_comment(
        str(with_comment_docx),
        "1",
        output_path=str(out),
    )
    assert res["removed"] >= 1
    assert extract_comments(str(out)) == []


def test_delete_unknown_comment_is_noop(with_comment_docx: Path, tmp_path: Path):
    """Deleting a non-existent ID leaves the file untouched (removed=0)
    and the original comment intact."""
    out = tmp_path / "unchanged.docx"
    res = redline.delete_comment(
        str(with_comment_docx),
        "Com:999",
        output_path=str(out),
    )
    assert res["removed"] == 0
    # The lone real comment must still be there.
    assert len(extract_comments(str(out))) == 1
