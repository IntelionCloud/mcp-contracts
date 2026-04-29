"""
Surgical DOCX editing — apply tracked changes, accept/reject existing changes,
reply to comments, sanitize for sharing.

Wraps `adeu.RedlineEngine` and `adeu.sanitize.sanitize_docx` (MIT-licensed
upstream — see LICENSE files in the installed package).
"""
from __future__ import annotations

import os
import zipfile
from io import BytesIO
from xml.etree import ElementTree as ET

from adeu import (
    AcceptChange,
    ModifyText,
    RedlineEngine,
    RejectChange,
    ReplyComment,
    extract_text_from_stream,
)
from adeu.sanitize import sanitize_docx as _adeu_sanitize
from adeu.sanitize.core import SanitizeError  # not re-exported in package __init__

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


_TYPE_MAP = {
    "modify": ModifyText,
    "accept": AcceptChange,
    "reject": RejectChange,
    "reply": ReplyComment,
}


def _load_engine(file_path: str, author: str):
    with open(file_path, "rb") as f:
        stream = BytesIO(f.read())
    return RedlineEngine(stream, author=author)


def _save_engine(engine, output_path: str) -> None:
    with open(output_path, "wb") as f:
        f.write(engine.save_to_stream().getvalue())


def _build_changes(edits: list[dict]) -> list:
    """Convert raw dicts (from MCP arguments) to typed adeu change models.

    Each dict needs `type`: one of {modify, accept, reject, reply}.
    The remaining fields match the corresponding pydantic model.
    """
    out = []
    for raw in edits:
        kind = raw.get("type")
        if kind not in _TYPE_MAP:
            raise ValueError(
                f"unknown change type: {kind!r}; expected one of "
                f"{sorted(_TYPE_MAP)}"
            )
        out.append(_TYPE_MAP[kind].model_validate(raw))
    return out


def read_with_changes(file_path: str, *, clean_view: bool = False) -> str:
    """Render DOCX text including tracked changes & comments markers.

    Tracked changes appear inline as CriticMarkup-style {++ins++}/{--del--};
    comments are anchored with [Com:N] markers. Use clean_view=True to get
    the "accepted" rendering (no markup).
    """
    with open(file_path, "rb") as f:
        stream = BytesIO(f.read())
    return extract_text_from_stream(
        stream, filename=os.path.basename(file_path), clean_view=clean_view
    )


def apply_changes(
    file_path: str,
    edits: list[dict],
    *,
    author: str = "AI Copilot",
    output_path: str | None = None,
) -> dict:
    """Apply a batch of edits and write the result.

    edits — list of dicts, each is one of:
      {type:"modify", target_text, new_text, comment?}
      {type:"accept", target_id, comment?}
      {type:"reject", target_id, comment?}
      {type:"reply",  target_id, text}

    Returns {output_path, applied, skipped, details} where `details` is
    adeu's per-item skip reasons (when present).
    """
    if output_path is None:
        stem, ext = os.path.splitext(file_path)
        output_path = f"{stem}_redlined{ext}"

    engine = _load_engine(file_path, author=author)
    changes = _build_changes(edits)

    # adeu's unified entry point: handles modify/accept/reject/reply
    # in one pass, rebuilds the doc map between phases, returns a
    # structured count dict.
    result = engine.process_batch(changes)
    applied = int(result.get("actions_applied", 0)) + int(result.get("edits_applied", 0))
    skipped = int(result.get("actions_skipped", 0)) + int(result.get("edits_skipped", 0))

    _save_engine(engine, output_path)
    return {
        "output_path": output_path,
        "applied": applied,
        "skipped": skipped,
        "details": list(result.get("skipped_details") or []),
    }


def accept_all_changes(file_path: str, *, output_path: str | None = None) -> str:
    """Accept every tracked change and write the result. Returns output path."""
    if output_path is None:
        stem, ext = os.path.splitext(file_path)
        output_path = f"{stem}_accepted{ext}"

    engine = _load_engine(file_path, author="AI Copilot")
    engine.accept_all_revisions()
    _save_engine(engine, output_path)
    return output_path


def sanitize(
    input_path: str,
    *,
    output_path: str | None = None,
    keep_markup: bool = False,
    accept_all: bool = False,
    author: str | None = None,
) -> dict:
    """Strip metadata/author IDs/internal tracking before sharing.

    Returns the structured SanitizeResult as a plain dict. Re-raises
    `SanitizeError` so the caller can give the user actionable feedback
    (typical case: unresolved tracked changes — fix is `accept_all=True`).
    """
    result = _adeu_sanitize(
        input_path,
        output_path=output_path,
        keep_markup=keep_markup,
        accept_all=accept_all,
        author=author,
    )
    return {
        "output_path": result.output_path,
        "status": result.status,
        "tracked_changes_found": result.tracked_changes_found,
        "tracked_changes_accepted": result.tracked_changes_accepted,
        "comments_removed": result.comments_removed,
        "comments_kept": result.comments_kept,
        "metadata_stripped": list(result.metadata_stripped),
        "warnings": list(result.warnings),
        "report_text": result.report_text,
    }


def delete_comment(
    file_path: str,
    comment_id: str,
    *,
    output_path: str | None = None,
) -> dict:
    """Remove a comment definition AND its anchors from a DOCX.

    `comment_id` accepts either the bare numeric ID (`"1"`) or the marker
    form (`"Com:1"`) returned by `read_with_changes`. adeu has no native
    delete operation — Word's UI distinguishes "resolve" (keeps thread)
    from "delete" (strips XML); this helper does the latter via direct
    OOXML surgery: removes `<w:comment w:id="N">` from word/comments*.xml
    and `<w:commentRangeStart|End|Reference w:id="N">` from word/document.xml.
    Replies that target the deleted comment as their parent are dropped too.

    Returns {output_path, removed} where `removed` is the count of comment
    elements actually deleted (0 if the ID wasn't found).
    """
    if output_path is None:
        stem, ext = os.path.splitext(file_path)
        output_path = f"{stem}_deleted{ext}"

    numeric_id = comment_id.split(":")[-1].strip()  # "Com:1" → "1", "1" → "1"

    removed = _strip_comment_from_zip(file_path, output_path, numeric_id)
    return {"output_path": output_path, "removed": removed}


def _strip_comment_from_zip(src: str, dst: str, numeric_id: str) -> int:
    """Copy zip src → dst, rewriting comments and document parts to drop
    the comment with `numeric_id`. Returns total elements removed."""
    removed = 0
    # Discover which part holds the comments (Word may name it
    # word/comments.xml, word/comments1.xml, etc.).
    with zipfile.ZipFile(src) as zin:
        comments_part = _find_part_by_content_type(
            zin,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        )

    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if comments_part and item.filename == comments_part:
                data, n = _remove_comment_def(data, numeric_id)
                removed += n
            elif item.filename == "word/document.xml":
                data, n = _remove_comment_anchors(data, numeric_id)
                removed += n
            zout.writestr(item, data)
    return removed


def _find_part_by_content_type(zf: zipfile.ZipFile, content_type: str) -> str | None:
    try:
        ct_root = ET.fromstring(zf.read("[Content_Types].xml"))
    except KeyError:
        return None
    for override in ct_root.findall(f"{{{_TYPES_NS}}}Override"):
        if override.get("ContentType") == content_type:
            return override.get("PartName", "").lstrip("/")
    return None


def _remove_comment_def(xml_bytes: bytes, numeric_id: str) -> tuple[bytes, int]:
    """In comments part: drop `<w:comment w:id="N">`. Returns (xml, count)."""
    ET.register_namespace("w", _W_NS)
    root = ET.fromstring(xml_bytes)
    id_attr = f"{{{_W_NS}}}id"
    removed = 0
    for comment in list(root):
        if comment.tag == f"{{{_W_NS}}}comment" and comment.get(id_attr) == numeric_id:
            root.remove(comment)
            removed += 1
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8"), removed


def _remove_comment_anchors(xml_bytes: bytes, numeric_id: str) -> tuple[bytes, int]:
    """In document part: drop commentRangeStart/End/Reference with given id.

    These are inline markers — we walk parents and drop matching children.
    """
    ET.register_namespace("w", _W_NS)
    root = ET.fromstring(xml_bytes)
    id_attr = f"{{{_W_NS}}}id"
    target_tags = {
        f"{{{_W_NS}}}commentRangeStart",
        f"{{{_W_NS}}}commentRangeEnd",
        f"{{{_W_NS}}}commentReference",
    }
    removed = 0
    # ET doesn't expose parent pointers; walk and patch each parent.
    for parent in root.iter():
        for child in list(parent):
            if child.tag in target_tags and child.get(id_attr) == numeric_id:
                parent.remove(child)
                removed += 1
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8"), removed


# Re-export so callers can `except redline.SanitizeError` without reaching
# into adeu directly.
__all__ = [
    "SanitizeError",
    "accept_all_changes",
    "apply_changes",
    "delete_comment",
    "read_with_changes",
    "sanitize",
]
