"""
Surgical DOCX editing — apply tracked changes, accept/reject existing changes,
reply to comments, sanitize for sharing.

Wraps `adeu.RedlineEngine` and `adeu.sanitize.sanitize_docx` (MIT-licensed
upstream — see LICENSE files in the installed package).
"""
from __future__ import annotations

import os
from io import BytesIO

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

    Returns {output_path, applied, skipped}.
    """
    if output_path is None:
        stem, ext = os.path.splitext(file_path)
        output_path = f"{stem}_redlined{ext}"

    engine = _load_engine(file_path, author=author)
    changes = _build_changes(edits)

    applied = 0
    skipped = 0

    # ModifyText edits go through apply_edits; structural changes
    # (accept/reject/reply) are applied via the engine's batch dispatcher
    # if available, falling back to per-item helpers.
    modify_edits = [c for c in changes if c.type == "modify"]
    structural = [c for c in changes if c.type != "modify"]

    if modify_edits:
        result = engine.apply_edits(modify_edits)
        # adeu returns (applied, skipped) tuple
        try:
            a, s = result
            applied += int(a)
            skipped += int(s)
        except (TypeError, ValueError):
            applied += len(modify_edits)

    for change in structural:
        method = {
            "accept": "accept_change",
            "reject": "reject_change",
            "reply": "reply_to_comment",
        }[change.type]
        if hasattr(engine, method):
            ok = getattr(engine, method)(change)
            if ok:
                applied += 1
            else:
                skipped += 1
        else:
            # Older API: skip with a warning rather than crash.
            skipped += 1

    _save_engine(engine, output_path)
    return {
        "output_path": output_path,
        "applied": applied,
        "skipped": skipped,
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


# Re-export so callers can `except redline.SanitizeError` without reaching
# into adeu directly.
__all__ = [
    "SanitizeError",
    "accept_all_changes",
    "apply_changes",
    "read_with_changes",
    "sanitize",
]
