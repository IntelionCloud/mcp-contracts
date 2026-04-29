"""One-shot generator for binary DOCX fixtures used by redline tests.

Re-run this only when you need to regenerate the binaries (e.g. after a
breaking change in adeu's comment format). Output:

  tests/fixtures/sample_plain.docx       — single paragraph, no comments
  tests/fixtures/sample_with_comment.docx — same paragraph + 1 reviewer comment

The fixtures are committed to git so tests don't need adeu/python-docx at
test time to seed state — only to *read* the binaries.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from docx import Document  # noqa: E402

from core import redline  # noqa: E402


def main() -> None:
    out_dir = Path(__file__).resolve().parent

    # 1. Plain doc — used as a starting point for "add comment" tests.
    plain = out_dir / "sample_plain.docx"
    doc = Document()
    doc.add_paragraph(
        "Compute power is calculated based on Akash hourly rates."
    )
    doc.save(str(plain))
    print(f"wrote {plain}")

    # 2. Doc with one reviewer comment — used as a starting point for
    # "find / reply / delete" tests. We seed it via redline.apply_changes
    # so the format matches what adeu produces in production.
    with_comment = out_dir / "sample_with_comment.docx"
    redline.apply_changes(
        str(plain),
        edits=[{
            "type": "modify",
            "target_text": "Akash hourly rates",
            "new_text": "Akash hourly rates (TBD)",
            "comment": "Need to clarify the rate basis.",
        }],
        author="Original Reviewer",
        output_path=str(with_comment),
    )
    print(f"wrote {with_comment}")


if __name__ == "__main__":
    main()
