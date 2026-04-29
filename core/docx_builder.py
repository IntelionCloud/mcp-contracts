"""
DOCX builder — generate Word documents from Markdown.
"""
import os
import tempfile

from core._md_to_docx import build_docx


def build_docx_from_md(md_path: str, output_path: str | None = None,
                       accept_changes: bool = False) -> str:
    """Convert MD to DOCX. Returns path to generated file."""
    if output_path is None:
        base = os.path.splitext(md_path)[0]
        suffix = '_clean' if accept_changes else ''
        output_path = f"{base}{suffix}.docx"

    build_docx(md_path, output_path, accept=accept_changes)
    return output_path


def build_docx_from_text(md_text: str, output_path: str,
                         accept_changes: bool = False) -> str:
    """Convert MD text (string) to DOCX. Returns path to generated file."""
    with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False,
                                     encoding='utf-8') as f:
        f.write(md_text)
        tmp_md = f.name

    try:
        return build_docx_from_md(tmp_md, output_path, accept_changes)
    finally:
        os.unlink(tmp_md)
