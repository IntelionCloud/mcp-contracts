"""
DOCX parser — extract text, tracked changes, and comments from Word documents.
Refactored from /shared/business/scripts/docx_to_md.py
"""
import re
import zipfile
from xml.etree import ElementTree as ET

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
}


def _parse_xml(zf: zipfile.ZipFile, path: str) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(path))
    except KeyError:
        return None


def extract_comments(docx_path: str) -> list[dict]:
    """Extract comments from DOCX as list of {id, author, date, text}."""
    with zipfile.ZipFile(docx_path) as zf:
        root = _parse_xml(zf, 'word/comments.xml')
    if root is None:
        return []
    comments = []
    for comment in root.findall('.//w:comment', NS):
        cid = comment.get(f'{{{NS["w"]}}}id')
        author = comment.get(f'{{{NS["w"]}}}author', '')
        date = comment.get(f'{{{NS["w"]}}}date', '')
        texts = []
        for t in comment.iter(f'{{{NS["w"]}}}t'):
            if t.text:
                texts.append(t.text)
        comments.append({
            'id': cid,
            'author': author,
            'date': date[:10] if date else '',
            'text': ''.join(texts),
        })
    return comments


def extract_tracked_changes(docx_path: str) -> list[dict]:
    """Extract tracked changes (insertions and deletions) from DOCX."""
    with zipfile.ZipFile(docx_path) as zf:
        doc_root = _parse_xml(zf, 'word/document.xml')
    if doc_root is None:
        return []

    w = NS['w']
    changes = []

    # Insertions
    for el in doc_root.iter(f'{{{w}}}ins'):
        cid = el.get(f'{{{w}}}id', '')
        author = el.get(f'{{{w}}}author', '')
        date = el.get(f'{{{w}}}date', '')
        texts = [t.text for t in el.iter(f'{{{w}}}t') if t.text]
        text = ''.join(texts)
        if text.strip():
            changes.append({
                'id': cid,
                'type': 'insertion',
                'author': author,
                'date': date[:10] if date else '',
                'text': text,
            })

    # Deletions
    for el in doc_root.iter(f'{{{w}}}del'):
        cid = el.get(f'{{{w}}}id', '')
        author = el.get(f'{{{w}}}author', '')
        date = el.get(f'{{{w}}}date', '')
        texts = [t.text for t in el.iter(f'{{{w}}}delText') if t.text]
        text = ''.join(texts)
        if text.strip():
            changes.append({
                'id': cid,
                'type': 'deletion',
                'author': author,
                'date': date[:10] if date else '',
                'text': text,
            })

    return changes


def convert_docx_to_md(docx_path: str, accept_changes: bool = False) -> str:
    """Convert DOCX to Markdown string. Delegates to the script's function."""
    import sys
    sys.path.insert(0, '/shared/business/scripts')
    from docx_to_md import convert_docx_to_md as _convert
    return _convert(docx_path, accept_changes=accept_changes)
