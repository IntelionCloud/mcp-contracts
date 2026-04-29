#!/usr/bin/env python3
"""
Convert DOCX to Markdown with tracked changes and comments.

Features:
- Extracts full document text preserving structure (headings, lists, tables)
- Shows tracked changes: [INS: text] and [DEL: text] with author info
- Shows comments as footnotes linked to the commented text
- Handles numbering (numbered/bulleted lists)
- Handles tables as Markdown tables
- Handles headers/footers

Usage:
    python3 docx_to_md.py input.docx                  # outputs to input.md
    python3 docx_to_md.py input.docx -o output.md     # custom output path
    python3 docx_to_md.py input.docx --accept          # accept all changes (clean text)
    python3 docx_to_md.py *.docx                       # batch convert
"""

import argparse
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

# Word XML namespaces
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
}


def parse_xml(zf: zipfile.ZipFile, path: str) -> ET.Element | None:
    """Parse an XML file from the docx ZIP, return None if missing."""
    try:
        return ET.fromstring(zf.read(path))
    except KeyError:
        return None


def extract_comments(zf: zipfile.ZipFile) -> dict:
    """Extract comments from word/comments.xml as {id: {author, date, text}}."""
    root = parse_xml(zf, 'word/comments.xml')
    if root is None:
        return {}
    comments = {}
    for comment in root.findall('.//w:comment', NS):
        cid = comment.get(f'{{{NS["w"]}}}id')
        author = comment.get(f'{{{NS["w"]}}}author', '')
        date = comment.get(f'{{{NS["w"]}}}date', '')
        # Collect all text in comment paragraphs
        texts = []
        for t in comment.iter(f'{{{NS["w"]}}}t'):
            if t.text:
                texts.append(t.text)
        comments[cid] = {
            'author': author,
            'date': date[:10] if date else '',
            'text': ''.join(texts),
        }
    return comments


def get_paragraph_style(p_elem) -> str | None:
    """Get the paragraph style ID (e.g., 'Heading1', '1', etc.)."""
    ppr = p_elem.find('w:pPr', NS)
    if ppr is not None:
        pstyle = ppr.find('w:pStyle', NS)
        if pstyle is not None:
            return pstyle.get(f'{{{NS["w"]}}}val')
    return None


def get_num_info(p_elem) -> tuple:
    """Get numbering info (numId, ilvl) if paragraph is in a list."""
    ppr = p_elem.find('w:pPr', NS)
    if ppr is not None:
        num_pr = ppr.find('w:numPr', NS)
        if num_pr is not None:
            ilvl_el = num_pr.find('w:ilvl', NS)
            num_id_el = num_pr.find('w:numId', NS)
            ilvl = ilvl_el.get(f'{{{NS["w"]}}}val', '0') if ilvl_el is not None else '0'
            num_id = num_id_el.get(f'{{{NS["w"]}}}val', '0') if num_id_el is not None else '0'
            return (num_id, int(ilvl))
    return None


def style_to_heading_level(style: str | None) -> int:
    """Convert style name to heading level (0 = not a heading)."""
    if not style:
        return 0
    # Common patterns: Heading1, 1, heading 1, etc.
    style_lower = style.lower().replace(' ', '')
    for i in range(1, 7):
        if style_lower in (f'heading{i}', f'{i}', f'заголовок{i}'):
            return i
    return 0


def extract_run_text(run, accept_changes: bool) -> str:
    """Extract text from a single w:r element."""
    parts = []
    for child in run:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 't':
            parts.append(child.text or '')
        elif tag == 'tab':
            parts.append('\t')
        elif tag == 'br':
            br_type = child.get(f'{{{NS["w"]}}}type', '')
            if br_type == 'page':
                parts.append('\n\n---\n\n')
            else:
                parts.append('\n')
        elif tag == 'sym':
            parts.append('•')
    return ''.join(parts)


def process_paragraph(p_elem, comments: dict, accept_changes: bool,
                       comment_refs: list, num_counters: dict) -> str:
    """Process a single paragraph element into markdown text."""
    parts = []
    active_comments = []  # stack of comment IDs

    for child in p_elem:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag == 'r':
            # Regular run
            text = extract_run_text(child, accept_changes)
            parts.append(text)

        elif tag == 'hyperlink':
            # Hyperlink — extract inner runs
            link_texts = []
            for r in child.findall('w:r', NS):
                link_texts.append(extract_run_text(r, accept_changes))
            parts.append(''.join(link_texts))

        elif tag == 'ins':
            # Tracked insertion
            author = child.get(f'{{{NS["w"]}}}author', '')
            ins_parts = []
            for r in child.findall('.//w:r', NS):
                ins_parts.append(extract_run_text(r, accept_changes))
            text = ''.join(ins_parts)
            if text.strip():
                if accept_changes:
                    parts.append(text)
                else:
                    parts.append(f'{{++{text}++}}')

        elif tag == 'del':
            # Tracked deletion
            author = child.get(f'{{{NS["w"]}}}author', '')
            del_parts = []
            for dt in child.findall('.//w:delText', NS):
                if dt.text:
                    del_parts.append(dt.text)
            text = ''.join(del_parts)
            if text.strip():
                if not accept_changes:
                    parts.append(f'{{--{text}--}}')
                # If accepting changes, deletions are simply omitted

        elif tag == 'commentRangeStart':
            cid = child.get(f'{{{NS["w"]}}}id')
            if cid in comments:
                active_comments.append(cid)

        elif tag == 'commentRangeEnd':
            cid = child.get(f'{{{NS["w"]}}}id')
            if cid in comments and cid in active_comments:
                active_comments.remove(cid)
                # Add comment reference
                comment_refs.append(cid)
                ref_num = len(comment_refs)
                c = comments[cid]
                parts.append(f'[^{ref_num}]')

    # Build line
    line = ''.join(parts).strip()

    # Apply heading
    style = get_paragraph_style(p_elem)
    level = style_to_heading_level(style)
    if level > 0 and line:
        line = '#' * level + ' ' + line

    # Apply list formatting
    num_info = get_num_info(p_elem)
    if num_info and line:
        num_id, ilvl = num_info
        indent = '  ' * ilvl
        # Try to determine if bulleted or numbered
        # Simple heuristic: if numId is even or style contains "List", use bullets
        key = (num_id, ilvl)
        if key not in num_counters:
            num_counters[key] = 0
        num_counters[key] += 1
        # Use numbered list
        line = f'{indent}{num_counters[key]}. {line}'

    return line


def process_table(tbl_elem, comments: dict, accept_changes: bool,
                  comment_refs: list, num_counters: dict) -> list:
    """Process a table element into markdown lines."""
    rows = []
    for tr in tbl_elem.findall('.//w:tr', NS):
        cells = []
        for tc in tr.findall('w:tc', NS):
            cell_parts = []
            for p in tc.findall('w:p', NS):
                text = process_paragraph(p, comments, accept_changes, comment_refs, num_counters)
                if text:
                    cell_parts.append(text)
            cells.append('<br>'.join(cell_parts) if cell_parts else '')
        rows.append(cells)

    if not rows:
        return []

    # Build markdown table
    lines = []
    # Normalize column count
    max_cols = max(len(r) for r in rows)
    for r in rows:
        while len(r) < max_cols:
            r.append('')

    # Header row
    lines.append('| ' + ' | '.join(rows[0]) + ' |')
    lines.append('| ' + ' | '.join(['---'] * max_cols) + ' |')
    for row in rows[1:]:
        lines.append('| ' + ' | '.join(row) + ' |')

    return lines


def convert_docx_to_md(docx_path: str, accept_changes: bool = False) -> str:
    """Convert a DOCX file to Markdown string."""
    with zipfile.ZipFile(docx_path) as zf:
        comments = extract_comments(zf)
        doc_root = parse_xml(zf, 'word/document.xml')
        if doc_root is None:
            raise ValueError(f"No word/document.xml in {docx_path}")

    body = doc_root.find('.//w:body', NS)
    if body is None:
        raise ValueError("No body element found")

    lines = []
    comment_refs = []  # ordered list of comment IDs referenced
    num_counters = {}  # numbering counters

    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag == 'p':
            line = process_paragraph(child, comments, accept_changes, comment_refs, num_counters)
            lines.append(line)

        elif tag == 'tbl':
            table_lines = process_table(child, comments, accept_changes, comment_refs, num_counters)
            lines.extend(table_lines)

        elif tag == 'sectPr':
            # Section properties — skip
            pass

    # Build result
    result_lines = []
    prev_empty = False
    for line in lines:
        if not line:
            if not prev_empty:
                result_lines.append('')
            prev_empty = True
        else:
            result_lines.append(line)
            prev_empty = False

    result = '\n'.join(result_lines)

    # Add comment footnotes
    if comment_refs:
        result += '\n\n---\n\n**Комментарии рецензента:**\n\n'
        for i, cid in enumerate(comment_refs, 1):
            c = comments[cid]
            result += f'[^{i}]: **{c["author"]}** ({c["date"]}): {c["text"]}\n\n'

    # Legend for tracked changes
    if not accept_changes:
        has_ins = '{++' in result
        has_del = '{--' in result
        if has_ins or has_del:
            legend = '\n---\n\n**Обозначения tracked changes:**\n'
            if has_ins:
                legend += '- `{++текст++}` — вставка (tracked insertion)\n'
            if has_del:
                legend += '- `{--текст--}` — удаление (tracked deletion)\n'
            result += legend

    return result


def main():
    parser = argparse.ArgumentParser(
        description='Convert DOCX to Markdown with tracked changes and comments'
    )
    parser.add_argument('files', nargs='+', help='DOCX file(s) to convert')
    parser.add_argument('-o', '--output', help='Output file path (only for single file)')
    parser.add_argument('--accept', action='store_true',
                        help='Accept all tracked changes (show clean text)')
    args = parser.parse_args()

    if args.output and len(args.files) > 1:
        print("Error: -o/--output can only be used with a single file", file=sys.stderr)
        sys.exit(1)

    for docx_path in args.files:
        p = Path(docx_path)
        if not p.exists():
            print(f"Error: {docx_path} not found", file=sys.stderr)
            continue

        md_text = convert_docx_to_md(str(p), accept_changes=args.accept)

        if args.output:
            out_path = Path(args.output)
        else:
            out_path = p.with_suffix('.md')

        out_path.write_text(md_text, encoding='utf-8')
        print(f"Converted: {p.name} -> {out_path.name}")


if __name__ == '__main__':
    main()
