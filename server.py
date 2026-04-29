#!/usr/bin/env python3
"""
MCP server for working with legal contracts (DOCX/MD).

Tools:
  - docx_to_md: Convert DOCX to Markdown (preserving tracked changes & comments)
  - md_to_docx: Convert Markdown to DOCX (with multilevel numbering)
  - docx_to_pdf: Convert DOCX to PDF via LibreOffice headless
  - read_contract: Read a contract file and return structured text
  - list_comments: Extract reviewer comments from DOCX
  - list_tracked_changes: Extract tracked changes (insertions/deletions)
"""
import json
import os
import sys

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

# Ensure our modules are importable
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core.docx_parser import (
    convert_docx_to_md,
    extract_comments,
    extract_tracked_changes,
)
from core.docx_builder import build_docx_from_md
from core.pdf_converter import (
    convert_docx_to_pdf,
    LibreOfficeNotInstalled,
    ConversionFailed,
)
from core.contract_model import (
    parse_contract,
    find_clause as _find_clause,
    contract_summary as _contract_summary,
    validate_references as _validate_references,
)

app = Server("docx-contracts")


# ---------------------------------------------------------------------------
# Tool definitions
# ---------------------------------------------------------------------------

TOOLS = [
    Tool(
        name="docx_to_md",
        description=(
            "Конвертировать DOCX-файл договора в Markdown. "
            "Сохраняет tracked changes ({++вставки++}, {--удаления--}) "
            "и комментарии рецензента как сноски [^N]. "
            "Результат сохраняется в .md файл рядом с оригиналом."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Абсолютный путь к DOCX-файлу",
                },
                "accept_changes": {
                    "type": "boolean",
                    "description": "Принять все правки (чистый текст без пометок)",
                    "default": False,
                },
                "output_path": {
                    "type": "string",
                    "description": "Путь для выходного MD-файла (опционально, по умолчанию рядом с DOCX)",
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="md_to_docx",
        description=(
            "Конвертировать Markdown-файл договора в DOCX. "
            "Создаёт Word-документ с автоматической multilevel нумерацией "
            "(1., 1.1., 1.1.1.), форматированием A4, Times New Roman. "
            "Поддерживает tracked changes и таблицы."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Абсолютный путь к MD-файлу",
                },
                "accept_changes": {
                    "type": "boolean",
                    "description": "Принять все правки (чистовик без пометок)",
                    "default": False,
                },
                "output_path": {
                    "type": "string",
                    "description": "Путь для выходного DOCX-файла (опционально)",
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="docx_to_pdf",
        description=(
            "Конвертировать DOCX-файл в PDF через LibreOffice headless. "
            "Сохраняет форматирование, нумерацию, таблицы. "
            "Требует установленный libreoffice (soffice) на хосте."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Абсолютный путь к DOCX-файлу",
                },
                "output_path": {
                    "type": "string",
                    "description": "Путь для выходного PDF-файла (опционально, по умолчанию рядом с DOCX)",
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="read_contract",
        description=(
            "Прочитать договор (DOCX или MD) и вернуть текст. "
            "Для DOCX — автоматически конвертирует в Markdown. "
            "Для MD — читает как есть."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Абсолютный путь к файлу (DOCX или MD)",
                },
                "accept_changes": {
                    "type": "boolean",
                    "description": "Принять все правки при чтении",
                    "default": False,
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="list_comments",
        description=(
            "Извлечь все комментарии рецензента из DOCX-файла. "
            "Возвращает список: автор, дата, текст комментария."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Абсолютный путь к DOCX-файлу",
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="list_tracked_changes",
        description=(
            "Извлечь все tracked changes (вставки и удаления) из DOCX-файла. "
            "Возвращает список: тип (insertion/deletion), автор, дата, текст."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Абсолютный путь к DOCX-файлу",
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="find_clause",
        description=(
            "Найти пункт(ы) договора по номеру или ключевому слову. "
            "Экономит токены — возвращает только найденные пункты, а не весь текст. "
            "Примеры: '4.2' → пункт 4.2, '7' → глава 7 со всеми пунктами, "
            "'оплата' → все пункты про оплату."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Абсолютный путь к файлу (DOCX или MD)",
                },
                "query": {
                    "type": "string",
                    "description": "Номер пункта (e.g. '4.2', '7') или ключевое слово",
                },
            },
            "required": ["file_path", "query"],
        },
    ),
    Tool(
        name="contract_summary",
        description=(
            "Структурное резюме договора: стороны, предмет, цена, сроки, "
            "оглавление всех разделов и пунктов. "
            "Экономит токены — компактный обзор вместо полного текста."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Абсолютный путь к файлу (DOCX или MD)",
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="validate_references",
        description=(
            "Проверить внутренние ссылки (п. X.Y) в договоре на корректность. "
            "Находит ссылки на несуществующие пункты."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Абсолютный путь к файлу (DOCX или MD)",
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="read_sections",
        description=(
            "Прочитать только указанные разделы договора. "
            "Экономит токены — возвращает только нужные главы. "
            "Пример: sections='2,4,9' вернёт только разделы 2, 4 и 9."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Абсолютный путь к файлу (DOCX или MD)",
                },
                "sections": {
                    "type": "string",
                    "description": "Номера разделов через запятую (e.g. '2,4,9')",
                },
            },
            "required": ["file_path", "sections"],
        },
    ),
]


# ---------------------------------------------------------------------------
# Handlers
# ---------------------------------------------------------------------------

@app.list_tools()
async def list_tools():
    return TOOLS


@app.call_tool()
async def call_tool(name: str, arguments: dict):
    try:
        if name == "docx_to_md":
            return _handle_docx_to_md(arguments)
        elif name == "md_to_docx":
            return _handle_md_to_docx(arguments)
        elif name == "docx_to_pdf":
            return _handle_docx_to_pdf(arguments)
        elif name == "read_contract":
            return _handle_read_contract(arguments)
        elif name == "list_comments":
            return _handle_list_comments(arguments)
        elif name == "list_tracked_changes":
            return _handle_list_tracked_changes(arguments)
        elif name == "find_clause":
            return _handle_find_clause(arguments)
        elif name == "contract_summary":
            return _handle_contract_summary(arguments)
        elif name == "validate_references":
            return _handle_validate_references(arguments)
        elif name == "read_sections":
            return _handle_read_sections(arguments)
        else:
            return [TextContent(type="text", text=f"Unknown tool: {name}")]
    except Exception as e:
        return [TextContent(type="text", text=f"Error: {e}")]


def _handle_docx_to_md(args: dict):
    file_path = args["file_path"]
    accept = args.get("accept_changes", False)
    output_path = args.get("output_path")

    if not os.path.exists(file_path):
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    md_text = convert_docx_to_md(file_path, accept_changes=accept)

    if output_path is None:
        output_path = os.path.splitext(file_path)[0] + '.md'

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(md_text)

    return [TextContent(
        type="text",
        text=f"Converted to: {output_path}\n\n{md_text}"
    )]


def _handle_md_to_docx(args: dict):
    file_path = args["file_path"]
    accept = args.get("accept_changes", False)
    output_path = args.get("output_path")

    if not os.path.exists(file_path):
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    result_path = build_docx_from_md(file_path, output_path, accept_changes=accept)

    return [TextContent(
        type="text",
        text=f"Generated DOCX: {result_path}"
    )]


def _handle_docx_to_pdf(args: dict):
    file_path = args["file_path"]
    output_path = args.get("output_path")

    if not os.path.exists(file_path):
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    if not file_path.lower().endswith('.docx'):
        return [TextContent(type="text", text=f"Expected .docx, got: {file_path}")]

    try:
        result_path = convert_docx_to_pdf(file_path, output_path)
    except LibreOfficeNotInstalled as e:
        return [TextContent(type="text", text=str(e))]
    except ConversionFailed as e:
        return [TextContent(type="text", text=f"Conversion failed: {e}")]

    return [TextContent(
        type="text",
        text=f"Generated PDF: {result_path}"
    )]


def _handle_read_contract(args: dict):
    file_path = args["file_path"]
    accept = args.get("accept_changes", False)

    if not os.path.exists(file_path):
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    if file_path.lower().endswith('.docx'):
        text = convert_docx_to_md(file_path, accept_changes=accept)
    elif file_path.lower().endswith('.md'):
        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read()
    else:
        return [TextContent(type="text", text=f"Unsupported format: {file_path}")]

    return [TextContent(type="text", text=text)]


def _handle_list_comments(args: dict):
    file_path = args["file_path"]

    if not os.path.exists(file_path):
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    comments = extract_comments(file_path)

    if not comments:
        return [TextContent(type="text", text="No comments found.")]

    lines = [f"Found {len(comments)} comment(s):\n"]
    for c in comments:
        lines.append(f"- [{c['date']}] **{c['author']}**: {c['text']}")

    return [TextContent(type="text", text='\n'.join(lines))]


def _handle_list_tracked_changes(args: dict):
    file_path = args["file_path"]

    if not os.path.exists(file_path):
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    changes = extract_tracked_changes(file_path)

    if not changes:
        return [TextContent(type="text", text="No tracked changes found.")]

    insertions = [c for c in changes if c['type'] == 'insertion']
    deletions = [c for c in changes if c['type'] == 'deletion']

    lines = [f"Found {len(changes)} change(s): {len(insertions)} insertions, {len(deletions)} deletions\n"]
    for c in changes:
        marker = "+++" if c['type'] == 'insertion' else "---"
        preview = c['text'][:80] + '...' if len(c['text']) > 80 else c['text']
        lines.append(f"[{marker}] {c['author']} ({c['date']}): {preview}")

    return [TextContent(type="text", text='\n'.join(lines))]


def _get_md_text(file_path: str) -> str:
    """Get MD text from DOCX or MD file."""
    if file_path.lower().endswith('.docx'):
        return convert_docx_to_md(file_path, accept_changes=False)
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read()


def _handle_find_clause(args: dict):
    file_path = args["file_path"]
    query = args["query"]

    if not os.path.exists(file_path):
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    md_text = _get_md_text(file_path)
    clauses = parse_contract(md_text)
    results = _find_clause(clauses, query)

    if not results:
        return [TextContent(type="text", text=f"No clauses found for query: {query}")]

    lines = [f"Found {len(results)} clause(s) for '{query}':\n"]
    for c in results:
        level_prefix = "  " * c.level
        lines.append(f"{level_prefix}**п. {c.ref}.** {c.clean_text()}\n")

    return [TextContent(type="text", text='\n'.join(lines))]


def _handle_contract_summary(args: dict):
    file_path = args["file_path"]

    if not os.path.exists(file_path):
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    md_text = _get_md_text(file_path)
    clauses = parse_contract(md_text)
    summary = _contract_summary(clauses, md_text)

    return [TextContent(type="text", text=summary)]


def _handle_validate_references(args: dict):
    file_path = args["file_path"]

    if not os.path.exists(file_path):
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    md_text = _get_md_text(file_path)
    clauses = parse_contract(md_text)
    issues = _validate_references(clauses, md_text)

    if not issues:
        return [TextContent(type="text", text="All internal references are valid.")]

    lines = [f"Found {len(issues)} reference issue(s):\n"]
    for issue in issues:
        lines.append(f"- **{issue['ref']}**: {issue['issue']}")
        lines.append(f"  Context: {issue['context']}")

    return [TextContent(type="text", text='\n'.join(lines))]


def _handle_read_sections(args: dict):
    file_path = args["file_path"]
    sections_str = args["sections"]

    if not os.path.exists(file_path):
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    target_sections = {s.strip() for s in sections_str.split(',')}

    md_text = _get_md_text(file_path)
    clauses = parse_contract(md_text)

    lines = []
    for c in clauses:
        # Chapter number is the first part of ref
        chapter = c.ref.split('.')[0]
        if chapter in target_sections:
            indent = "  " * c.level
            if c.level == 0:
                lines.append(f"\n**{c.ref}. {c.clean_text()}**\n")
            else:
                lines.append(f"{indent}{c.ref}. {c.clean_text()}")

    if not lines:
        return [TextContent(type="text", text=f"No sections found for: {sections_str}")]

    return [TextContent(type="text", text='\n'.join(lines))]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

async def main():
    async with stdio_server() as (read_stream, write_stream):
        await app.run(read_stream, write_stream, app.create_initialization_options())


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
