#!/usr/bin/env python3
"""
MCP server for working with legal contracts (DOCX/MD).

Tools:
  - docx_to_md: Convert DOCX to Markdown (preserving tracked changes & comments)
  - md_to_docx: Convert Markdown to DOCX (with multilevel numbering)
  - docx_to_pdf: Convert DOCX to PDF via LibreOffice headless
  - read_contract: Read a contract file and return structured text
  - read_with_changes: Render DOCX with track-changes markers inline
  - apply_changes: Apply a batch of redline edits (modify/accept/reject/reply)
  - accept_all_changes: Accept every tracked change in a DOCX
  - sanitize_docx: Strip metadata/author IDs before sharing externally
  - list_comments: Extract reviewer comments from DOCX
  - list_tracked_changes: Extract tracked changes (insertions/deletions)
"""
import json
import logging
import os
import sys

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

# MCP servers communicate over stdout, so logs go to stderr.
logging.basicConfig(
    stream=sys.stderr,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("mcp-contracts")

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
from core.i18n import Lang, labels_for as _labels_for
from core import redline as _redline
from core.doc_compat import resolve_path as _resolve_path


def _exists(path: str) -> tuple[bool, str]:
    """Check existence with NFC/NFD path normalization. Returns (ok, real_path).

    macOS-copied filenames may be in NFD form on disk; user input is NFC.
    Resolving here at the boundary lets every handler accept either form.
    """
    real = _resolve_path(path)
    return os.path.exists(real), real


# Reusable JSON-Schema fragment for the optional language override.
# Enum tracks `core.i18n.Lang` — single source of truth.
_LANGUAGE_SCHEMA = {
    "type": "string",
    "enum": list(Lang.__args__),  # ("ru", "en", "ru+en")
    "description": (
        "Override auto-detected contract language. Affects parsing patterns "
        "(parties / price / deadline / cross-refs) and the language of output "
        "labels. Use 'ru+en' to force bilingual mode. "
        "Default: auto-detect by character ratio."
    ),
}

app = Server("docx-contracts")


# ---------------------------------------------------------------------------
# Tool definitions
# ---------------------------------------------------------------------------

TOOLS = [
    Tool(
        name="docx_to_md",
        description=(
            "Convert a DOCX contract to Markdown. Preserves tracked changes "
            "({++insertions++}, {--deletions--}) and reviewer comments as [^N] "
            "footnotes. Result is written to a .md file next to the original."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to the DOCX file",
                },
                "accept_changes": {
                    "type": "boolean",
                    "description": "Accept all tracked changes (clean text without markup)",
                    "default": False,
                },
                "output_path": {
                    "type": "string",
                    "description": "Path for the output MD file (optional, defaults next to DOCX)",
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="md_to_docx",
        description=(
            "Convert a Markdown contract to DOCX. Produces a Word document with "
            "automatic multilevel numbering (1., 1.1., 1.1.1.), A4 layout, "
            "Times New Roman. Supports tracked changes and tables."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to the MD file",
                },
                "accept_changes": {
                    "type": "boolean",
                    "description": "Accept all tracked changes (clean output without markup)",
                    "default": False,
                },
                "output_path": {
                    "type": "string",
                    "description": "Path for the output DOCX file (optional)",
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="docx_to_pdf",
        description=(
            "Convert DOCX to PDF via LibreOffice headless. Preserves formatting, "
            "numbering, tables. Requires libreoffice (soffice) on the host."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to the DOCX file",
                },
                "output_path": {
                    "type": "string",
                    "description": "Path for the output PDF (optional, defaults next to DOCX)",
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="read_contract",
        description=(
            "Read a contract (DOCX or MD) and return its text. DOCX is "
            "auto-converted to Markdown; MD is read as-is."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to the file (DOCX or MD)",
                },
                "accept_changes": {
                    "type": "boolean",
                    "description": "Accept all tracked changes when reading",
                    "default": False,
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="list_comments",
        description=(
            "Extract every reviewer comment from a DOCX file. Returns a list "
            "with author, date, and comment text."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to the DOCX file",
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="list_tracked_changes",
        description=(
            "Extract every tracked change (insertions and deletions) from a "
            "DOCX file. Returns a list with type (insertion/deletion), "
            "author, date, text."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to the DOCX file",
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="find_clause",
        description=(
            "Find clause(s) in a contract by number or keyword. Token-saving — "
            "returns only matching clauses, not the whole text. "
            "Examples: '4.2' → clause 4.2; '7' → chapter 7 with all sub-clauses; "
            "'payment' / 'оплата' → every clause about payment."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to the file (DOCX or MD)",
                },
                "query": {
                    "type": "string",
                    "description": "Clause number (e.g. '4.2', '7') or keyword",
                },
                "language": _LANGUAGE_SCHEMA,
            },
            "required": ["file_path", "query"],
        },
    ),
    Tool(
        name="contract_summary",
        description=(
            "Structural overview of a contract: parties, subject, price, term, "
            "table of contents. Token-saving — a compact view instead of the "
            "full text. Auto-detects RU/EN."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to the file (DOCX or MD)",
                },
                "language": _LANGUAGE_SCHEMA,
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="validate_references",
        description=(
            "Validate internal cross-references (e.g. 'см. п. X.Y' / 'see clause "
            "X.Y') in a contract. Reports references that point to non-existent "
            "clauses. Auto-detects RU/EN."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to the file (DOCX or MD)",
                },
                "language": _LANGUAGE_SCHEMA,
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="read_sections",
        description=(
            "Read only the specified chapters of a contract. Token-saving — "
            "returns just the requested chapters. "
            "Example: sections='2,4,9' returns only chapters 2, 4 and 9."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to the file (DOCX or MD)",
                },
                "sections": {
                    "type": "string",
                    "description": "Comma-separated chapter numbers (e.g. '2,4,9')",
                },
                "language": _LANGUAGE_SCHEMA,
            },
            "required": ["file_path", "sections"],
        },
    ),
    Tool(
        name="read_with_changes",
        description=(
            "Render DOCX text with inline tracked-changes and comment markers. "
            "Insertions/deletions appear as CriticMarkup ({++ins++}, {--del--}); "
            "comments are anchored as [Com:N]. Use clean_view=True for the "
            "'accepted' rendering without markup. Useful to see which reviewer "
            "suggestions are still pending."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to the DOCX file",
                },
                "clean_view": {
                    "type": "boolean",
                    "description": "Hide tracked-changes markup (show 'accepted' text)",
                    "default": False,
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="apply_changes",
        description=(
            "Apply a batch of edits to a DOCX without losing styles or numbering. "
            "Edits land in the file as **native Word Track Changes** (w:ins / "
            "w:del) — lawyers see them as the author's insertions/deletions. "
            "\n\nEach edit in the array is one of:\n"
            "• {type:'modify', target_text, new_text, comment?} — find-and-replace (fuzzy match)\n"
            "• {type:'accept', target_id, comment?} — accept an existing change by ID (Chg:N)\n"
            "• {type:'reject', target_id, comment?} — reject an existing change by ID\n"
            "• {type:'reply',  target_id, text}    — reply to a comment (Com:N)\n"
            "\nIDs are obtained via read_with_changes — they appear inline as [Chg:N]/[Com:N]."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to the DOCX file",
                },
                "edits": {
                    "type": "array",
                    "description": "List of edits (see tool description)",
                    "items": {
                        "type": "object",
                        "properties": {
                            "type": {
                                "type": "string",
                                "enum": ["modify", "accept", "reject", "reply"],
                            },
                            "target_text": {
                                "type": "string",
                                "description": "For type=modify: exact text to find",
                            },
                            "new_text": {
                                "type": "string",
                                "description": "For type=modify: replacement text",
                            },
                            "target_id": {
                                "type": "string",
                                "description": "For accept/reject/reply: ID like 'Chg:12' or 'Com:5'",
                            },
                            "text": {
                                "type": "string",
                                "description": "For type=reply: the reply body",
                            },
                            "comment": {
                                "type": "string",
                                "description": "Optional rationale (not for reply)",
                            },
                        },
                        "required": ["type"],
                    },
                },
                "author": {
                    "type": "string",
                    "description": "Author name visible in Word",
                    "default": "AI Copilot",
                },
                "output_path": {
                    "type": "string",
                    "description": "Output DOCX path (defaults to <stem>_redlined.docx)",
                },
            },
            "required": ["file_path", "edits"],
        },
    ),
    Tool(
        name="accept_all_changes",
        description=(
            "Accept every tracked change in a DOCX and save a clean copy. "
            "Use before publishing to PDF or sending to a counterparty once "
            "all edits are agreed."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to the DOCX file",
                },
                "output_path": {
                    "type": "string",
                    "description": "Output DOCX path (defaults to <stem>_accepted.docx)",
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="sanitize_docx",
        description=(
            "Strip metadata, author names, and internal tracking IDs before "
            "sending the file outside. Default (full mode) accepts all changes "
            "(if accept_all=True) and removes all comments. With keep_markup="
            "True the markup stays, only metadata is cleaned."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to the DOCX file",
                },
                "output_path": {
                    "type": "string",
                    "description": "Output DOCX path (defaults to <stem>_sanitized.docx)",
                },
                "keep_markup": {
                    "type": "boolean",
                    "description": "Keep tracked changes and open comments",
                    "default": False,
                },
                "accept_all": {
                    "type": "boolean",
                    "description": "Accept unresolved tracked changes (full mode only)",
                    "default": False,
                },
                "author": {
                    "type": "string",
                    "description": "Replace every author name with this value",
                },
            },
            "required": ["file_path"],
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
        elif name == "read_with_changes":
            return _handle_read_with_changes(arguments)
        elif name == "apply_changes":
            return _handle_apply_changes(arguments)
        elif name == "accept_all_changes":
            return _handle_accept_all_changes(arguments)
        elif name == "sanitize_docx":
            return _handle_sanitize_docx(arguments)
        else:
            return [TextContent(type="text", text=f"Unknown tool: {name}")]
    except _redline.SanitizeError as e:
        # Actionable: tells the user how to unblock the sanitize.
        return [TextContent(
            type="text",
            text=(
                f"Sanitize blocked: {e}\n"
                "Hint: pass accept_all=True to accept unresolved tracked "
                "changes, or keep_markup=True to leave them in place while "
                "still scrubbing metadata."
            ),
        )]
    except Exception as e:
        # Full traceback to stderr (visible in Claude's MCP logs); short
        # message to the agent.
        logger.exception("tool %s raised", name)
        return [TextContent(type="text", text=f"Error: {type(e).__name__}: {e}")]


def _handle_docx_to_md(args: dict):
    file_path = args["file_path"]
    accept = args.get("accept_changes", False)
    output_path = args.get("output_path")

    ok, file_path = _exists(file_path)
    if not ok:
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

    ok, file_path = _exists(file_path)
    if not ok:
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    result_path = build_docx_from_md(file_path, output_path, accept_changes=accept)

    return [TextContent(
        type="text",
        text=f"Generated DOCX: {result_path}"
    )]


def _handle_docx_to_pdf(args: dict):
    file_path = args["file_path"]
    output_path = args.get("output_path")

    ok, file_path = _exists(file_path)
    if not ok:
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

    ok, file_path = _exists(file_path)
    if not ok:
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    low = file_path.lower()
    if low.endswith('.docx') or low.endswith('.doc'):
        # .doc is auto-converted to .docx inside convert_docx_to_md.
        text = convert_docx_to_md(file_path, accept_changes=accept)
    elif low.endswith('.md'):
        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read()
    else:
        return [TextContent(type="text", text=f"Unsupported format: {file_path}")]

    return [TextContent(type="text", text=text)]


def _handle_list_comments(args: dict):
    file_path = args["file_path"]

    ok, file_path = _exists(file_path)
    if not ok:
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

    ok, file_path = _exists(file_path)
    if not ok:
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
    """Get MD text from DOCX/DOC or MD file.

    Legacy `.doc` is transparently routed through `convert_docx_to_md`,
    which calls `ensure_docx` to materialize a `.docx` via LibreOffice.
    Without this branch, `.doc` would slip into the text-read fallback
    and crash on the OLE compound-file header (0xD0CF11E0 → invalid UTF-8).
    """
    low = file_path.lower()
    if low.endswith('.docx') or low.endswith('.doc'):
        return convert_docx_to_md(file_path, accept_changes=False)
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read()


def _handle_find_clause(args: dict):
    file_path = args["file_path"]
    query = args["query"]
    language = args.get("language")

    ok, file_path = _exists(file_path)
    if not ok:
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    md_text = _get_md_text(file_path)
    clauses = parse_contract(md_text)
    results = _find_clause(clauses, query)

    if not results:
        return [TextContent(type="text", text=f"No clauses found for query: {query}")]

    L = _labels_for(language, md_text)
    lines = [f"Found {len(results)} clause(s) for '{query}':\n"]
    for c in results:
        level_prefix = "  " * c.level
        lines.append(
            f"{level_prefix}**{L['clause_prefix']} {c.ref}.** {c.clean_text()}\n"
        )

    return [TextContent(type="text", text='\n'.join(lines))]


def _handle_contract_summary(args: dict):
    file_path = args["file_path"]
    language = args.get("language")

    ok, file_path = _exists(file_path)
    if not ok:
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    md_text = _get_md_text(file_path)
    clauses = parse_contract(md_text)
    summary = _contract_summary(clauses, md_text, language=language)

    return [TextContent(type="text", text=summary)]


def _handle_validate_references(args: dict):
    file_path = args["file_path"]
    language = args.get("language")

    ok, file_path = _exists(file_path)
    if not ok:
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    md_text = _get_md_text(file_path)
    clauses = parse_contract(md_text)
    issues = _validate_references(clauses, md_text, language=language)

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
    # language is accepted but currently unused — read_sections has no
    # language-dependent rendering. Reserved for future "Section"/"Раздел"
    # heading style.
    _ = args.get("language")

    ok, file_path = _exists(file_path)
    if not ok:
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    target_sections = {s.strip() for s in sections_str.split(',')}

    md_text = _get_md_text(file_path)
    clauses = parse_contract(md_text)

    lines = []
    for c in clauses:
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


def _handle_read_with_changes(args: dict):
    file_path = args["file_path"]
    clean_view = args.get("clean_view", False)

    ok, file_path = _exists(file_path)
    if not ok:
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    text = _redline.read_with_changes(file_path, clean_view=clean_view)
    return [TextContent(type="text", text=text)]


def _handle_apply_changes(args: dict):
    file_path = args["file_path"]
    edits = args["edits"]
    author = args.get("author", "AI Copilot")
    output_path = args.get("output_path")

    ok, file_path = _exists(file_path)
    if not ok:
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    result = _redline.apply_changes(
        file_path, edits, author=author, output_path=output_path
    )
    lines = [
        f"Wrote: {result['output_path']}",
        f"Applied: {result['applied']} | Skipped: {result['skipped']}",
    ]
    if result.get("details"):
        lines.append("Skip reasons:")
        for d in result["details"]:
            lines.append(f"  - {d}")
    return [TextContent(type="text", text="\n".join(lines))]


def _handle_accept_all_changes(args: dict):
    file_path = args["file_path"]
    output_path = args.get("output_path")

    ok, file_path = _exists(file_path)
    if not ok:
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    out = _redline.accept_all_changes(file_path, output_path=output_path)
    return [TextContent(type="text", text=f"Wrote (all changes accepted): {out}")]


def _handle_sanitize_docx(args: dict):
    file_path = args["file_path"]
    output_path = args.get("output_path")
    keep_markup = args.get("keep_markup", False)
    accept_all = args.get("accept_all", False)
    author = args.get("author")

    ok, file_path = _exists(file_path)
    if not ok:
        return [TextContent(type="text", text=f"File not found: {file_path}")]

    result = _redline.sanitize(
        file_path,
        output_path=output_path,
        keep_markup=keep_markup,
        accept_all=accept_all,
        author=author,
    )
    summary = (
        f"Wrote: {result['output_path']}\n"
        f"Status: {result['status']}\n"
        f"Tracked changes — found: {result['tracked_changes_found']}, "
        f"accepted: {result['tracked_changes_accepted']}\n"
        f"Comments — removed: {result['comments_removed']}, "
        f"kept: {result['comments_kept']}\n"
        f"Metadata stripped: {', '.join(result['metadata_stripped']) or '—'}"
    )
    if result["warnings"]:
        summary += "\nWarnings:\n  " + "\n  ".join(result["warnings"])
    return [TextContent(type="text", text=summary)]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

async def main():
    async with stdio_server() as (read_stream, write_stream):
        await app.run(read_stream, write_stream, app.create_initialization_options())


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
