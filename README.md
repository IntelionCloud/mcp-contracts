# mcp-contracts

[Model Context Protocol](https://modelcontextprotocol.io/) server that lets
AI agents work with legal contracts (DOCX/MD) without paying the round-trip
cost of dumping a 30-page document into the context window.

## Why

LLM agents are good at reasoning over contract text, but the file formats are
hostile: DOCX is binary with tracked changes and reviewer comments, MD is
plaintext with no structural awareness. Naively feeding either to a model is
expensive and noisy.

This server bridges that gap. It exposes targeted tools — read a single
clause, list reviewer comments, validate cross-references, summarize the
structure — so the agent fetches just what it needs. It also handles the
common authoring tasks (DOCX ↔ MD, DOCX → PDF) so the same agent can edit
and publish without leaving its tooling.

## Use cases

### 1. Review a DOCX with reviewer comments

Counterparty sends back a DOCX with their comments and tracked changes.
You want an agent to extract the comments, find the corresponding clauses,
and propose responses.

```
agent: list_comments(file_path)
       → "5 comments by lawyer@counterparty.com"
agent: find_clause(file_path, query="4.2")  # the clause being commented on
agent: list_tracked_changes(file_path)      # what they edited inline
→ agent drafts a reply per comment, you review and reply
```

No need to load the whole document into context — the agent pulls just the
clauses being argued about.

### 2. Iterative editing via AI agent

When a contract sees frequent edits, keep it in MD (a format the agent
edits cleanly) and convert to DOCX only when sending to colleagues.

```
1. agent edits contract.md (clauses, prices, dates — plain text diff)
2. md_to_docx(contract.md)                  # → contract.docx
3. you send contract.docx to counterparty
4. they reply with their commented version  → see use case #1
5. agent applies accepted changes back into contract.md
6. loop
```

Source of truth stays text-based and diffable; DOCX is the wire format.

### 3. MD-in-repo, PDF for publication

For public-facing legal docs (Terms of Service, privacy policy, conditions
of service), keep the master in MD inside your repo and auto-generate the
PDF that's actually served on the website.

```
1. You edit terms.md and commit to git
2. CI / pre-deploy hook calls:
     md_to_docx(terms.md)   → terms.docx  (intermediate, with proper styling)
     docx_to_pdf(terms.docx) → terms.pdf  (the file users download)
3. Diffs are reviewable in PRs (text), but readers get a polished PDF
```

Versioning, blame, branch reviews — all work as on regular code; the binary
DOCX/PDF artifacts are byproducts.

## Tools

### Conversion

| Tool | Direction | Notes |
|------|-----------|-------|
| `docx_to_md` | DOCX → MD | Preserves tracked changes (`{++ins++}`, `{--del--}`) and reviewer comments as `[^N]` footnotes. |
| `md_to_docx` | MD → DOCX | Multilevel numbering (1., 1.1., 1.1.1.), A4, Times New Roman. |
| `docx_to_pdf` | DOCX → PDF | LibreOffice headless. Tries local `soffice`; falls back to a docker image (see [Setup](#setup)). |

### Token-saving readers

| Tool | What it returns |
|------|-----------------|
| `read_contract` | Full text (DOCX auto-converted to MD). |
| `read_sections` | Only the requested chapters, e.g. `sections="2,4,9"`. |
| `find_clause` | A specific clause by number (`"4.2"`) or by keyword (`"payment"`). |
| `contract_summary` | Parties, subject, price, term, table of contents — compact overview. |

### Review helpers

| Tool | What it does |
|------|--------------|
| `list_comments` | Reviewer comments with author, date, text. |
| `list_tracked_changes` | Insertions and deletions with author, date, text. |
| `validate_references` | Finds broken internal references (e.g. `см. п. 7.4` when 7.4 doesn't exist). |

All tools take an absolute `file_path`. See `server.py` for full schemas.

## Installation

### 1. Clone

```bash
git clone https://github.com/IntelionCloud/mcp-contracts.git
cd mcp-contracts
```

### 2. Python dependencies

Requires Python 3.10+.

```bash
pip install -r requirements.txt
```

### 3. PDF conversion (optional, only if you use `docx_to_pdf`)

Either install LibreOffice on the host:

```bash
# Debian/Ubuntu
sudo apt install libreoffice-core libreoffice-writer fonts-open-sans
# macOS
brew install --cask libreoffice
```

…or build the bundled docker image once (auto-used as fallback when no
local `soffice` is found):

```bash
docker build -t mcp-docx-soffice:latest docker/
```

### 4. Register in your MCP client

The server runs as a stdio process. Add it to your client config.

**Claude Code / Claude Desktop** — add to the project's `.mcp.json`
(or to `~/Library/Application Support/Claude/claude_desktop_config.json`
for Claude Desktop):

```json
{
  "mcpServers": {
    "docx-contracts": {
      "command": "python3",
      "args": ["/absolute/path/to/mcp-contracts/server.py"]
    }
  }
}
```

Restart the MCP client; the tools (`docx_to_md`, `md_to_docx`,
`docx_to_pdf`, …) should now appear in the tool list.

### 5. Smoke test

```bash
echo '{"jsonrpc":"2.0","id":1,"method":"tools/list"}' \
  | python3 server.py
```

You should see a JSON response listing all tools. If the process complains
about missing modules, re-run step 2 inside the right virtualenv.

## Localization

Tools auto-detect contract language by Cyrillic-vs-Latin character ratio
over the first 50 KB of text. Three modes are recognized:

| Detected | When | Effect |
|----------|------|--------|
| `ru` | Cyrillic ≥ 70% of letters | Russian patterns + Russian labels (`Цена`, `Срок`, `Структура`, `п.`) |
| `en` | Cyrillic ≤ 30% | English patterns + English labels (`Price`, `Deadline`, `Structure`, `Clause`) |
| `ru+en` | mid-range (typical of two-column EN-RU contracts) | Both pattern sets applied with dedupe; bilingual labels (`Цена / Price`, `Структура / Structure`) |

Override auto-detection by passing `language="ru" | "en" | "ru+en"` to
any of: `find_clause`, `contract_summary`, `validate_references`,
`read_sections`.

To add a third language: extend `PATTERNS` and `LABELS` in
`core/i18n.py`, update the `Lang` `Literal`, and adjust the threshold
in `detect_lang` if needed.

## Legacy `.doc` support

All read/conversion tools (`docx_to_md`, `read_contract`, `find_clause`, …)
accept legacy binary `.doc` files (Word 97–2003) transparently — they are
auto-converted to `.docx` via LibreOffice headless and cached in
`tempfile.gettempdir()` keyed by file mtime, so repeated tool calls don't
re-invoke `soffice`. Requires either a local `soffice` install or the
`mcp-docx-soffice` docker image (see PDF setup above).

Cyrillic filenames stored in NFD form (common when copied from macOS)
are also handled — paths are normalized to NFC at the boundary.

## Tests

```bash
pip install -r requirements-dev.txt
pytest
```

## License

This wrapper code is **non-commercial use only**: free for personal,
academic, or internal-tooling purposes. Commercial use (reselling, SaaS,
or use in revenue-generating products) requires written permission from
Intelion Cloud.

Bundled dependencies retain their original licenses — most notably
[`adeu`](https://github.com/dealfluence/adeu) is **MIT-licensed** and is
*not* affected by this restriction. The non-commercial clause covers only
the wrapper layer in this repository (the MCP server, tool schemas, and
helper modules under `core/`).
