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

## Setup

The server runs as a stdio process; register it in your MCP client config:

```json
{
  "mcpServers": {
    "docx-contracts": {
      "command": "python3",
      "args": ["/path/to/mcp-contracts/server.py"]
    }
  }
}
```

### Dependencies

```bash
pip install -r requirements.txt
```

### PDF conversion

`docx_to_pdf` needs LibreOffice. Either install `soffice` on the host:

```bash
# Debian/Ubuntu
apt install libreoffice-core libreoffice-writer fonts-open-sans
# macOS
brew install --cask libreoffice
```

…or build the bundled docker image once (used as fallback):

```bash
docker build -t mcp-docx-soffice:latest docker/
```

## Tests

```bash
pip install -r requirements-dev.txt
pytest
```

## License

Private — internal Intelion Cloud tooling.
