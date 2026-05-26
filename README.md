# docx-comments-to-text

Extract reviewer comments from `.docx` **and `.xlsx`** files and insert them inline with the text/cell they reference. Output is **Markdown** (with proper headings, formatting, and tables) and each comment is wrapped in an unambiguous `<comment>` XML tag so downstream tools (especially LLMs) can recognise it cleanly. File type is detected automatically from the extension.

> This is a fork of [platelminto/docx-comments-to-text](https://github.com/platelminto/docx-comments-to-text). See [Differences from upstream](#differences-from-upstream) for what's new.

Hosted version of the original (plain-text) tool is available at [https://docx-comment.app/](https://docx-comment.app/).

## Installation

### Run directly from this fork with `uvx` (no install)
```bash
# One-shot run — uvx fetches the repo into an isolated env, runs the CLI, then cleans up
uvx --from git+https://github.com/do-me/docx-comments-to-text docx-comments-to-text path/to/file.docx
uvx --from git+https://github.com/do-me/docx-comments-to-text docx-comments-to-text path/to/file.xlsx --sheet "Sheet1" -o out.md
```

### Install persistently with `uv tool`
```bash
uv tool install --from git+https://github.com/do-me/docx-comments-to-text docx-comments-to-text

# Then run from anywhere:
docx-comments-to-text path/to/file.docx
```

### From PyPI (upstream plain-text version)
> The PyPI release is the **upstream** project, which emits plain text and lacks the xlsx / markdown features in this fork.
```bash
pip install docx-comments-to-text
```

### From source
```bash
# Clone this fork
git clone https://github.com/do-me/docx-comments-to-text
cd docx-comments-to-text

# Install in development mode
uv sync --dev
# or: pip install -e .
```

## Usage

### Command Line Interface

```bash
# Basic usage - output to stdout (Word or Excel)
docx-comments-to-text document.docx
docx-comments-to-text workbook.xlsx

# Save to file (Markdown)
docx-comments-to-text document.docx -o output.md

# Control author display
docx-comments-to-text document.docx --authors always   # Always show authors (default)
docx-comments-to-text document.docx --authors auto     # Show authors only when multiple exist
docx-comments-to-text document.docx --authors never    # Hide authors

# Control comment placement (--placement applies to .docx only)
docx-comments-to-text document.docx --placement inline         # Inline with text (default)
docx-comments-to-text document.docx --placement end-paragraph  # At end of each paragraph
docx-comments-to-text document.docx --placement comments-only  # Comments only with context

# XLSX: pick a single worksheet (default: render every sheet)
docx-comments-to-text workbook.xlsx --list-sheets               # List worksheet names
docx-comments-to-text workbook.xlsx --sheet "EO Products"       # Render just one sheet
```

### Development Usage

If working from source:
```bash
# Run with uv
uv run docx-comments-to-text document.docx

# Or use module syntax
uv run python -m docx_comments_to_text.cli document.docx
```

### Example Output

#### Inline placement (default)
```markdown
# Section heading

Original text with [reviewer feedback] <comment author="Jane">This needs clarification</comment> continues here.
More content [needs examples] <comment author="John">Consider adding examples</comment> and final text.

| Col A | Col B |
| --- | --- |
| **bold cell** | *italic cell* |
```

#### End-paragraph placement
```markdown
Original text with reviewer feedback[1] continues here.
More content needs examples[2] and final text.

Comments:
1. <comment author="Jane">This needs clarification</comment>
2. <comment author="John">Consider adding examples</comment>
```

#### Comments-only placement
```markdown
"reviewer feedback": <comment author="Jane">This needs clarification</comment>
"needs examples": <comment author="John">Consider adding examples</comment>
```

#### XLSX output
Each sheet becomes a Markdown table; comments are placed **inside the cell that owns them** with a `cell="..."` attribute so the LLM sees the comment next to its content:
```markdown
## Scores

| Item | Value | Note |
| --- | --- | --- |
| Apples<comment author="Reviewer1" cell="A2">Check spelling</comment> | 10 | fresh<comment author="Jane" cell="C2">Confirm with supplier</comment> |
| Bananas | 5<comment author="John" cell="B3">Why so few?<br>Needs restock</comment> |   |
```

## Features

- **Markdown output**: headings, bold/italic, bullet/numbered lists, and tables
- **XLSX support**: each worksheet rendered as a Markdown table with comments embedded inline in their cells. Both legacy comments (`xl/comments*.xml`) and modern Office 365 threaded comments (`xl/threadedComments/*.xml` + `xl/persons/*.xml`) are recognised; threaded comments take priority when both are present.
- **XML-tagged comments**: `<comment author="..." cell="B7">…</comment>` so LLMs can pick them out (and locate them) unambiguously
- **Sheet filtering**: `--list-sheets` to discover names, `--sheet "Name"` to convert a single worksheet
- Accurate comment positioning and text preservation
- Handles overlapping comments and multiple comment types
- Configurable author display (authors shown by default)
- Multiple comment placement styles for `.docx` (inline, end-of-paragraph, comments-only)
- **Zero runtime third-party deps for parsing** — `zipfile` + `xml.etree.ElementTree` from the stdlib do all the heavy lifting; only `click` is needed for the CLI

## Differences from upstream

This fork (`do-me/docx-comments-to-text`) extends [platelminto/docx-comments-to-text](https://github.com/platelminto/docx-comments-to-text) along several axes:

| Area | Upstream | This fork |
| --- | --- | --- |
| Output format | Plain text | **Markdown** — headings, bold/italic, bullet/numbered lists, and proper pipe tables |
| Comment syntax | `[COMMENT John: "feedback"]` | **XML tag**: `<comment author="John">feedback</comment>` (LLM-friendly; XML-escaped body) |
| Author display default | `auto` (hidden when single author) | **`always`** — authors shown unless `--authors never` |
| DOCX tables | Each cell on its own line | **Markdown pipe tables** with header separator |
| DOCX headings | Plain text | `# … ######` based on `Heading1`–`Heading6` (plus `Title` / `Subtitle`) |
| DOCX lists | Plain text | `- ` bullets and `1. ` numbered items, resolved from `numbering.xml`, with indent per level |
| DOCX formatting | Lost | Bold (`**…**`) and italic (`*…*`); adjacent same-style runs are coalesced so output renders cleanly |
| XLSX support | None | **Full support**: each sheet → Markdown table with comments embedded inside their owning cells (e.g. `Apples<comment author="…" cell="A2">…</comment>`) |
| XLSX comments | n/a | Both legacy (`xl/comments*.xml`) and modern threaded (`xl/threadedComments/*.xml` + `xl/persons/*.xml`) — threaded preferred, legacy fallback |
| XLSX sheet filtering | n/a | `--list-sheets` to discover names, `--sheet "Name"` to render one worksheet |
| Runtime dependencies | `python-docx`, `lxml`, `click` | **`click` only** — DOCX/XLSX parsing is pure stdlib (`zipfile` + `xml.etree.ElementTree`) |

The behavioural changes are covered by the existing test suite (73 tests).

## Technical Details

### DOCX Structure
- DOCX files are ZIP archives containing XML files
- `word/document.xml` - main document content
- `word/comments.xml` - comment definitions
- Comment ranges marked with `<w:commentRangeStart>` and `<w:commentRangeEnd>`

### XLSX Structure
- `xl/workbook.xml` - sheet list (resolved via `xl/_rels/workbook.xml.rels`)
- `xl/worksheets/sheetN.xml` - cell values and refs
- `xl/sharedStrings.xml` - shared string pool
- `xl/comments/commentN.xml` - legacy comment threads (author list + per-cell `ref`)
- `xl/threadedComments/*.xml` + `xl/persons/*.xml` - modern threaded comments

### Comment Insertion Strategy
**DOCX:**
1. Walk the document XML, rendering paragraphs / tables / lists as Markdown
2. Track character positions in the Markdown stream
3. Map comment ranges to their start/end positions in that stream
4. Wrap commented text in brackets: `[commented text]`
5. Insert the comment as an XML tag after the bracketed text: `<comment author="…">feedback</comment>`

**XLSX:**
1. Walk each sheet's cells (resolving shared strings, inline strings, numbers, booleans)
2. Resolve each sheet's comments file via its relationships — threaded comments are preferred when present (they carry richer author info via `xl/persons/*.xml`), with legacy comments as a fallback
3. Append the comment XML tag directly to the owning cell's text, including a `cell="..."` attribute for traceability

## Dependencies

- `click` — command line interface

The DOCX and XLSX parsers are implemented with only the Python standard library (`zipfile`, `xml.etree.ElementTree`, `xml.sax.saxutils`), so installation is fast and there are no native build steps.
