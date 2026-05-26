# docx-comments-to-text

Extract reviewer comments from `.docx` **and `.xlsx`** files and insert them inline with the text/cell they reference. Output is **Markdown** (with proper headings, formatting, and tables) and each comment is wrapped in an unambiguous `<comment>` XML tag so downstream tools (especially LLMs) can recognise it cleanly. File type is detected automatically from the extension.

Hosted version available at [https://docx-comment.app/](https://docx-comment.app/).

## Installation

### From PyPI (recommended)
```bash
pip install docx-comments-to-text
```

### From source
```bash
# Clone the repository
git clone https://github.com/platelminto/docx-comments-to-text
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
- **XLSX support**: each worksheet rendered as a Markdown table with comments embedded inline in their cells (legacy + threaded comments)
- **XML-tagged comments**: `<comment author="...">…</comment>` so LLMs can pick them out unambiguously
- Accurate comment positioning and text preservation
- Handles overlapping comments and multiple comment types
- Configurable author display (authors shown by default)
- Multiple comment placement styles for `.docx` (inline, end-of-paragraph, comments-only)

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
2. Resolve each sheet's comments file via its relationships
3. Append the comment XML tag directly to the owning cell's text, including a `cell="..."` attribute for traceability

## Dependencies

- `python-docx` - DOCX file handling
- `lxml` - XML parsing
- `click` - Command line interface
