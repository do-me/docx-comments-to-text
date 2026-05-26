# docx-comments-to-text

Extract reviewer comments from `.docx` files and insert them inline with the text they reference. Output is **Markdown** (with proper headings, formatting, and tables) and each comment is wrapped in an unambiguous `<comment>` XML tag so downstream tools (especially LLMs) can recognise it cleanly.

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
# Basic usage - output to stdout
docx-comments-to-text document.docx

# Save to file (Markdown)
docx-comments-to-text document.docx -o output.md

# Control author display
docx-comments-to-text document.docx --authors always   # Always show authors (default)
docx-comments-to-text document.docx --authors auto     # Show authors only when multiple exist
docx-comments-to-text document.docx --authors never    # Hide authors

# Control comment placement
docx-comments-to-text document.docx --placement inline         # Inline with text (default)
docx-comments-to-text document.docx --placement end-paragraph  # At end of each paragraph
docx-comments-to-text document.docx --placement comments-only  # Comments only with context
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

## Features

- **Markdown output**: headings, bold/italic, bullet/numbered lists, and tables
- **XML-tagged comments**: `<comment author="...">…</comment>` so LLMs can pick them out unambiguously
- Accurate comment positioning and text preservation
- Handles overlapping comments and multiple comment types
- Configurable author display (authors shown by default)
- Multiple comment placement styles (inline, end-of-paragraph, comments-only)

## Technical Details

### DOCX Structure
- DOCX files are ZIP archives containing XML files
- `word/document.xml` - main document content
- `word/comments.xml` - comment definitions
- Comment ranges marked with `<w:commentRangeStart>` and `<w:commentRangeEnd>`

### Comment Insertion Strategy
1. Walk the document XML, rendering paragraphs / tables / lists as Markdown
2. Track character positions in the Markdown stream
3. Map comment ranges to their start/end positions in that stream
4. Wrap commented text in brackets: `[commented text]`
5. Insert the comment as an XML tag after the bracketed text: `<comment author="…">feedback</comment>`

## Dependencies

- `python-docx` - DOCX file handling
- `lxml` - XML parsing
- `click` - Command line interface
