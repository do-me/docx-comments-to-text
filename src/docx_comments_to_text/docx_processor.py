from pathlib import Path
from .docx_parser import DocxParser
from .text_formatter import format_text_with_comments
from .xlsx_parser import XlsxParser


def process_docx(input_file: str | Path, authors: str = 'always', placement: str = 'inline') -> str:
    """
    Extract comments from a DOCX or XLSX file and return formatted text with comments.

    The output is Markdown. Comments are wrapped in `<comment author="...">...</comment>`
    XML tags so downstream consumers (especially LLMs) can recognise them unambiguously.

    Args:
        input_file: Path to the .docx or .xlsx file
        authors: How to display authors ('never', 'always', 'auto')
        placement: Comment placement style — only meaningful for .docx
                   ('inline', 'end-paragraph', 'comments-only')

    Returns:
        Markdown text with comments inserted.

    Raises:
        FileNotFoundError: If input file doesn't exist
        ValueError: If the file extension isn't supported
    """
    path = Path(input_file)
    suffix = path.suffix.lower()

    if suffix == '.docx':
        parser = DocxParser(str(path))
        text, comments, ranges = parser.extract_text_and_comments()
        return format_text_with_comments(text, comments, ranges, show_authors=authors, placement=placement)

    if suffix == '.xlsx':
        return XlsxParser(str(path)).render_markdown(show_authors=authors)

    raise ValueError(
        f"Unsupported file type: {suffix or '(no extension)'}. "
        "Supported types: .docx, .xlsx"
    )


if __name__ == "__main__":
    # Example usage
    input_docx = "tests/docs/simple_comment.docx"
    output = process_docx(input_docx, authors="always", placement="inline")
    print(output)
