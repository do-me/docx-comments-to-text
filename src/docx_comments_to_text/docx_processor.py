from pathlib import Path
from .docx_parser import DocxParser
from .text_formatter import format_text_with_comments


def process_docx(input_file: str | Path, authors: str = 'always', placement: str = 'inline') -> str:
    """
    Extract comments from a DOCX file and return formatted text with comments.

    Args:
        input_file: Path to the DOCX file
        authors: How to display authors ('never', 'always', 'auto')
        placement: Comment placement style ('inline', 'end-paragraph', 'comments-only')
        
    Returns:
        Formatted text with comments according to placement style
        
    Raises:
        FileNotFoundError: If input file doesn't exist
        Exception: If processing fails
    """
    # Parse the DOCX file
    parser = DocxParser(str(input_file))
    text, comments, ranges = parser.extract_text_and_comments()
    
    # Format text with comments
    formatted_text = format_text_with_comments(text, comments, ranges, show_authors=authors, placement=placement)
    
    return formatted_text


if __name__ == "__main__":
    # Example usage
    input_docx = "tests/docs/simple_comment.docx"
    output = process_docx(input_docx, authors="always", placement="inline")
    print(output)