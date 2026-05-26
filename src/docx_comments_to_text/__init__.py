"""Extract reviewer comments from .docx / .xlsx files and emit Markdown with XML-tagged comments."""

from .docx_parser import DocxParser, Comment, CommentRange
from .docx_processor import process_docx
from .text_formatter import format_text_with_comments
from .xlsx_parser import XlsxParser, XlsxComment

__all__ = [
    'DocxParser', 'Comment', 'CommentRange',
    'XlsxParser', 'XlsxComment',
    'process_docx', 'format_text_with_comments',
]
