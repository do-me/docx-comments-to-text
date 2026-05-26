import os
from docx_comments_to_text.xlsx_parser import XlsxParser
from docx_comments_to_text.docx_processor import process_docx

FIXTURES_DIR = os.path.join("tests", "docs")
SAMPLE = os.path.join(FIXTURES_DIR, "sample_with_comments.xlsx")


class TestXlsxParser:
    def test_renders_sheet_name_as_heading(self):
        out = XlsxParser(SAMPLE).render_markdown()
        assert "## Scores" in out
        assert "## Summary" in out

    def test_renders_table_with_header_separator(self):
        out = XlsxParser(SAMPLE).render_markdown()
        # First row becomes header, second line is the markdown separator
        lines = out.splitlines()
        scores_idx = lines.index("## Scores")
        # next non-empty is the header row, then the separator
        header = lines[scores_idx + 2]
        separator = lines[scores_idx + 3]
        assert "| Item | Value | Note |" == header
        assert separator == "| --- | --- | --- |"

    def test_comment_appears_in_correct_cell(self):
        out = XlsxParser(SAMPLE).render_markdown()
        # Cell A2 has "Apples" + a comment by Reviewer1
        assert 'Apples<comment author="Reviewer1" cell="A2">Check spelling</comment>' in out

    def test_comment_with_newlines_uses_br(self):
        out = XlsxParser(SAMPLE).render_markdown()
        # Cell B3 comment has a newline — escaped to <br>
        assert "Why so few?<br>Needs restock" in out

    def test_authors_never_drops_author_attribute(self):
        out = XlsxParser(SAMPLE).render_markdown(show_authors="never")
        assert "author=" not in out
        # Cell ref is still emitted so the LLM can locate the cell
        assert '<comment cell="A2">' in out

    def test_authors_auto_with_multiple_authors_keeps_them(self):
        # Fixture has 3 distinct authors → auto should include them
        out = XlsxParser(SAMPLE).render_markdown(show_authors="auto")
        assert 'author="Reviewer1"' in out
        assert 'author="John"' in out
        assert 'author="Jane"' in out

    def test_dispatcher_recognizes_xlsx(self):
        # process_docx should dispatch by extension and return the markdown
        out = process_docx(SAMPLE)
        assert "## Scores" in out
        assert "<comment" in out

    def test_dispatcher_rejects_unknown_extension(self, tmp_path):
        bogus = tmp_path / "thing.pdf"
        bogus.write_text("hi")
        try:
            process_docx(bogus)
            assert False, "Expected ValueError"
        except ValueError as e:
            assert "Unsupported file type" in str(e)


class TestXlsxParserHelpers:
    def test_parse_ref(self):
        from docx_comments_to_text.xlsx_parser import _parse_ref
        assert _parse_ref("A1") == (0, 0)
        assert _parse_ref("B7") == (1, 6)
        assert _parse_ref("AA1") == (26, 0)
        assert _parse_ref("") == (-1, -1)

    def test_escape_cell(self):
        from docx_comments_to_text.xlsx_parser import _escape_cell
        assert _escape_cell("a|b") == "a\\|b"
        assert _escape_cell("line1\nline2") == "line1<br>line2"
        assert _escape_cell("") == " "

    def test_normalize_target(self):
        from docx_comments_to_text.xlsx_parser import _normalize_target
        assert _normalize_target("xl/workbook.xml", "worksheets/sheet1.xml") == "xl/worksheets/sheet1.xml"
        assert _normalize_target("xl/workbook.xml", "/xl/comments/comment1.xml") == "xl/comments/comment1.xml"
        assert _normalize_target("xl/worksheets/sheet1.xml", "../comments/comment1.xml") == "xl/comments/comment1.xml"
