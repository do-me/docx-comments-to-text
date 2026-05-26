import os
from docx_comments_to_text.xlsx_parser import XlsxParser
from docx_comments_to_text.docx_processor import process_docx

FIXTURES_DIR = os.path.join("tests", "docs")
SAMPLE = os.path.join(FIXTURES_DIR, "sample_with_comments.xlsx")
THREADED_SAMPLE = os.path.join(FIXTURES_DIR, "sample_threaded_comments.xlsx")


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

    def test_list_sheets(self):
        names = XlsxParser(SAMPLE).list_sheets()
        assert names == ["Scores", "Summary"]

    def test_filter_to_single_sheet(self):
        out = XlsxParser(SAMPLE).render_markdown(sheet_name="Summary")
        assert "## Summary" in out
        assert "## Scores" not in out

    def test_filter_to_missing_sheet_raises(self):
        try:
            XlsxParser(SAMPLE).render_markdown(sheet_name="Nope")
            assert False, "Expected ValueError"
        except ValueError as e:
            assert "Sheet 'Nope' not found" in str(e)
            assert "Scores" in str(e)
            assert "Summary" in str(e)

    def test_process_docx_threads_sheet_filter(self):
        out = process_docx(SAMPLE, sheet="Summary")
        assert "## Summary" in out
        assert "## Scores" not in out


class TestXlsxThreadedComments:
    """Threaded comments must be parsed (correct 2018 namespace) and take
    priority over legacy comments when both are present."""

    def test_threaded_comments_recognized(self):
        out = XlsxParser(THREADED_SAMPLE).render_markdown()
        # Threaded author names beat the legacy author names
        assert 'author="Alice (threaded)"' in out
        assert 'author="Bob (threaded)"' in out
        assert "Threaded check spelling" in out

    def test_threaded_takes_priority_over_legacy(self):
        # Sheet "Scores" has BOTH threaded and legacy comments; legacy should be hidden
        out = XlsxParser(THREADED_SAMPLE).render_markdown(sheet_name="Scores")
        # Legacy 'Check spelling' / 'Why so few?' / 'Confirm with supplier' must not show
        assert "Check spelling" not in out
        assert "Confirm with supplier" not in out
        assert "Why so few?" not in out

    def test_fallback_to_legacy_when_no_threaded(self):
        # Sheet "Summary" only has legacy comments — they should still appear
        out = XlsxParser(THREADED_SAMPLE).render_markdown(sheet_name="Summary")
        assert "Sum of column B" in out
        assert 'author="Reviewer1"' in out


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
