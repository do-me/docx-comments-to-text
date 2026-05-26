import os
from docx_comments_to_text.docx_parser import DocxParser, Comment, CommentRange
from docx_comments_to_text.text_formatter import format_text_with_comments, _format_comment

FIXTURES_DIR = os.path.join("tests", "docs")


class TestTextFormatter:
    def test_simple_range_comment(self):
        """Test formatting a single range comment"""
        text = "Hello world, this is a test."
        comments = [Comment(id="1", author="Reviewer", text="This word needs clarification")]
        ranges = [CommentRange(comment_id="1", start_pos=6, end_pos=11)]

        result = format_text_with_comments(text, comments, ranges)
        # Authors are shown by default ("always")
        expected = 'Hello [world] <comment author="Reviewer">This word needs clarification</comment>, this is a test.'
        assert result == expected

    def test_point_comment(self):
        """Test formatting a point comment (no text range)"""
        text = "Insert example here. More text follows."
        comments = [Comment(id="1", author="Reviewer", text="Add specific example")]
        ranges = [CommentRange(comment_id="1", start_pos=19, end_pos=19)]

        result = format_text_with_comments(text, comments, ranges)
        expected = 'Insert example here<comment author="Reviewer">Add specific example</comment>. More text follows.'
        assert result == expected

    def test_multiple_comments_different_ranges(self):
        """Test multiple comments on different text ranges"""
        text = "The quick brown fox jumps over the lazy dog."
        comments = [
            Comment(id="1", author="Reviewer", text="Too informal"),
            Comment(id="2", author="Reviewer", text="Be more specific"),
            Comment(id="3", author="Reviewer", text="Negative connotation"),
        ]
        ranges = [
            CommentRange(comment_id="1", start_pos=4, end_pos=9),
            CommentRange(comment_id="2", start_pos=16, end_pos=19),
            CommentRange(comment_id="3", start_pos=35, end_pos=39),
        ]

        result = format_text_with_comments(text, comments, ranges)
        expected = (
            'The [quick] <comment author="Reviewer">Too informal</comment> '
            'brown [fox] <comment author="Reviewer">Be more specific</comment> '
            'jumps over the [lazy] <comment author="Reviewer">Negative connotation</comment> dog.'
        )
        assert result == expected

    def test_overlapping_comments_same_range(self):
        """Test multiple comments on the same text range"""
        text = "This phrase needs work badly."
        comments = [
            Comment(id="1", author="Reviewer1", text="Unclear"),
            Comment(id="2", author="Reviewer2", text="This phrase requires improvement"),
        ]
        ranges = [
            CommentRange(comment_id="1", start_pos=12, end_pos=22),
            CommentRange(comment_id="2", start_pos=12, end_pos=22),
        ]

        result = format_text_with_comments(text, comments, ranges)
        assert "[needs work]" in result
        assert '<comment author="Reviewer1">Unclear</comment>' in result
        assert '<comment author="Reviewer2">This phrase requires improvement</comment>' in result

    def test_mixed_point_and_range_comments(self):
        """Test combination of point and range comments"""
        text = "Start here. End there."
        comments = [
            Comment(id="1", author="Reviewer", text="Range comment"),
            Comment(id="2", author="Reviewer", text="Point comment"),
        ]
        ranges = [
            CommentRange(comment_id="1", start_pos=0, end_pos=5),
            CommentRange(comment_id="2", start_pos=11, end_pos=11),
        ]

        result = format_text_with_comments(text, comments, ranges)
        expected = (
            '[Start] <comment author="Reviewer">Range comment</comment> here.'
            '<comment author="Reviewer">Point comment</comment> End there.'
        )
        assert result == expected

    def test_no_comments(self):
        """Test formatting text with no comments"""
        text = "Plain text with no comments."
        result = format_text_with_comments(text, [], [])
        assert result == text

    def test_comment_not_found(self):
        """Test handling of ranges with missing comments"""
        text = "Hello world"
        comments = []
        ranges = [CommentRange(comment_id="missing", start_pos=6, end_pos=11)]

        result = format_text_with_comments(text, comments, ranges)
        assert result == text

    def test_xml_escaping_in_comment_text(self):
        """Comment text with XML-special chars must be escaped so the tag stays parseable."""
        text = "Hello world"
        comments = [Comment(id="1", author="A&B", text="check <tag> & \"quotes\"")]
        ranges = [CommentRange(comment_id="1", start_pos=6, end_pos=11)]

        result = format_text_with_comments(text, comments, ranges)
        assert 'author="A&amp;B"' in result
        assert "&lt;tag&gt;" in result
        assert "&amp;" in result

    def test_integration_with_simple_comment_docx(self):
        docx_path = os.path.join(FIXTURES_DIR, "simple_comment.docx")
        parser = DocxParser(docx_path)
        text, comments, ranges = parser.extract_text_and_comments()

        result = format_text_with_comments(text, comments, ranges)

        assert "[world]" in result
        assert "<comment" in result
        assert "</comment>" in result
        assert "This word needs clarification" in result

    def test_integration_with_multiple_comments_docx(self):
        docx_path = os.path.join(FIXTURES_DIR, "multiple_comments.docx")
        parser = DocxParser(docx_path)
        text, comments, ranges = parser.extract_text_and_comments()

        result = format_text_with_comments(text, comments, ranges)
        # Three commented text regions
        assert result.count("</comment>") == 3

    def test_integration_with_point_comment_docx(self):
        docx_path = os.path.join(FIXTURES_DIR, "point_comment.docx")
        parser = DocxParser(docx_path)
        text, comments, ranges = parser.extract_text_and_comments()

        result = format_text_with_comments(text, comments, ranges)

        assert "<comment" in result
        assert "Add specific example" in result
        # Point comment should produce no surrounding bracket
        bracket_count = result.count('[')
        assert bracket_count == 0

    def test_integration_with_nested_comments_docx(self):
        docx_path = os.path.join(FIXTURES_DIR, "nested_comments.docx")
        parser = DocxParser(docx_path)
        text, comments, ranges = parser.extract_text_and_comments()

        result = format_text_with_comments(text, comments, ranges)

        # Outer range "really important section" and inner range "important" both bracketed
        assert "[really [important]" in result or "[really important]" in result
        assert "Define importance" in result
        assert "Key part of document" in result
        assert result.count("</comment>") == 2

    def test_integration_with_nested_comments_end_paragraph(self):
        docx_path = os.path.join(FIXTURES_DIR, "nested_comments.docx")
        parser = DocxParser(docx_path)
        text, comments, ranges = parser.extract_text_and_comments()

        result = format_text_with_comments(text, comments, ranges, placement="end-paragraph")

        assert "sect[1]ion" not in result
        assert "The really important[1] section[2] needs attention." in result
        assert "Comments:" in result
        assert "1." in result and "2." in result
        # Comments are now XML-tagged in the footer
        assert "<comment" in result
        assert "Define importance" in result
        assert "Key part of document" in result


class TestAuthorDisplay:
    def test_show_authors_never(self):
        text = "Hello world"
        comments = [Comment(id="1", author="John", text="needs work")]
        ranges = [CommentRange(comment_id="1", start_pos=6, end_pos=11)]

        result = format_text_with_comments(text, comments, ranges, show_authors="never")
        expected = 'Hello [world] <comment>needs work</comment>'
        assert result == expected

    def test_show_authors_always_single_author(self):
        text = "Hello world"
        comments = [Comment(id="1", author="John", text="needs work")]
        ranges = [CommentRange(comment_id="1", start_pos=6, end_pos=11)]

        result = format_text_with_comments(text, comments, ranges, show_authors="always")
        expected = 'Hello [world] <comment author="John">needs work</comment>'
        assert result == expected

    def test_show_authors_always_multiple_authors(self):
        text = "Hello world"
        comments = [
            Comment(id="1", author="John", text="needs work"),
            Comment(id="2", author="Jane", text="unclear"),
        ]
        ranges = [
            CommentRange(comment_id="1", start_pos=6, end_pos=11),
            CommentRange(comment_id="2", start_pos=6, end_pos=11),
        ]

        result = format_text_with_comments(text, comments, ranges, show_authors="always")
        assert '<comment author="John">needs work</comment>' in result
        assert '<comment author="Jane">unclear</comment>' in result

    def test_show_authors_auto_single_author(self):
        """Auto mode hides the author when only one author has commented."""
        text = "Hello world"
        comments = [Comment(id="1", author="John", text="needs work")]
        ranges = [CommentRange(comment_id="1", start_pos=6, end_pos=11)]

        result = format_text_with_comments(text, comments, ranges, show_authors="auto")
        assert '<comment>needs work</comment>' in result
        assert "John" not in result

    def test_show_authors_auto_multiple_authors(self):
        text = "Hello world"
        comments = [
            Comment(id="1", author="John", text="needs work"),
            Comment(id="2", author="Jane", text="unclear"),
        ]
        ranges = [
            CommentRange(comment_id="1", start_pos=6, end_pos=11),
            CommentRange(comment_id="2", start_pos=6, end_pos=11),
        ]

        result = format_text_with_comments(text, comments, ranges, show_authors="auto")
        assert '<comment author="John">needs work</comment>' in result
        assert '<comment author="Jane">unclear</comment>' in result

    def test_show_authors_auto_same_author_multiple_comments(self):
        text = "Hello world test"
        comments = [
            Comment(id="1", author="John", text="needs work"),
            Comment(id="2", author="John", text="also unclear"),
        ]
        ranges = [
            CommentRange(comment_id="1", start_pos=6, end_pos=11),
            CommentRange(comment_id="2", start_pos=12, end_pos=16),
        ]

        result = format_text_with_comments(text, comments, ranges, show_authors="auto")
        assert "John" not in result
        assert "needs work" in result
        assert "also unclear" in result

    def test_default_show_authors_parameter(self):
        """Default is now 'always' — author appears even with a single commenter."""
        text = "Hello world"
        comments = [Comment(id="1", author="John", text="needs work")]
        ranges = [CommentRange(comment_id="1", start_pos=6, end_pos=11)]

        result = format_text_with_comments(text, comments, ranges)
        expected = 'Hello [world] <comment author="John">needs work</comment>'
        assert result == expected

    def test_show_authors_point_comments(self):
        text = "Insert here."
        comments = [Comment(id="1", author="Jane", text="Add example")]
        ranges = [CommentRange(comment_id="1", start_pos=7, end_pos=7)]

        result_never = format_text_with_comments(text, comments, ranges, show_authors="never")
        result_always = format_text_with_comments(text, comments, ranges, show_authors="always")

        assert '<comment>Add example</comment>' in result_never
        assert "Jane" not in result_never
        assert '<comment author="Jane">Add example</comment>' in result_always


class TestPlacementOptions:
    def test_end_paragraph_placement_single_paragraph(self):
        text = "This is a sentence. This is another sentence."
        comments = [
            Comment(id="1", author="Reviewer", text="First comment"),
            Comment(id="2", author="Reviewer", text="Second comment"),
        ]
        ranges = [
            CommentRange(comment_id="1", start_pos=5, end_pos=7),
            CommentRange(comment_id="2", start_pos=28, end_pos=35),
        ]

        result = format_text_with_comments(text, comments, ranges, placement="end-paragraph")

        assert "This is[1] a sentence. This is another[2] sentence." in result
        assert "Comments:" in result
        assert '1. <comment author="Reviewer">First comment</comment>' in result
        assert '2. <comment author="Reviewer">Second comment</comment>' in result

    def test_end_paragraph_placement_multiple_paragraphs(self):
        text = "First paragraph with comment.\n\nSecond paragraph also has feedback."
        comments = [
            Comment(id="1", author="Reviewer", text="Paragraph 1 feedback"),
            Comment(id="2", author="Reviewer", text="Paragraph 2 feedback"),
        ]
        ranges = [
            CommentRange(comment_id="1", start_pos=21, end_pos=28),
            CommentRange(comment_id="2", start_pos=57, end_pos=65),
        ]

        result = format_text_with_comments(text, comments, ranges, placement="end-paragraph")

        assert "First paragraph with comment[1]." in result
        assert "Second paragraph also has feedback[2]." in result
        assert "Comments:" in result
        assert '1. <comment author="Reviewer">Paragraph 1 feedback</comment>' in result
        assert '2. <comment author="Reviewer">Paragraph 2 feedback</comment>' in result

    def test_end_paragraph_placement_point_comments(self):
        text = "Insert example here. More text follows."
        comments = [Comment(id="1", author="Reviewer", text="Add specific example")]
        ranges = [CommentRange(comment_id="1", start_pos=19, end_pos=19)]

        result = format_text_with_comments(text, comments, ranges, placement="end-paragraph")

        assert "Insert example here[1]. More text follows." in result
        assert "Comments:" in result
        assert '1. <comment author="Reviewer">Add specific example</comment>' in result

    def test_end_paragraph_placement_with_authors(self):
        text = "Text with multiple reviewers."
        comments = [
            Comment(id="1", author="John", text="First feedback"),
            Comment(id="2", author="Jane", text="Second feedback"),
        ]
        ranges = [
            CommentRange(comment_id="1", start_pos=5, end_pos=9),
            CommentRange(comment_id="2", start_pos=19, end_pos=28),
        ]

        result = format_text_with_comments(text, comments, ranges, placement="end-paragraph", show_authors="always")

        assert "Text with[1] multiple reviewers[2]." in result
        assert "Comments:" in result
        assert '1. <comment author="John">First feedback</comment>' in result
        assert '2. <comment author="Jane">Second feedback</comment>' in result

    def test_comments_only_placement(self):
        text = "This is original text with comments."
        comments = [
            Comment(id="1", author="Reviewer1", text="First feedback"),
            Comment(id="2", author="Reviewer2", text="Second feedback"),
            Comment(id="3", author="Reviewer1", text="Third feedback"),
        ]
        ranges = [
            CommentRange(comment_id="1", start_pos=5, end_pos=7),
            CommentRange(comment_id="2", start_pos=17, end_pos=21),
            CommentRange(comment_id="3", start_pos=27, end_pos=35),
        ]

        result = format_text_with_comments(text, comments, ranges, placement="comments-only")

        assert "This is original text with comments." not in result
        assert "First feedback" in result
        assert "Second feedback" in result
        assert "Third feedback" in result

    def test_comments_only_with_context(self):
        text = "Context for understanding feedback."
        comments = [
            Comment(id="1", author="Reviewer", text="Needs more detail"),
            Comment(id="2", author="Reviewer", text="Unclear phrasing"),
        ]
        ranges = [
            CommentRange(comment_id="1", start_pos=0, end_pos=7),
            CommentRange(comment_id="2", start_pos=26, end_pos=34),
        ]

        result = format_text_with_comments(text, comments, ranges, placement="comments-only")

        assert '"Context": <comment author="Reviewer">Needs more detail</comment>' in result
        assert '"feedback": <comment author="Reviewer">Unclear phrasing</comment>' in result

    def test_comments_only_point_comments(self):
        text = "Text with insertion point."
        comments = [Comment(id="1", author="Reviewer", text="Add example here")]
        ranges = [CommentRange(comment_id="1", start_pos=15, end_pos=15)]

        result = format_text_with_comments(text, comments, ranges, placement="comments-only")

        assert '[Position 15]: <comment author="Reviewer">Add example here</comment>' in result

    def test_comments_only_with_authors_always(self):
        text = "Text with single author comment."
        comments = [Comment(id="1", author="John", text="Single author feedback")]
        ranges = [CommentRange(comment_id="1", start_pos=5, end_pos=9)]

        result = format_text_with_comments(text, comments, ranges, placement="comments-only", show_authors="always")

        assert '"with": <comment author="John">Single author feedback</comment>' in result

    def test_comments_only_with_authors_never(self):
        text = "Text with multiple author comments."
        comments = [
            Comment(id="1", author="John", text="First feedback"),
            Comment(id="2", author="Jane", text="Second feedback"),
        ]
        ranges = [
            CommentRange(comment_id="1", start_pos=5, end_pos=9),
            CommentRange(comment_id="2", start_pos=19, end_pos=25),
        ]

        result = format_text_with_comments(text, comments, ranges, placement="comments-only", show_authors="never")

        assert '"with": <comment>First feedback</comment>' in result
        assert '"author": <comment>Second feedback</comment>' in result
        assert "John" not in result
        assert "Jane" not in result

    def test_inline_placement_unchanged(self):
        """Inline placement uses the default 'always' authors with XML tags."""
        text = "Hello world"
        comments = [Comment(id="1", author="Reviewer", text="needs work")]
        ranges = [CommentRange(comment_id="1", start_pos=6, end_pos=11)]

        result_explicit = format_text_with_comments(text, comments, ranges, placement="inline")
        result_default = format_text_with_comments(text, comments, ranges)

        assert result_explicit == result_default
        assert result_explicit == 'Hello [world] <comment author="Reviewer">needs work</comment>'

    def test_end_paragraph_point_comment_at_paragraph_end(self):
        text = "First paragraph.\n\nSecond paragraph ends here."
        comments = [
            Comment(id="1", author="Reviewer", text="End of first para"),
            Comment(id="2", author="Reviewer", text="End of second para"),
        ]
        ranges = [
            CommentRange(comment_id="1", start_pos=16, end_pos=16),
            CommentRange(comment_id="2", start_pos=45, end_pos=45),
        ]

        result = format_text_with_comments(text, comments, ranges, placement="end-paragraph")

        assert "First paragraph.[1]" in result
        assert "Second paragraph ends here.[2]" in result
        assert "Comments:" in result
        assert '1. <comment author="Reviewer">End of first para</comment>' in result
        assert '2. <comment author="Reviewer">End of second para</comment>' in result

    def test_end_paragraph_consistent_formatting_point_and_range(self):
        text = "This is text. More text here."
        comments = [
            Comment(id="1", author="Reviewer", text="Range comment text"),
            Comment(id="2", author="Reviewer", text="Point comment text"),
        ]
        ranges = [
            CommentRange(comment_id="1", start_pos=5, end_pos=7),
            CommentRange(comment_id="2", start_pos=15, end_pos=15),
        ]

        result = format_text_with_comments(text, comments, ranges, placement="end-paragraph")

        assert '1. <comment author="Reviewer">Range comment text</comment>' in result
        assert '2. <comment author="Reviewer">Point comment text</comment>' in result

        # No [Position]: prefix for point comments anymore
        assert "[Position" not in result
