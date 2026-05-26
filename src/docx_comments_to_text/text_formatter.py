from typing import List
from xml.sax.saxutils import escape, quoteattr
from .docx_parser import Comment, CommentRange


def format_text_with_comments(text: str, comments: List[Comment], ranges: List[CommentRange], show_authors: str = "always", placement: str = "inline") -> str:
    """
    Format text by inserting comments with configurable placement options.

    Args:
        text: Original document text
        comments: List of Comment objects
        ranges: List of CommentRange objects mapping comments to text positions
        show_authors: "never", "always", or "auto" (default: "always")
                     - "never": never show authors
                     - "always": always include author= attribute on the XML tag
                     - "auto": include authors only when multiple authors exist
        placement: "inline", "end-paragraph", or "comments-only" (default: "inline")
                  - "inline": `<comment author="...">feedback</comment>` placed after the bracketed referenced text
                  - "end-paragraph": numbered markers in text with XML comments grouped at end of each paragraph
                  - "comments-only": extract only comments with text context

    Returns:
        Formatted text according to placement option
    """
    if not ranges:
        return text

    if placement == "inline":
        return _format_inline(text, comments, ranges, show_authors)
    elif placement == "end-paragraph":
        return _format_end_paragraph(text, comments, ranges, show_authors)
    elif placement == "comments-only":
        return _format_comments_only(text, comments, ranges, show_authors)
    else:
        raise ValueError(f"Unknown placement option: {placement}. Valid options: inline, end-paragraph, comments-only")


def _should_show_authors(show_authors: str, comments: List[Comment]) -> bool:
    if show_authors == "never":
        return False
    if show_authors == "always":
        return True
    # auto
    authors = {comment.author for comment in comments if comment.author}
    return len(authors) > 1


def _format_comment(comment: Comment, show_author: bool) -> str:
    """Wrap a comment in a <comment> XML tag (optionally with author attribute)."""
    body = escape(comment.text or '')
    if show_author and comment.author:
        return f'<comment author={quoteattr(comment.author)}>{body}</comment>'
    return f'<comment>{body}</comment>'


def _format_inline(text: str, comments: List[Comment], ranges: List[CommentRange], show_authors: str) -> str:
    if not ranges:
        return text

    comment_map = {c.id: c for c in comments}
    should_show_authors = _should_show_authors(show_authors, comments)
    return _build_nested_inline(text, ranges, comment_map, should_show_authors)


def _build_nested_inline(text: str, ranges: List[CommentRange], comment_map: dict, should_show_authors: bool) -> str:
    if not ranges:
        return text

    events = []
    for range_obj in ranges:
        comment = comment_map.get(range_obj.comment_id)
        if not comment:
            continue
        comment_text = _format_comment(comment, should_show_authors)
        if range_obj.start_pos == range_obj.end_pos:
            events.append((range_obj.start_pos, 'point', comment_text))
        else:
            events.append((range_obj.start_pos, 'start', range_obj))
            events.append((range_obj.end_pos, 'end', range_obj, comment_text))

    events.sort(key=lambda x: (x[0], x[1] == 'end'))

    result = []
    last_pos = 0

    for event in events:
        pos = event[0]
        event_type = event[1]

        result.append(text[last_pos:pos])

        if event_type == 'point':
            result.append(event[2])
        elif event_type == 'start':
            result.append('[')
        elif event_type == 'end':
            result.append(f'] {event[3]}')

        last_pos = pos

    result.append(text[last_pos:])
    return ''.join(result)


def _format_end_paragraph(text: str, comments: List[Comment], ranges: List[CommentRange], show_authors: str) -> str:
    comment_map = {c.id: c for c in comments}
    should_show_authors = _should_show_authors(show_authors, comments)

    paragraphs = text.split('\n\n')
    char_offset = 0
    comment_counter = 1
    result_paragraphs = []

    for paragraph in paragraphs:
        para_start = char_offset
        para_end = char_offset + len(paragraph)

        para_ranges = []
        for range_obj in ranges:
            comment = comment_map.get(range_obj.comment_id)
            if comment and para_start <= range_obj.start_pos <= para_end:
                para_ranges.append(range_obj)

        para_ranges_by_end = sorted(para_ranges, key=lambda r: r.end_pos)
        range_to_number = {}
        para_comments = []

        for range_obj in para_ranges_by_end:
            comment = comment_map.get(range_obj.comment_id)
            if not comment:
                continue
            range_to_number[range_obj] = comment_counter
            para_comments.append((comment_counter, _format_comment(comment, should_show_authors)))
            comment_counter += 1

        events = []
        for range_obj in para_ranges:
            if range_obj not in range_to_number:
                continue
            number = range_to_number[range_obj]
            marker = f'[{number}]'
            rel_start = range_obj.start_pos - para_start
            rel_end = range_obj.end_pos - para_start

            if range_obj.start_pos == range_obj.end_pos:
                events.append((rel_start, 'point', marker, number))
            else:
                events.append((rel_end, 'marker', marker, number))

        events.sort(key=lambda x: (x[0], x[3]))

        modified_paragraph = ""
        last_pos = 0
        for pos, _event_type, marker, _number in events:
            modified_paragraph += paragraph[last_pos:pos]
            modified_paragraph += marker
            last_pos = pos
        modified_paragraph += paragraph[last_pos:]

        if para_comments:
            para_comments.sort(key=lambda x: x[0])
            comment_list = '\n'.join(f'{num}. {xml}' for num, xml in para_comments)
            result_paragraphs.append(modified_paragraph + '\n\nComments:\n' + comment_list)
        else:
            result_paragraphs.append(modified_paragraph)

        char_offset = para_end + 2

    return '\n\n'.join(result_paragraphs)


def _format_comments_only(text: str, comments: List[Comment], ranges: List[CommentRange], show_authors: str) -> str:
    comment_map = {c.id: c for c in comments}
    should_show_authors = _should_show_authors(show_authors, comments)

    sorted_ranges = sorted(ranges, key=lambda r: r.start_pos)

    comment_lines = []
    for range_obj in sorted_ranges:
        comment = comment_map.get(range_obj.comment_id)
        if not comment:
            continue
        xml = _format_comment(comment, should_show_authors)
        if range_obj.start_pos == range_obj.end_pos:
            comment_lines.append(f'[Position {range_obj.start_pos}]: {xml}')
        else:
            commented_text = text[range_obj.start_pos:range_obj.end_pos]
            comment_lines.append(f'"{commented_text}": {xml}')

    return '\n'.join(comment_lines)
