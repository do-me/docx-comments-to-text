import zipfile
from xml.etree import ElementTree as ET
from dataclasses import dataclass, field
from typing import List


@dataclass
class Comment:
    id: str
    author: str
    text: str


@dataclass(frozen=True)
class CommentRange:
    comment_id: str
    start_pos: int
    end_pos: int


@dataclass
class _RenderState:
    parts: list = field(default_factory=list)
    position: int = 0
    comment_starts: dict = field(default_factory=dict)
    ranges: list = field(default_factory=list)
    list_counters: dict = field(default_factory=dict)
    in_table: bool = False
    # Style we currently have an "open" markdown marker for
    open_bold: bool = False
    open_italic: bool = False
    # Pending trailing whitespace to defer until we know the next style
    pending_ws: str = ''

    def append(self, text: str) -> None:
        if not text:
            return
        self.parts.append(text)
        self.position += len(text)

    def text(self) -> str:
        return ''.join(self.parts)


class DocxParser:
    WORD_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

    SKIP_STYLES = {'TOC1', 'TOC2', 'TOC3', 'TOC4', 'TOC5', 'TOC6', 'TOC7', 'TOC8', 'TOC9', 'TOCHeading'}

    def __init__(self, docx_path: str):
        self.docx_path = docx_path
        self._numbering_cache: dict | None = None

    def extract_text_and_comments(self) -> tuple[str, List[Comment], List[CommentRange]]:
        try:
            with zipfile.ZipFile(self.docx_path, 'r') as docx:
                document_xml = docx.read('word/document.xml')

                comments_xml = None
                try:
                    comments_xml = docx.read('word/comments.xml')
                except KeyError:
                    pass

                numbering_xml = None
                try:
                    numbering_xml = docx.read('word/numbering.xml')
                except KeyError:
                    pass

                self._load_numbering(numbering_xml)

                text, ranges = self._render_document(document_xml, with_ranges=comments_xml is not None)
                comments = self._extract_comments(comments_xml) if comments_xml else []

                return text, comments, ranges
        except FileNotFoundError:
            raise FileNotFoundError(f"Could not find file: {self.docx_path}")

    # -- comments.xml ----------------------------------------------------------

    def _extract_comments(self, comments_xml: bytes) -> List[Comment]:
        tree = ET.fromstring(comments_xml)
        comments = []

        for comment_elem in tree.iter():
            if comment_elem.tag != f'{self.WORD_NS}comment':
                continue
            comment_id = comment_elem.get(f'{self.WORD_NS}id')
            author = comment_elem.get(f'{self.WORD_NS}author', '')

            text_parts: list[str] = []
            paragraphs_seen = 0
            for sub in comment_elem.iter():
                if sub.tag == f'{self.WORD_NS}p':
                    if paragraphs_seen > 0 and text_parts:
                        text_parts.append('\n')
                    paragraphs_seen += 1
                elif sub.tag == f'{self.WORD_NS}t' and sub.text:
                    text_parts.append(sub.text)
                elif sub.tag == f'{self.WORD_NS}tab':
                    text_parts.append('\t')
                elif sub.tag == f'{self.WORD_NS}br':
                    text_parts.append('\n')

            comments.append(Comment(id=comment_id, author=author, text=''.join(text_parts)))

        return comments

    # -- numbering.xml ---------------------------------------------------------

    def _load_numbering(self, numbering_xml: bytes | None) -> None:
        """Build a map numId -> {ilvl: 'bullet' | 'decimal'} so we can pick markers."""
        self._numbering_cache = {}
        if numbering_xml is None:
            return
        try:
            tree = ET.fromstring(numbering_xml)
        except ET.ParseError:
            return

        # abstractNumId -> {ilvl: fmt}
        abstract_formats: dict[str, dict[int, str]] = {}
        for abstract_num in tree.findall(f'{self.WORD_NS}abstractNum'):
            abs_id = abstract_num.get(f'{self.WORD_NS}abstractNumId', '')
            levels: dict[int, str] = {}
            for lvl in abstract_num.findall(f'{self.WORD_NS}lvl'):
                ilvl = int(lvl.get(f'{self.WORD_NS}ilvl', '0'))
                num_fmt = lvl.find(f'{self.WORD_NS}numFmt')
                fmt = num_fmt.get(f'{self.WORD_NS}val', 'bullet') if num_fmt is not None else 'bullet'
                levels[ilvl] = fmt
            abstract_formats[abs_id] = levels

        # numId -> abstractNumId
        for num in tree.findall(f'{self.WORD_NS}num'):
            num_id = num.get(f'{self.WORD_NS}numId', '')
            abs_ref = num.find(f'{self.WORD_NS}abstractNumId')
            if abs_ref is None:
                continue
            abs_id = abs_ref.get(f'{self.WORD_NS}val', '')
            self._numbering_cache[num_id] = abstract_formats.get(abs_id, {})

    def _list_marker(self, num_id: str, ilvl: int, state: _RenderState) -> str:
        levels = (self._numbering_cache or {}).get(num_id, {})
        fmt = levels.get(ilvl, 'bullet')
        indent = '  ' * ilvl
        if fmt in {'decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman'}:
            key = (num_id, ilvl)
            n = state.list_counters.get(key, 0) + 1
            state.list_counters[key] = n
            return f'{indent}{n}. '
        # bullet (or unknown) → use markdown bullets
        return f'{indent}- '

    # -- document.xml ----------------------------------------------------------

    def _render_document(self, document_xml: bytes, with_ranges: bool) -> tuple[str, List[CommentRange]]:
        tree = ET.fromstring(document_xml)
        body = tree.find(f'{self.WORD_NS}body')
        if body is None:
            return '', []

        state = _RenderState()
        last_was_list = False

        for child in body:
            tag = child.tag
            if tag == f'{self.WORD_NS}p':
                is_list = self._is_list_paragraph(child)
                # Reset list counters when leaving a list region
                if last_was_list and not is_list:
                    state.list_counters.clear()
                self._render_paragraph(child, state, with_ranges)
                last_was_list = is_list
            elif tag == f'{self.WORD_NS}tbl':
                state.list_counters.clear()
                last_was_list = False
                self._render_table(child, state, with_ranges)
            # ignore sectPr and other body-level metadata

        # Any unclosed comment range starts become point comments
        if with_ranges:
            for comment_id, start_pos in state.comment_starts.items():
                state.ranges.append(CommentRange(comment_id=comment_id, start_pos=start_pos, end_pos=start_pos))

        text = state.text().rstrip('\n') + '\n' if state.parts else ''
        return text.rstrip('\n'), state.ranges

    def _is_list_paragraph(self, p) -> bool:
        return p.find(f'{self.WORD_NS}pPr/{self.WORD_NS}numPr') is not None

    def _paragraph_style(self, p) -> str:
        pStyle = p.find(f'{self.WORD_NS}pPr/{self.WORD_NS}pStyle')
        if pStyle is None:
            return ''
        return pStyle.get(f'{self.WORD_NS}val', '')

    def _paragraph_prefix(self, p, state: _RenderState) -> str:
        style = self._paragraph_style(p)

        # Numbered/bulleted lists
        numPr = p.find(f'{self.WORD_NS}pPr/{self.WORD_NS}numPr')
        if numPr is not None:
            ilvl_elem = numPr.find(f'{self.WORD_NS}ilvl')
            num_id_elem = numPr.find(f'{self.WORD_NS}numId')
            ilvl = int(ilvl_elem.get(f'{self.WORD_NS}val', '0')) if ilvl_elem is not None else 0
            num_id = num_id_elem.get(f'{self.WORD_NS}val', '') if num_id_elem is not None else ''
            return self._list_marker(num_id, ilvl, state)

        # Headings
        if style.startswith('Heading'):
            suffix = style[len('Heading'):]
            if suffix.isdigit():
                level = max(1, min(6, int(suffix)))
                return '#' * level + ' '
        if style in {'Title'}:
            return '# '
        if style in {'Subtitle'}:
            return '## '
        return ''

    def _render_paragraph(self, p, state: _RenderState, with_ranges: bool) -> None:
        style = self._paragraph_style(p)
        if style in self.SKIP_STYLES:
            return

        prefix = self._paragraph_prefix(p, state)
        state.append(prefix)

        self._render_inline_children(p, state, with_ranges)
        self._close_inline_style(state)

        # Paragraph break — blank line between blocks
        state.append('\n\n')

    def _render_inline_children(self, parent, state: _RenderState, with_ranges: bool) -> None:
        for child in parent:
            tag = child.tag
            if tag == f'{self.WORD_NS}r':
                self._render_run(child, state, with_ranges)
            elif tag == f'{self.WORD_NS}hyperlink':
                self._render_inline_children(child, state, with_ranges)
            elif tag == f'{self.WORD_NS}smartTag' or tag == f'{self.WORD_NS}sdt':
                # smart tags and structured doc tags act as transparent wrappers
                # for sdt, recurse into sdtContent
                sdt_content = child.find(f'{self.WORD_NS}sdtContent')
                target = sdt_content if sdt_content is not None else child
                self._render_inline_children(target, state, with_ranges)
            elif with_ranges and tag == f'{self.WORD_NS}commentRangeStart':
                cid = child.get(f'{self.WORD_NS}id')
                if cid is not None:
                    state.comment_starts[cid] = state.position
            elif with_ranges and tag == f'{self.WORD_NS}commentRangeEnd':
                cid = child.get(f'{self.WORD_NS}id')
                if cid is not None and cid in state.comment_starts:
                    start = state.comment_starts.pop(cid)
                    state.ranges.append(CommentRange(comment_id=cid, start_pos=start, end_pos=state.position))
            elif with_ranges and tag == f'{self.WORD_NS}commentReference':
                cid = child.get(f'{self.WORD_NS}id')
                if cid is not None and not any(r.comment_id == cid for r in state.ranges):
                    state.ranges.append(CommentRange(comment_id=cid, start_pos=state.position, end_pos=state.position))

    def _render_run(self, r, state: _RenderState, with_ranges: bool) -> None:
        rPr = r.find(f'{self.WORD_NS}rPr')
        is_bold = False
        is_italic = False
        if rPr is not None:
            b = rPr.find(f'{self.WORD_NS}b')
            i = rPr.find(f'{self.WORD_NS}i')
            is_bold = b is not None and b.get(f'{self.WORD_NS}val', 'true') not in {'0', 'false'}
            is_italic = i is not None and i.get(f'{self.WORD_NS}val', 'true') not in {'0', 'false'}

        # Gather plain text (with possible inline comment markers)
        text = ''
        for child in r:
            ctag = child.tag
            if ctag == f'{self.WORD_NS}t':
                text += child.text or ''
            elif ctag == f'{self.WORD_NS}tab':
                text += '\t' if not state.in_table else ' '
            elif ctag == f'{self.WORD_NS}br':
                text += ' ' if state.in_table else '\n'
            elif with_ranges and ctag == f'{self.WORD_NS}commentReference':
                # Flush accumulated text first, then record point comment
                if text:
                    self._emit_styled_text(text, is_bold, is_italic, state)
                    text = ''
                cid = child.get(f'{self.WORD_NS}id')
                if cid is not None and not any(r2.comment_id == cid for r2 in state.ranges):
                    state.ranges.append(CommentRange(comment_id=cid, start_pos=state.position, end_pos=state.position))

        if text:
            self._emit_styled_text(text, is_bold, is_italic, state)

    def _emit_styled_text(self, text: str, is_bold: bool, is_italic: bool, state: _RenderState) -> None:
        """Emit text, opening/closing markdown bold/italic markers across run boundaries.

        Whitespace-only chunks are emitted without markers and do not close any
        currently-open style, so that adjacent same-style runs merge cleanly.
        """
        if state.in_table:
            text = text.replace('|', '\\|').replace('\n', ' ')

        if not text:
            return

        # Whitespace-only chunks: hold them until we know the next style. If the
        # next styled chunk has the same style, the whitespace stays inside the
        # currently-open markers; if the style differs, we'll close the markers
        # first and emit the whitespace outside.
        if not text.strip():
            state.pending_ws += text
            return

        # Split text into leading/core/trailing whitespace so markdown markers
        # never wrap whitespace (which would break rendering).
        leading = text[:len(text) - len(text.lstrip())]
        trailing = text[len(text.rstrip()):]
        core = text.strip()

        # If style matches what's currently open, just continue inside it
        same_style = (is_bold == state.open_bold) and (is_italic == state.open_italic)
        if same_style:
            # Flush pending whitespace inside the open style
            state.append(state.pending_ws)
            state.pending_ws = ''
            state.append(leading + core + trailing)
            return

        # Style is changing — close current, open new
        self._close_inline_style(state)
        # Pending whitespace now sits *outside* the previous style
        state.append(state.pending_ws)
        state.pending_ws = ''
        state.append(leading)
        marker = self._marker_for(is_bold, is_italic)
        state.append(marker)
        state.append(core)
        state.open_bold = is_bold
        state.open_italic = is_italic
        # Defer the trailing whitespace so we can decide whether it belongs
        # inside or outside the marker based on the next run
        state.pending_ws = trailing

    def _close_inline_style(self, state: _RenderState) -> None:
        """Close any currently-open bold/italic markers and flush deferred whitespace."""
        if state.open_bold or state.open_italic:
            marker = self._marker_for(state.open_bold, state.open_italic)
            state.append(marker)
            state.open_bold = False
            state.open_italic = False
        if state.pending_ws:
            state.append(state.pending_ws)
            state.pending_ws = ''

    @staticmethod
    def _marker_for(is_bold: bool, is_italic: bool) -> str:
        if is_bold and is_italic:
            return '***'
        if is_bold:
            return '**'
        if is_italic:
            return '*'
        return ''

    # -- tables ----------------------------------------------------------------

    def _render_table(self, tbl, state: _RenderState, with_ranges: bool) -> None:
        rows = list(tbl.findall(f'{self.WORD_NS}tr'))
        if not rows:
            return

        state.in_table = True
        try:
            for row_idx, tr in enumerate(rows):
                cells = list(tr.findall(f'{self.WORD_NS}tc'))
                if not cells:
                    continue

                state.append('| ')
                for cell_idx, tc in enumerate(cells):
                    if cell_idx > 0:
                        state.append(' | ')
                    self._render_cell(tc, state, with_ranges)
                state.append(' |\n')

                if row_idx == 0:
                    state.append('|')
                    for _ in cells:
                        state.append(' --- |')
                    state.append('\n')

            state.append('\n')
        finally:
            state.in_table = False

    def _render_cell(self, tc, state: _RenderState, with_ranges: bool) -> None:
        cell_start = state.position
        paragraphs = list(tc.findall(f'{self.WORD_NS}p'))
        for i, p in enumerate(paragraphs):
            if i > 0:
                self._close_inline_style(state)
                state.append(' ')
            self._render_inline_children(p, state, with_ranges)
        self._close_inline_style(state)
        # Ensure non-empty cells so markdown rendering isn't ambiguous
        if state.position == cell_start:
            state.append(' ')
