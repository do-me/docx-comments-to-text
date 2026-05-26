"""Parse .xlsx files: render each sheet as a Markdown table with comment XML tags
embedded inside the cell content that owns the comment."""

import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape, quoteattr
from dataclasses import dataclass, field
from typing import Dict, List, Tuple


SHEET_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
PKG_REL_NS = '{http://schemas.openxmlformats.org/package/2006/relationships}'
THREAD_NS = '{http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments}'


@dataclass
class XlsxComment:
    cell_ref: str
    author: str
    text: str


@dataclass
class _Sheet:
    name: str
    path: str
    rels_path: str
    comments: Dict[str, List[XlsxComment]] = field(default_factory=dict)


class XlsxParser:
    def __init__(self, xlsx_path: str):
        self.xlsx_path = xlsx_path

    def render_markdown(self, show_authors: str = 'always', sheet_name: str | None = None) -> str:
        try:
            with zipfile.ZipFile(self.xlsx_path, 'r') as z:
                shared_strings = self._read_shared_strings(z)
                persons = self._read_persons(z)
                sheets = self._read_sheets(z)

                if sheet_name is not None:
                    matched = [s for s in sheets if s.name == sheet_name]
                    if not matched:
                        available = ', '.join(s.name for s in sheets)
                        raise ValueError(
                            f"Sheet '{sheet_name}' not found. Available sheets: {available}"
                        )
                    sheets = matched

                # Resolve comments for each sheet
                for sheet in sheets:
                    sheet.comments = self._read_sheet_comments(z, sheet, persons)

                # Decide author-display mode
                all_comments = [c for s in sheets for cs in s.comments.values() for c in cs]
                should_show = _should_show_authors(show_authors, all_comments)

                parts: List[str] = []
                for sheet in sheets:
                    parts.append(self._render_sheet(z, sheet, shared_strings, should_show))

                return '\n\n'.join(p for p in parts if p)
        except FileNotFoundError:
            raise FileNotFoundError(f"Could not find file: {self.xlsx_path}")

    def list_sheets(self) -> List[str]:
        """Return the sheet names in workbook order."""
        with zipfile.ZipFile(self.xlsx_path, 'r') as z:
            return [s.name for s in self._read_sheets(z)]

    # -- workbook structure ----------------------------------------------------

    def _read_sheets(self, z: zipfile.ZipFile) -> List[_Sheet]:
        wb_xml = z.read('xl/workbook.xml')
        wb_tree = ET.fromstring(wb_xml)

        rels = self._read_rels(z, 'xl/_rels/workbook.xml.rels')

        sheets: List[_Sheet] = []
        for sheet_elem in wb_tree.iter(f'{SHEET_NS}sheet'):
            name = sheet_elem.get('name', '')
            rid = sheet_elem.get(f'{REL_NS}id', '')
            target = rels.get(rid, '')
            path = _normalize_target('xl/workbook.xml', target)
            rels_path = self._sibling_rels_path(path)
            sheets.append(_Sheet(name=name, path=path, rels_path=rels_path))
        return sheets

    def _read_rels(self, z: zipfile.ZipFile, rels_path: str) -> Dict[str, str]:
        try:
            data = z.read(rels_path)
        except KeyError:
            return {}
        tree = ET.fromstring(data)
        result: Dict[str, str] = {}
        for rel in tree.findall(f'{PKG_REL_NS}Relationship'):
            result[rel.get('Id', '')] = rel.get('Target', '')
        return result

    def _read_rels_with_types(self, z: zipfile.ZipFile, rels_path: str) -> List[Tuple[str, str, str]]:
        try:
            data = z.read(rels_path)
        except KeyError:
            return []
        tree = ET.fromstring(data)
        out: List[Tuple[str, str, str]] = []
        for rel in tree.findall(f'{PKG_REL_NS}Relationship'):
            out.append((rel.get('Id', ''), rel.get('Type', ''), rel.get('Target', '')))
        return out

    @staticmethod
    def _sibling_rels_path(path: str) -> str:
        if '/' in path:
            dir_part, file_part = path.rsplit('/', 1)
            return f'{dir_part}/_rels/{file_part}.rels'
        return f'_rels/{path}.rels'

    # -- shared strings & persons ---------------------------------------------

    def _read_shared_strings(self, z: zipfile.ZipFile) -> List[str]:
        try:
            data = z.read('xl/sharedStrings.xml')
        except KeyError:
            return []
        tree = ET.fromstring(data)
        result: List[str] = []
        for si in tree.findall(f'{SHEET_NS}si'):
            result.append(_join_text(si))
        return result

    def _read_persons(self, z: zipfile.ZipFile) -> Dict[str, str]:
        """Threaded comments reference authors by personId — collect displayName per id."""
        persons: Dict[str, str] = {}
        for name in z.namelist():
            if not (name.startswith('xl/persons/') and name.endswith('.xml')):
                continue
            try:
                data = z.read(name)
            except KeyError:
                continue
            try:
                tree = ET.fromstring(data)
            except ET.ParseError:
                continue
            for person in tree.iter():
                # Match <person> by local-name; namespace varies between Office builds
                if not person.tag.endswith('}person') and person.tag != 'person':
                    continue
                pid = person.get('id', '')
                display = person.get('displayName', '')
                if pid:
                    persons[pid] = display
        return persons

    # -- per-sheet comments ----------------------------------------------------

    def _read_sheet_comments(self, z: zipfile.ZipFile, sheet: _Sheet, persons: Dict[str, str]) -> Dict[str, List[XlsxComment]]:
        result: Dict[str, List[XlsxComment]] = {}
        rels = self._read_rels_with_types(z, sheet.rels_path)
        if not rels:
            return result

        threaded_target = None
        legacy_target = None
        for _rid, rtype, target in rels:
            if 'threadedComment' in rtype:
                threaded_target = target
            elif rtype.endswith('/comments'):
                legacy_target = target

        # Prefer threaded comments when present (richer author info); fall
        # back to legacy if the threaded file is missing or yields nothing.
        if threaded_target:
            tpath = _normalize_target(sheet.path, threaded_target)
            try:
                data = z.read(tpath)
                for c in self._parse_threaded_comments(data, persons):
                    result.setdefault(c.cell_ref, []).append(c)
            except KeyError:
                pass

        if not result and legacy_target:
            lpath = _normalize_target(sheet.path, legacy_target)
            try:
                data = z.read(lpath)
                for c in self._parse_legacy_comments(data):
                    result.setdefault(c.cell_ref, []).append(c)
            except KeyError:
                pass

        return result

    def _parse_legacy_comments(self, data: bytes) -> List[XlsxComment]:
        tree = ET.fromstring(data)
        authors: List[str] = []
        for author_elem in tree.findall(f'{SHEET_NS}authors/{SHEET_NS}author'):
            authors.append(author_elem.text or '')

        comments: List[XlsxComment] = []
        for comment_elem in tree.findall(f'{SHEET_NS}commentList/{SHEET_NS}comment'):
            ref = comment_elem.get('ref', '')
            author_id_str = comment_elem.get('authorId', '0')
            try:
                author_id = int(author_id_str)
            except ValueError:
                author_id = 0
            author = authors[author_id] if 0 <= author_id < len(authors) else ''
            text = _join_text(comment_elem.find(f'{SHEET_NS}text')) if comment_elem.find(f'{SHEET_NS}text') is not None else ''
            comments.append(XlsxComment(cell_ref=ref, author=author, text=text))
        return comments

    def _parse_threaded_comments(self, data: bytes, persons: Dict[str, str]) -> List[XlsxComment]:
        tree = ET.fromstring(data)
        comments: List[XlsxComment] = []
        for c in tree.iter(f'{THREAD_NS}threadedComment'):
            ref = c.get('ref', '')
            person_id = c.get('personId', '')
            author = persons.get(person_id, '')
            text_elem = c.find(f'{THREAD_NS}text')
            text = (text_elem.text or '') if text_elem is not None else ''
            if ref:
                comments.append(XlsxComment(cell_ref=ref, author=author, text=text))
        return comments

    # -- sheet rendering -------------------------------------------------------

    def _render_sheet(self, z: zipfile.ZipFile, sheet: _Sheet, shared_strings: List[str], show_authors: bool) -> str:
        try:
            data = z.read(sheet.path)
        except KeyError:
            return ''
        tree = ET.fromstring(data)

        # Build a sparse map (row_idx, col_idx) -> text
        cells: Dict[Tuple[int, int], str] = {}
        max_row = 0
        max_col = 0

        for row in tree.iter(f'{SHEET_NS}row'):
            for cell in row.findall(f'{SHEET_NS}c'):
                ref = cell.get('r', '')
                col_idx, row_idx = _parse_ref(ref)
                if col_idx < 0 or row_idx < 0:
                    continue
                value = _cell_value(cell, shared_strings)
                cells[(row_idx, col_idx)] = value
                max_row = max(max_row, row_idx)
                max_col = max(max_col, col_idx)

        # Merge cell comments into the cell content
        for ref, comment_list in sheet.comments.items():
            col_idx, row_idx = _parse_ref(ref)
            if col_idx < 0 or row_idx < 0:
                continue
            existing = cells.get((row_idx, col_idx), '')
            tags = ''.join(_format_xlsx_comment(c, show_authors) for c in comment_list)
            cells[(row_idx, col_idx)] = existing + tags
            max_row = max(max_row, row_idx)
            max_col = max(max_col, col_idx)

        title = f'## {sheet.name}'.rstrip()
        if not cells:
            return f'{title}\n\n*(empty sheet)*'

        # Render markdown table — first row is the header
        lines: List[str] = [title, '']
        header_cells = [cells.get((0, c), '') for c in range(max_col + 1)]
        lines.append('| ' + ' | '.join(_escape_cell(v) for v in header_cells) + ' |')
        lines.append('|' + ' --- |' * (max_col + 1))
        for r in range(1, max_row + 1):
            row_cells = [cells.get((r, c), '') for c in range(max_col + 1)]
            lines.append('| ' + ' | '.join(_escape_cell(v) for v in row_cells) + ' |')

        return '\n'.join(lines)


# -- helpers -------------------------------------------------------------------

def _join_text(elem) -> str:
    if elem is None:
        return ''
    parts: List[str] = []
    for t in elem.iter(f'{SHEET_NS}t'):
        parts.append(t.text or '')
    return ''.join(parts)


def _parse_ref(ref: str) -> Tuple[int, int]:
    """'B7' -> (col_idx=1, row_idx=6) — both 0-indexed."""
    if not ref:
        return -1, -1
    letters = ''
    digits = ''
    for ch in ref:
        if ch.isalpha():
            letters += ch
        elif ch.isdigit():
            digits += ch
    if not letters or not digits:
        return -1, -1
    col = 0
    for c in letters.upper():
        col = col * 26 + (ord(c) - ord('A') + 1)
    return col - 1, int(digits) - 1


def _cell_value(cell, shared_strings: List[str]) -> str:
    t = cell.get('t', '')
    if t == 'inlineStr':
        is_elem = cell.find(f'{SHEET_NS}is')
        return _join_text(is_elem) if is_elem is not None else ''
    if t == 's':
        v = cell.find(f'{SHEET_NS}v')
        if v is not None and v.text is not None:
            try:
                idx = int(v.text)
                if 0 <= idx < len(shared_strings):
                    return shared_strings[idx]
            except ValueError:
                pass
        return ''
    if t == 'b':
        v = cell.find(f'{SHEET_NS}v')
        return 'TRUE' if (v is not None and v.text == '1') else 'FALSE'
    if t in {'str', 'e'}:
        # Formula string result or error
        v = cell.find(f'{SHEET_NS}v')
        return (v.text or '') if v is not None else ''
    # Number or date (we don't reformat dates — Excel stores them as serial numbers)
    v = cell.find(f'{SHEET_NS}v')
    if v is not None and v.text is not None:
        return v.text
    is_elem = cell.find(f'{SHEET_NS}is')
    if is_elem is not None:
        return _join_text(is_elem)
    return ''


def _escape_cell(value: str) -> str:
    """Escape pipes and replace newlines so the cell stays on one markdown row."""
    if not value:
        return ' '
    return value.replace('|', '\\|').replace('\r\n', '<br>').replace('\n', '<br>').replace('\r', '<br>')


def _format_xlsx_comment(comment: XlsxComment, show_author: bool) -> str:
    body = escape(comment.text or '')
    attrs = f' cell={quoteattr(comment.cell_ref)}' if comment.cell_ref else ''
    if show_author and comment.author:
        return f'<comment author={quoteattr(comment.author)}{attrs}>{body}</comment>'
    return f'<comment{attrs}>{body}</comment>'


def _should_show_authors(show_authors: str, comments: List[XlsxComment]) -> bool:
    if show_authors == 'never':
        return False
    if show_authors == 'always':
        return True
    authors = {c.author for c in comments if c.author}
    return len(authors) > 1


def _normalize_target(source_path: str, target: str) -> str:
    """Resolve a relationship Target relative to the file containing the relationship."""
    if target.startswith('/'):
        return target.lstrip('/')
    if '/' in source_path:
        base = source_path.rsplit('/', 1)[0]
    else:
        base = ''
    parts = (base.split('/') if base else []) + target.split('/')
    out: List[str] = []
    for p in parts:
        if p in ('', '.'):
            continue
        if p == '..':
            if out:
                out.pop()
        else:
            out.append(p)
    return '/'.join(out)
