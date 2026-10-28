"""Turn a PDF into the semantic model (see model.py), keeping its layout.

A PDF only has positioned characters, so the structure is rebuilt from geometry:
lines that continue at the same indent join into paragraphs, a bullet glyph
followed by a gap starts a list item (its indent gives the level), and bigger
bold text becomes a heading. The fonts, sizes, indents and spacing are recorded
too, so one copy can look like the PDF and another can take the user's styles
(see `strip_layout`).
"""

from __future__ import annotations

import copy
import io
import itertools
import logging
import re
import statistics
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

from .docx_reader import _convert_typed_numbers, _promote_numbered_headings
from .model import (
    Block,
    Cell,
    Document,
    Image,
    PageSetup,
    Paragraph,
    Run,
    Table,
)
from .reader import _append_run

log = logging.getLogger(__name__)

# Symbol/Wingdings glyphs without a Unicode mapping come out in the private use area.
PRIVATE_GLYPHS = {
    "": "•", "": "▪", "": "➢", "": "❖", "": "✓",
    "": "➢", "": "▪", "": "▪", "": "▪", "": "•",
    "": "-", "": "•",
}  # fmt: skip
BULLET_GLYPHS = set("•●◦○▪■□▫◆◇❖➢➤►▶→⇒✓✔-–—*o§")
_NUMBER_TOKEN = re.compile(r"^(\()?(\d{1,3}|[a-zA-Z]|[ivxIVX]{1,4})([.)])$")
_ROMAN = re.compile(r"^[ivx]+$|^[IVX]+$")
_BOLD = re.compile(r"bold|black|heavy|semibold|demi", re.IGNORECASE)
_ITALIC = re.compile(r"italic|oblique", re.IGNORECASE)
_FONT_SUFFIX = re.compile(r"(PSMT|PS|MT)$")

LIST_GAP = 4.0  # points between a bullet glyph and its text, at least
INDENT_TOLERANCE = 4.0
CELL_MARGIN = 5.4  # Word's default left/right cell padding


@dataclass
class _Line:
    chars: list[dict]
    page: int
    x0: float
    x1: float
    top: float
    bottom: float
    size: float
    bullet: str | None = None  # glyph, or number token for numbered items
    text_x: float = 0.0  # where the text after the bullet starts
    text_chars: list[dict] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "".join(c["text"] for c in self.text_chars)


def _glyph(text: str) -> str:
    if "" <= text <= "":  # other symbol-font glyphs: some bullet shape
        return PRIVATE_GLYPHS.get(text, "•")
    return text


def font_family(fontname: str) -> str:
    """'ABCDEE+TimesNewRomanPS-BoldMT' -> 'Times New Roman'."""
    name = fontname.split("+", 1)[-1]
    name = re.split(r"[-,]", name, maxsplit=1)[0]
    name = _FONT_SUFFIX.sub("", name)
    return re.sub(r"(?<=[a-z])(?=[A-Z])", " ", name) or fontname


def _color(value) -> str | None:
    if not isinstance(value, list | tuple) or not all(
        isinstance(v, int | float) for v in value
    ):
        return None
    if len(value) == 1:
        rgb = (value[0],) * 3
    elif len(value) == 3:
        rgb = tuple(value)
    elif len(value) == 4:
        c, m, y, k = value
        rgb = ((1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k))
    else:
        return None
    hexed = "".join(f"{round(max(0, min(1, v)) * 255):02X}" for v in rgb)
    return None if hexed == "000000" else hexed


def _inside(obj: dict, box: tuple[float, float, float, float]) -> bool:
    cx, cy = (obj["x0"] + obj["x1"]) / 2, (obj["top"] + obj["bottom"]) / 2
    return box[0] <= cx <= box[2] and box[1] <= cy <= box[3]


class PdfReader:
    def __init__(self, path: Path):
        self.path = Path(path)
        self._skipped: Counter[str] = Counter()
        self._last_columns = 0  # of the last table, to join one split by a page break

    def read(self) -> Document:
        import pdfplumber

        with pdfplumber.open(str(self.path)) as pdf:
            if not pdf.pages:
                return Document([], None)
            items: list[tuple[int, float, object]] = []  # (page, top, line/table/image)
            for number, page in enumerate(pdf.pages):
                items.extend(self._page(number, page))
            first = pdf.pages[0]
            width, height = float(first.width), float(first.height)

        lines = [it for _, _, it in items if isinstance(it, _Line)]
        self._drop_page_furniture(lines, height)
        items = [
            it
            for it in items
            if not (isinstance(it[2], _Line) and it[2].bullet == "drop")
        ]
        lines = [it for _, _, it in items if isinstance(it, _Line)]
        if not lines:
            log.warning("PDF neobsahuje text (naskenovaný dokument?)")
            return Document([], None)

        self.left = _percentile([ln.x0 for ln in lines], 0.02)
        self.right = _percentile([ln.x1 for ln in lines], 0.98)
        self.body = Counter(
            round(ln.size) for ln in lines for _ in ln.chars
        ).most_common(1)[0][0]
        blocks = self._blocks(items)
        self._headings([b for b in blocks if isinstance(b, Paragraph)])
        _promote_numbered_headings([b for b in blocks if isinstance(b, Paragraph)])
        _convert_typed_numbers(blocks)

        page = PageSetup(
            width=width,
            height=height,
            left=max(self.left, 18),
            right=max(width - self.right, 18),
            top=max(min(ln.top for ln in lines) - 2, 18),
            bottom=max(height - max(ln.bottom for ln in lines) - 2, 18),
        )
        for what, count in self._skipped.items():
            log.warning("Vynecháno: %s (%d×)", what, count)
        return Document(blocks, _guess_language(lines), page)

    # ---- one page -----------------------------------------------------------------

    def _page(self, number: int, page) -> list[tuple[int, float, object]]:
        out: list[tuple[int, float, object]] = []
        boxes: list[tuple[float, float, float, float]] = []
        if page.rects or page.lines:
            for found in page.find_tables():
                table = self._table(page, found)
                if table is not None:
                    boxes.append(found.bbox)
                    out.append((number, found.bbox[1], table))
        for img in page.images:
            box = (img["x0"], img["top"], img["x1"], img["bottom"])
            image = self._image(page, box, f"pdf-p{number + 1}-{len(out)}.png")
            if image is not None:
                boxes.append(box)
                out.append((number, img["top"], Paragraph([image])))

        highlights = [
            (
                r["x0"],
                r["top"],
                r["x1"],
                r["bottom"],
                _color(r.get("non_stroking_color")),
            )
            for r in page.rects
            if r.get("fill")
            and r["bottom"] - r["top"] < 40
            and _color(r.get("non_stroking_color")) not in (None, "FFFFFF")
        ]

        def keep(obj) -> bool:
            return obj.get("object_type") != "char" or not any(
                _inside(obj, b) for b in boxes
            )

        raw_lines = page.filter(keep).extract_text_lines(return_chars=True, strip=True)
        for raw in _attach_lone_glyphs(raw_lines):
            chars = [c for c in raw["chars"] if c["text"].strip() or c["text"] == " "]
            if not chars:
                continue
            for c in chars:
                c["text"] = _glyph(c["text"])
                for x0, top, x1, bottom, color in highlights:
                    if _inside(c, (x0, top, x1, bottom)):
                        c["highlight"] = color
            size = Counter(
                round(c["size"], 1) for c in chars if c["text"].strip()
            ).most_common(1)[0][0]
            line = _Line(
                chars, number, raw["x0"], raw["x1"], raw["top"], raw["bottom"], size
            )
            _find_bullet(line)
            out.append((number, line.top, line))
        out.sort(key=lambda it: it[1])
        return out

    def _table(self, page, found) -> Table | None:
        columns = max(len(row.cells) for row in found.rows)
        # the rest of a table that started on the previous page may be one row
        continued = columns == self._last_columns and found.bbox[1] < page.height * 0.15
        if not continued and not _is_grid(found.extract()):
            return None
        self._last_columns = columns
        out: list[list[Cell]] = []
        for row in found.rows:
            cells: list[Cell] = []
            for box in row.cells:
                if box is None and cells:  # covered by the cell to the left
                    cells.append(Cell(merged=True))
                    origin = next(c for c in reversed(cells) if not c.merged)
                    origin.col_span += 1
                else:
                    cells.append(Cell(_cell_blocks(page, box) if box else []))
            cells.extend(Cell() for _ in range(columns - len(cells)))
            out.append(cells)
        return Table(out, col_widths=_column_widths(found.rows, columns))

    def _image(self, page, box, name: str) -> Image | None:
        x0, top, x1, bottom = box
        if x1 - x0 < 2 or bottom - top < 2:
            return None
        try:
            bbox = (
                max(x0, 0),
                max(top, 0),
                min(x1, page.width),
                min(bottom, page.height),
            )
            rendered = page.crop(bbox).to_image(resolution=150).original
            data = io.BytesIO()
            rendered.save(data, format="PNG")
        except Exception as exc:  # noqa: BLE001 - one picture must not stop the conversion
            log.warning(
                "Obrázek na straně %d se nepodařilo převést: %s", page.page_number, exc
            )
            return None
        return Image(data.getvalue(), name, x1 - x0, bottom - top)

    def _drop_page_furniture(self, lines: list[_Line], height: float):
        """Page numbers and running headers/footers repeat in the page margins."""
        pages = len({ln.page for ln in lines})
        margin = [
            ln for ln in lines if ln.top < height * 0.08 or ln.bottom > height * 0.92
        ]
        repeated = Counter(
            (round(ln.top), re.sub(r"\d+", "#", ln.text)) for ln in margin
        )
        for ln in margin:
            text = ln.text.strip()
            numeric = re.fullmatch(
                r"[-–—\s]*(strana|page|str\.)?\s*\d+(\s*(/|z|of)\s*\d+)?[-–—\s]*",
                text,
                re.IGNORECASE,
            )
            if numeric or (
                pages > 2
                and repeated[(round(ln.top), re.sub(r"\d+", "#", ln.text))] > pages / 3
            ):
                ln.bullet = "drop"

    # ---- paragraphs ---------------------------------------------------------------

    def _blocks(self, items) -> list[Block]:
        blocks: list[Block] = []
        groups: list[tuple[Paragraph, list[_Line]]] = []
        current: tuple[Paragraph, list[_Line]] | None = None
        stack: list[float] = []  # bullet x positions of the open list levels

        for _page, _top, item in items:
            if not isinstance(item, _Line):
                last = blocks[-1] if blocks else None
                if (
                    isinstance(item, Table)
                    and isinstance(last, Table)
                    and current is None
                    and len(last.rows[0]) == len(item.rows[0])
                ):
                    last.rows.extend(item.rows)  # continued after a page break
                else:
                    blocks.append(item)
                current = None
                continue
            line = item
            if current is not None and self._continues(current[1], line, current[0]):
                _add_text(current[0], line, separator=" ")
                current[1].append(line)
                continue

            p = Paragraph()
            if line.bullet:
                while stack and line.x0 < stack[-1] - INDENT_TOLERANCE:
                    stack.pop()
                if not stack or line.x0 > stack[-1] + INDENT_TOLERANCE:
                    stack.append(line.x0)
                p.level = min(len(stack) - 1, 8)
                m = _NUMBER_TOKEN.match(line.bullet)
                if m:
                    paren, token, end = m.groups()
                    p.list_kind = "number"
                    p.number_suffix = "()" if paren else end
                    p.number_format = _number_format(token)
                    p.restart = token in ("1", "a", "A", "i", "I")
                else:
                    p.list_kind = "bullet"
                p.left_indent = line.text_x - self.left
                p.first_line_indent = line.x0 - line.text_x
            else:
                # text at the margin, or a heading, ends the open lists
                if (
                    line.x0 <= self.left + INDENT_TOLERANCE
                    or abs(line.size - self.body) > 1
                ):
                    stack = []
                p.left_indent = max(line.x0 - self.left, 0)
            _add_text(p, line, separator="")
            blocks.append(p)
            current = (p, [line])
            groups.append(current)

        for p, lines in groups:
            self._finish(p, lines)
        self._spacing(groups)
        return blocks

    def _centered(self, line: _Line) -> bool:
        center = (self.left + self.right) / 2
        return abs((line.x0 + line.x1) / 2 - center) < 6

    def _continues(self, lines: list[_Line], line: _Line, p: Paragraph) -> bool:
        prev = lines[-1]
        if line.bullet or abs(line.size - prev.size) > 0.6:
            return False
        if line.page == prev.page:
            if line.top - prev.bottom > line.size * 0.8:
                return False
        elif line.page != prev.page + 1:
            return False
        # centred lines (titles) of which at least one is visibly indented
        both_centered = (
            self._centered(prev)
            and self._centered(line)
            and (prev.x0 > self.left + 10 or line.x0 > self.left + 10)
        )
        if p.list_kind and len(lines) == 1:
            # wrapped lines hang under the text, or a bit further in when a tab
            # stop put the first line's text closer to the bullet
            aligned = lines[0].x0 + INDENT_TOLERANCE < line.x0 <= prev.text_x + 24
        elif len(lines) == 1:
            # the first line may be indented, the rest start further left
            aligned = (
                self.left - INDENT_TOLERANCE <= line.x0 <= prev.x0 + INDENT_TOLERANCE
            )
        else:
            aligned = abs(line.x0 - prev.x0) <= INDENT_TOLERANCE
        if not both_centered and not aligned:
            return False
        # The previous line ended early although the next word would have fit: new paragraph.
        word = _first_word_width(line)
        return both_centered or prev.x1 + word + line.size * 0.3 >= self.right - 2

    def _finish(self, p: Paragraph, lines: list[_Line]):
        width = self.right - self.left
        if p.list_kind:
            if len(lines) > 1:
                p.left_indent = lines[1].x0 - self.left
                p.first_line_indent = lines[0].x0 - lines[1].x0
        elif all(self._centered(ln) for ln in lines) and any(
            ln.x0 > self.left + 10 and ln.x1 < self.right - 10 for ln in lines
        ):
            p.align = "center"
            p.left_indent = p.first_line_indent = 0.0
        elif (
            len(lines) == 1
            and lines[0].x1 >= self.right - 3
            and lines[0].x1 - lines[0].x0 < width * 0.6
        ):
            p.align = "right"
            p.left_indent = p.first_line_indent = 0.0
        if len(lines) > 1 and all(ln.x1 >= self.right - 3 for ln in lines[:-1]):
            p.align = p.align or "justify"
        if not p.list_kind and p.align not in ("center", "right") and len(lines) > 1:
            p.left_indent = max(lines[1].x0 - self.left, 0)
            p.first_line_indent = lines[0].x0 - lines[1].x0
        pitches = [
            b.top - a.top for a, b in itertools.pairwise(lines) if a.page == b.page
        ]
        if pitches:
            p.line_spacing = round(statistics.median(pitches), 1)

    def _spacing(self, groups: list[tuple[Paragraph, list[_Line]]]):
        """Exact line pitch and the gap above each paragraph, as in the PDF."""
        typical: dict[float, list[float]] = {}
        for p, lines in groups:
            if p.line_spacing:
                typical.setdefault(round(lines[0].size), []).append(p.line_spacing)
        pitch = {size: statistics.median(v) for size, v in typical.items()}
        for i, (p, lines) in enumerate(groups):
            if p.line_spacing is None:
                p.line_spacing = round(
                    pitch.get(round(lines[0].size), lines[0].size * 1.2), 1
                )
                nxt = groups[i + 1][1][0] if i + 1 < len(groups) else None
                if nxt is not None and nxt.page == lines[-1].page:
                    # single lines set closer than usual (list items with no spacing)
                    p.line_spacing = round(
                        min(p.line_spacing, nxt.top - lines[-1].top), 1
                    )
            if i:
                prev = groups[i - 1][1][-1]
                if prev.page == lines[0].page:
                    delta = lines[0].top - prev.top
                    p.space_before = round(max(delta - p.line_spacing, 0.0), 1)

    def _headings(self, paragraphs: list[Paragraph]):
        """Bigger text than the body, bold or not, is a heading; one level per size."""

        def info(p: Paragraph):
            runs = [r for r in p.items if isinstance(r, Run) and r.text.strip()]
            if not runs or p.list_kind or len(p.text) > 200:
                return None
            size = max(r.size or 0 for r in runs)
            bold = all(r.bold for r in runs)
            if size >= self.body + 1.5 and (bold or size >= self.body + 3):
                return round(size * 2) / 2
            return None

        sizes = {id(p): s for p in paragraphs if (s := info(p)) is not None}
        if not sizes:
            return
        levels = sorted(set(sizes.values()), reverse=True)
        title = None
        biggest = [p for p in paragraphs if sizes.get(id(p)) == levels[0]]
        if len(levels) > 1 and len(biggest) == 1 and paragraphs.index(biggest[0]) < 3:
            title = levels.pop(0)
        for p in paragraphs:
            size = sizes.get(id(p))
            if size is None:
                continue
            p.heading = 0 if size == title else min(levels.index(size) + 1, 6)
            p.align = p.align if p.align == "center" else None


# ---- helpers --------------------------------------------------------------------------


def _percentile(values: list[float], q: float) -> float:
    values = sorted(values)
    return values[min(int(len(values) * q), len(values) - 1)]


def _number_format(token: str) -> str:
    if token.isdigit():
        return "decimal"
    if len(token) > 1 or _ROMAN.match(token) and token in ("i", "I"):
        return "lowerRoman" if token.islower() else "upperRoman"
    return "lowerLetter" if token.islower() else "upperLetter"


def _find_bullet(line: _Line):
    chars = [c for c in line.chars if c["text"].strip()]
    line.text_chars = line.chars
    line.text_x = line.x0
    if len(chars) < 2:
        return
    # the first "word": a glyph, or a number token like "1." / "a)" / "(3)"
    end = 1
    while end < len(chars) and chars[end]["x0"] - chars[end - 1]["x1"] < 1 and end < 5:
        end += 1
    token = "".join(c["text"] for c in chars[:end])
    gap = chars[end]["x0"] - chars[end - 1]["x1"] if end < len(chars) else 0
    symbol_font = re.search(
        r"symbol|wingding|courier|dingbat", chars[0]["fontname"], re.IGNORECASE
    )
    if (
        token in BULLET_GLYPHS
        and (gap >= LIST_GAP or (symbol_font and gap > 1))
        or _NUMBER_TOKEN.match(token)
        and gap >= LIST_GAP * 1.5
    ):
        line.bullet = token
    else:
        return
    first = chars[end]
    line.text_x = first["x0"]
    line.text_chars = line.chars[line.chars.index(first) :]


def _attach_lone_glyphs(lines: list[dict]) -> list[dict]:
    """A bullet glyph set a little lower or higher than its text is found as a
    line of its own; put it back in front of the text beside it."""
    glyphs = [
        ln
        for ln in lines
        if len(ln["chars"]) == 1 and _glyph(ln["text"]) in BULLET_GLYPHS
    ]
    for g in glyphs:
        middle = (g["top"] + g["bottom"]) / 2
        target = next(
            (
                ln
                for ln in lines
                if ln is not g
                and ln not in glyphs
                and ln["top"] - 4 <= middle <= ln["bottom"] + 4
                and 0 <= ln["x0"] - g["x1"] <= 40
            ),
            None,
        )
        if target is not None:
            target["chars"] = g["chars"] + target["chars"]
            target["x0"] = g["x0"]
            target["text"] = g["text"] + " " + target["text"]
            g["chars"] = []
    return [ln for ln in lines if ln["chars"]]


def _first_word_width(line: _Line) -> float:
    chars = [c for c in line.text_chars if c["text"].strip()]
    if not chars:
        return 0.0
    end = 0  # words end at a gap; the extracted characters have no spaces
    for _ in range(2):
        start, end = end, end + 1
        while (
            end < len(chars)
            and chars[end]["x0"] - chars[end - 1]["x1"] < chars[end]["size"] * 0.2
        ):
            end += 1
        # Czech binds one-letter words to the next one ("s právním")
        if end - start > 1 or not chars[start]["text"].isalpha() or end >= len(chars):
            break
    return chars[end - 1]["x1"] - chars[0]["x0"]


def _add_text(p: Paragraph, line: _Line, separator: str):
    if separator:
        last = p.items[-1] if p.items else None
        if isinstance(last, Run) and not last.text.endswith((" ", "-", "­")):
            last.text += separator
        elif isinstance(last, Run) and last.text.endswith("­"):
            last.text = last.text[:-1]  # soft hyphen at a line break
    previous = None
    for c in line.text_chars:
        text = c["text"]
        if (
            previous is not None
            and text != " "
            and previous["text"] != " "
            and c["x0"] - previous["x1"] > c["size"] * 0.2
        ):
            _append_run(p, _run_format(previous, line), " ")
        _append_run(p, _run_format(c, line), text)
        previous = c


def _run_format(c: dict, line: _Line) -> Run:
    name = c["fontname"].split("+", 1)[-1]
    run = Run("")
    run.bold = bool(_BOLD.search(name))
    run.italic = bool(_ITALIC.search(name))
    run.font = font_family(c["fontname"])
    run.size = round(c["size"] * 2) / 2
    run.color = _color(c.get("non_stroking_color"))
    run.highlight = c.get("highlight")
    if c["size"] < line.size * 0.8 and c["bottom"] < line.bottom - line.size * 0.15:
        run.script = "super"
    return run


def _is_grid(rows: list[list[str | None]]) -> bool:
    """A real table: most rows have text in two or more cells (boxes around a
    highlighted paragraph are detected as one-column "tables")."""
    if len(rows) < 2 or max(len(r) for r in rows) < 2:
        return False
    filled = sum(1 for r in rows if sum(1 for c in r if c and c.strip()) >= 2)
    return filled >= len(rows) / 2


def _column_widths(rows, columns: int) -> list[float] | None:
    edges: list[float] = []
    for x in sorted(
        v for row in rows for box in row.cells if box for v in (box[0], box[2])
    ):
        if not edges or x - edges[-1] > 2:
            edges.append(x)
    if len(edges) != columns + 1:
        return None
    return [b - a for a, b in itertools.pairwise(edges)]


def _cell_blocks(page, box) -> list[Block]:
    """Paragraphs of one table cell, with their formatting and bullets."""
    cell = page.within_bbox(box)
    blocks: list[Block] = []
    current: Paragraph | None = None
    text_x = 0.0
    for raw in _attach_lone_glyphs(
        cell.extract_text_lines(return_chars=True, strip=True)
    ):
        chars = [c for c in raw["chars"] if c["text"].strip()]
        if not chars:
            continue
        for c in chars:
            c["text"] = _glyph(c["text"])
        size = Counter(round(c["size"], 1) for c in chars).most_common(1)[0][0]
        line = _Line(
            chars,
            page.page_number,
            raw["x0"],
            raw["x1"],
            raw["top"],
            raw["bottom"],
            size,
        )
        _find_bullet(line)
        if (
            current is not None
            and not line.bullet
            and abs(line.x0 - text_x) <= INDENT_TOLERANCE
        ):
            _add_text(current, line, separator=" ")
            continue
        current = Paragraph()
        if line.bullet:
            current.list_kind = "bullet"
            current.left_indent = max(line.text_x - box[0] - CELL_MARGIN, 0)
            current.first_line_indent = line.x0 - line.text_x
        text_x = line.text_x
        _add_text(current, line, separator="")
        blocks.append(current)
    return blocks


def _guess_language(lines: list[_Line]) -> str | None:
    text = "".join(ln.text for ln in lines[:400]).lower()
    if any(ch in text for ch in "ěřů"):
        return "cs-CZ"
    if any(ch in text for ch in "ľĺŕô"):
        return "sk-SK"
    return None


def strip_layout(doc: Document) -> Document:
    """A copy without the PDF's own look, ready for the user's styles."""
    doc = copy.deepcopy(doc)
    doc.page = None

    def clean(blocks: list[Block]):
        for b in blocks:
            if isinstance(b, Paragraph):
                if b.align == "justify" or b.list_kind or b.heading is not None:
                    b.align = None
                b.left_indent = b.first_line_indent = 0.0
                b.space_before = b.line_spacing = None
                for item in b.items:
                    if isinstance(item, Run):
                        item.font = None
            elif isinstance(b, Table):
                for row in b.rows:
                    for cell in row:
                        clean(cell.blocks)

    clean(doc.blocks)
    return doc
