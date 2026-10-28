"""Turn a Word document (.docx, or .doc via Word) into the semantic model (see model.py).

Used to restyle a messy Word document: only meaning is kept (headings, lists,
bold/italic/…, colours, language, images, tables, text boxes, footnotes). Fonts,
sizes, indents, justification and the document's own list definitions are
dropped, so the writer can apply a style document's look to all of it.
"""

from __future__ import annotations

import contextlib
import logging
import os
import re
import subprocess
import sys
import tempfile
from collections import Counter
from collections.abc import Iterator
from pathlib import Path

import docx
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import parse_xml
from docx.oxml.ns import qn

from .model import (
    Block,
    Cell,
    Document,
    Footnote,
    Image,
    PageBreak,
    Paragraph,
    Run,
    Table,
    TextBox,
)
from .reader import (
    _append_run,
    _convert_typed_bullets,
    _detect_headings,
    _dominant_language,
    _fill_language,
    _split_out_blocks,
    _strip_prefix,
    word_language,
)
from .wordstyles import Styles, _on, _val

log = logging.getLogger(__name__)

WORD_SUFFIXES = (".docx", ".doc")
EMU_PER_PT = 12700
MC_CHOICE = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Choice"

NUMBER_FORMATS = {"decimal", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman"}

# Word's named highlight colours
HIGHLIGHTS = {
    "yellow": "FFFF00", "green": "00FF00", "cyan": "00FFFF", "magenta": "FF00FF",
    "blue": "0000FF", "red": "FF0000", "darkBlue": "000080", "darkCyan": "008080",
    "darkGreen": "008000", "darkMagenta": "800080", "darkRed": "800000",
    "darkYellow": "808000", "darkGray": "808080", "lightGray": "C0C0C0", "black": "000000",
}  # fmt: skip

# Symbol/Wingdings characters (w:sym) used as typed bullets: dot, square, arrow.
SYMBOL_BULLETS = {"F0B7": "•", "F0A7": "▪", "F0D8": "•", "F076": "•", "F0FC": "•"}


# ---- .doc -------------------------------------------------------------------------

# Word started by automation runs document macros by default: force them off
# (AutomationSecurity 3 = msoAutomationSecurityForceDisable) before opening.
_DOC_TO_DOCX = (
    "$ErrorActionPreference = 'Stop'; $w = New-Object -ComObject Word.Application; "
    "try { $w.Visible = $false; $w.DisplayAlerts = 0; $w.AutomationSecurity = 3; "
    "$w.WordBasic.DisableAutoMacros(1); "
    "$d = $w.Documents.Open($env:P2W_SRC, $false, $true, $false); "
    "$d.SaveAs2($env:P2W_DST, 16); $d.Close(0) } finally { $w.Quit() }"
)


@contextlib.contextmanager
def open_word_file(path: Path) -> Iterator[Path]:
    """Yield a .docx for `path`; an old .doc is converted by Word into a temp folder."""
    if path.suffix.lower() != ".doc":
        yield path
        return
    with tempfile.TemporaryDirectory(prefix="pages2word-") as tmp:
        out = Path(tmp) / (path.stem + ".docx")
        env = {**os.environ, "P2W_SRC": str(path.resolve()), "P2W_DST": str(out)}
        flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        try:
            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-NonInteractive",
                    "-Command",
                    _DOC_TO_DOCX,
                ],
                env=env,
                check=False,  # a missing output file is the error we report
                capture_output=True,
                timeout=600,
                creationflags=flags,
            )
        except OSError, subprocess.TimeoutExpired:
            pass
        if not out.exists():
            raise RuntimeError(
                "starý formát .doc se nepodařilo otevřít ve Wordu "
                "(uložte ho ve Wordu jako .docx)"
            )
        yield out


# ---- numbering ----------------------------------------------------------------------


class _ListDefinitions:
    """numbering.xml: which number format a list paragraph shows at each level."""

    def __init__(self, document, styles: Styles):
        self.styles = styles
        self.nums: dict[str, object] = {}
        self.abstracts: dict[str, object] = {}
        for rel in document.part.rels.values():
            if rel.reltype == RT.NUMBERING and not rel.is_external:
                root = rel.target_part.element
                self.nums = {n.get(qn("w:numId")): n for n in root.findall(qn("w:num"))}
                self.abstracts = {
                    a.get(qn("w:abstractNumId")): a
                    for a in root.findall(qn("w:abstractNum"))
                }

    def level(self, num_id: str, ilvl: int, depth: int = 0):
        num = self.nums.get(num_id)
        if num is None or depth > 3:
            return None
        for override in num.findall(qn("w:lvlOverride")):
            lvl = override.find(qn("w:lvl"))
            if override.get(qn("w:ilvl")) == str(ilvl) and lvl is not None:
                return lvl
        abstract = self.abstracts.get(_val(num.find(qn("w:abstractNumId"))))
        if abstract is None:
            return None
        link = _val(abstract.find(qn("w:numStyleLink")))
        if link:  # the list is defined by a numbering style ("MiskaStyl")
            num_pr = self.styles.find(link, "pPr", "numPr")
            linked = _val(num_pr.find(qn("w:numId"))) if num_pr is not None else None
            return self.level(linked, ilvl, depth + 1) if linked else None
        for lvl in abstract.findall(qn("w:lvl")):
            if lvl.get(qn("w:ilvl")) == str(ilvl):
                return lvl
        return None


def _number_suffix(text: str) -> str:
    if text.startswith("("):
        return "()"
    return ")" if text.endswith(")") else "."


# ---- reader -------------------------------------------------------------------------


class DocxReader:
    def __init__(self, path: Path):
        self.path = Path(path)
        self._skipped: Counter[str] = Counter()

    def read(self) -> Document:
        with open_word_file(self.path) as path:
            self.document = docx.Document(str(path))
            self.part = self.document.part
            self.styles = Styles(self.document.styles.element)
            self.lists = _ListDefinitions(self.document, self.styles)
            self.notes = self._load_notes()
            self._last_num: str | None = None
            blocks = self._blocks(self.document.element.body)

        language = _dominant_language(blocks)
        _fill_language(blocks, language)
        _convert_typed_bullets(blocks, merge_continuations=False)
        _promote_numbered_headings([b for b in blocks if isinstance(b, Paragraph)])
        _convert_typed_numbers(blocks)
        _detect_headings([b for b in blocks if isinstance(b, Paragraph)])

        for what, count in self._skipped.items():
            log.warning("Vynecháno: %s (%d×)", what, count)
        return Document(blocks, language)

    def _load_notes(self) -> dict[tuple[str, str], object]:
        notes = {}
        for rel in self.part.rels.values():
            kind = {RT.FOOTNOTES: "footnote", RT.ENDNOTES: "endnote"}.get(rel.reltype)
            if kind is None or rel.is_external:
                continue
            root = parse_xml(rel.target_part.blob)
            for note in root.findall(qn(f"w:{kind}")):
                notes[(kind, note.get(qn("w:id")))] = note
        return notes

    # ---- blocks ---------------------------------------------------------------------

    def _blocks(self, container) -> list[Block]:
        out: list[Block] = []
        for child in container:
            if child.tag == qn("w:p"):
                out.extend(_split_out_blocks(self._paragraph(child)))
            elif child.tag == qn("w:tbl"):
                out.append(self._table(child))
            elif child.tag in (qn("w:sdt"), qn("w:customXml")):
                content = child.find(qn("w:sdtContent"))
                out.extend(self._blocks(child if content is None else content))
        return out

    def _table(self, tbl) -> Table:
        rows: list[list[Cell]] = []
        header_rows = 0
        open_merges: dict[int, Cell] = {}  # grid column -> cell that spans down
        for tr in tbl.findall(qn("w:tr")):
            trpr = tr.find(qn("w:trPr"))
            if (
                trpr is not None
                and _on(trpr.find(qn("w:tblHeader")))
                and len(rows) == header_rows
            ):
                header_rows += 1
            row: list[Cell] = []
            for tc in tr.iter(qn("w:tc")):
                if tc.getparent() is not tr and tc.getparent().tag != qn(
                    "w:sdtContent"
                ):
                    continue  # a cell of a nested table
                tcpr = tc.find(qn("w:tcPr"))
                span = _val(tcpr.find(qn("w:gridSpan"))) if tcpr is not None else None
                span = int(span) if span and span.isdigit() else 1
                vmerge = tcpr.find(qn("w:vMerge")) if tcpr is not None else None
                col = len(row)
                if (
                    vmerge is not None
                    and _val(vmerge) != "restart"
                    and col in open_merges
                ):
                    open_merges[col].row_span += 1
                    cell = Cell(merged=True)
                else:
                    cell = Cell(self._blocks(tc), col_span=span)
                    if vmerge is not None:
                        open_merges[col] = cell
                    else:
                        open_merges.pop(col, None)
                row.append(cell)
                row.extend(Cell(merged=True) for _ in range(span - 1))
            rows.append(row)
        width = max((len(r) for r in rows), default=0)
        for row in rows:
            row.extend(Cell() for _ in range(width - len(row)))
        return Table(rows, header_rows)

    def _paragraph(self, el) -> Paragraph:
        p = Paragraph()
        sid = self.styles.paragraph_style(el)
        ppr = el.find(qn("w:pPr"))
        heading = self.styles.heading_level(sid)
        outline = ppr.find(qn("w:outlineLvl")) if ppr is not None else None
        if outline is not None and (_val(outline) or "").isdigit():
            heading = int(_val(outline)) + 1 if int(_val(outline)) < 9 else None

        num_id, ilvl = self.styles.num_pr(el, sid)
        lvl = self.lists.level(num_id, ilvl) if num_id else None
        fmt = _val(lvl.find(qn("w:numFmt"))) if lvl is not None else None
        if heading is not None:
            p.heading = heading
        elif fmt == "bullet":
            p.list_kind, p.level = "bullet", min(ilvl, 8)
        elif fmt and fmt != "none":
            p.list_kind, p.level = "number", min(ilvl, 8)
            p.number_format = fmt if fmt in NUMBER_FORMATS else "decimal"
            p.number_suffix = _number_suffix(_val(lvl.find(qn("w:lvlText"))) or "")
            p.restart = num_id != self._last_num
            self._last_num = num_id

        if ppr is not None:
            jc = _val(ppr.find(qn("w:jc")))
            if (
                jc in ("center", "right", "end")
                and not p.list_kind
                and p.heading is None
            ):
                p.align = "center" if jc == "center" else "right"
            p.page_break_before = _on(ppr.find(qn("w:pageBreakBefore")))

        self._inline(el, p, sid)
        if p.heading is not None and p.is_empty:
            p.heading = None
        return p

    # ---- inline content ---------------------------------------------------------------

    def _inline(self, container, p: Paragraph, sid: str | None):
        for child in container:
            tag = child.tag
            if tag == qn("w:r"):
                self._run(child, p, sid)
            elif tag in (qn("w:hyperlink"), qn("w:ins"), qn("w:smartTag"), qn("w:fldSimple"),
                         qn("w:customXml"), qn("w:bdo"), qn("w:dir")):  # fmt: skip
                self._inline(child, p, sid)
            elif tag == qn("w:sdt"):
                content = child.find(qn("w:sdtContent"))
                if content is not None:
                    self._inline(content, p, sid)

    def _run(self, r, p: Paragraph, sid: str | None):
        proto = self._format(r, sid)

        def text(value: str):
            _append_run(p, proto, value)

        for child in r:
            tag = child.tag
            if tag == qn("w:t"):
                text(child.text or "")
            elif tag == qn("w:tab"):
                text("\t")
            elif tag == qn("w:br"):
                if child.get(qn("w:type")) == "page":
                    p.items.append(PageBreak())
                else:
                    text("\n")
            elif tag == qn("w:cr"):
                text("\n")
            elif tag == qn("w:noBreakHyphen"):
                text("-")
            elif tag == qn("w:sym"):
                char = (child.get(qn("w:char")) or "").upper()
                if char in SYMBOL_BULLETS:
                    text(SYMBOL_BULLETS[char])
            elif tag in (qn("w:footnoteReference"), qn("w:endnoteReference")):
                kind = "footnote" if tag == qn("w:footnoteReference") else "endnote"
                note = self.notes.get((kind, child.get(qn("w:id"))))
                if note is not None:
                    p.items.append(Footnote(self._note_blocks(note)))
            elif tag in (qn("w:drawing"), qn("w:pict")):
                self._object(child, p)
            elif tag.endswith("}AlternateContent"):
                choice = child.find(MC_CHOICE)
                for inner in choice if choice is not None else []:
                    if inner.tag in (qn("w:drawing"), qn("w:pict")):
                        self._object(inner, p)
            elif tag == qn("w:object"):
                self._skipped["vložený objekt"] += 1

    def _note_blocks(self, note) -> list[Block]:
        blocks = self._blocks(note)
        first = next((b for b in blocks if isinstance(b, Paragraph)), None)
        if first and first.items and isinstance(first.items[0], Run):
            first.items[0].text = first.items[0].text.lstrip()
        return blocks

    def _object(self, el, p: Paragraph):
        """A picture or a text box (DrawingML or legacy VML)."""
        box = next(el.iter(qn("w:txbxContent")), None)
        if box is not None:
            p.items.append(TextBox(self._blocks(box)))
            return
        blip = next((e for e in el.iter() if e.tag.endswith("}blip")), None)
        if blip is not None:
            rid = blip.get(qn("r:embed"))
            extent = next((e for e in el.iter() if e.tag.endswith("}extent")), None)
            size = None
            if extent is not None:
                size = (
                    int(extent.get("cx", 0)) / EMU_PER_PT,
                    int(extent.get("cy", 0)) / EMU_PER_PT,
                )
        else:
            data = next((e for e in el.iter() if e.tag.endswith("}imagedata")), None)
            rid = data.get(qn("r:id")) if data is not None else None
            shape = data.getparent() if data is not None else None
            style = shape.get("style", "") if shape is not None else ""
            dims = dict(re.findall(r"(width|height):([\d.]+)pt", style))
            size = (
                (float(dims["width"]), float(dims["height"]))
                if len(dims) == 2
                else None
            )
        part = self.part.related_parts.get(rid) if rid else None
        if part is None or not size or not all(size):
            self._skipped["objekt bez obrázku"] += 1
            return
        name = str(part.partname).rsplit("/", 1)[-1]
        p.items.append(Image(part.blob, name, *size))

    def _format(self, r, sid: str | None) -> Run:
        rpr = r.find(qn("w:rPr"))
        char_style = _val(rpr.find(qn("w:rStyle"))) if rpr is not None else None

        def prop(name: str):
            if rpr is not None and (el := rpr.find(qn(f"w:{name}"))) is not None:
                return el
            return self.styles.find(char_style, "rPr", name)

        def para_prop(name: str):
            el = self.styles.find(sid, "rPr", name)
            return el if el is not None else self.styles.default_run(name)

        run = Run("")
        run.bold = _on(prop("b"))
        run.italic = _on(prop("i"))
        u = prop("u")
        run.underline = u is not None and _val(u) not in ("none", "0")
        run.strike = _on(prop("strike")) or _on(prop("dstrike"))
        run.caps = _on(prop("caps"))
        script = _val(prop("vertAlign"))
        run.script = {"superscript": "super", "subscript": "sub"}.get(script or "")

        color = prop("color")
        value = _val(color)
        if (
            value
            and value not in ("auto", "000000")
            and not (color.get(qn("w:themeColor")) in ("text1", "dark1"))
        ):
            run.color = value.upper()
        highlight = _val(prop("highlight"))
        shading = prop("shd")
        fill = shading.get(qn("w:fill")) if shading is not None else None
        if highlight in HIGHLIGHTS:
            run.highlight = HIGHLIGHTS[highlight]
        elif fill and fill.lower() not in ("auto", "ffffff"):
            run.highlight = fill.upper()

        lang = prop("lang")
        if lang is None:
            lang = para_prop("lang")
        run.lang = word_language(_val(lang))
        size = prop("sz")
        if size is None:
            size = para_prop("sz")
        if size is not None and (_val(size) or "").isdigit():
            run.size = int(_val(size)) / 2
        return run


# ---- typed numbering ------------------------------------------------------------------

_NUMBERED_TITLE = re.compile(r"^\s*(\d{1,3})\.\s+\S")


def _title_number(p: Paragraph) -> int | None:
    m = _NUMBERED_TITLE.match(p.text)
    return int(m.group(1)) if m else None


def _promote_numbered_headings(paragraphs: list[Paragraph]):
    """'23. Vláda…' typed as plain text between the headings '22. …' and '24. …'
    is one of them: give it the level of the numbered headings it continues."""

    def candidate(p: Paragraph) -> bool:
        return p.heading is None and not p.list_kind and len(p.text.strip()) <= 150

    for order in (paragraphs, paragraphs[::-1]):
        step = 1 if order is paragraphs else -1
        last: tuple[int, int] | None = None  # (number, heading level)
        for p in order:
            number = _title_number(p)
            if number is None:
                continue
            if p.heading:
                last = (number, p.heading)
            elif candidate(p) and last and number == last[0] + step:
                p.heading = last[1]
                last = (number, last[1])


# "a) text", "B. text", "(3) text": lists typed by hand. Each must be part of a
# sequence (a, b, c…) of at least two items, so "A. Novák" or "1. ledna" stay text.
_TYPED_NUMBER = re.compile(r"^[ \t]*(\()?([a-zA-Z]|\d{1,2})([.)])[ \t ]+(?=\S)")


def _typed_label(p: Paragraph):
    if not p.items or not isinstance(p.items[0], Run):
        return None
    m = _TYPED_NUMBER.match(p.text)
    if not m:
        return None
    paren, token, end = m.groups()
    if paren and end != ")":
        return None
    if token.isdigit():
        kind, value = "decimal", int(token)
    elif token.islower():
        kind, value = "lowerLetter", ord(token) - ord("a") + 1
    else:
        kind, value = "upperLetter", ord(token) - ord("A") + 1
    suffix = "()" if paren else end
    return (kind, suffix), value, m.end()


def _convert_typed_numbers(blocks: list[Block]):
    open_runs: dict[tuple[str, str], list[tuple[Paragraph, int]]] = {}

    def finish(key):
        items = open_runs.pop(key, [])
        if len(items) < 2:
            return
        kind, suffix = key
        for i, (p, prefix) in enumerate(items):
            _strip_prefix(p, prefix)
            p.list_kind, p.level = "number", 0
            p.number_format, p.number_suffix = kind, suffix
            p.restart = i == 0

    for p in blocks:
        if not isinstance(p, Paragraph):
            continue
        if p.heading is not None:
            if p.heading <= 2:  # typed lists don't run across sections
                for key in list(open_runs):
                    finish(key)
            continue
        if p.list_kind:
            continue
        label = _typed_label(p)
        if label is None:
            continue
        key, value, prefix = label
        if value == 1:
            finish(key)
            open_runs[key] = [(p, prefix)]
        elif key in open_runs and value == len(open_runs[key]) + 1:
            open_runs[key].append((p, prefix))
    for key in list(open_runs):
        finish(key)
