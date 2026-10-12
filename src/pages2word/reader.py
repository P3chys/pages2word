"""Turn a .pages document into the semantic model (see model.py).

Only meaning is carried over (headings, lists, bold/italic/…, colours,
language, images, tables, text boxes, footnotes, pencil drawings). Fonts,
sizes and spacing are deliberately dropped so the Word side can use its own
styles.
"""

from __future__ import annotations

import bisect
import itertools
import logging
import re
from collections import Counter
from typing import Any

from google.protobuf.internal.decoder import _DecodeVarint
from keynote_parser.versions.v14_5.generated import TSDArchives_pb2 as TSD
from keynote_parser.versions.v14_5.generated import TSSArchives_pb2 as TSS
from keynote_parser.versions.v14_5.generated import TSWPArchives_pb2 as TSWP

from .iwa import Package
from .model import (
    Block,
    Document,
    Footnote,
    Image,
    PageBreak,
    Paragraph,
    Pencil,
    Run,
    Table,
    TextBox,
)
from .tables import TableReader

__all__ = [
    "Document",
    "Image",
    "PageBreak",
    "PagesReader",
    "Paragraph",
    "Run",
    "word_language",
]

log = logging.getLogger(__name__)

T_STORAGE = 2001
T_DRAWABLE_ATTACHMENT = 2003
T_FOOTNOTE = 2008
T_SHAPE_INFO = 2011
T_PENCIL = 2016
T_SHAPE = 3004
T_IMAGE = 3005
T_GROUP = 3008
T_TABLE_INFO = 6000
T_FLOATING_DRAWABLES = 10010  # TP.FloatingDrawablesArchive (no published proto)

UNSUPPORTED = {
    T_SHAPE: "tvar bez textu",
    3007: "video/zvuk",
    5021: "graf",
    2015: "rovnice",
}

PARAGRAPH_BREAKS = "\n "
LINE_BREAK = " "
PAGE_BREAKS = "\x0c\x04\x05"
OBJECT_REPLACEMENT = "￼"

# Pages stylesheet keys of the built-in paragraph styles -> heading level (0 = Title).
NAMED_HEADINGS = {
    "paragraphstyle-Title": 0,
    "paragraphstyle-Heading": 1,
    "paragraphstyle-Heading 1": 1,
    "paragraphstyle-Heading 2": 2,
    "paragraphstyle-Heading 3": 3,
    "paragraphstyle-Heading 4": 4,
}

# Pages alignment enum: 0 left, 1 right, 2 center, 3 justify, 4 natural.
ALIGNMENTS = {1: "right", 2: "center", 3: "justify"}

NUMBER_FORMATS = {
    "Numeric": "decimal",
    "RomanUpper": "upperRoman",
    "RomanLower": "lowerRoman",
    "AlphaUpper": "upperLetter",
    "AlphaLower": "lowerLetter",
}

LANGUAGE_REGIONS = {
    "cs": "cs-CZ", "sk": "sk-SK", "en": "en-US", "de": "de-DE", "fr": "fr-FR",
    "es": "es-ES", "it": "it-IT", "pl": "pl-PL", "pt": "pt-PT", "ru": "ru-RU",
    "uk": "uk-UA", "hu": "hu-HU", "nl": "nl-NL", "la": "la",
}  # fmt: skip


# Language codes come from the file and end up in XML attributes: accept only BCP 47 shapes.
_LANGUAGE_CODE = re.compile(r"[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*")


def word_language(code: str | None) -> str | None:
    if not code:
        return None
    code = code.replace("_", "-")
    if not _LANGUAGE_CODE.fullmatch(code):
        return None
    if "-" in code:
        return code
    return LANGUAGE_REGIONS.get(code, code)


class _Table:
    """A TSWP attribute table: value in effect from each character index on."""

    def __init__(self, entries, value, carry_forward: bool = False):
        self.idx = [e.character_index for e in entries]
        self.val = [value(e) for e in entries]
        if carry_forward:  # a null entry means "unchanged from the previous paragraph"
            for k in range(1, len(self.val)):
                if self.val[k] is None:
                    self.val[k] = self.val[k - 1]

    def at(self, pos: int):
        k = bisect.bisect_right(self.idx, pos) - 1
        return self.val[k] if k >= 0 else None

    def changes_in(self, start: int, end: int) -> list[int]:
        lo = bisect.bisect_right(self.idx, start)
        hi = bisect.bisect_left(self.idx, end)
        return self.idx[lo:hi]


def _hex(color) -> str:
    return "".join(
        f"{round(max(0.0, min(1.0, c)) * 255):02X}" for c in (color.r, color.g, color.b)
    )


def _raw_fields(data: bytes) -> list[tuple[int, Any]]:
    """Minimal protobuf wire decoder for archives without a published .proto."""
    pos, out = 0, []
    while pos < len(data):
        key, pos = _DecodeVarint(data, pos)
        number, wire = key >> 3, key & 7
        if wire == 0:
            value, pos = _DecodeVarint(data, pos)
        elif wire == 1:
            value, pos = data[pos : pos + 8], pos + 8
        elif wire == 5:
            value, pos = data[pos : pos + 4], pos + 4
        elif wire == 2:
            length, pos = _DecodeVarint(data, pos)
            value, pos = data[pos : pos + length], pos + length
        else:
            raise ValueError(f"unsupported wire type {wire}")
        out.append((number, value))
    return out


def _reference(data: bytes) -> int:
    """Follow field 1 through nested messages down to the object identifier."""
    for _ in range(4):
        value = next((v for n, v in _raw_fields(data) if n == 1), None)
        if isinstance(value, int):
            return value
        if not isinstance(value, bytes):
            return 0
        data = value
    return 0


class PagesReader:
    def __init__(self, path):
        self.pkg = Package(path)
        self._styles: dict[int, Any] = {}
        self._skipped: Counter[str] = Counter()
        self._attached: set[int] = set()  # drawables placed in the text flow
        self._reading: set[int] = set()  # storages being read (cycle guard)
        self.body = self._find_body()
        sheet = self.pkg.get(self.body.style_sheet.identifier, TSS.StylesheetArchive)
        self.style_keys = (
            {e.style.identifier: e.identifier for e in sheet.identifier_to_style_map}
            if sheet
            else {}
        )
        key_to_id = {v: k for k, v in self.style_keys.items()}
        self.default_para_style = key_to_id.get("paragraph-style-default")
        self.language: str | None = None
        self.tables = TableReader(self.pkg, self.read_storage_id, None)

    # ---- objects ------------------------------------------------------------

    def _find_body(self):
        best = None
        for _, storage in self.pkg.of_type(T_STORAGE, TSWP.StorageArchive):
            if storage.kind == TSWP.StorageArchive.BODY:
                size = sum(len(t) for t in storage.text)
                if best is None or size > best[0]:
                    best = (size, storage)
        if best is None:
            raise ValueError("No body text found (is this a page-layout document?)")
        return best[1]

    def style(self, ident: int | None):
        if not ident:
            return None
        if ident not in self._styles:
            t = self.pkg.type_of(ident)
            cls = {
                2021: TSWP.CharacterStyleArchive,
                2022: TSWP.ParagraphStyleArchive,
                2023: TSWP.ListStyleArchive,
            }.get(t)
            self._styles[ident] = self.pkg.get(ident, cls) if cls else None
        return self._styles[ident]

    def chain(self, ident: int | None):
        seen = set()
        while ident and ident not in seen:
            seen.add(ident)
            style = self.style(ident)
            if style is None:
                return
            yield ident, style
            ident = (
                style.super.parent.identifier
                if style.super.HasField("parent")
                else None
            )

    def char_prop(self, name: str, char_style: int | None, para_style: int | None):
        for sid in (char_style, para_style or self.default_para_style):
            for _, style in self.chain(sid):
                props = style.char_properties
                null = f"{name}_null"
                if (
                    null in props.DESCRIPTOR.fields_by_name
                    and props.HasField(null)
                    and getattr(props, null)
                ):
                    return None
                if props.HasField(name):
                    return getattr(props, name)
        return None

    def para_prop(self, name: str, para_style: int | None):
        for _, style in self.chain(para_style or self.default_para_style):
            props = style.para_properties
            if props.HasField(name):
                return getattr(props, name)
        return None

    def named_heading(self, para_style):
        for ident, _ in self.chain(para_style):
            key = self.style_keys.get(ident, "")
            for suffix, level in NAMED_HEADINGS.items():
                if key.endswith(suffix):
                    return level
        return None

    def skip(self, what: str):
        self._skipped[what] += 1

    # ---- conversion ---------------------------------------------------------

    def read(self) -> Document:
        blocks = _StorageReader(self, self.body, body=True).read()
        self.language = _dominant_language(blocks)
        _fill_language(blocks, self.language)

        paragraphs = [b for b in blocks if isinstance(b, Paragraph)]
        _normalize_list_levels(paragraphs)
        _convert_typed_bullets(blocks)
        _detect_headings([b for b in blocks if isinstance(b, Paragraph)])
        blocks += self._floating_objects()

        for what, count in self._skipped.items():
            log.warning("Vynecháno: %s (%d×)", what, count)
        return Document(blocks, self.language)

    def read_storage_id(self, ident: int) -> list[Block]:
        storage = self.pkg.get(ident, TSWP.StorageArchive)
        if (
            storage is None or ident in self._reading
        ):  # a malformed file could nest a storage in itself
            return []
        self._reading.add(ident)
        try:
            return _StorageReader(self, storage).read()
        finally:
            self._reading.discard(ident)

    def drawable(self, ident: int) -> list[Image | Table | TextBox]:
        """Convert a drawable (image, table, text box, group) into model objects."""
        self._attached.add(ident)
        t = self.pkg.type_of(ident)
        if t == T_IMAGE:
            image = self._image(self.pkg.get(ident, TSD.ImageArchive))
            return [image] if image else []
        if t == T_TABLE_INFO:
            self.tables.language = self.language
            table = self.tables.read(ident)
            return [table] if table else []
        if t == T_SHAPE_INFO:
            shape = self.pkg.get(ident, TSWP.ShapeInfoArchive)
            storage_id = (
                shape.owned_storage.identifier or shape.deprecated_storage.identifier
                if shape
                else 0
            )
            blocks = self.read_storage_id(storage_id) if storage_id else []
            if any(not (isinstance(b, Paragraph) and b.is_empty) for b in blocks):
                return [TextBox(blocks)]
            self.skip(UNSUPPORTED[T_SHAPE])
            return []
        if t == T_GROUP:
            group = self.pkg.get(ident, TSD.GroupArchive)
            children = group.children if group else []
            return [
                obj
                for child in children
                if child.identifier not in self._attached
                for obj in self.drawable(child.identifier)
            ]
        self.skip(UNSUPPORTED.get(t, f"objekt typu {t}"))
        return []

    def _image(self, image) -> Image | None:
        found = self.pkg.read_data(image.data.identifier) if image is not None else None
        if not found:
            self.skip("obrázek bez dat")
            return None
        size = image.super.geometry.size
        return Image(found[1], found[0], size.width, size.height)

    def pencil(self, ident: int) -> Pencil | None:
        annotation = self.pkg.get(ident, TSWP.PencilAnnotationArchive)
        storage = (
            self.pkg.get(
                annotation.pencil_annotation_storage.identifier,
                TSD.PencilAnnotationStorageArchive,
            )
            if annotation
            else None
        )
        found = (
            self.pkg.read_data(storage.rasterized_image.identifier) if storage else None
        )
        if not found:
            self.skip("kresba Apple Pencil bez obrázku")
            return None
        size = storage.strokes_bounding_box_size
        image = Image(found[1], found[0], size.width, size.height)
        return Pencil(image, storage.markup_offset.x, storage.markup_offset.y)

    def _floating_objects(self) -> list[Block]:
        """Objects placed on a page (not in the text flow), listed after the text by page."""
        pages: dict[int, list[int]] = {}
        for msgs in self.pkg.objects.values():
            if not msgs or msgs[0][0] != T_FLOATING_DRAWABLES:
                continue
            try:
                for number, group in _raw_fields(msgs[0][1]):
                    if number != 1:
                        continue
                    fields = _raw_fields(group)
                    page = next(
                        (v for n, v in fields if n == 1 and isinstance(v, int)), 0
                    )
                    for n, value in fields:
                        ident = _reference(value) if n == 4 else 0
                        if ident:
                            pages.setdefault(page, []).append(ident)
            except ValueError, IndexError:
                log.warning("Could not read floating objects list")
        blocks: list[Block] = []
        for page in sorted(pages):
            objects = [
                obj
                for ident in pages[page]
                if ident not in self._attached
                for obj in self.drawable(ident)
            ]
            if not objects:
                continue
            blocks.append(
                Paragraph(
                    items=[
                        Run(
                            f"Objekty ze strany {page + 1}:",
                            italic=True,
                            lang=self.language,
                        )
                    ]
                )
            )
            for obj in objects:
                blocks.append(Paragraph(items=[obj]) if isinstance(obj, Image) else obj)
        if blocks:
            blocks[0].page_break_before = True
        return blocks


class _StorageReader:
    """Reads one TSWP text storage (body, cell, text box or footnote) into blocks."""

    def __init__(self, owner: PagesReader, storage, body: bool = False):
        self.owner = owner
        self.storage = storage
        self.body = body
        self.text = "".join(storage.text)
        obj = lambda e: e.object.identifier or None
        self.t_para_style = _Table(
            storage.table_para_style.entries, obj, carry_forward=True
        )
        self.t_char_style = _Table(storage.table_char_style.entries, obj)
        self.t_list_style = _Table(storage.table_list_style.entries, obj)
        self.t_level = _Table(storage.table_para_data.entries, lambda e: e.first)
        self.t_starts = _Table(storage.table_para_starts.entries, lambda e: e.first)
        self.t_lang = _Table(storage.table_language.entries, lambda e: e.object or None)
        self.attachments = {
            e.character_index: e.object.identifier
            for e in storage.table_attachment.entries
        }
        self.footnotes = {
            e.character_index: e.object.identifier
            for e in storage.table_footnote.entries
            if e.object.identifier
        }
        self.pencils = {
            e.range.location: e.field.identifier
            for e in storage.table_pencil_annotation.entries
        }

    def read(self) -> list[Block]:
        blocks: list[Block] = []
        start = 0
        for i, ch in enumerate(self.text):
            if ch in PARAGRAPH_BREAKS:
                blocks += self._paragraph(start, i)
                start = i + 1
        if start < len(self.text):
            blocks += self._paragraph(start, len(self.text))
        return blocks

    def _paragraph(self, start: int, end: int) -> list[Block]:
        owner = self.owner
        para_style = self.t_para_style.at(start)
        p = Paragraph()

        heading = owner.named_heading(para_style)
        if heading is not None:
            p.heading = heading
        else:
            outline = owner.para_prop("outline_level", para_style)
            if outline is not None and outline < 9:
                p.heading = outline + 1

        p.align = ALIGNMENTS.get(owner.para_prop("alignment", para_style))
        p.left_indent = owner.para_prop("left_indent", para_style) or 0.0
        p.page_break_before = (
            bool(owner.para_prop("page_break_before", para_style)) and self.body
        )
        self._apply_list(p, start)

        points = {start, end}
        for t in (self.t_char_style, self.t_lang):
            points.update(t.changes_in(start, end))
        for pos in itertools.chain(self.attachments, self.footnotes, self.pencils):
            if start <= pos < end:
                points.update((pos, pos + 1))
        for a, b in itertools.pairwise(sorted(points)):
            self._emit(p, a, b, para_style)
        return _split_out_blocks(p)

    def _apply_list(self, p: Paragraph, start: int):
        ls = self.owner.style(self.t_list_style.at(start))
        level = self.t_level.at(start) or 0
        p.level = level
        p.restart = bool(self.t_starts.at(start))
        if ls is None or not ls.label_types:
            return
        label = ls.label_types[min(level, len(ls.label_types) - 1)]
        if label == TSWP.ListStyleArchive.kNumber:
            p.list_kind = "number"
            if ls.number_types:
                nt = TSWP.ListStyleArchive.NumberType.Name(
                    ls.number_types[min(level, len(ls.number_types) - 1)]
                )
                for prefix, fmt in NUMBER_FORMATS.items():
                    if nt.startswith("k" + prefix):
                        p.number_format = fmt
                        p.number_suffix = (
                            "()"
                            if "DoubleParen" in nt
                            else ")"
                            if "RightParen" in nt
                            else "."
                        )
        elif label in (TSWP.ListStyleArchive.kString, TSWP.ListStyleArchive.kImage):
            p.list_kind = "bullet"

    def _emit(self, p: Paragraph, a: int, b: int, para_style):
        owner = self.owner
        chunk = self.text[a:b]
        if a in self.pencils:
            pencil = owner.pencil(self.pencils[a])
            if pencil:
                p.items.append(pencil)
        if a in self.attachments:
            att = owner.pkg.get(self.attachments[a], TSWP.DrawableAttachmentArchive)
            if att is not None:
                p.items += owner.drawable(att.drawable.identifier)
            else:
                owner.skip(
                    f"vložený objekt typu {owner.pkg.type_of(self.attachments[a])}"
                )
            chunk = chunk.removeprefix(OBJECT_REPLACEMENT)
        if a in self.footnotes:
            note = owner.pkg.get(
                self.footnotes[a], TSWP.FootnoteReferenceAttachmentArchive
            )
            if note is not None and note.contained_storage.identifier:
                p.items.append(
                    Footnote(owner.read_storage_id(note.contained_storage.identifier))
                )
            chunk = chunk.removeprefix(OBJECT_REPLACEMENT)
        if not chunk:
            return

        cs = self.t_char_style.at(a)
        prop = lambda n: owner.char_prop(n, cs, para_style)
        C = TSWP.CharacterStylePropertiesArchive
        color = prop("font_color")
        background = prop("background_color")
        run_proto = Run(
            text="",
            bold=bool(prop("bold")),
            italic=bool(prop("italic")),
            underline=(prop("underline") or 0) != C.kNoUnderline,
            strike=(prop("strikethru") or 0) != C.kNoStrikethru,
            script={C.kSuperscript: "super", C.kSubscript: "sub"}.get(
                prop("superscript")
            ),
            caps=prop("capitalization") == C.kAllCaps,
            color=_hex(color)
            if color is not None and getattr(color, "a", 1) > 0
            else None,
            highlight=_hex(background)
            if background is not None and getattr(background, "a", 1) > 0
            else None,
            lang=word_language(self.t_lang.at(a)),
            size=prop("font_size"),
        )
        if run_proto.color == "000000":
            run_proto.color = None
        if run_proto.highlight == "FFFFFF":
            run_proto.highlight = None

        buf = []
        for ch in chunk:
            if ch in PAGE_BREAKS:
                if buf:
                    _append_run(p, run_proto, "".join(buf))
                    buf = []
                if self.body:
                    p.items.append(PageBreak())
            elif ch == LINE_BREAK:
                buf.append("\n")
            elif ch == OBJECT_REPLACEMENT or (ord(ch) < 32 and ch != "\t"):
                continue
            else:
                buf.append(ch)
        if buf:
            _append_run(p, run_proto, "".join(buf))


def _split_out_blocks(p: Paragraph) -> list[Block]:
    """A table or text box anchored inside a paragraph becomes its own block."""
    if not any(isinstance(i, Table | TextBox) for i in p.items):
        return [p]
    out: list[Block] = []
    current = _copy_props(p)
    for item in p.items:
        if isinstance(item, Table | TextBox):
            if not current.is_empty:
                out.append(current)
            out.append(item)
            current = _copy_props(p)
            current.list_kind = None
        else:
            current.items.append(item)
    if not current.is_empty:
        out.append(current)
    return out


def _copy_props(p: Paragraph) -> Paragraph:
    return Paragraph(
        list_kind=p.list_kind,
        level=p.level,
        number_format=p.number_format,
        number_suffix=p.number_suffix,
        restart=p.restart,
        heading=p.heading,
        align=p.align,
        left_indent=p.left_indent,
        page_break_before=p.page_break_before,
    )


def _walk_runs(blocks: list[Block]):
    for b in blocks:
        if isinstance(b, Paragraph):
            for item in b.items:
                if isinstance(item, Run):
                    yield item
                elif isinstance(item, Footnote):
                    yield from _walk_runs(item.blocks)
        elif isinstance(b, TextBox):
            yield from _walk_runs(b.blocks)
        elif isinstance(b, Table):
            for row in b.rows:
                for cell in row:
                    yield from _walk_runs(cell.blocks)


def _dominant_language(blocks: list[Block]) -> str | None:
    langs: Counter[str] = Counter()
    for r in _walk_runs(blocks):
        if r.lang:
            langs[r.lang] += len(r.text)
    return langs.most_common(1)[0][0] if langs else None


def _fill_language(blocks: list[Block], language: str | None):
    for r in _walk_runs(blocks):
        if r.lang is None:
            r.lang = language


def _append_run(p: Paragraph, proto: Run, text: str):
    last = p.items[-1] if p.items else None
    if isinstance(last, Run) and last.same_format(proto):
        last.text += text
    else:
        p.items.append(Run(**{**proto.__dict__, "text": text}))


def _normalize_list_levels(paragraphs: list[Paragraph]):
    """Pages users often indent by two levels per step (0, 2, 4…); compact that."""
    levels = [p.level for p in paragraphs if p.list_kind]
    if not levels:
        return
    odd = sum(1 for lv in levels if lv % 2)
    halve = odd <= len(levels) * 0.05 and any(lv >= 2 for lv in levels)
    for p in paragraphs:
        if p.list_kind:
            p.level = min(p.level // 2 if halve else p.level, 8)


def _detect_headings(paragraphs: list[Paragraph]):
    """Pages docs often fake headings with big bold text; promote them to real headings."""
    sizes = Counter()
    for p in paragraphs:
        for r in p.items:
            if isinstance(r, Run) and r.size:
                sizes[round(r.size)] += len(r.text)
    base = sizes.most_common(1)[0][0] if sizes else 12

    for p in paragraphs:
        runs = [r for r in p.items if isinstance(r, Run) and r.text.strip()]
        if p.heading is not None or p.list_kind or not runs:
            continue
        text = p.text.strip()
        all_bold = all(r.bold for r in runs)
        size = max((r.size or base) for r in runs)
        if len(text) > 150:
            continue
        if size >= base + 8:
            p.heading = 1
        elif size >= base + 2 and all_bold:
            p.heading = 2
        elif (
            all_bold
            and size >= base - 1
            and len(text) <= 100
            and not text.endswith(".")
        ):
            p.heading = 3


# Bullets that were flattened to plain text (pasted from Word or a PDF): "•" is
# level 1, "o" level 2, "§"/"▪" level 3 (Word's default glyphs). "o" and "§" are
# also ordinary Czech text, so they count only inside a run of bullets or when
# followed by a capital letter ("o Vysílání", not "o tom").
_TYPED_BULLET = re.compile(r"^[ \t]*(\d{1,3})?([•●◦▪■·*\-–o§])[ \t\u00a0]+(?=\S)")
_ALWAYS_BULLET = {
    "•": 0,
    "●": 0,
    "·": 0,
    "*": 0,
    "-": 0,
    "–": 0,
    "◦": 1,
    "▪": 2,
    "■": 2,
}
_CONTEXT_BULLET = {"o": 1, "§": 2}


def _convert_typed_bullets(blocks: list[Block]):
    out: list[Block] = []
    in_run = False  # previous paragraph was a bullet
    for p in blocks:
        if not isinstance(p, Paragraph):
            in_run = False
            out.append(p)
            continue
        text = p.text
        prev = out[-1] if out and isinstance(out[-1], Paragraph) else None
        if p.list_kind or p.heading is not None:
            in_run = p.list_kind == "bullet"
            out.append(p)
            continue

        m = (
            _TYPED_BULLET.match(text)
            if p.items and isinstance(p.items[0], Run)
            else None
        )
        level = None
        if m:
            page_number, glyph = (
                m.groups()
            )  # PDF pastes can glue a page number on: "13o Text"
            rest = text[m.end() :]
            if glyph in _ALWAYS_BULLET:
                level = _ALWAYS_BULLET[glyph]
            elif (
                glyph in _CONTEXT_BULLET
                and (in_run or rest[:1].isupper())
                and not rest[:1].isdigit()
            ):
                level = _CONTEXT_BULLET[glyph]
            if page_number and not in_run:
                level = None

        if level is not None:
            _strip_prefix(p, m.end())
            p.list_kind, p.level = "bullet", level
            in_run = True
        elif (
            in_run
            and prev is not None
            and text[:1]
            and (text[:1].islower() or text[:1] in "(…,;)")
        ):
            # hard-wrapped continuation of the previous bullet
            _merge_into(prev, p)
            continue
        else:
            in_run = False
        out.append(p)
    blocks[:] = out


def _merge_into(target: Paragraph, p: Paragraph):
    last = target.items[-1] if target.items else None
    if isinstance(last, Run) and not last.text.endswith((" ", "\u00a0")):
        last.text += " "
    for item in p.items:
        if isinstance(item, Run):
            _append_run(target, item, item.text)
        else:
            target.items.append(item)


def _strip_prefix(p: Paragraph, count: int):
    while count and p.items and isinstance(p.items[0], Run):
        run = p.items[0]
        if len(run.text) <= count:
            count -= len(run.text)
            p.items.pop(0)
        else:
            run.text = run.text[count:]
            count = 0
