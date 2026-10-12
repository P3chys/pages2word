"""Turn the body text of a .pages document into a simple semantic model.

Only meaning is carried over (headings, lists, bold/italic/…, colours,
language, images). Fonts, sizes and spacing are deliberately dropped so the
Word side can use its own styles.
"""

from __future__ import annotations

import bisect
import itertools
import logging
import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Any

from keynote_parser.versions.v14_5.generated import TSDArchives_pb2 as TSD
from keynote_parser.versions.v14_5.generated import TSSArchives_pb2 as TSS
from keynote_parser.versions.v14_5.generated import TSWPArchives_pb2 as TSWP

from .iwa import Package

log = logging.getLogger(__name__)

T_STORAGE = 2001
T_DRAWABLE_ATTACHMENT = 2003
T_IMAGE = 3005

PARAGRAPH_BREAKS = "\n "
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


def word_language(code: str | None) -> str | None:
    if not code:
        return None
    code = code.replace("_", "-")
    if "-" in code:
        return code
    return LANGUAGE_REGIONS.get(code, code)


@dataclass
class Run:
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    script: str | None = None  # "super" | "sub"
    caps: bool = False
    color: str | None = None  # RRGGBB
    highlight: str | None = None  # RRGGBB
    lang: str | None = None
    size: float | None = None

    def same_format(self, other: Run) -> bool:
        return all(
            getattr(self, f) == getattr(other, f)
            for f in (
                "bold",
                "italic",
                "underline",
                "strike",
                "script",
                "caps",
                "color",
                "highlight",
                "lang",
                "size",
            )
        )


@dataclass
class Image:
    data: bytes
    filename: str
    width: float  # points
    height: float


@dataclass
class PageBreak:
    pass


@dataclass
class Paragraph:
    items: list[Run | Image | PageBreak] = field(default_factory=list)
    list_kind: str | None = None  # "bullet" | "number"
    level: int = 0
    number_format: str = "decimal"
    number_suffix: str = "."  # "." | ")" | "()"
    restart: bool = False
    heading: int | None = None  # 0 = title, 1..n
    align: str | None = None
    left_indent: float = 0.0
    page_break_before: bool = False

    @property
    def text(self) -> str:
        return "".join(i.text for i in self.items if isinstance(i, Run))

    @property
    def is_empty(self) -> bool:
        return not self.text.strip() and not any(
            isinstance(i, Image) for i in self.items
        )


@dataclass
class Document:
    paragraphs: list[Paragraph]
    language: str | None


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


class PagesReader:
    def __init__(self, path):
        self.pkg = Package(path)
        self._styles: dict[int, Any] = {}
        body = self._find_body()
        self.body = body
        self.text = "".join(body.text)
        sheet = self.pkg.get(body.style_sheet.identifier, TSS.StylesheetArchive)
        self.style_keys = (
            {e.style.identifier: e.identifier for e in sheet.identifier_to_style_map}
            if sheet
            else {}
        )
        key_to_id = {v: k for k, v in self.style_keys.items()}
        self.default_para_style = key_to_id.get("paragraph-style-default")

        obj = lambda e: e.object.identifier or None
        self.t_para_style = _Table(
            body.table_para_style.entries, obj, carry_forward=True
        )
        self.t_char_style = _Table(body.table_char_style.entries, obj)
        self.t_list_style = _Table(body.table_list_style.entries, obj)
        self.t_level = _Table(body.table_para_data.entries, lambda e: e.first)
        self.t_starts = _Table(body.table_para_starts.entries, lambda e: e.first)
        self.t_lang = _Table(body.table_language.entries, lambda e: e.object or None)
        self.attachments = {
            e.character_index: e.object.identifier
            for e in body.table_attachment.entries
        }

    # ---- objects ------------------------------------------------------------

    def _find_body(self):
        best = None
        for _, storage in self.pkg.of_type(T_STORAGE, TSWP.StorageArchive):
            if storage.kind == 0:  # BODY
                size = sum(len(t) for t in storage.text)
                if best is None or size > best[0]:
                    best = (size, storage)
        if best is None:
            raise ValueError("No body text found (is this a page-layout document?)")
        return best[1]

    def _style(self, ident: int | None):
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

    def _chain(self, ident: int | None):
        seen = set()
        while ident and ident not in seen:
            seen.add(ident)
            style = self._style(ident)
            if style is None:
                return
            yield ident, style
            ident = (
                style.super.parent.identifier
                if style.super.HasField("parent")
                else None
            )

    def _char_prop(self, name: str, char_style: int | None, para_style: int | None):
        for sid in (char_style, para_style or self.default_para_style):
            for _, style in self._chain(sid):
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

    def _para_prop(self, name: str, para_style: int | None):
        for _, style in self._chain(para_style or self.default_para_style):
            props = style.para_properties
            if props.HasField(name):
                return getattr(props, name)
        return None

    # ---- conversion ---------------------------------------------------------

    def read(self) -> Document:
        paragraphs = []
        start = 0
        text = self.text
        for i, ch in enumerate(text):
            if ch in PARAGRAPH_BREAKS:
                paragraphs.append(self._paragraph(start, i))
                start = i + 1
        if start < len(text):
            paragraphs.append(self._paragraph(start, len(text)))

        langs = Counter()
        for p in paragraphs:
            for r in p.items:
                if isinstance(r, Run) and r.lang:
                    langs[r.lang] += len(r.text)
        language = langs.most_common(1)[0][0] if langs else None
        for p in paragraphs:
            for r in p.items:
                if isinstance(r, Run) and r.lang is None:
                    r.lang = language

        _normalize_list_levels(paragraphs)
        _convert_typed_bullets(paragraphs)
        _detect_headings(paragraphs)
        return Document(paragraphs, language)

    def _paragraph(self, start: int, end: int) -> Paragraph:
        para_style = self.t_para_style.at(start)
        p = Paragraph()

        heading = self._named_heading(para_style)
        if heading is not None:
            p.heading = heading
        else:
            outline = self._para_prop("outline_level", para_style)
            if outline is not None and outline < 9:
                p.heading = outline + 1

        align = self._para_prop("alignment", para_style)
        p.align = ALIGNMENTS.get(align)
        p.left_indent = self._para_prop("left_indent", para_style) or 0.0
        p.page_break_before = bool(self._para_prop("page_break_before", para_style))

        self._apply_list(p, start)

        points = {start, end}
        for t in (self.t_char_style, self.t_lang):
            points.update(t.changes_in(start, end))
        for pos in self.attachments:
            if start <= pos < end:
                points.update((pos, pos + 1))
        bounds = sorted(points)
        for a, b in itertools.pairwise(bounds):
            self._emit(p, a, b, para_style)
        return p

    def _named_heading(self, para_style):
        for ident, _ in self._chain(para_style):
            key = self.style_keys.get(ident, "")
            for suffix, level in NAMED_HEADINGS.items():
                if key.endswith(suffix):
                    return level
        return None

    def _apply_list(self, p: Paragraph, start: int):
        ls = self._style(self.t_list_style.at(start))
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
        chunk = self.text[a:b]
        if a in self.attachments:
            img = self._image(self.attachments[a])
            if img:
                p.items.append(img)
            chunk = chunk[1:] if chunk.startswith(OBJECT_REPLACEMENT) else chunk
            if not chunk:
                return

        cs = self.t_char_style.at(a)
        prop = lambda n: self._char_prop(n, cs, para_style)
        C = TSWP.CharacterStylePropertiesArchive
        color = prop("font_color")
        background = prop("background_color")
        superscript = prop("superscript")
        run_proto = Run(
            text="",
            bold=bool(prop("bold")),
            italic=bool(prop("italic")),
            underline=(prop("underline") or 0) != C.kNoUnderline,
            strike=(prop("strikethru") or 0) != C.kNoStrikethru,
            script={C.kSuperscript: "super", C.kSubscript: "sub"}.get(superscript),
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
                p.items.append(PageBreak())
            elif ch == " ":
                buf.append("\n")
            elif ch == OBJECT_REPLACEMENT or (ord(ch) < 32 and ch != "\t"):
                continue
            else:
                buf.append(ch)
        if buf:
            _append_run(p, run_proto, "".join(buf))

    def _image(self, attachment_id: int) -> Image | None:
        att = self.pkg.get(attachment_id, TSWP.DrawableAttachmentArchive)
        if att is None:
            return None
        drawable_id = att.drawable.identifier
        if self.pkg.type_of(drawable_id) != T_IMAGE:
            log.warning(
                "Skipping unsupported embedded object (type %s)",
                self.pkg.type_of(drawable_id),
            )
            return None
        image = self.pkg.get(drawable_id, TSD.ImageArchive)
        found = self.pkg.read_data(image.data.identifier)
        if not found:
            log.warning("Image data %s missing from package", image.data.identifier)
            return None
        size = image.super.geometry.size
        return Image(found[1], found[0], size.width, size.height)


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


def _convert_typed_bullets(paragraphs: list[Paragraph]):
    out: list[Paragraph] = []
    in_run = False  # previous paragraph was a bullet
    for p in paragraphs:
        text = p.text
        prev = out[-1] if out else None
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
    paragraphs[:] = out


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
