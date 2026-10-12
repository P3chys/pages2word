"""Semantic document model shared by the reader and the writer."""

from __future__ import annotations

from dataclasses import dataclass, field

RUN_FORMAT_FIELDS = (
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
        return all(getattr(self, f) == getattr(other, f) for f in RUN_FORMAT_FIELDS)


@dataclass
class Image:
    data: bytes
    filename: str
    width: float  # points
    height: float


@dataclass
class Pencil:
    """Apple Pencil drawing, floating next to the character it is attached to."""

    image: Image
    dx: float  # offset from the anchor character, points
    dy: float


@dataclass
class PageBreak:
    pass


@dataclass
class Footnote:
    blocks: list[Block]


@dataclass
class Paragraph:
    items: list[Inline] = field(default_factory=list)
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
            not isinstance(i, Run) for i in self.items
        )


@dataclass
class Cell:
    blocks: list[Block] = field(default_factory=list)
    row_span: int = 1
    col_span: int = 1
    merged: bool = False  # covered by another cell's span


@dataclass
class Table:
    rows: list[list[Cell]]
    header_rows: int = 0


@dataclass
class TextBox:
    blocks: list[Block]


Inline = Run | Image | Pencil | PageBreak | Footnote | Table | TextBox
Block = Paragraph | Table | TextBox


@dataclass
class Document:
    blocks: list[Block]
    language: str | None

    @property
    def paragraphs(self) -> list[Paragraph]:
        return [b for b in self.blocks if isinstance(b, Paragraph)]
