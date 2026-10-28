"""The editable style template: a sample .docx showing every style the converter uses.

The user restyles it in Word (Home > Styles > right-click > Modify) and saves it;
conversions then take their styles, page size and margins from it. It can also be
extracted from a document the user already styled the way they like.
"""

from __future__ import annotations

from pathlib import Path

from .docx_reader import open_word_file
from .model import Cell, Document, Footnote, Paragraph, Run, Table
from .writer import write_docx


def _p(text: str, **kw) -> Paragraph:
    return Paragraph(items=[Run(text, lang="cs-CZ")], **kw)


SAMPLE = [
    _p("Vzor stylů pro převod z Pages", heading=0),
    _p(
        "Tento dokument určuje vzhled všech převedených souborů. Změňte styly přímo ve Wordu: "
        "karta Domů → Styly → klikněte pravým tlačítkem na styl → Upravit. "
        "Pak dokument uložte a zavřete. Text zde se nikam nepřenáší, důležité jsou jen styly, "
        "velikost stránky a okraje."
    ),
    _p("Nadpis 1 – název kapitoly", heading=1),
    _p("Běžný text (styl Normální). Takhle bude vypadat většina poznámek."),
    _p("Nadpis 2 – oddíl", heading=2),
    _p("Odrážky používají styl Odstavec se seznamem:"),
    _p("První úroveň odrážky", list_kind="bullet", level=0),
    _p("Druhá úroveň odrážky", list_kind="bullet", level=1),
    _p("Třetí úroveň odrážky", list_kind="bullet", level=2),
    _p("Nadpis 3 – pododdíl", heading=3),
    _p("Číslovaný seznam", list_kind="number", level=0, restart=True),
    _p("Druhá položka", list_kind="number", level=0),
    _p(
        "Tabulky používají styl tabulky Mřížka tabulky (Návrh tabulky → Styly tabulky):"
    ),
    Table(
        rows=[
            [Cell([_p("Záhlaví")]), Cell([_p("Druhý sloupec")])],
            [Cell([_p("Buňka")]), Cell([_p("Buňka")])],
        ],
        header_rows=1,
    ),
    Paragraph(
        items=[
            Run(
                "Poznámky pod čarou používají styly Text poznámky pod čarou a Znak poznámky pod čarou.",
                lang="cs-CZ",
            ),
            Footnote([_p("Ukázka poznámky pod čarou.")]),
        ]
    ),
]


def make_style_template(path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    write_docx(Document(SAMPLE, "cs-CZ"), path, empty_lines_to_page_break=0)
    return path


def extract_styles(src: Path, dst: Path) -> Path:
    """Style document from `src`: its styles, list bullets, page setup, and which
    styles it uses for headings, lists and text; the content is replaced by a sample."""
    dst.parent.mkdir(parents=True, exist_ok=True)
    with open_word_file(src) as path:
        write_docx(Document(SAMPLE, "cs-CZ"), dst, path, empty_lines_to_page_break=0)
    return dst
