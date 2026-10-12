"""Write the semantic model into a .docx that uses Word's own styles and lists."""

from __future__ import annotations

import io
import logging
from pathlib import Path

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor

from .reader import Document, Image, PageBreak, Paragraph, Run

log = logging.getLogger(__name__)

W_NS = nsdecls("w")

# Word 365 defaults: Aptos 12 pt, 8 pt after, 1.15 lines; headings in accent blue.
STYLES_XML = f"""
<w:styles {W_NS}>
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
      <w:kern w:val="2"/><w:sz w:val="24"/><w:szCs w:val="24"/>
      <w:lang w:val="{{lang}}" w:eastAsia="en-US" w:bidi="ar-SA"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="278" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
  <w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
  {{headings}}
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/>
    <w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr>
    <w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/>
    <w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
</w:styles>
"""

# (size half-points, before, after, major font?, italic?)
HEADING_SPECS = {
    1: (40, 360, 80, True, False),
    2: (32, 160, 80, True, False),
    3: (28, 160, 80, False, False),
    4: (24, 80, 40, False, True),
    5: (24, 80, 40, False, False),
    6: (24, 40, 0, False, True),
}


def _heading_style(level: int) -> str:
    size, before, after, major, italic = HEADING_SPECS[min(level, 6)]
    theme = "major" if major else "minor"
    return f"""
  <w:style w:type="paragraph" w:styleId="Heading{level}"><w:name w:val="heading {level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="{before}" w:after="{after}"/><w:outlineLvl w:val="{level - 1}"/></w:pPr>
    <w:rPr><w:rFonts w:asciiTheme="{theme}HAnsi" w:eastAsiaTheme="{theme}EastAsia" w:hAnsiTheme="{theme}HAnsi" w:cstheme="{theme}Bidi"/>
      {"<w:i/><w:iCs/>" if italic else ""}<w:color w:val="0F4761" w:themeColor="accent1" w:themeShade="BF"/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr></w:style>"""


# Word's default bullet glyphs per level: filled dot, "o", filled square.
BULLETS = [("", "Symbol"), ("o", "Courier New"), ("", "Wingdings")]


def _bullet_abstract(abstract_id: int) -> str:
    levels = []
    for lvl in range(9):
        glyph, font = BULLETS[lvl % 3]
        left = 720 * (lvl + 1)
        levels.append(
            f'<w:lvl w:ilvl="{lvl}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="{glyph}"/>'
            f'<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{left}" w:hanging="360"/></w:pPr>'
            f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:hint="default"/></w:rPr></w:lvl>'
        )
    return (
        f'<w:abstractNum {W_NS} w:abstractNumId="{abstract_id}"><w:multiLevelType w:val="hybridMultilevel"/>'
        + "".join(levels)
        + "</w:abstractNum>"
    )


def _number_abstract(abstract_id: int, formats: dict[int, tuple[str, str]]) -> str:
    """formats: level -> (numFmt, suffix) taken from Pages; other levels use Word's 1. a. i. cycle."""
    cycle = ["decimal", "lowerLetter", "lowerRoman"]
    levels = []
    for lvl in range(9):
        fmt, suffix = formats.get(lvl, (cycle[lvl % 3], "."))
        text = f"({'%'}{lvl + 1})" if suffix == "()" else f"%{lvl + 1}{suffix}"
        roman = fmt == "lowerRoman"
        left = 720 * (lvl + 1)
        levels.append(
            f'<w:lvl w:ilvl="{lvl}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="{text}"/>'
            f'<w:lvlJc w:val="{"right" if roman else "left"}"/>'
            f'<w:pPr><w:ind w:left="{left}" w:hanging="{180 if roman else 360}"/></w:pPr></w:lvl>'
        )
    return (
        f'<w:abstractNum {W_NS} w:abstractNumId="{abstract_id}"><w:multiLevelType w:val="hybridMultilevel"/>'
        + "".join(levels)
        + "</w:abstractNum>"
    )


class _Numbering:
    def __init__(self, document):
        self.root = document.part.numbering_part.element
        for child in list(self.root):
            self.root.remove(child)
        self._abstract_count = 0
        self._num_count = 0
        self.bullet_num = self._num(self._abstract(_bullet_abstract))

    def _abstract(self, factory, *args) -> int:
        aid = self._abstract_count
        self._abstract_count += 1
        el = parse_xml(factory(aid, *args))
        # abstractNum elements must precede num elements
        nums = self.root.findall(qn("w:num"))
        if nums:
            nums[0].addprevious(el)
        else:
            self.root.append(el)
        return aid

    def _num(self, abstract_id: int, restart: bool = False) -> int:
        self._num_count += 1
        nid = self._num_count
        override = (
            '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>'
            if restart
            else ""
        )
        self.root.append(
            parse_xml(
                f'<w:num {W_NS} w:numId="{nid}"><w:abstractNumId w:val="{abstract_id}"/>{override}</w:num>'
            )
        )
        return nid

    def new_numbered_list(self, formats: dict[int, tuple[str, str]]) -> int:
        return self._num(self._abstract(_number_abstract, formats), restart=True)


def _patch_theme_fonts(document):
    """Swap the template's Calibri/Cambria theme fonts for Word 365's Aptos."""
    for rel in document.part.rels.values():
        if rel.reltype.endswith("/theme"):
            part = rel.target_part
            xml = part.blob.decode("utf-8")
            xml = xml.replace('typeface="Cambria"', 'typeface="Aptos Display"').replace(
                'typeface="Calibri"', 'typeface="Aptos"'
            )
            part._blob = xml.encode("utf-8")


def _new_document(language: str, template: Path | None):
    if template:
        document = docx.Document(str(template))
        body = document.element.body
        for child in list(body):
            if child.tag != qn("w:sectPr"):
                body.remove(child)
        return document

    document = docx.Document()
    headings = "".join(_heading_style(n) for n in range(1, 7))
    styles = parse_xml(
        STYLES_XML.replace("{lang}", language).replace("{headings}", headings)
    )
    root = document.styles.element
    latent = root.find(qn("w:latentStyles"))
    for child in list(root):
        root.remove(child)
    root.append(styles.find(qn("w:docDefaults")))
    if latent is not None:
        root.append(latent)
    for style in styles.findall(qn("w:style")):
        root.append(style)
    _patch_theme_fonts(document)

    # A4 with Word's default 2.5 cm margins
    section = document.sections[0]
    section.page_width, section.page_height = docx.shared.Mm(210), docx.shared.Mm(297)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, side, docx.shared.Cm(2.5))
    return document


def _set_lang(run, lang: str):
    rpr = run._r.get_or_add_rPr()
    el = rpr.find(qn("w:lang"))
    if el is None:
        el = parse_xml(f"<w:lang {W_NS}/>")
        rpr.append(el)
    el.set(qn("w:val"), lang)


def _set_shading(run, color: str):
    rpr = run._r.get_or_add_rPr()
    rpr.append(
        parse_xml(f'<w:shd {W_NS} w:val="clear" w:color="auto" w:fill="{color}"/>')
    )


def _set_numbering(paragraph, num_id: int, level: int):
    ppr = paragraph._p.get_or_add_pPr()
    ppr.append(
        parse_xml(
            f'<w:numPr {W_NS}><w:ilvl w:val="{level}"/><w:numId w:val="{num_id}"/></w:numPr>'
        )
    )


ALIGN = {
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


class WordWriter:
    def __init__(
        self,
        doc: Document,
        template: Path | None = None,
        empty_lines_to_page_break: int = 4,
    ):
        self.doc = doc
        self.language = doc.language or "cs-CZ"
        self.document = _new_document(self.language, template)
        self.numbering = _Numbering(self.document)
        self.empty_limit = empty_lines_to_page_break
        # With a user template its heading styles decide colours, not Pages.
        self.heading_colors = template is None
        section = self.document.sections[0]
        self.max_image_width = (
            section.page_width - section.left_margin - section.right_margin
        ) / 12700  # pt

    def write(self, out: Path):
        paragraphs = self._collapse_empty_runs(self.doc.paragraphs)
        number_list: int | None = None
        number_formats: dict[int, tuple[str, str]] = {}

        for p in paragraphs:
            if p.list_kind == "number":
                if number_list is None or p.restart:
                    number_formats = {}
                    number_list = None
                number_formats.setdefault(p.level, (p.number_format, p.number_suffix))
                if number_list is None:
                    number_list = self.numbering.new_numbered_list(number_formats)
            self._paragraph(p, number_list)

        self.document.save(str(out))

    def _collapse_empty_runs(self, paragraphs: list[Paragraph]) -> list[Paragraph]:
        """Long runs of blank lines were used to push text to a new page; use a page break instead."""
        if not self.empty_limit:
            return paragraphs
        out: list[Paragraph] = []
        blanks: list[Paragraph] = []
        for p in paragraphs:
            if (
                p.is_empty
                and not p.list_kind
                and not any(isinstance(i, PageBreak) for i in p.items)
            ):
                blanks.append(p)
                continue
            if len(blanks) >= self.empty_limit and out:
                p.page_break_before = True
            else:
                out.extend(blanks)
            blanks = []
            out.append(p)
        if len(blanks) < self.empty_limit:
            out.extend(blanks)
        return out

    def _paragraph(self, p: Paragraph, number_list: int | None):
        para = self.document.add_paragraph()
        if p.heading is not None:
            self._style(
                para, "Title" if p.heading == 0 else f"Heading {min(p.heading, 6)}"
            )
        elif p.list_kind:
            self._style(para, "List Paragraph")
            _set_numbering(
                para,
                self.numbering.bullet_num if p.list_kind == "bullet" else number_list,
                p.level,
            )
        elif p.left_indent > 0:
            para.paragraph_format.left_indent = Pt(p.left_indent)

        if p.align in ALIGN:
            para.alignment = ALIGN[p.align]
        if p.page_break_before:
            para.paragraph_format.page_break_before = True

        for item in p.items:
            if isinstance(item, Run):
                self._run(para, item, heading=p.heading is not None)
            elif isinstance(item, Image):
                self._image(para, item)
            elif isinstance(item, PageBreak):
                para.add_run().add_break(WD_BREAK.PAGE)

    def _style(self, para, name: str):
        try:
            para.style = self.document.styles[name]
        except KeyError:
            log.warning("Style %r missing from template, using Normal", name)

    def _run(self, para, r: Run, heading: bool):
        run = para.add_run(r.text)
        if r.bold and not heading:
            run.bold = True
        if r.italic:
            run.italic = True
        if r.underline:
            run.underline = True
        if r.strike:
            run.font.strike = True
        if r.caps:
            run.font.all_caps = True
        if r.script == "super":
            run.font.superscript = True
        elif r.script == "sub":
            run.font.subscript = True
        if r.color and (self.heading_colors or not heading):
            run.font.color.rgb = RGBColor.from_string(r.color)
        if r.highlight:
            _set_shading(run, r.highlight)
        if r.lang and r.lang != self.language:
            _set_lang(run, r.lang)

    def _image(self, para, img: Image):
        width, height = img.width, img.height
        if width > self.max_image_width:
            height *= self.max_image_width / width
            width = self.max_image_width
        try:
            para.add_run().add_picture(
                io.BytesIO(img.data), width=Pt(width), height=Pt(height)
            )
        except Exception as exc:  # noqa: BLE001 - unsupported format (e.g. HEIC, PDF)
            log.warning("Could not embed image %s: %s", img.filename, exc)


def write_docx(
    doc: Document,
    out: Path,
    template: Path | None = None,
    empty_lines_to_page_break: int = 4,
):
    WordWriter(doc, template, empty_lines_to_page_break).write(out)
