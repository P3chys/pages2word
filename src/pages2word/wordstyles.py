"""Style profiles: which styles of a Word document play which role.

A styled .docx (for example the user's own notes) is analysed for the styles it
actually uses for headings, list items and body text. Formatting the user applied
by hand to almost all text of a style (e.g. every list item set to Calibri 12) is
folded into that style, so it survives when the text is re-created.

The roles are stored as Word document variables (`pages2word.heading1` = style
name, …) in settings.xml. Word keeps them when the user edits and saves the file,
so an extracted style document stays usable as a template.
"""

from __future__ import annotations

import copy
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

VAR_PREFIX = "pages2word."
ROLES = ("title", *(f"heading{n}" for n in range(1, 7)), "list", "body")

# "Heading 2", "Nadpis 2", custom "Nadpis 2_M", German/French/Polish built-ins.
_HEADING_NAME = re.compile(
    r"(?:heading|nadpis|überschrift|titre|nagłówek)\W*_?\s*(\d)", re.IGNORECASE
)
_TITLE_NAMES = {"title", "název", "titel"}

# Direct formatting shared by at least this share of a style's text becomes part of the style.
FOLD_SHARE = 0.6
FOLD_MIN_CHARS = 200

# Schema order of the children we insert into (unknown elements count as "last").
_ORDER = {
    "style": "name aliases basedOn next link autoRedefine hidden uiPriority semiHidden "
    "unhideWhenUsed qFormat locked personal personalCompose personalReply rsid pPr rPr",
    "pPr": "pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr "
    "suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct "
    "topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind "
    "contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment "
    "textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange",
    "rPr": "rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss "
    "imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs "
    "highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath",
    "settings": "writeProtection view zoom removePersonalInformation removeDateAndTime "
    "doNotDisplayPageBoundaries displayBackgroundShape printPostScriptOverText "
    "printFractionalCharacterWidth printFormsData embedTrueTypeFonts embedSystemFonts "
    "saveSubsetFonts saveFormsData mirrorMargins alignBordersAndEdges "
    "bordersDoNotSurroundHeader bordersDoNotSurroundFooter gutterAtTop hideSpellingErrors "
    "hideGrammaticalErrors activeWritingStyle proofState formsDesign attachedTemplate "
    "linkStyles stylePaneFormatFilter stylePaneSortMethod documentType mailMerge "
    "revisionView trackRevisions doNotTrackMoves doNotTrackFormatting documentProtection "
    "autoFormatOverride styleLockTheme styleLockQFSet defaultTabStop autoHyphenation "
    "consecutiveHyphenLimit hyphenationZone doNotHyphenateCaps showEnvelope "
    "summaryLength clickAndTypeStyle defaultTableStyle evenAndOddHeaders "
    "bookFoldRevPrinting bookFoldPrinting bookFoldPrintingSheets "
    "drawingGridHorizontalSpacing drawingGridVerticalSpacing "
    "displayHorizontalDrawingGridEvery displayVerticalDrawingGridEvery "
    "doNotUseMarginsForDrawingGridOrigin drawingGridHorizontalOrigin "
    "drawingGridVerticalOrigin doNotShadeFormData noPunctuationKerning "
    "characterSpacingControl printTwoOnOne strictFirstAndLastChars noLineBreaksAfter "
    "noLineBreaksBefore savePreviewPicture doNotValidateAgainstSchema "
    "saveInvalidXml ignoreMixedContent alwaysShowPlaceholderText doNotDemarcateInvalidXml "
    "saveXmlDataOnly useXSLTWhenSaving saveThroughXslt showXMLTags "
    "alwaysMergeEmptyNamespace updateFields hdrShapeDefaults footnotePr endnotePr "
    "compat docVars rsids",
}
_ORDER_INDEX = {k: {t: i for i, t in enumerate(v.split())} for k, v in _ORDER.items()}


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _child(parent, name: str, create: bool = False):
    """Find w:<name> under parent, creating it at its schema position if asked."""
    el = parent.find(qn(f"w:{name}"))
    if el is not None or not create:
        return el
    el = parse_xml(f"<w:{name} {nsdecls('w')}/>")
    order = _ORDER_INDEX.get(_local(parent.tag), {})
    rank = order.get(name, len(order))
    for sibling in parent:
        if order.get(_local(sibling.tag), len(order) + 1) > rank:
            sibling.addprevious(el)
            return el
    parent.append(el)
    return el


def _val(el) -> str | None:
    return None if el is None else el.get(qn("w:val"))


def _on(el) -> bool:
    """OOXML toggle property: present and not switched off."""
    return el is not None and _val(el) not in ("0", "false", "off", "none")


class Styles:
    """The style sheet of a document, with inheritance (w:basedOn) lookups."""

    def __init__(self, root):
        self.root = root
        self.by_id = {s.get(qn("w:styleId")): s for s in root.findall(qn("w:style"))}
        self.by_name = {self.name(sid): sid for sid in self.by_id}
        self.default_paragraph = next(
            (
                sid
                for sid, s in self.by_id.items()
                if s.get(qn("w:type")) == "paragraph"
                and s.get(qn("w:default")) in ("1", "true", "on")
            ),
            None,
        )
        self.defaults = root.find(qn("w:docDefaults"))

    def name(self, sid: str | None) -> str:
        el = self.by_id.get(sid)
        name = el.find(qn("w:name")) if el is not None else None
        return _val(name) if name is not None else (sid or "")

    def chain(self, sid: str | None):
        seen = set()
        while sid in self.by_id and sid not in seen:
            seen.add(sid)
            el = self.by_id[sid]
            yield el
            sid = _val(el.find(qn("w:basedOn")))

    def find(self, sid: str | None, *path: str):
        """First w:path element along the style's inheritance chain."""
        xpath = "/".join(qn(f"w:{p}") for p in path)
        for el in self.chain(sid):
            hit = el.find(xpath)
            if hit is not None:
                return hit
        return None

    def default_run(self, name: str):
        if self.defaults is None:
            return None
        return self.defaults.find(
            f"{qn('w:rPrDefault')}/{qn('w:rPr')}/{qn(f'w:{name}')}"
        )

    def heading_level(self, sid: str | None) -> int | None:
        """0 = title, 1..9 = heading level, None = not a heading style."""
        if sid is None:
            return None
        name = self.name(sid)
        if name.lower() in _TITLE_NAMES:
            return 0
        outline = self.find(sid, "pPr", "outlineLvl")
        if outline is not None and (_val(outline) or "9").isdigit():
            level = int(_val(outline))
            return level + 1 if level < 9 else None
        m = _HEADING_NAME.search(name)
        return int(m.group(1)) if m and int(m.group(1)) > 0 else None

    def paragraph_style(self, p) -> str | None:
        ppr = p.find(qn("w:pPr"))
        sid = _val(ppr.find(qn("w:pStyle"))) if ppr is not None else None
        return sid if sid in self.by_id else self.default_paragraph

    def num_pr(self, p, sid: str | None) -> tuple[str | None, int]:
        """(numId, level) of a paragraph, from its own pPr or its style; numId None = no list."""
        ppr = p.find(qn("w:pPr"))
        direct = ppr.find(qn("w:numPr")) if ppr is not None else None
        style = self.find(sid, "pPr", "numPr")
        num_id = ilvl = None
        for num_pr in (direct, style):
            if num_pr is None:
                continue
            num_id = num_id or _val(num_pr.find(qn("w:numId")))
            ilvl = ilvl or _val(num_pr.find(qn("w:ilvl")))
        if num_id in (None, "0"):
            return None, 0
        return num_id, int(ilvl) if ilvl and ilvl.isdigit() else 0


def _text(p) -> str:
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def _runs(p):
    """Text runs of a paragraph, including those inside hyperlinks and fields."""
    for r in p.iter(qn("w:r")):
        text = "".join(t.text or "" for t in r.findall(qn("w:t")))
        if text:
            yield r, text


@dataclass
class Profile:
    roles: dict[str, str] = field(default_factory=dict)  # role -> style name
    bullets: object | None = None  # w:abstractNum to copy for bullet lists

    def style(self, role: str, default: str) -> str:
        return self.roles.get(role) or default


# ---- learning a profile from a styled document -------------------------------------


def learn(document) -> Profile:
    """Work out the style roles of a document and fold hand formatting into its styles."""
    styles = Styles(document.styles.element)
    body = document.element.body
    usage: dict[str, Counter] = defaultdict(Counter)
    fonts: dict[str | None, Counter] = defaultdict(Counter)  # style -> font -> chars
    sizes: dict[str | None, Counter] = defaultdict(Counter)
    chars: Counter = Counter()
    bullet_lists: Counter = Counter()

    for p in body.iter(qn("w:p")):
        if not _text(p).strip():
            continue
        sid = styles.paragraph_style(p)
        level = styles.heading_level(sid)
        num_id, _ = styles.num_pr(p, sid)
        if level is not None and level <= 6:
            usage["title" if level == 0 else f"heading{level}"][sid] += 1
        elif num_id:
            usage["list"][sid] += 1
            bullet_lists[num_id] += 1
        else:
            usage["body"][sid] += 1
        for r, text in _runs(p):
            rpr = r.find(qn("w:rPr"))
            rfonts = rpr.find(qn("w:rFonts")) if rpr is not None else None
            font = rfonts.get(qn("w:ascii")) if rfonts is not None else None
            size = _val(rpr.find(qn("w:sz"))) if rpr is not None else None
            chars[sid] += len(text)
            fonts[sid][font] += len(text)
            sizes[sid][size] += len(text)
            if level is None:
                fonts["*"][font] += len(text)
                sizes["*"][size] += len(text)
                chars["*"] += len(text)

    def dominant(counter: Counter, total: int) -> str | None:
        value, count = counter.most_common(1)[0] if counter else (None, 0)
        if value and total >= FOLD_MIN_CHARS and count >= total * FOLD_SHARE:
            return value
        return None

    # Hand formatting of almost all body text goes into the default paragraph
    # style, so plain paragraphs and every style based on it match the lists.
    targets = [("*", styles.default_paragraph)] + [
        (sid, sid) for sid in chars if sid != "*"
    ]
    for source, sid in targets:
        if sid not in styles.by_id:
            continue
        font = dominant(fonts[source], chars[source])
        size = dominant(sizes[source], chars[source])
        if font and _style_font(styles, sid) != font:
            _set_style_font(styles.by_id[sid], font)
        if size and _style_size(styles, sid) != size:
            _set_style_size(styles.by_id[sid], size)

    roles = {
        role: styles.name(counter.most_common(1)[0][0])
        for role, counter in usage.items()
        if counter.most_common(1)[0][0] in styles.by_id
    }
    # Custom heading styles often lack an outline level; without it Word's
    # Navigation pane and table of contents ignore them.
    for n in range(1, 7):
        sid = styles.by_name.get(roles.get(f"heading{n}"))
        if sid and styles.find(sid, "pPr", "outlineLvl") is None:
            ppr = _child(styles.by_id[sid], "pPr", create=True)
            _child(ppr, "outlineLvl", create=True).set(qn("w:val"), str(n - 1))

    return Profile(roles, _dominant_bullets(document, bullet_lists))


def _style_font(styles: Styles, sid: str) -> str | None:
    rfonts = styles.find(sid, "rPr", "rFonts")
    return rfonts.get(qn("w:ascii")) if rfonts is not None else None


def _style_size(styles: Styles, sid: str) -> str | None:
    return _val(styles.find(sid, "rPr", "sz")) or _val(styles.default_run("sz"))


def _set_style_font(style, font: str):
    rfonts = _child(_child(style, "rPr", create=True), "rFonts", create=True)
    for attr in ("asciiTheme", "hAnsiTheme", "cstheme"):
        rfonts.attrib.pop(qn(f"w:{attr}"), None)
    for attr in ("ascii", "hAnsi", "cs"):
        rfonts.set(qn(f"w:{attr}"), font)


def _set_style_size(style, half_points: str):
    rpr = _child(style, "rPr", create=True)
    for name in ("sz", "szCs"):
        _child(rpr, name, create=True).set(qn("w:val"), half_points)


def _dominant_bullets(document, used: Counter):
    """The bullet list definition behind most list paragraphs, if any."""
    numbering = _numbering_root(document)
    if numbering is None:
        return None
    abstract_of = {
        n.get(qn("w:numId")): _val(n.find(qn("w:abstractNumId")))
        for n in numbering.findall(qn("w:num"))
    }
    abstracts = {
        a.get(qn("w:abstractNumId")): a for a in numbering.findall(qn("w:abstractNum"))
    }
    counts: Counter = Counter()
    for num_id, n in used.items():
        abstract = abstracts.get(abstract_of.get(num_id))
        first = abstract.find(qn("w:lvl")) if abstract is not None else None
        if first is not None and _val(first.find(qn("w:numFmt"))) == "bullet":
            counts[abstract.get(qn("w:abstractNumId"))] += n
    if not counts:
        return None
    return copy.deepcopy(abstracts[counts.most_common(1)[0][0]])


def _numbering_root(document):
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    for rel in document.part.rels.values():
        if rel.reltype == RT.NUMBERING and not rel.is_external:
            return rel.target_part.element
    return None


# ---- storing the profile in the document -------------------------------------------


def load_profile(document) -> Profile:
    """The profile stored in a template, or one learned from how it is styled."""
    settings = document.settings.element
    doc_vars = settings.find(qn("w:docVars"))
    stored = {}
    if doc_vars is not None:
        for var in doc_vars.findall(qn("w:docVar")):
            name = var.get(qn("w:name")) or ""
            if name.startswith(VAR_PREFIX):
                stored[name.removeprefix(VAR_PREFIX)] = var.get(qn("w:val"))

    if stored:
        profile = Profile({k: v for k, v in stored.items() if k in ROLES})
        bullets_id = stored.get("bullets")
        numbering = _numbering_root(document)
        if bullets_id is not None and numbering is not None:
            for a in numbering.findall(qn("w:abstractNum")):
                if a.get(qn("w:abstractNumId")) == bullets_id:
                    profile.bullets = copy.deepcopy(a)
    else:
        profile = learn(document)

    # Fill roles the template has no style for with the next level up.
    names = Styles(document.styles.element).by_name
    if "title" not in profile.roles and "Title" not in names:
        profile.roles["title"] = profile.roles.get("heading1", "heading 1")
    for n in range(2, 7):
        role = f"heading{n}"
        if role not in profile.roles and f"heading {n}" not in names:
            profile.roles[role] = profile.roles.get(
                f"heading{n - 1}", f"heading {n - 1}"
            )
    return profile


def save_profile(document, profile: Profile, bullets_id: int | None):
    settings = document.settings.element
    doc_vars = _child(settings, "docVars", create=True)
    for var in doc_vars.findall(qn("w:docVar")):
        if (var.get(qn("w:name")) or "").startswith(VAR_PREFIX):
            doc_vars.remove(var)
    values = dict(profile.roles)
    if bullets_id is not None:
        values["bullets"] = str(bullets_id)
    for key, value in values.items():
        var = parse_xml(f"<w:docVar {nsdecls('w')}/>")
        var.set(qn("w:name"), VAR_PREFIX + key)
        var.set(qn("w:val"), value)
        doc_vars.append(var)
