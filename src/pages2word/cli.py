from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from .docx_reader import WORD_SUFFIXES, DocxReader
from .reader import PagesReader
from .writer import write_docx

log = logging.getLogger("pages2word")

SUFFIXES = (".pages", ".pdf", *WORD_SUFFIXES)


def pdf_layout_path(dst: Path) -> Path:
    """Where the look-alike copy of a PDF goes, next to the restyled `dst`."""
    return dst.with_name(f"{dst.stem} (PDF).docx")


def outputs(src: Path, dst: Path) -> list[Path]:
    """Every file converting `src` to `dst` writes."""
    if src.suffix.lower() == ".pdf":
        return [pdf_layout_path(dst), dst]
    return [dst]


def convert(
    src: Path, dst: Path, template: Path | None = None, blank_lines: int = 4
) -> Path:
    """Convert a .pages file, restyle a Word document, or rebuild a PDF, into `dst`.

    A PDF also gets a copy that keeps its own look (see `pdf_layout_path`)."""
    suffix = src.suffix.lower()
    if suffix == ".pdf":
        from .pdf_reader import PdfReader, strip_layout

        doc = PdfReader(src).read()
        write_docx(doc, pdf_layout_path(dst), None, 0, keep_layout=True)
        write_docx(strip_layout(doc), dst, template, 0)
        return dst
    if suffix in WORD_SUFFIXES:
        doc = DocxReader(src).read()
    else:
        doc = PagesReader(src).read()
    write_docx(doc, dst, template, blank_lines)
    return dst


def extract_main(argv: list[str]) -> int:
    from .styles import extract_styles

    parser = argparse.ArgumentParser(
        prog="pages2word extract-styles",
        description="Save the styles of a Word document as a style document for -t.",
    )
    parser.add_argument(
        "source", type=Path, help=".docx or .doc styled the way you like"
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="style document to write (default: '<source> styles.docx')",
    )
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    dst = args.output or args.source.with_name(f"{args.source.stem} styles.docx")
    if dst.resolve() == args.source.resolve():
        parser.error("output would overwrite the source document")
    try:
        extract_styles(args.source, dst)
    except Exception as exc:  # noqa: BLE001 - one-line error for the user
        log.error("%s: %s", args.source, exc)
        return 1
    print(f"{args.source} -> {dst}")
    return 0


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if argv[:1] == ["extract-styles"]:
        return extract_main(argv[1:])

    parser = argparse.ArgumentParser(
        prog="pages2word",
        description="Convert Apple Pages (.pages) files to Word (.docx) using Word's own styles "
        "and lists, or restyle Word documents (.docx/.doc) with a style document. "
        "A PDF gives two documents: '<name> (PDF).docx' looks like the PDF, "
        "'<name>.docx' uses the styles. "
        "'pages2word extract-styles DOC' saves a document's styles for -t.",
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        type=Path,
        help=".pages/.docx/.doc/.pdf files or folders containing them",
    )
    parser.add_argument(
        "-o", "--out-dir", type=Path, help="output folder (default: next to each input)"
    )
    parser.add_argument(
        "-t",
        "--template",
        type=Path,
        help="use styles from this .docx instead of Word 365 defaults",
    )
    parser.add_argument(
        "--blank-lines",
        type=int,
        default=4,
        metavar="N",
        help="N or more consecutive empty lines become a page break (0 = keep them; default 4)",
    )
    parser.add_argument(
        "-f", "--force", action="store_true", help="overwrite existing .docx files"
    )
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    files: list[Path] = []
    for path in args.inputs:
        if path.is_dir():
            files.extend(
                sorted(
                    f
                    for f in path.iterdir()
                    if f.suffix.lower() in SUFFIXES and not f.name.startswith("~$")
                )
            )
        else:
            files.append(path)

    failed = 0
    for src in files:
        dst = (args.out_dir or src.parent) / (src.stem + ".docx")
        # never overwrite a Word document with its own restyled copy
        exists = any(out.exists() for out in outputs(src, dst))
        if (exists and not args.force) or dst.resolve() == src.resolve():
            dst = dst.with_stem(dst.stem + " (converted)")
            if any(out.exists() for out in outputs(src, dst)) and not args.force:
                log.error("%s exists, use --force to overwrite", dst)
                failed += 1
                continue
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            convert(src, dst, args.template, args.blank_lines)
            print(f"{src} -> {', '.join(str(out) for out in outputs(src, dst))}")
        except Exception as exc:  # noqa: BLE001 - keep converting the other files
            log.error("%s: %s", src, exc)
            failed += 1
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
