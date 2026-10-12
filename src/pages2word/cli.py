from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from .reader import PagesReader
from .writer import write_docx

log = logging.getLogger("pages2word")


def convert(
    src: Path, dst: Path, template: Path | None = None, blank_lines: int = 4
) -> Path:
    doc = PagesReader(src).read()
    write_docx(doc, dst, template, blank_lines)
    return dst


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="pages2word",
        description="Convert Apple Pages (.pages) files to Word (.docx) using Word's own styles and lists.",
    )
    parser.add_argument(
        "inputs", nargs="+", type=Path, help=".pages files or folders containing them"
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
        files.extend(sorted(path.glob("*.pages")) if path.is_dir() else [path])

    failed = 0
    for src in files:
        dst = (args.out_dir or src.parent) / (src.stem + ".docx")
        if dst.exists() and not args.force:
            dst = dst.with_stem(dst.stem + " (converted)")
            if dst.exists():
                log.error("%s exists, use --force to overwrite", dst)
                failed += 1
                continue
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            convert(src, dst, args.template, args.blank_lines)
            print(f"{src} -> {dst}")
        except Exception as exc:  # noqa: BLE001 - keep converting the other files
            log.error("%s: %s", src, exc)
            failed += 1
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
