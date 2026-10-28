"""Windows front end (Pages2Word.exe).

- Drop .pages files on the exe: each is converted to a .docx next to it.
  Word documents (.docx/.doc) dropped on it get a restyled copy instead.
- Double-click the exe: a small window with buttons.
- `--open FILE`: convert and open in Word (used by the file association).
- `--output OUT --quiet FILE`: convert to an exact path (used by the Word macro).
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
from pathlib import Path

from .cli import SUFFIXES, convert

APP_NAME = "Pages2Word"
PROG_ID = "Pages2Word.Document"
APP_DIR = Path(os.environ.get("APPDATA", Path.home())) / APP_NAME
STYLES_FILE = Path(os.environ.get("PAGES2WORD_STYLES") or APP_DIR / "styly.docx")


class _Collect(logging.Handler):
    def __init__(self):
        super().__init__(logging.WARNING)
        self.messages: list[str] = []

    def emit(self, record):
        self.messages.append(record.getMessage())


def style_template() -> Path | None:
    return STYLES_FILE if STYLES_FILE.exists() else None


def unique_output(src: Path) -> Path:
    dst = src.with_suffix(".docx")
    n = 2
    while dst.exists():
        dst = src.with_name(f"{src.stem} ({n}).docx")
        n += 1
    return dst


def convert_one(src: Path, dst: Path | None = None) -> tuple[Path, list[str]]:
    collector = _Collect()
    logging.getLogger("pages2word").addHandler(collector)
    try:
        dst = dst or unique_output(src)
        convert(src, dst, style_template())
        return dst, collector.messages
    finally:
        logging.getLogger("pages2word").removeHandler(collector)


# ---- dialogs ------------------------------------------------------------------


def _tk():
    import tkinter as tk

    root = tk.Tk()
    root.withdraw()
    return root


def info(text: str):
    from tkinter import messagebox

    _tk()
    messagebox.showinfo(APP_NAME, text)


def error(text: str):
    from tkinter import messagebox

    _tk()
    messagebox.showerror(APP_NAME, text)


def convert_files(
    files: list[Path], open_after: bool = False, quiet: bool = False
) -> int:
    done, failed, notes = [], [], []
    for src in files:
        if src.suffix.lower() not in SUFFIXES:
            failed.append(f"{src.name}: není soubor .pages ani dokument Wordu")
            continue
        try:
            dst, warnings = convert_one(src)
            done.append(dst)
            notes += [f"{src.name}: {w}" for w in warnings]
            if open_after:
                os.startfile(dst)
        except Exception as exc:  # noqa: BLE001 - report every failure to the user
            failed.append(f"{src.name}: {exc}")

    if not quiet:
        lines = []
        if done and not open_after:
            lines.append(
                "Hotovo, vytvořeno:\n" + "\n".join(f"• {d.name}" for d in done)
            )
            lines.append(f"\nSložka: {done[0].parent}")
        if notes:
            lines.append(
                "\nUpozornění (tyto části se nepřevedly):\n" + "\n".join(notes[:10])
            )
        if failed:
            error("\n\n".join(lines + ["Nepodařilo se převést:\n" + "\n".join(failed)]))
        elif lines:
            info("\n".join(lines))
    return 1 if failed else 0


# ---- Windows integration ----------------------------------------------------------


def _exe() -> str:
    return (
        sys.executable
        if getattr(sys, "frozen", False)
        else f'{sys.executable}" -m "pages2word.app'
    )


def install_integration():
    import winreg

    exe = _exe()
    classes = r"Software\Classes"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{classes}\.pages") as key:
        try:
            previous = winreg.QueryValueEx(key, "")[0]
        except OSError:
            previous = ""
        if previous and previous != PROG_ID:
            winreg.SetValueEx(key, f"{APP_NAME}.Previous", 0, winreg.REG_SZ, previous)
        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, PROG_ID)
    with winreg.CreateKey(
        winreg.HKEY_CURRENT_USER, rf"{classes}\.pages\OpenWithProgids"
    ) as key:
        winreg.SetValueEx(key, PROG_ID, 0, winreg.REG_NONE, b"")

    base = rf"{classes}\{PROG_ID}"
    values = {
        base: "Dokument Pages",
        rf"{base}\DefaultIcon": f'"{sys.executable}",0',
        rf"{base}\shell": "open",
        rf"{base}\shell\open": "Otevřít ve Wordu",
        rf"{base}\shell\open\command": f'"{exe}" --open "%1"',
        rf"{base}\shell\convert": "Převést do Wordu (.docx)",
        rf"{base}\shell\convert\command": f'"{exe}" "%1"',
    }
    for path, value in values.items():
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, path) as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, value)
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"Software\{APP_NAME}") as key:
        winreg.SetValueEx(key, "ExePath", 0, winreg.REG_SZ, sys.executable)
    _notify_shell()


def uninstall_integration():
    import winreg

    classes = r"Software\Classes"
    _delete_tree(winreg.HKEY_CURRENT_USER, rf"{classes}\{PROG_ID}")
    _delete_tree(winreg.HKEY_CURRENT_USER, rf"Software\{APP_NAME}")
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, rf"{classes}\.pages", 0, winreg.KEY_ALL_ACCESS
        ) as key:
            try:
                previous = winreg.QueryValueEx(key, f"{APP_NAME}.Previous")[0]
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, previous)
                winreg.DeleteValue(key, f"{APP_NAME}.Previous")
            except OSError:
                if winreg.QueryValueEx(key, "")[0] == PROG_ID:
                    winreg.DeleteValue(key, "")
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            rf"{classes}\.pages\OpenWithProgids",
            0,
            winreg.KEY_ALL_ACCESS,
        ) as key:
            winreg.DeleteValue(key, PROG_ID)
    except OSError:
        pass
    _notify_shell()


def integration_installed() -> bool:
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, rf"Software\{APP_NAME}") as key:
            return bool(winreg.QueryValueEx(key, "ExePath")[0])
    except OSError:
        return False


def _delete_tree(root, path: str):
    import winreg

    try:
        with winreg.OpenKey(root, path, 0, winreg.KEY_ALL_ACCESS) as key:
            while True:
                try:
                    child = winreg.EnumKey(key, 0)
                except OSError:
                    break
                _delete_tree(root, rf"{path}\{child}")
        winreg.DeleteKey(root, path)
    except FileNotFoundError:
        pass


def _notify_shell():
    import ctypes

    SHCNE_ASSOCCHANGED, SHCNF_IDLIST = 0x08000000, 0
    ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None)


# ---- styles ---------------------------------------------------------------------


def edit_styles():
    from .styles import make_style_template

    if not STYLES_FILE.exists():
        make_style_template(STYLES_FILE)
    os.startfile(STYLES_FILE)


def reset_styles():
    STYLES_FILE.unlink(missing_ok=True)


def styles_from_document(src: Path):
    from .styles import extract_styles

    extract_styles(src, STYLES_FILE)


# ---- main window ----------------------------------------------------------------


def window():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title(APP_NAME)
    root.resizable(False, False)
    frame = ttk.Frame(root, padding=20)
    frame.grid()

    ttk.Label(
        frame, text="Převod dokumentů Pages do Wordu", font=("Segoe UI", 13, "bold")
    ).grid(sticky="w")
    ttk.Label(
        frame,
        text="Tip: soubory .pages i .docx stačí přetáhnout myší na ikonu Pages2Word.exe.\n"
        "Dokumenty Wordu dostanou vaše styly (uloží se jako kopie).",
        foreground="#555",
    ).grid(sticky="w", pady=(2, 14))

    def pick():
        names = filedialog.askopenfilenames(
            title="Vyberte soubory Pages nebo Wordu",
            filetypes=[
                ("Pages a Word", "*.pages *.docx *.doc"),
                ("Dokumenty Pages", "*.pages"),
                ("Dokumenty Wordu", "*.docx *.doc"),
            ],
        )
        if names:
            convert_files([Path(n) for n in names])

    def styles():
        edit_styles()
        messagebox.showinfo(
            APP_NAME,
            "Ve Wordu se otevřel vzor stylů.\n\n"
            "Změňte styly (Domů → Styly → pravým tlačítkem → Upravit), "
            "dokument uložte a zavřete. Další převody už použijí vaše styly.",
        )

    def take_styles():
        name = filedialog.askopenfilename(
            title="Vyberte dokument s vašimi styly",
            filetypes=[("Dokumenty Wordu", "*.docx *.doc")],
        )
        if not name:
            return
        if style_template() and not messagebox.askyesno(
            APP_NAME, "Nahradit současné vlastní styly styly z tohoto dokumentu?"
        ):
            return
        try:
            styles_from_document(Path(name))
        except Exception as exc:  # noqa: BLE001 - report to the user
            messagebox.showerror(APP_NAME, f"{Path(name).name}: {exc}")
            return
        refresh()
        messagebox.showinfo(
            APP_NAME,
            "Styly převzaty. Nadpisy, odrážky a písmo z tohoto dokumentu "
            "se teď použijí pro všechny převody.\n\n"
            "Pro kontrolu je můžete zobrazit tlačítkem Upravit styly.",
        )

    def reset():
        if messagebox.askyesno(
            APP_NAME, "Vrátit výchozí styly Wordu? Vaše úpravy stylů se smažou."
        ):
            reset_styles()
            refresh()

    def toggle():
        if integration_installed():
            uninstall_integration()
            messagebox.showinfo(APP_NAME, "Propojení s Wordem bylo zrušeno.")
        else:
            install_integration()
            messagebox.showinfo(
                APP_NAME,
                "Hotovo. Soubory .pages se teď otevřou dvojklikem přímo ve Wordu "
                "(převedený .docx se uloží vedle původního souboru).\n\n"
                "Pravým tlačítkem na soubor → Převést do Wordu jen vytvoří .docx.\n\n"
                "Pokud Windows nabídne výběr aplikace, zvolte Pages2Word.",
            )
        refresh()

    buttons = ttk.Frame(frame)
    buttons.grid(sticky="ew")
    ttk.Button(buttons, text="Převést soubory…", command=pick, width=34).grid(
        sticky="ew", pady=3
    )
    ttk.Button(buttons, text="Převzít styly z dokumentu…", command=take_styles).grid(
        sticky="ew", pady=3
    )
    ttk.Button(buttons, text="Upravit styly…", command=styles).grid(sticky="ew", pady=3)
    reset_btn = ttk.Button(buttons, text="Obnovit výchozí styly", command=reset)
    reset_btn.grid(sticky="ew", pady=3)
    link_btn = ttk.Button(buttons, command=toggle)
    link_btn.grid(sticky="ew", pady=3)
    status = ttk.Label(frame, foreground="#555")
    status.grid(sticky="w", pady=(12, 0))

    def refresh():
        link_btn.configure(
            text="Zrušit propojení s Wordem"
            if integration_installed()
            else "Propojit s Wordem"
        )
        reset_btn.state(["!disabled"] if style_template() else ["disabled"])
        status.configure(
            text="Styly: vlastní" if style_template() else "Styly: výchozí Word"
        )

    refresh()
    root.mainloop()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog=APP_NAME)
    parser.add_argument("files", nargs="*", type=Path)
    parser.add_argument("--open", action="store_true", help="open the result in Word")
    parser.add_argument("--quiet", action="store_true", help="no dialogs")
    parser.add_argument("--output", type=Path, help="exact output path (single input)")
    parser.add_argument("--install", action="store_true")
    parser.add_argument("--uninstall", action="store_true")
    args = parser.parse_args(argv)

    if args.install:
        install_integration()
        return 0
    if args.uninstall:
        uninstall_integration()
        return 0
    if args.output:
        if len(args.files) != 1:
            parser.error("--output needs exactly one input file")
        try:
            args.output.unlink(missing_ok=True)
            convert_one(args.files[0], args.output)
            return 0
        except Exception as exc:  # noqa: BLE001 - surfaced to the caller
            if not args.quiet:
                error(f"{args.files[0].name}: {exc}")
            return 1
    if args.files:
        return convert_files(args.files, open_after=args.open, quiet=args.quiet)
    window()
    return 0


if __name__ == "__main__":
    sys.exit(main())
