# pages2word

Converts Apple Pages (`.pages`) files to Word (`.docx`) by reading the Pages file format directly. The output uses Word's own styles, so it looks like a document you wrote in Word.

```bash
uv run pages2word *.pages -o out        # or: just convert *.pages -o out
```

What gets carried over:
- **Headings:** big or bold title lines become Word `Heading 1/2/3`, so the Navigation pane and Table of Contents work.
- **Lists:** these become real Word lists using Word's default bullets (•, o, ▪) and default indents. Bullets that were pasted in as plain text ("• …", "o …", "§ …") are turned back into real lists.
- **Text formatting:** bold, italic, underline, strikethrough, super/subscript, text colour and highlight. Images are carried over too.
- **Language:** each run keeps its proofing language (e.g. `cs-CZ`), so Word's spellchecker uses the right dictionary.

What is dropped on purpose: Pages fonts, font sizes and spacing. Word 365 defaults are used instead (Aptos 12 pt, A4).

Options:
- `-t my.docx`: use the styles from your own Word document instead of the defaults.
- `--blank-lines N`: N or more empty lines in a row become a page break (default 4; `0` keeps them).
- `-f`: overwrite existing output.

Limitations: only the main body text is converted. Tables, text boxes, shapes, footnotes and Apple Pencil drawings are skipped (a warning is printed for embedded objects).

## Windows app (Pages2Word.exe)

`just build` creates `dist/Pages2Word.exe` and `release/Pages2Word.zip`. The zip holds the exe, the Word macro and a Czech guide (`NAVOD.txt`).

- **Drag and drop:** drop `.pages` files on the exe and a `.docx` appears next to each one. Existing files are never overwritten.
- **The window** (double-click the exe) has these buttons:
  - Convert files
  - Edit styles
  - Reset styles
  - Link with Word: per-user file association, so double-clicking a `.pages` file opens it in Word. Right-click → convert also works.
- **Word button:** `packaging/ImportPages.bas` is a VBA macro (imported once into Normal). It inserts a converted `.pages` file into the current document, using that document's styles.
- **Custom styles:** "Edit styles" opens `%APPDATA%\Pages2Word\styly.docx` in Word. Whatever styles, margins and page size are saved there are used for every later conversion. Set `PAGES2WORD_STYLES` to point to a different file.
