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

What else is converted:
- **Tables:** become Word tables (style "Table Grid"). Cell text keeps its formatting, header rows repeat on each page, and merged cells stay merged. Numbers, dates and checkboxes are shown as plain values.
- **Text boxes and shapes with text:** become a framed box at the same spot in the text.
- **Footnotes:** become real Word footnotes.
- **Apple Pencil drawings:** float next to the text they were drawn on, as pictures.
- **Objects pinned to a page rather than placed in the text:** Pages doesn't store which text they sit next to, so they are collected at the end of the document under "Objekty ze strany N".

Still skipped, with a warning: shapes without text, charts, video and equations. Headers and footers are also not converted.

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
