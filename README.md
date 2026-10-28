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

## Restyling Word documents

Messy Word documents can be given the look of a document you styled yourself:

```bash
uv run pages2word extract-styles myNotes.docx -o "my styles.docx"   # once
uv run pages2word messy.docx -t "my styles.docx"                    # -> messy (converted).docx
```

- **Extracting styles** finds the styles your document actually uses for headings, list items and text, even custom ones such as "Nadpis 1_M". It also keeps your bullet list definition, page size and margins.
  - A font or size you set by hand on almost all text (e.g. every list item set to Calibri 12) is moved into the style.
  - Custom heading styles get an outline level, so the Navigation pane and table of contents pick them up.
  - The roles are stored as Word document variables, so you can edit the style document in Word afterwards. Any styled `.docx` also works directly as `-t`.
- **Restyling** keeps the text and its meaning: headings, lists and their levels, bold, italic, colours, tables, images and footnotes. Everything else is replaced by your styles:
  - Paragraph styles pasted in from elsewhere ("Výchozí", "Default", …) are replaced.
  - Fonts and sizes set by hand are removed, and so are justification and indents.
  - The many list definitions such documents collect become your single bullet list.
  - Bullets typed by hand ("• …") and lettered lists typed by hand ("a) …", "b) …") become real Word lists.
  - Numbered titles typed as plain text between numbered headings ("23. …" between "22. …" and "24. …") become headings.
- Old `.doc` files are converted through Word, so this needs Word to be installed.
- The input is never overwritten. If the output would replace it, the copy gets ` (converted)` in its name.

## PDF to Word

```bash
uv run pages2word notes.pdf -t "my styles.docx"
```

A PDF gives two documents:

- **`notes (PDF).docx`** looks like the PDF: the same fonts, sizes, colours, alignment, indents, line spacing, page size and margins. It is still built from real Word structure: headings, real lists and real tables.
- **`notes.docx`** has the same content in your styles (`-t`, or Word's defaults), the same way as when restyling a Word document.

A PDF only stores positioned characters, so the structure is rebuilt from the layout:
- Lines continuing at the same indent are joined into paragraphs.
- A bullet glyph or number followed by a gap starts a list item, and its indent gives the level.
- Bigger text becomes a heading, with one heading level per font size.
- Page numbers and repeating headers or footers are removed.
- Tables with text in several columns become Word tables with the PDF's column widths. Pictures are copied as images.

Limitations:
- Scanned PDFs (images without text) are not supported.
- Footnotes stay as ordinary text at the end of their page.
- A table that breaks across pages can become two tables.
- Page breaks are not copied, so the copy's pages may drift a little from the PDF's.

The PDF reading uses [pdfplumber](https://github.com/jsvine/pdfplumber) (MIT).

## Windows app (Pages2Word.exe)

`just build` creates `dist/Pages2Word.exe` and `release/Pages2Word.zip`. The zip holds the exe, the Word macro and a Czech guide (`NAVOD.txt`).

- **Drag and drop:** drop `.pages` files on the exe and a `.docx` appears next to each one. A Word document (`.docx`/`.doc`) dropped on it gets a restyled copy instead. A PDF gives the two documents described above. Existing files are never overwritten.
- **The window** (double-click the exe) has these buttons:
  - Convert files (`.pages`, `.docx`, `.doc`, `.pdf`)
  - Take styles from a document: extracts the styles of a Word document you pick (see above) into the custom styles file
  - Edit styles
  - Reset styles
  - Link with Word: per-user file association, so double-clicking a `.pages` file opens it in Word. Right-click → convert also works.
- **Word button:** `packaging/ImportPages.bas` is a VBA macro (imported once into Normal). It inserts a converted `.pages` file into the current document, using that document's styles.
- **Custom styles:** "Edit styles" opens `%APPDATA%\Pages2Word\styly.docx` in Word. Whatever styles, margins and page size are saved there are used for every later conversion. Set `PAGES2WORD_STYLES` to point to a different file.
