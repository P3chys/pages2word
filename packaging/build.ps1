# Builds dist\Pages2Word.exe (one file, no console window).
$ErrorActionPreference = "Stop"
Set-Location (Split-Path $PSScriptRoot)
$versions = uv run python -c "import os, keynote_parser; print(os.path.join(os.path.dirname(keynote_parser.__file__), 'versions'))"
uv run pyinstaller --noconfirm --onefile --windowed --name Pages2Word `
  --collect-submodules keynote_parser `
  --add-data "$versions;keynote_parser/versions" `
  --collect-all cramjam `
  --distpath dist --workpath build --specpath build packaging/entry.py
