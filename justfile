# list recipes
default:
    @just --list

# convert .pages files or folders: just convert *.pages
convert *args:
    uv run pages2word {{args}}

test:
    uv run pytest -q

lint:
    uv run ruff check --fix src tests
    uv run ruff format src tests

# build dist/Pages2Word.exe and the hand-off zip in release/
build:
    powershell -NoProfile -ExecutionPolicy Bypass -File packaging/build.ps1
    powershell -NoProfile -Command "New-Item -ItemType Directory -Force release | Out-Null; Compress-Archive -Force -Path dist/Pages2Word.exe,packaging/ImportPages.bas,packaging/NAVOD.txt -DestinationPath release/Pages2Word.zip"
