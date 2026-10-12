"""Decode iWork tables (TST.TableModelArchive) into the semantic model.

Cell contents live in per-row byte buffers inside tiles. Two layouts exist:
storage version 5 (current) and version 4 ("pre-BNC", Pages 2013–2016).
Byte layouts follow numbers-parser (MIT, github.com/masaccio/numbers-parser).
"""

from __future__ import annotations

import logging
import struct
from array import array
from collections.abc import Callable
from datetime import datetime, timedelta
from decimal import Decimal

from keynote_parser.versions.v14_5.generated import TSTArchives_pb2 as TST

from .iwa import Package
from .model import Block, Cell, Paragraph, Run, Table

log = logging.getLogger(__name__)

EPOCH = datetime(2001, 1, 1)
DEFAULT_TILE_SIZE = 256
DECIMAL128_BIAS = 0x1820

TEXT, NUMBER, DATE, BOOL, DURATION, ERROR, RICH, CURRENCY = 3, 2, 5, 6, 7, 8, 9, 10

# v5 optional fields in buffer order: (flag, byte size, name)
_V5_FIELDS = [
    (0x1, 16, "decimal"), (0x2, 8, "double"), (0x4, 8, "seconds"), (0x8, 4, "string"),
    (0x10, 4, "rich"), (0x20, 4, None), (0x40, 4, None), (0x80, 4, None), (0x100, 4, None),
    (0x200, 4, None), (0x400, 4, None), (0x800, 4, None), (0x1000, 4, None), (0x2000, 4, None),
    (0x4000, 4, None), (0x8000, 4, None), (0x10000, 4, None), (0x20000, 4, None), (0x40000, 4, None),
]  # fmt: skip


def _decimal128(b: bytes) -> Decimal:
    exp = (((b[15] & 0x7F) << 7) | (b[14] >> 1)) - DECIMAL128_BIAS
    mantissa = b[14] & 1
    for i in range(13, -1, -1):
        mantissa = mantissa * 256 + b[i]
    value = Decimal(mantissa).scaleb(exp)
    return -value if b[15] & 0x80 else value


def decode_cell(buf: bytes) -> dict:
    """Return the cell type and raw values stored in one cell buffer."""
    version, cell_type = buf[0], buf[1]
    values: dict = {"type": cell_type}
    if version >= 5:
        flags = struct.unpack_from("<I", buf, 8)[0]
        offset = 12
        for flag, size, name in _V5_FIELDS:
            if flags & flag:
                if name == "decimal":
                    values[name] = _decimal128(buf[offset : offset + 16])
                elif name in ("double", "seconds"):
                    values[name] = struct.unpack_from("<d", buf, offset)[0]
                elif name:
                    values[name] = struct.unpack_from("<i", buf, offset)[0]
                offset += size
    else:
        flags = struct.unpack_from("<I", buf, 4)[0]
        offset = 12 + (flags & 0x0D8E).bit_count() * 4
        if flags & 0x200:
            values["rich"] = struct.unpack_from("<i", buf, offset)[0]
            offset += 4
        offset += (flags & 0x3000).bit_count() * 4
        if flags & 0x10:
            values["string"] = struct.unpack_from("<i", buf, offset)[0]
            offset += 4
        if flags & 0x20:
            values["double"] = struct.unpack_from("<d", buf, offset)[0]
            offset += 8
        if flags & 0x40:
            values["seconds"] = struct.unpack_from("<d", buf, offset)[0]
    return values


def split_row(
    buffer: bytes, offsets: bytes, num_cols: int, wide: bool
) -> list[bytes | None]:
    """Cut a row's storage buffer into per-column cell buffers (None = empty cell)."""
    offs = array("h", offsets).tolist()
    if wide:
        offs = [o * 4 for o in offs]
    cells: list[bytes | None] = []
    for col in range(min(num_cols, len(offs))):
        start = offs[col]
        if start < 0:
            cells.append(None)
            continue
        end = next((o for o in offs[col + 1 :] if o >= 0), len(buffer))
        cells.append(buffer[start:end])
    return cells


def format_number(value: float | Decimal, language: str | None) -> str:
    text = f"{value:f}" if isinstance(value, Decimal) else f"{value:.10g}"
    if "." in text and "e" not in text:
        text = text.rstrip("0").rstrip(".")
    if language and language.split("-")[0] in (
        "cs",
        "sk",
        "de",
        "fr",
        "pl",
        "es",
        "it",
        "ru",
        "uk",
        "hu",
        "nl",
    ):
        text = text.replace(".", ",")
    return text


def cell_text(values: dict, strings: dict[int, str], language: str | None) -> str:
    t = values["type"]
    if t == TEXT:
        return strings.get(values.get("string", -1), "")
    if t in (NUMBER, CURRENCY):
        value = values.get("decimal", values.get("double"))
        return "" if value is None else format_number(value, language)
    if t == DATE and "seconds" in values:
        date = EPOCH + timedelta(seconds=values["seconds"])
        return date.strftime(
            "%d. %m. %Y" if (language or "").startswith("cs") else "%Y-%m-%d"
        )
    if t == BOOL:
        return "☑" if values.get("double", 0) > 0 else "☐"
    if t == DURATION:
        total = int(values.get("double", 0))
        return f"{total // 3600}:{total % 3600 // 60:02d}:{total % 60:02d}"
    if t == ERROR:
        return "#CHYBA"
    return ""


class TableReader:
    def __init__(
        self,
        pkg: Package,
        read_storage: Callable[[int], list[Block]],
        language: str | None,
    ):
        self.pkg = pkg
        self.read_storage = read_storage  # storage id -> blocks
        self.language = language

    def read(self, info_id: int) -> Table | None:
        info = self.pkg.get(info_id, TST.TableInfoArchive)
        model = (
            self.pkg.get(info.tableModel.identifier, TST.TableModelArchive)
            if info
            else None
        )
        if model is None:
            return None
        rows, cols = model.number_of_rows, model.number_of_columns
        grid = [[Cell() for _ in range(cols)] for _ in range(rows)]
        ds = model.base_data_store
        strings = self._list(ds.stringTable.identifier, lambda e: e.string)
        rich = self._list(
            ds.rich_text_table.identifier, lambda e: e.rich_text_payload.identifier
        )

        tile_size = ds.tiles.tile_size or DEFAULT_TILE_SIZE
        for tile_ref in ds.tiles.tiles:
            tile = self.pkg.get(tile_ref.tile.identifier, TST.Tile)
            if tile is None:
                continue
            for info_row in tile.rowInfos:
                row = tile_ref.tileid * tile_size + info_row.tile_row_index
                if row >= rows:
                    continue
                if info_row.cell_storage_buffer:
                    buffer, offsets = (
                        info_row.cell_storage_buffer,
                        info_row.cell_offsets,
                    )
                    wide = info_row.has_wide_offsets
                else:
                    buffer, offsets, wide = (
                        info_row.cell_storage_buffer_pre_bnc,
                        info_row.cell_offsets_pre_bnc,
                        False,
                    )
                for col, buf in enumerate(split_row(buffer, offsets, cols, wide)):
                    if buf:
                        grid[row][col].blocks = self._cell_blocks(buf, strings, rich)

        self._apply_merges(ds.merge_region_map.identifier, grid)
        return Table(grid, header_rows=min(model.number_of_header_rows, rows))

    def _list(self, ident: int, value) -> dict:
        table = self.pkg.get(ident, TST.TableDataList) if ident else None
        return {e.key: value(e) for e in table.entries} if table else {}

    def _cell_blocks(self, buf: bytes, strings: dict, rich: dict) -> list[Block]:
        try:
            values = decode_cell(buf)
        except struct.error, IndexError:
            log.warning("Unreadable table cell skipped")
            return []
        payload_id = rich.get(values.get("rich", -1))
        if payload_id:
            payload = self.pkg.get(payload_id, TST.RichTextPayloadArchive)
            if payload is not None:
                return self.read_storage(payload.storage.identifier)
        text = cell_text(values, strings, self.language)
        return [Paragraph(items=[Run(text)])] if text else []

    def _apply_merges(self, ident: int, grid: list[list[Cell]]):
        merges = self.pkg.get(ident, TST.MergeRegionMapArchive) if ident else None
        if merges is None:
            return
        for cell_range in merges.cell_range:
            col, row = (
                cell_range.origin.packedData >> 16,
                cell_range.origin.packedData & 0xFFFF,
            )
            ncols, nrows = (
                cell_range.size.packedData >> 16,
                cell_range.size.packedData & 0xFFFF,
            )
            if row >= len(grid) or col >= len(grid[0]):
                continue
            nrows, ncols = min(nrows, len(grid) - row), min(ncols, len(grid[0]) - col)
            grid[row][col].row_span, grid[row][col].col_span = nrows, ncols
            for r in range(row, row + nrows):
                for c in range(col, col + ncols):
                    if (r, c) != (row, col):
                        grid[r][c].merged = True
