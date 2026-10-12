"""Low-level reader for Apple iWork (.pages) packages.

A .pages file is a zip holding Index/*.iwa archives: snappy-framed streams of
protobuf messages, each object prefixed by a TSP.ArchiveInfo header.
"""

from __future__ import annotations

import struct
import unicodedata
import zipfile
from pathlib import Path

import snappy
from google.protobuf.internal.decoder import _DecodeVarint32
from keynote_parser.versions.v14_5.generated import TSPArchiveMessages_pb2 as TSP


def _decompress(raw: bytes) -> bytes:
    out = []
    while raw:
        if raw[0] != 0:
            raise ValueError("IWA chunk does not start with 0x00")
        length = struct.unpack("<I", raw[1:4] + b"\0")[0]
        chunk = raw[4 : 4 + length]
        try:
            out.append(snappy.uncompress(chunk))
        except Exception:  # noqa: BLE001 - chunk was stored uncompressed
            out.append(chunk)
        raw = raw[4 + length :]
    return b"".join(out)


def _parse_iwa(raw: bytes) -> dict[int, list[tuple[int, bytes]]]:
    data = _decompress(raw)
    objects: dict[int, list[tuple[int, bytes]]] = {}
    pos = 0
    while pos < len(data):
        length, pos = _DecodeVarint32(data, pos)
        info = TSP.ArchiveInfo.FromString(data[pos : pos + length])
        pos += length
        messages = []
        for mi in info.message_infos:
            messages.append((mi.type, data[pos : pos + mi.length]))
            pos += mi.length
        objects[info.identifier] = messages
    return objects


class Package:
    """Decoded .pages package: object store plus embedded data files."""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        if self.path.is_dir():
            raise ValueError(
                f"{self.path} is a package folder (old Pages format); zip it or re-save it in Pages first"
            )
        self._zip = zipfile.ZipFile(self.path)
        # Names inside the zip come from macOS and are NFD-normalized.
        # Pages often omits the zip UTF-8 flag, so Python decodes them as cp437.
        self._names = {}
        for info in self._zip.infolist():
            name = info.filename
            if not info.flag_bits & 0x800:
                try:
                    name = name.encode("cp437").decode("utf-8")
                except UnicodeError:
                    pass
            self._names[unicodedata.normalize("NFC", name)] = info.filename
        self.objects: dict[int, list[tuple[int, bytes]]] = {}
        for name in self._zip.namelist():
            if name.endswith(".iwa"):
                self.objects.update(_parse_iwa(self._zip.read(name)))
        self._data_files = self._load_data_map()

    def type_of(self, ident: int) -> int | None:
        msgs = self.objects.get(ident)
        return msgs[0][0] if msgs else None

    def get(self, ident: int, cls):
        msgs = self.objects.get(ident)
        if not msgs:
            return None
        return cls.FromString(msgs[0][1])

    def of_type(self, type_id: int, cls):
        for ident, msgs in self.objects.items():
            if msgs and msgs[0][0] == type_id:
                yield ident, cls.FromString(msgs[0][1])

    def _load_data_map(self) -> dict[int, str]:
        files = {}
        for _, meta in self.of_type(11006, TSP.PackageMetadata):
            for d in meta.datas:
                name = d.file_name or d.preferred_file_name
                files[d.identifier] = unicodedata.normalize("NFC", name)
        return files

    def read_data(self, data_id: int) -> tuple[str, bytes] | None:
        name = self._data_files.get(data_id)
        if not name:
            return None
        real = self._names.get(f"Data/{name}")
        if real is None:
            return None
        return name, self._zip.read(real)
