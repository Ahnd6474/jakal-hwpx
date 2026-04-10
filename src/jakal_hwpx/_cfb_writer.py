from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC
NOSTREAM = 0xFFFFFFFF

STGTY_EMPTY = 0
STGTY_STORAGE = 1
STGTY_STREAM = 2
STGTY_ROOT = 5

BLACK = 1

HEADER_SIGNATURE = bytes.fromhex("D0CF11E0A1B11AE1")
HEADER_CLSID = b"\x00" * 16
BYTE_ORDER = 0xFFFE
MINI_STREAM_CUTOFF = 4096
SECTOR_SIZE = 512
MINI_SECTOR_SIZE = 64
DIR_ENTRY_SIZE = 128


@dataclass
class _CfbEntry:
    name: str
    entry_type: int
    parent: int | None
    data: bytes = b""
    children: list[int] = field(default_factory=list)
    left: int = NOSTREAM
    right: int = NOSTREAM
    child: int = NOSTREAM
    start_sector: int = ENDOFCHAIN
    size: int = 0


def write_compound_file(path: str | Path, streams: dict[str, bytes]) -> Path:
    target_path = Path(path).expanduser().resolve()
    target_path.parent.mkdir(parents=True, exist_ok=True)

    entries = _build_entries(streams)
    mini_stream_data = _assign_mini_stream(entries)

    regular_objects: list[tuple[int, bytes]] = []
    for index, entry in enumerate(entries):
        if entry.entry_type == STGTY_STREAM and (entry.size == 0 or entry.size >= MINI_STREAM_CUTOFF):
            regular_objects.append((index, entry.data))
    if mini_stream_data:
        regular_objects.append((-1, mini_stream_data))

    directory_size = _directory_stream_size(len(entries))

    object_chains: dict[int, list[int]] = {}
    sectors: list[bytes] = []
    for entry_id, data in regular_objects:
        chain = _append_chain(sectors, data)
        object_chains[entry_id] = chain

    if mini_stream_data:
        mini_stream_chain = object_chains[-1]
        entries[0].start_sector = mini_stream_chain[0]
        entries[0].size = len(mini_stream_data)
    else:
        entries[0].start_sector = ENDOFCHAIN
        entries[0].size = 0

    for index, entry in enumerate(entries):
        if entry.entry_type != STGTY_STREAM or entry.size == 0 or entry.size < MINI_STREAM_CUTOFF:
            continue
        chain = object_chains[index]
        entry.start_sector = chain[0]

    mini_sector_count = len(mini_stream_data) // MINI_SECTOR_SIZE
    mini_fat_entries = _build_mini_fat(entries, mini_sector_count)
    mini_fat_bytes = b"".join(value.to_bytes(4, "little") for value in mini_fat_entries)
    mini_fat_chain = _append_chain(sectors, mini_fat_bytes) if mini_fat_entries else []

    directory_chain = _append_chain(sectors, b"\x00" * directory_size)
    directory_start_sector = directory_chain[0]

    data_sector_count = len(sectors)
    fat_sector_count = _compute_fat_sector_count(data_sector_count)
    fat_sector_ids = list(range(data_sector_count, data_sector_count + fat_sector_count))

    fat_entries = [FREESECT] * (data_sector_count + fat_sector_count)
    for chain in object_chains.values():
        _mark_chain(fat_entries, chain)
    if mini_fat_chain:
        _mark_chain(fat_entries, mini_fat_chain)
    _mark_chain(fat_entries, directory_chain)
    for fat_sector_id in fat_sector_ids:
        fat_entries[fat_sector_id] = FATSECT

    fat_bytes = b"".join(value.to_bytes(4, "little") for value in fat_entries)
    fat_bytes = fat_bytes.ljust(fat_sector_count * SECTOR_SIZE, b"\xFF")
    for index in range(fat_sector_count):
        start = index * SECTOR_SIZE
        end = start + SECTOR_SIZE
        sectors.append(fat_bytes[start:end])

    header = _build_header(
        directory_start_sector=directory_start_sector,
        fat_sector_ids=fat_sector_ids,
        mini_fat_chain=mini_fat_chain,
    )

    directory_bytes = _build_directory_stream(entries)
    _store_chain_payload(sectors, directory_chain, directory_bytes)

    target_path.write_bytes(header + b"".join(sectors))
    return target_path


def _append_chain(sectors: list[bytes], data: bytes) -> list[int]:
    if not data:
        return []
    padded = data + (b"\x00" * ((SECTOR_SIZE - (len(data) % SECTOR_SIZE)) % SECTOR_SIZE))
    start = len(sectors)
    for index in range(0, len(padded), SECTOR_SIZE):
        sectors.append(padded[index : index + SECTOR_SIZE])
    return list(range(start, len(sectors)))


def _store_chain_payload(sectors: list[bytes], chain: list[int], data: bytes) -> None:
    if not chain:
        return
    padded = data + (b"\x00" * ((SECTOR_SIZE - (len(data) % SECTOR_SIZE)) % SECTOR_SIZE))
    for chain_index, sector_id in enumerate(chain):
        start = chain_index * SECTOR_SIZE
        end = start + SECTOR_SIZE
        sectors[sector_id] = padded[start:end]


def _compute_fat_sector_count(data_sector_count: int) -> int:
    fat_sector_count = 1
    while True:
        needed = (data_sector_count + fat_sector_count + 127) // 128
        if needed == fat_sector_count:
            return fat_sector_count
        fat_sector_count = needed


def _mark_chain(fat_entries: list[int], chain: list[int]) -> None:
    if not chain:
        return
    for current, next_sector in zip(chain, chain[1:]):
        fat_entries[current] = next_sector
    fat_entries[chain[-1]] = ENDOFCHAIN


def _build_entries(streams: dict[str, bytes]) -> list[_CfbEntry]:
    entries: list[_CfbEntry] = [_CfbEntry(name="Root Entry", entry_type=STGTY_ROOT, parent=None)]
    index_by_path: dict[str, int] = {"": 0}

    def ensure_storage(storage_path: str) -> int:
        if storage_path in index_by_path:
            return index_by_path[storage_path]
        parent_path, _, name = storage_path.rpartition("/")
        parent_index = ensure_storage(parent_path)
        entry = _CfbEntry(name=name, entry_type=STGTY_STORAGE, parent=parent_index)
        entries.append(entry)
        entry_index = len(entries) - 1
        entries[parent_index].children.append(entry_index)
        index_by_path[storage_path] = entry_index
        return entry_index

    for stream_path, data in sorted(streams.items()):
        parent_path, _, name = stream_path.rpartition("/")
        parent_index = ensure_storage(parent_path)
        entry = _CfbEntry(name=name, entry_type=STGTY_STREAM, parent=parent_index, data=data, size=len(data))
        entries.append(entry)
        entry_index = len(entries) - 1
        entries[parent_index].children.append(entry_index)
        index_by_path[stream_path] = entry_index

    for entry in entries:
        if entry.entry_type in (STGTY_ROOT, STGTY_STORAGE):
            _assign_btree(entries, entry.children, lambda entry_id: entries[entry_id].name.casefold())
            entry.child = entry.children[0] if len(entry.children) == 1 else _btree_root(entry.children, entries)

    return entries


def _btree_root(children: list[int], entries: list[_CfbEntry]) -> int:
    ordered = sorted(children, key=lambda entry_id: entries[entry_id].name.casefold())
    return ordered[len(ordered) // 2]


def _assign_btree(entries: list[_CfbEntry], child_ids: list[int], key) -> None:
    ordered = sorted(child_ids, key=key)

    def build(ids: list[int]) -> int:
        if not ids:
            return NOSTREAM
        middle = len(ids) // 2
        entry_id = ids[middle]
        entries[entry_id].left = build(ids[:middle])
        entries[entry_id].right = build(ids[middle + 1 :])
        return entry_id

    root = build(ordered)
    if root == NOSTREAM:
        return
    for entry_id in ordered:
        entries[entry_id].child = entries[entry_id].child if entries[entry_id].children else NOSTREAM


def _assign_mini_stream(entries: list[_CfbEntry]) -> bytes:
    mini_stream_chunks: list[bytes] = []
    mini_sector_cursor = 0
    for entry in entries:
        if entry.entry_type != STGTY_STREAM:
            continue
        if entry.size == 0:
            entry.start_sector = ENDOFCHAIN
            continue
        if entry.size >= MINI_STREAM_CUTOFF:
            continue
        padded = entry.data + (b"\x00" * ((MINI_SECTOR_SIZE - (len(entry.data) % MINI_SECTOR_SIZE)) % MINI_SECTOR_SIZE))
        entry.start_sector = mini_sector_cursor
        mini_sector_cursor += len(padded) // MINI_SECTOR_SIZE
        mini_stream_chunks.append(padded)
    return b"".join(mini_stream_chunks)


def _build_mini_fat(entries: list[_CfbEntry], mini_sector_count: int) -> list[int]:
    if mini_sector_count == 0:
        return []
    entries_map = [FREESECT] * mini_sector_count
    for entry in entries:
        if entry.entry_type != STGTY_STREAM or entry.size == 0 or entry.size >= MINI_STREAM_CUTOFF:
            continue
        first = entry.start_sector
        count = (entry.size + MINI_SECTOR_SIZE - 1) // MINI_SECTOR_SIZE
        for offset in range(count - 1):
            entries_map[first + offset] = first + offset + 1
        entries_map[first + count - 1] = ENDOFCHAIN
    return entries_map


def _build_directory_stream(entries: list[_CfbEntry]) -> bytes:
    data = bytearray()
    for index, entry in enumerate(entries):
        data.extend(_build_directory_entry(entry))
    padding = (SECTOR_SIZE - (len(data) % SECTOR_SIZE)) % SECTOR_SIZE
    if padding:
        data.extend(b"\x00" * padding)
    return bytes(data)


def _directory_stream_size(entry_count: int) -> int:
    raw_size = entry_count * DIR_ENTRY_SIZE
    return raw_size + ((SECTOR_SIZE - (raw_size % SECTOR_SIZE)) % SECTOR_SIZE)


def _build_directory_entry(entry: _CfbEntry) -> bytes:
    name = entry.name[:31]
    encoded_name = name.encode("utf-16-le") + b"\x00\x00"
    encoded_name = encoded_name.ljust(64, b"\x00")
    name_length = min(len(name) + 1, 32) * 2

    data = bytearray(128)
    data[0:64] = encoded_name
    data[64:66] = name_length.to_bytes(2, "little")
    data[66] = entry.entry_type
    data[67] = BLACK
    data[68:72] = entry.left.to_bytes(4, "little", signed=False)
    data[72:76] = entry.right.to_bytes(4, "little", signed=False)
    data[76:80] = entry.child.to_bytes(4, "little", signed=False)
    data[80:96] = b"\x00" * 16
    data[96:100] = (0).to_bytes(4, "little")
    data[100:108] = (0).to_bytes(8, "little")
    data[108:116] = (0).to_bytes(8, "little")
    data[116:120] = entry.start_sector.to_bytes(4, "little", signed=False)
    data[120:128] = entry.size.to_bytes(8, "little", signed=False)
    return bytes(data)


def _build_header(*, directory_start_sector: int, fat_sector_ids: list[int], mini_fat_chain: list[int]) -> bytes:
    header = bytearray(512)
    header[0:8] = HEADER_SIGNATURE
    header[8:24] = HEADER_CLSID
    header[24:26] = (0x003E).to_bytes(2, "little")
    header[26:28] = (0x0003).to_bytes(2, "little")
    header[28:30] = BYTE_ORDER.to_bytes(2, "little")
    header[30:32] = (9).to_bytes(2, "little")
    header[32:34] = (6).to_bytes(2, "little")
    header[34:40] = b"\x00" * 6
    header[40:44] = (0).to_bytes(4, "little")
    header[44:48] = len(fat_sector_ids).to_bytes(4, "little")
    header[48:52] = directory_start_sector.to_bytes(4, "little")
    header[52:56] = (0).to_bytes(4, "little")
    header[56:60] = MINI_STREAM_CUTOFF.to_bytes(4, "little")
    if mini_fat_chain:
        header[60:64] = mini_fat_chain[0].to_bytes(4, "little")
        header[64:68] = len(mini_fat_chain).to_bytes(4, "little")
    else:
        header[60:64] = ENDOFCHAIN.to_bytes(4, "little")
        header[64:68] = (0).to_bytes(4, "little")
    header[68:72] = ENDOFCHAIN.to_bytes(4, "little")
    header[72:76] = (0).to_bytes(4, "little")
    difat_offset = 76
    for index in range(109):
        value = fat_sector_ids[index] if index < len(fat_sector_ids) else FREESECT
        header[difat_offset + index * 4 : difat_offset + (index + 1) * 4] = value.to_bytes(4, "little")
    return bytes(header)
