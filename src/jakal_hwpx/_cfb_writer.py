from __future__ import annotations

from dataclasses import dataclass, field
import heapq
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
    clsid: bytes = b"\x00" * 16
    color: int = BLACK
    state_bits: int = 0
    create_time: int = 0
    modify_time: int = 0


@dataclass(frozen=True)
class CfbLayoutEntry:
    sid: int
    path: str | None
    name: str
    entry_type: int
    left: int
    right: int
    child: int
    start_sector: int
    size: int
    clsid: bytes
    color: int
    state_bits: int
    create_time: int
    modify_time: int
    sector_chain: tuple[int, ...] = ()
    mini_sector_chain: tuple[int, ...] = ()


@dataclass(frozen=True)
class CfbLayout:
    entries: tuple[CfbLayoutEntry, ...]
    sector_size: int = SECTOR_SIZE
    mini_sector_size: int = MINI_SECTOR_SIZE
    directory_chain: tuple[int, ...] = ()
    mini_stream_chain: tuple[int, ...] = ()
    mini_fat_chain: tuple[int, ...] = ()
    fat_entries: tuple[int, ...] = ()
    mini_fat_entries: tuple[int, ...] = ()
    fat_sector_ids: tuple[int, ...] = ()
    difat_sector_ids: tuple[int, ...] = ()
    num_difat_sectors: int = 0
    original_file_bytes: bytes | None = None


def write_compound_file(path: str | Path, streams: dict[str, bytes], layout: CfbLayout | None = None) -> Path:
    target_path = Path(path).expanduser().resolve()
    target_path.parent.mkdir(parents=True, exist_ok=True)

    if _can_patch_original_compound_file(layout, streams):
        target_path.write_bytes(_patch_original_compound_file(layout, streams))
        return target_path
    layout_bytes = _write_compound_file_using_layout(streams, layout)
    if layout_bytes is not None:
        target_path.write_bytes(layout_bytes)
        return target_path

    entries = _build_entries_from_layout(streams, layout) if _layout_matches_streams(layout, streams) else _build_entries(streams)
    mini_stream_data = _assign_mini_stream(entries)

    regular_objects: list[tuple[int, bytes]] = []
    for index, entry in enumerate(entries):
        if entry.entry_type == STGTY_STREAM and (entry.size == 0 or entry.size >= MINI_STREAM_CUTOFF):
            regular_objects.append((index, entry.data))
    regular_objects.sort(key=lambda item: _stream_sector_order(entries[item[0]]))
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
    fat_sector_count, difat_sector_count = _compute_fat_and_difat_sector_counts(data_sector_count)
    fat_sector_ids = list(range(data_sector_count, data_sector_count + fat_sector_count))
    difat_sector_ids = list(range(data_sector_count + fat_sector_count, data_sector_count + fat_sector_count + difat_sector_count))

    fat_entries = [FREESECT] * (data_sector_count + fat_sector_count + difat_sector_count)
    for chain in object_chains.values():
        _mark_chain(fat_entries, chain)
    if mini_fat_chain:
        _mark_chain(fat_entries, mini_fat_chain)
    _mark_chain(fat_entries, directory_chain)
    for fat_sector_id in fat_sector_ids:
        fat_entries[fat_sector_id] = FATSECT
    for difat_sector_id in difat_sector_ids:
        fat_entries[difat_sector_id] = DIFSECT

    fat_bytes = b"".join(value.to_bytes(4, "little") for value in fat_entries)
    fat_bytes = fat_bytes.ljust(fat_sector_count * SECTOR_SIZE, b"\xFF")
    for index in range(fat_sector_count):
        start = index * SECTOR_SIZE
        end = start + SECTOR_SIZE
        sectors.append(fat_bytes[start:end])
    sectors.extend(_build_difat_sector_payloads(fat_sector_ids, difat_sector_ids, SECTOR_SIZE))

    header = _build_header(
        directory_start_sector=directory_start_sector,
        fat_sector_ids=fat_sector_ids,
        mini_fat_chain=mini_fat_chain,
        difat_sector_ids=difat_sector_ids,
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


def _compute_fat_and_difat_sector_counts(data_sector_count: int) -> tuple[int, int]:
    fat_sector_count = max(1, _compute_fat_sector_count(data_sector_count))
    difat_sector_count = max(0, (fat_sector_count - 109 + 126) // 127)
    while True:
        needed_fat_sector_count = max(1, _compute_fat_sector_count(data_sector_count + difat_sector_count))
        needed_difat_sector_count = max(0, (needed_fat_sector_count - 109 + 126) // 127)
        if needed_fat_sector_count == fat_sector_count and needed_difat_sector_count == difat_sector_count:
            return fat_sector_count, difat_sector_count
        fat_sector_count = needed_fat_sector_count
        difat_sector_count = needed_difat_sector_count


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


def _layout_matches_streams(layout: CfbLayout | None, streams: dict[str, bytes]) -> bool:
    if layout is None:
        return False
    layout_stream_paths = {entry.path for entry in layout.entries if entry.entry_type == STGTY_STREAM and entry.path is not None}
    return layout_stream_paths == set(streams)


def _can_patch_original_compound_file(layout: CfbLayout | None, streams: dict[str, bytes]) -> bool:
    if layout is None or layout.original_file_bytes is None:
        return False
    if not _layout_matches_streams(layout, streams):
        return False
    for entry in layout.entries:
        if entry.entry_type != STGTY_STREAM or entry.path is None:
            continue
        if len(streams[entry.path]) != entry.size:
            return False
    return True


def _patch_original_compound_file(layout: CfbLayout, streams: dict[str, bytes]) -> bytes:
    assert layout.original_file_bytes is not None
    data = bytearray(layout.original_file_bytes)
    mini_stream_bytes = bytearray(_read_regular_chain_bytes(data, layout.mini_stream_chain, layout.sector_size))
    for entry in layout.entries:
        if entry.entry_type != STGTY_STREAM or entry.path is None:
            continue
        stream_data = streams[entry.path]
        if entry.size == 0:
            continue
        if entry.size >= MINI_STREAM_CUTOFF:
            _patch_regular_stream_bytes(data, entry.sector_chain, stream_data, layout.sector_size)
        else:
            _patch_mini_stream_bytes(mini_stream_bytes, entry.mini_sector_chain, stream_data, layout.mini_sector_size)
    if layout.mini_stream_chain:
        _patch_regular_stream_bytes(data, layout.mini_stream_chain, bytes(mini_stream_bytes), layout.sector_size)
    return bytes(data)


def _write_compound_file_using_layout(streams: dict[str, bytes], layout: CfbLayout | None) -> bytes | None:
    if layout is None or layout.original_file_bytes is None:
        return None
    if not layout.fat_entries or not layout.fat_sector_ids:
        return None

    entries, sid_to_path = _build_entries_with_layout_hints(streams, layout)
    templates_by_sid = {entry.sid: entry for entry in layout.entries}
    layout_difat_sector_ids = set(layout.difat_sector_ids or tuple(sector_id for sector_id, value in enumerate(layout.fat_entries) if value == DIFSECT))
    regular_free_pool = {
        sector_id
        for sector_id, value in enumerate(layout.fat_entries)
        if value == FREESECT and sector_id not in layout.fat_sector_ids and sector_id not in layout_difat_sector_ids
    }
    mini_free_pool = {sector_id for sector_id, value in enumerate(layout.mini_fat_entries) if value == FREESECT}

    for template in layout.entries:
        if template.entry_type != STGTY_STREAM or template.path is None:
            continue
        stream_data = streams.get(template.path)
        stream_size = len(stream_data) if stream_data is not None else 0
        use_mini_stream = 0 < stream_size < MINI_STREAM_CUTOFF
        if template.sector_chain and stream_size > 0 and not use_mini_stream:
            continue
        if template.sector_chain:
            regular_free_pool.update(template.sector_chain)
        if template.mini_sector_chain and stream_size > 0 and use_mini_stream:
            continue
        if template.mini_sector_chain:
            mini_free_pool.update(template.mini_sector_chain)

    regular_heap = list(sorted(regular_free_pool))
    heapq.heapify(regular_heap)
    next_regular_sector_id = len(layout.fat_entries)
    mini_heap = list(sorted(mini_free_pool))
    heapq.heapify(mini_heap)
    next_mini_sector_id = len(layout.mini_fat_entries)

    def allocate_regular(count: int) -> list[int]:
        nonlocal next_regular_sector_id
        allocated: list[int] = []
        for _ in range(count):
            if regular_heap:
                allocated.append(heapq.heappop(regular_heap))
            else:
                allocated.append(next_regular_sector_id)
                next_regular_sector_id += 1
        return allocated

    def allocate_mini(count: int) -> list[int]:
        nonlocal next_mini_sector_id
        allocated: list[int] = []
        for _ in range(count):
            if mini_heap:
                allocated.append(heapq.heappop(mini_heap))
            else:
                allocated.append(next_mini_sector_id)
                next_mini_sector_id += 1
        return allocated

    regular_stream_chains: dict[int, list[int]] = {}
    mini_stream_chains: dict[int, list[int]] = {}
    max_used_mini_sector = -1

    for sid, entry in enumerate(entries):
        if entry.entry_type != STGTY_STREAM:
            continue
        template = templates_by_sid.get(sid)
        stream_path = sid_to_path.get(sid)
        if stream_path is None or stream_path not in streams:
            continue
        stream_data = streams[stream_path]
        stream_size = len(stream_data)
        use_mini_stream = 0 < stream_size < MINI_STREAM_CUTOFF
        if use_mini_stream:
            base_chain = list(template.mini_sector_chain) if template is not None else []
            desired_count = _sector_count_for_size(stream_size, layout.mini_sector_size)
            target_count = max(desired_count, len(base_chain))
            if target_count > len(base_chain):
                base_chain.extend(allocate_mini(target_count - len(base_chain)))
            mini_stream_chains[sid] = base_chain
            entry.start_sector = base_chain[0] if base_chain else ENDOFCHAIN
            entry.size = stream_size
            if base_chain:
                max_used_mini_sector = max(max_used_mini_sector, base_chain[-1])
            continue
        if stream_size <= 0:
            entry.start_sector = ENDOFCHAIN
            entry.size = 0
            continue
        base_chain = list(template.sector_chain) if template is not None else []
        desired_count = _sector_count_for_size(stream_size, layout.sector_size)
        target_count = max(desired_count, len(base_chain))
        if target_count > len(base_chain):
            base_chain.extend(allocate_regular(target_count - len(base_chain)))
        regular_stream_chains[sid] = base_chain
        entry.start_sector = base_chain[0] if base_chain else ENDOFCHAIN
        entry.size = stream_size

    preserve_existing_mini_layout = bool(layout.mini_fat_entries or layout.mini_stream_chain or layout.mini_fat_chain)
    target_mini_sector_count = max_used_mini_sector + 1
    if preserve_existing_mini_layout:
        target_mini_sector_count = max(target_mini_sector_count, len(layout.mini_fat_entries))
    if target_mini_sector_count < 0:
        target_mini_sector_count = 0

    root_entry = entries[0]
    root_chain = list(layout.mini_stream_chain)
    root_sector_count = _sector_count_for_size(target_mini_sector_count * layout.mini_sector_size, layout.sector_size)
    if root_chain:
        root_sector_count = max(root_sector_count, len(root_chain))
    if root_sector_count > len(root_chain):
        root_chain.extend(allocate_regular(root_sector_count - len(root_chain)))
    root_entry.start_sector = root_chain[0] if root_chain else ENDOFCHAIN
    root_template = templates_by_sid.get(0)
    root_entry.size = max(root_template.size if root_template is not None else 0, target_mini_sector_count * layout.mini_sector_size)

    mini_fat_entries = _build_layout_aware_mini_fat(mini_stream_chains, target_mini_sector_count)
    mini_fat_chain = list(layout.mini_fat_chain)
    mini_fat_sector_count = _sector_count_for_size(len(mini_fat_entries) * 4, layout.sector_size)
    if mini_fat_chain:
        mini_fat_sector_count = max(mini_fat_sector_count, len(mini_fat_chain))
    if mini_fat_sector_count > len(mini_fat_chain):
        mini_fat_chain.extend(allocate_regular(mini_fat_sector_count - len(mini_fat_chain)))

    directory_chain = list(layout.directory_chain)
    directory_size = _directory_stream_size(len(entries))
    directory_sector_count = _sector_count_for_size(directory_size, layout.sector_size)
    if directory_chain:
        directory_sector_count = max(directory_sector_count, len(directory_chain))
    if directory_sector_count > len(directory_chain):
        directory_chain.extend(allocate_regular(directory_sector_count - len(directory_chain)))

    fat_sector_ids = list(layout.fat_sector_ids)
    difat_sector_ids = list(sorted(layout_difat_sector_ids))
    max_nonfat_sector_id = _max_sector_id(
        list(regular_stream_chains.values()) + list(mini_stream_chains.values()) + [root_chain, mini_fat_chain, directory_chain]
    )
    original_nonfat_sector_count = max(0, len(layout.fat_entries) - len(layout.fat_sector_ids) - len(difat_sector_ids))
    nonfat_sector_count = max(original_nonfat_sector_count, next_regular_sector_id, max_nonfat_sector_id + 1)
    needed_fat_sector_count, needed_difat_sector_count = _compute_fat_and_difat_sector_counts(nonfat_sector_count)
    if needed_fat_sector_count > len(fat_sector_ids):
        fat_sector_ids.extend(allocate_regular(needed_fat_sector_count - len(fat_sector_ids)))
    if needed_difat_sector_count > len(difat_sector_ids):
        difat_sector_ids.extend(allocate_regular(needed_difat_sector_count - len(difat_sector_ids)))

    total_sector_count = max(
        len(layout.fat_entries),
        next_regular_sector_id,
        max_nonfat_sector_id + 1,
        (max(fat_sector_ids) + 1) if fat_sector_ids else 0,
        (max(difat_sector_ids) + 1) if difat_sector_ids else 0,
    )

    fat_entries = [FREESECT] * total_sector_count
    for chain in regular_stream_chains.values():
        _mark_chain(fat_entries, chain)
    if root_chain:
        _mark_chain(fat_entries, root_chain)
    if mini_fat_chain:
        _mark_chain(fat_entries, mini_fat_chain)
    if directory_chain:
        _mark_chain(fat_entries, directory_chain)
    for fat_sector_id in fat_sector_ids:
        if fat_sector_id >= len(fat_entries):
            return None
        fat_entries[fat_sector_id] = FATSECT
    for difat_sector_id in difat_sector_ids:
        if difat_sector_id >= len(fat_entries):
            return None
        fat_entries[difat_sector_id] = DIFSECT

    container = bytearray(layout.original_file_bytes)
    required_size = 512 + total_sector_count * layout.sector_size
    if len(container) < required_size:
        container.extend(b"\x00" * (required_size - len(container)))

    if root_chain:
        mini_stream_capacity = len(root_chain) * layout.sector_size
        mini_stream_bytes = bytearray(mini_stream_capacity)
        original_mini_stream = _read_regular_chain_bytes(container, layout.mini_stream_chain, layout.sector_size)
        mini_stream_bytes[: min(len(original_mini_stream), mini_stream_capacity)] = original_mini_stream[:mini_stream_capacity]
    else:
        mini_stream_bytes = bytearray()

    for sid, chain in regular_stream_chains.items():
        stream_path = sid_to_path.get(sid)
        if stream_path is None:
            continue
        _patch_regular_stream_bytes(container, tuple(chain), streams[stream_path], layout.sector_size)
    for sid, chain in mini_stream_chains.items():
        stream_path = sid_to_path.get(sid)
        if stream_path is None:
            continue
        _patch_mini_stream_bytes(mini_stream_bytes, tuple(chain), streams[stream_path], layout.mini_sector_size)
    if root_chain:
        _write_regular_chain_payload(container, tuple(root_chain), bytes(mini_stream_bytes), layout.sector_size, pad_byte=b"\x00")

    directory_bytes = _build_directory_stream(entries)
    _write_regular_chain_payload(container, tuple(directory_chain), directory_bytes, layout.sector_size, pad_byte=b"\x00")

    mini_fat_bytes = b"".join(value.to_bytes(4, "little") for value in mini_fat_entries)
    _write_regular_chain_payload(container, tuple(mini_fat_chain), mini_fat_bytes, layout.sector_size, pad_byte=b"\xFF")

    fat_bytes = b"".join(value.to_bytes(4, "little") for value in fat_entries)
    fat_bytes = fat_bytes.ljust(len(fat_sector_ids) * layout.sector_size, b"\xFF")
    for sector_index, fat_sector_id in enumerate(fat_sector_ids):
        start = sector_index * layout.sector_size
        end = start + layout.sector_size
        sector_start = 512 + fat_sector_id * layout.sector_size
        sector_end = sector_start + layout.sector_size
        container[sector_start:sector_end] = fat_bytes[start:end]
    difat_payloads = _build_difat_sector_payloads(fat_sector_ids, difat_sector_ids, layout.sector_size)
    for sector_index, difat_sector_id in enumerate(difat_sector_ids):
        sector_start = 512 + difat_sector_id * layout.sector_size
        sector_end = sector_start + layout.sector_size
        container[sector_start:sector_end] = difat_payloads[sector_index]

    header = _build_header(
        directory_start_sector=directory_chain[0] if directory_chain else ENDOFCHAIN,
        fat_sector_ids=fat_sector_ids,
        mini_fat_chain=mini_fat_chain,
        difat_sector_ids=difat_sector_ids,
    )
    container[:512] = header
    return bytes(container[:required_size])


def _patch_regular_stream_bytes(container: bytearray, sector_chain: tuple[int, ...], stream_data: bytes, sector_size: int) -> None:
    remaining = len(stream_data)
    cursor = 0
    for sector_id in sector_chain:
        if remaining <= 0:
            break
        chunk_size = min(remaining, sector_size)
        start = 512 + sector_id * sector_size
        end = start + chunk_size
        container[start:end] = stream_data[cursor : cursor + chunk_size]
        cursor += chunk_size
        remaining -= chunk_size


def _patch_mini_stream_bytes(mini_stream: bytearray, mini_sector_chain: tuple[int, ...], stream_data: bytes, mini_sector_size: int) -> None:
    remaining = len(stream_data)
    cursor = 0
    for sector_id in mini_sector_chain:
        if remaining <= 0:
            break
        chunk_size = min(remaining, mini_sector_size)
        start = sector_id * mini_sector_size
        end = start + chunk_size
        mini_stream[start:end] = stream_data[cursor : cursor + chunk_size]
        cursor += chunk_size
        remaining -= chunk_size


def _write_regular_chain_payload(
    container: bytearray,
    sector_chain: tuple[int, ...],
    payload: bytes,
    sector_size: int,
    *,
    pad_byte: bytes,
) -> None:
    if not sector_chain:
        return
    full_length = len(sector_chain) * sector_size
    padded = payload[:full_length].ljust(full_length, pad_byte)
    for chain_index, sector_id in enumerate(sector_chain):
        start = 512 + sector_id * sector_size
        end = start + sector_size
        payload_start = chain_index * sector_size
        payload_end = payload_start + sector_size
        container[start:end] = padded[payload_start:payload_end]


def _read_regular_chain_bytes(container: bytearray | bytes, sector_chain: tuple[int, ...], sector_size: int) -> bytes:
    chunks: list[bytes] = []
    for sector_id in sector_chain:
        start = 512 + sector_id * sector_size
        end = start + sector_size
        chunks.append(bytes(container[start:end]))
    return b"".join(chunks)


def _build_entries_from_layout(streams: dict[str, bytes], layout: CfbLayout) -> list[_CfbEntry]:
    if not layout.entries:
        return _build_entries(streams)
    max_sid = max(entry.sid for entry in layout.entries)
    entries = [_CfbEntry(name="", entry_type=STGTY_EMPTY, parent=None) for _ in range(max_sid + 1)]
    templates = {entry.sid: entry for entry in layout.entries}
    for sid in range(max_sid + 1):
        template = templates.get(sid)
        if template is None:
            continue
        data = streams.get(template.path or "", b"") if template.entry_type == STGTY_STREAM and template.path is not None else b""
        size = len(data) if template.entry_type == STGTY_STREAM and template.path is not None else template.size
        entries[sid] = _CfbEntry(
            name=template.name,
            entry_type=template.entry_type,
            parent=None,
            data=data,
            left=template.left,
            right=template.right,
            child=template.child,
            start_sector=template.start_sector,
            size=size,
            clsid=template.clsid,
            color=template.color,
            state_bits=template.state_bits,
            create_time=template.create_time,
            modify_time=template.modify_time,
        )
    return entries


def _build_entries_with_layout_hints(streams: dict[str, bytes], layout: CfbLayout) -> tuple[list[_CfbEntry], dict[int, str]]:
    if not layout.entries:
        entries = _build_entries(streams)
        return entries, _entry_paths_from_tree(entries)

    max_sid = max(entry.sid for entry in layout.entries)
    entries = [_CfbEntry(name="", entry_type=STGTY_EMPTY, parent=None) for _ in range(max_sid + 1)]
    templates = {entry.sid: entry for entry in layout.entries}
    sid_to_path: dict[int, str] = {}
    path_to_sid: dict[str, int] = {"": 0}

    for sid in range(max_sid + 1):
        template = templates.get(sid)
        if template is None:
            continue
        if sid == 0:
            entries[sid] = _CfbEntry(
                name=template.name or "Root Entry",
                entry_type=template.entry_type,
                parent=None,
                left=template.left,
                right=template.right,
                child=template.child,
                start_sector=template.start_sector,
                size=template.size,
                clsid=template.clsid,
                color=template.color,
                state_bits=template.state_bits,
                create_time=template.create_time,
                modify_time=template.modify_time,
            )
            continue
        path = template.path
        if path:
            path_to_sid[path] = sid
            sid_to_path[sid] = path
        if template.entry_type == STGTY_STREAM and (path is None or path not in streams):
            entries[sid] = _CfbEntry(name="", entry_type=STGTY_EMPTY, parent=None)
            continue
        data = streams.get(path or "", b"") if template.entry_type == STGTY_STREAM and path is not None else b""
        size = len(data) if template.entry_type == STGTY_STREAM and path is not None else template.size
        entries[sid] = _CfbEntry(
            name=template.name,
            entry_type=template.entry_type,
            parent=None,
            data=data,
            left=template.left,
            right=template.right,
            child=template.child,
            start_sector=template.start_sector,
            size=size,
            clsid=template.clsid,
            color=template.color,
            state_bits=template.state_bits,
            create_time=template.create_time,
            modify_time=template.modify_time,
        )

    def ensure_storage(storage_path: str) -> int:
        if storage_path in path_to_sid:
            return path_to_sid[storage_path]
        parent_path, _, name = storage_path.rpartition("/")
        parent_sid = ensure_storage(parent_path)
        entry = _CfbEntry(name=name, entry_type=STGTY_STORAGE, parent=parent_sid)
        entries.append(entry)
        sid = len(entries) - 1
        path_to_sid[storage_path] = sid
        sid_to_path[sid] = storage_path
        return sid

    for stream_path, data in sorted(streams.items()):
        existing_sid = path_to_sid.get(stream_path)
        if existing_sid is not None and entries[existing_sid].entry_type == STGTY_STREAM:
            entries[existing_sid].data = data
            entries[existing_sid].size = len(data)
            continue
        parent_path, _, name = stream_path.rpartition("/")
        parent_sid = ensure_storage(parent_path)
        entry = _CfbEntry(name=name, entry_type=STGTY_STREAM, parent=parent_sid, data=data, size=len(data))
        entries.append(entry)
        sid = len(entries) - 1
        path_to_sid[stream_path] = sid
        sid_to_path[sid] = stream_path

    for sid, entry in enumerate(entries):
        entry.children = []
        entry.left = NOSTREAM
        entry.right = NOSTREAM
        entry.child = NOSTREAM
        if sid == 0:
            entry.parent = None
            path_to_sid[""] = 0
            continue
        path = sid_to_path.get(sid)
        if entry.entry_type == STGTY_EMPTY or path is None:
            entry.parent = None
            continue
        parent_path, _, _ = path.rpartition("/")
        parent_sid = path_to_sid.get(parent_path, 0)
        entry.parent = parent_sid
        entries[parent_sid].children.append(sid)

    for entry in entries:
        if entry.entry_type in (STGTY_ROOT, STGTY_STORAGE):
            _assign_btree(entries, entry.children, lambda entry_id: entries[entry_id].name.casefold())
            entry.child = entry.children[0] if len(entry.children) == 1 else _btree_root(entry.children, entries)

    return entries, sid_to_path


def _entry_paths_from_tree(entries: list[_CfbEntry]) -> dict[int, str]:
    path_by_sid: dict[int, str] = {0: ""}
    pending = [0]
    while pending:
        sid = pending.pop()
        parent_path = path_by_sid.get(sid, "")
        for child_sid in entries[sid].children:
            name = entries[child_sid].name
            child_path = name if not parent_path else f"{parent_path}/{name}"
            path_by_sid[child_sid] = child_path
            pending.append(child_sid)
    return path_by_sid


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
    mini_entries = [
        entry
        for entry in entries
        if entry.entry_type == STGTY_STREAM and entry.size != 0 and entry.size < MINI_STREAM_CUTOFF
    ]
    mini_entries.sort(key=_stream_sector_order)
    for entry in mini_entries:
        padded = entry.data + (b"\x00" * ((MINI_SECTOR_SIZE - (len(entry.data) % MINI_SECTOR_SIZE)) % MINI_SECTOR_SIZE))
        entry.start_sector = mini_sector_cursor
        mini_sector_cursor += len(padded) // MINI_SECTOR_SIZE
        mini_stream_chunks.append(padded)
    return b"".join(mini_stream_chunks)


def _stream_sector_order(entry: _CfbEntry) -> tuple[int, int]:
    if entry.start_sector in (ENDOFCHAIN, FREESECT):
        return (1, 0x7FFFFFFF)
    return (0, int(entry.start_sector))


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


def _build_layout_aware_mini_fat(stream_chains: dict[int, list[int]], mini_sector_count: int) -> list[int]:
    if mini_sector_count <= 0:
        return []
    entries_map = [FREESECT] * mini_sector_count
    for chain in stream_chains.values():
        if not chain:
            continue
        for current, next_sector in zip(chain, chain[1:]):
            entries_map[current] = next_sector
        entries_map[chain[-1]] = ENDOFCHAIN
    return entries_map


def _max_sector_id(chains: list[list[int]]) -> int:
    maximum = -1
    for chain in chains:
        if chain:
            maximum = max(maximum, chain[-1])
    return maximum


def _build_difat_sector_payloads(fat_sector_ids: list[int], difat_sector_ids: list[int], sector_size: int) -> list[bytes]:
    if not difat_sector_ids:
        return []
    payloads: list[bytes] = []
    remaining_fat_sector_ids = list(fat_sector_ids[109:])
    entries_per_difat_sector = (sector_size // 4) - 1
    for index, _sector_id in enumerate(difat_sector_ids):
        payload = bytearray(sector_size)
        chunk = remaining_fat_sector_ids[:entries_per_difat_sector]
        remaining_fat_sector_ids = remaining_fat_sector_ids[entries_per_difat_sector:]
        for entry_index in range(entries_per_difat_sector):
            value = chunk[entry_index] if entry_index < len(chunk) else FREESECT
            offset = entry_index * 4
            payload[offset : offset + 4] = int(value).to_bytes(4, "little", signed=False)
        next_sector_id = difat_sector_ids[index + 1] if index + 1 < len(difat_sector_ids) else ENDOFCHAIN
        payload[sector_size - 4 : sector_size] = int(next_sector_id).to_bytes(4, "little", signed=False)
        payloads.append(bytes(payload))
    return payloads


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
    data[67] = entry.color & 0xFF
    data[68:72] = entry.left.to_bytes(4, "little", signed=False)
    data[72:76] = entry.right.to_bytes(4, "little", signed=False)
    data[76:80] = entry.child.to_bytes(4, "little", signed=False)
    data[80:96] = entry.clsid[:16].ljust(16, b"\x00")
    data[96:100] = int(entry.state_bits).to_bytes(4, "little", signed=False)
    data[100:108] = int(entry.create_time).to_bytes(8, "little", signed=False)
    data[108:116] = int(entry.modify_time).to_bytes(8, "little", signed=False)
    data[116:120] = entry.start_sector.to_bytes(4, "little", signed=False)
    data[120:128] = entry.size.to_bytes(8, "little", signed=False)
    return bytes(data)


def capture_compound_file_layout(ole, *, original_file_bytes: bytes | None = None) -> CfbLayout:
    path_by_sid: dict[int, str | None] = {}
    visited: set[int] = set()

    def walk(entry, parent_path: str) -> None:
        if entry is None or entry.sid in visited:
            return
        visited.add(entry.sid)
        current_path = "" if entry.sid == 0 else f"{parent_path}/{entry.name}" if parent_path else entry.name
        path_by_sid[entry.sid] = current_path
        for child in getattr(entry, "kids", []) or []:
            walk(child, current_path)

    walk(getattr(ole, "root", None), "")
    minifat = getattr(ole, "minifat", None)
    if minifat is None and hasattr(ole, "loadminifat"):
        try:
            ole.loadminifat()
        except Exception:
            minifat = None
        else:
            minifat = getattr(ole, "minifat", None)
    fat_entries = tuple(int(value) for value in (getattr(ole, "fat", []) or []))
    mini_fat_entries = tuple(int(value) for value in ((minifat or []) if minifat is not None else ()))
    directory_chain = _collect_sector_chain(getattr(ole, "first_dir_sector", ENDOFCHAIN), fat_entries)
    mini_stream_chain = _collect_sector_chain(getattr(getattr(ole, "root", None), "isectStart", ENDOFCHAIN), fat_entries)
    mini_fat_chain = _collect_sector_chain(getattr(ole, "first_mini_fat_sector", ENDOFCHAIN), fat_entries)
    fat_sector_ids = _extract_fat_sector_ids(
        original_file_bytes,
        sector_size=int(getattr(ole, "sector_size", SECTOR_SIZE)),
        fat=fat_entries,
        num_fat_sectors=int(getattr(ole, "num_fat_sectors", 0)),
        num_difat_sectors=int(getattr(ole, "num_difat_sectors", 0)),
    )
    difat_sector_ids = tuple(sector_id for sector_id, value in enumerate(fat_entries) if value == DIFSECT)
    templates: list[CfbLayoutEntry] = []
    for sid, entry in enumerate(getattr(ole, "direntries", ())):
        if entry is None:
            templates.append(
                CfbLayoutEntry(
                    sid=sid,
                    path=None,
                    name="",
                    entry_type=STGTY_EMPTY,
                    left=NOSTREAM,
                    right=NOSTREAM,
                    child=NOSTREAM,
                    start_sector=ENDOFCHAIN,
                    size=0,
                    clsid=b"\x00" * 16,
                    color=BLACK,
                    state_bits=0,
                    create_time=0,
                    modify_time=0,
                    sector_chain=(),
                    mini_sector_chain=(),
                )
            )
            continue
        clsid = getattr(entry, "clsid", "")
        clsid_bytes = clsid.bytes_le if hasattr(clsid, "bytes_le") else b"\x00" * 16
        sector_chain: tuple[int, ...] = ()
        mini_sector_chain: tuple[int, ...] = ()
        if int(entry.entry_type) == STGTY_STREAM and int(entry.size) > 0:
            if int(entry.size) >= MINI_STREAM_CUTOFF:
                sector_chain = _collect_sector_chain(int(entry.isectStart), getattr(ole, "fat", []), _sector_count_for_size(int(entry.size), SECTOR_SIZE))
            else:
                mini_sector_chain = _collect_sector_chain(int(entry.isectStart), minifat or [], _sector_count_for_size(int(entry.size), MINI_SECTOR_SIZE))
        templates.append(
            CfbLayoutEntry(
                sid=sid,
                path=path_by_sid.get(sid),
                name=str(entry.name),
                entry_type=int(entry.entry_type),
                left=int(entry.sid_left),
                right=int(entry.sid_right),
                child=int(entry.sid_child),
                start_sector=int(entry.isectStart),
                size=int(entry.size),
                clsid=clsid_bytes,
                color=int(getattr(entry, "color", BLACK)),
                state_bits=int(getattr(entry, "dwUserFlags", 0)),
                create_time=int(getattr(entry, "createTime", 0)),
                modify_time=int(getattr(entry, "modifyTime", 0)),
                sector_chain=sector_chain,
                mini_sector_chain=mini_sector_chain,
            )
        )
    return CfbLayout(
        entries=tuple(templates),
        sector_size=int(getattr(ole, "sector_size", SECTOR_SIZE)),
        mini_sector_size=int(getattr(ole, "mini_sector_size", MINI_SECTOR_SIZE)),
        directory_chain=directory_chain,
        mini_stream_chain=mini_stream_chain,
        mini_fat_chain=mini_fat_chain,
        fat_entries=fat_entries,
        mini_fat_entries=mini_fat_entries,
        fat_sector_ids=fat_sector_ids,
        difat_sector_ids=difat_sector_ids,
        num_difat_sectors=int(getattr(ole, "num_difat_sectors", 0)),
        original_file_bytes=original_file_bytes,
    )


def _collect_sector_chain(start_sector: int, fat: list[int] | tuple[int, ...] | None, limit: int | None = None) -> tuple[int, ...]:
    if fat is None or start_sector in (ENDOFCHAIN, FREESECT, NOSTREAM):
        return ()
    chain: list[int] = []
    current = int(start_sector)
    while current not in (ENDOFCHAIN, FREESECT, NOSTREAM):
        if current < 0 or current >= len(fat):
            break
        if current in chain:
            break
        chain.append(current)
        if limit is not None and len(chain) >= limit:
            break
        current = int(fat[current])
    return tuple(chain)


def _sector_count_for_size(size: int, sector_size: int) -> int:
    if size <= 0:
        return 0
    return (size + sector_size - 1) // sector_size


def _extract_fat_sector_ids(
    original_file_bytes: bytes | None,
    *,
    sector_size: int,
    fat: tuple[int, ...],
    num_fat_sectors: int,
    num_difat_sectors: int,
) -> tuple[int, ...]:
    if original_file_bytes is None or num_fat_sectors <= 0:
        return ()
    header = original_file_bytes[:512]
    fat_sector_ids: list[int] = []
    for offset in range(76, 512, 4):
        value = int.from_bytes(header[offset : offset + 4], "little")
        if value not in (FREESECT, ENDOFCHAIN):
            fat_sector_ids.append(value)
            if len(fat_sector_ids) >= num_fat_sectors:
                return tuple(fat_sector_ids[:num_fat_sectors])
    if num_difat_sectors <= 0:
        return tuple(fat_sector_ids[:num_fat_sectors])
    current_difat_sector = int.from_bytes(header[68:72], "little")
    visited: set[int] = set()
    while (
        current_difat_sector not in (ENDOFCHAIN, FREESECT, NOSTREAM)
        and current_difat_sector not in visited
        and len(fat_sector_ids) < num_fat_sectors
    ):
        visited.add(current_difat_sector)
        sector_start = 512 + current_difat_sector * sector_size
        sector_end = sector_start + sector_size
        sector_bytes = original_file_bytes[sector_start:sector_end]
        if len(sector_bytes) != sector_size:
            break
        for offset in range(0, sector_size - 4, 4):
            value = int.from_bytes(sector_bytes[offset : offset + 4], "little")
            if value not in (FREESECT, ENDOFCHAIN):
                fat_sector_ids.append(value)
                if len(fat_sector_ids) >= num_fat_sectors:
                    return tuple(fat_sector_ids[:num_fat_sectors])
        if current_difat_sector >= len(fat):
            break
        current_difat_sector = int.from_bytes(sector_bytes[sector_size - 4 : sector_size], "little")
    return tuple(fat_sector_ids[:num_fat_sectors])


def _build_header(
    *,
    directory_start_sector: int,
    fat_sector_ids: list[int],
    mini_fat_chain: list[int],
    difat_sector_ids: list[int],
) -> bytes:
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
    header[68:72] = (difat_sector_ids[0] if difat_sector_ids else ENDOFCHAIN).to_bytes(4, "little")
    header[72:76] = len(difat_sector_ids).to_bytes(4, "little")
    difat_offset = 76
    for index in range(109):
        value = fat_sector_ids[index] if index < len(fat_sector_ids) else FREESECT
        header[difat_offset + index * 4 : difat_offset + (index + 1) * 4] = value.to_bytes(4, "little")
    return bytes(header)
