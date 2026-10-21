from __future__ import annotations

import zipfile
from collections.abc import Iterable, Mapping
from pathlib import PurePosixPath

from .parts import HwpxPart, infer_part_class


def normalize_package_path(path: str) -> str:
    return PurePosixPath(path).as_posix()


def read_archive_state(
    archive: zipfile.ZipFile,
) -> tuple[dict[str, HwpxPart], list[str], dict[str, zipfile.ZipInfo], list[str]]:
    parts: dict[str, HwpxPart] = {}
    part_order: list[str] = []
    zip_infos: dict[str, zipfile.ZipInfo] = {}
    duplicates: list[str] = []

    for info in archive.infolist():
        normalized = normalize_package_path(info.filename)
        raw_bytes = archive.read(info)
        part_cls = infer_part_class(normalized, raw_bytes)
        part = part_cls(normalized, raw_bytes)
        if normalized in parts:
            duplicates.append(normalized)
        parts[normalized] = part
        zip_infos[normalized] = info
        if normalized not in part_order:
            part_order.append(normalized)
    return parts, part_order, zip_infos, duplicates


def write_archive_state(
    archive: zipfile.ZipFile,
    *,
    part_paths: Iterable[str],
    parts: Mapping[str, HwpxPart],
    zip_infos: Mapping[str, zipfile.ZipInfo],
) -> None:
    for path in part_paths:
        part = parts[path]
        data = part.to_bytes()
        source_info = zip_infos.get(path)
        if source_info is None:
            info = zipfile.ZipInfo(filename=path)
            info.compress_type = zipfile.ZIP_STORED if path == "mimetype" else zipfile.ZIP_DEFLATED
        else:
            info = clone_zip_info(source_info, path)
            if path == "mimetype":
                info.compress_type = zipfile.ZIP_STORED
        archive.writestr(info, data)


def clone_zip_info(source: zipfile.ZipInfo, filename: str) -> zipfile.ZipInfo:
    clone = zipfile.ZipInfo(filename=filename, date_time=source.date_time)
    clone.comment = source.comment
    clone.extra = source.extra
    clone.create_system = source.create_system
    clone.create_version = source.create_version
    clone.extract_version = source.extract_version
    clone.reserved = source.reserved
    clone.flag_bits = source.flag_bits
    clone.volume = source.volume
    clone.internal_attr = source.internal_attr
    clone.external_attr = source.external_attr
    clone.compress_type = source.compress_type
    return clone
