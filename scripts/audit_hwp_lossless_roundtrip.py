from __future__ import annotations

import argparse
import hashlib
import json
import sys
from dataclasses import asdict, dataclass
from pathlib import Path

import olefile

REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT_DIR = REPO_ROOT / "hwp_collection"
SRC_ROOT = REPO_ROOT / "src"
DEFAULT_TEMP_ROOT = REPO_ROOT / ".codex-temp" / "lossless_audit"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from jakal_hwpx import HwpBinaryDocument


@dataclass
class RoundtripResult:
    path: str
    writer_mode: str
    byte_identical: bool
    raw_streams_identical: bool
    logical_streams_identical: bool
    parsed_record_streams_identical: bool
    original_sha256: str
    roundtrip_sha256: str
    differing_raw_streams: list[str]
    differing_logical_streams: list[str]
    differing_parsed_record_streams: list[str]
    error: str | None = None


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _stable_temp_path(path: Path) -> Path:
    digest = hashlib.sha1(str(path).encode("utf-8")).hexdigest()
    return DEFAULT_TEMP_ROOT / f"{digest}.hwp"


def _stream_map(path: Path) -> dict[str, bytes]:
    streams: dict[str, bytes] = {}
    with olefile.OleFileIO(str(path)) as ole:
        for entry in ole.listdir(streams=True, storages=False):
            stream_path = "/".join(entry)
            streams[stream_path] = ole.openstream(entry).read()
    return streams


def _logical_stream_map(document: HwpBinaryDocument) -> dict[str, bytes]:
    return {stream_path: document.read_stream(stream_path) for stream_path in document.list_stream_paths()}


def _diff_stream_maps(base: dict[str, bytes], other: dict[str, bytes]) -> list[str]:
    differing: list[str] = []
    for path in sorted(set(base) | set(other)):
        if base.get(path) != other.get(path):
            differing.append(path)
    return differing


def _parsed_record_stream_diffs(document: HwpBinaryDocument) -> list[str]:
    differing: list[str] = []

    docinfo_raw = document.read_stream("DocInfo")
    if docinfo_raw != document.docinfo_model().to_bytes():
        differing.append("DocInfo")

    for section_index, stream_path in enumerate(document.section_stream_paths()):
        section_raw = document.read_stream(stream_path)
        if section_raw != document.section_model(section_index).to_bytes():
            differing.append(stream_path)

    return differing


def _roundtrip_once(path: Path, *, writer_mode: str) -> RoundtripResult:
    DEFAULT_TEMP_ROOT.mkdir(parents=True, exist_ok=True)
    temp_path = _stable_temp_path(path)
    if temp_path.exists():
        temp_path.unlink()

    original_doc = HwpBinaryDocument.open(path)
    original_logical_streams = _logical_stream_map(original_doc)
    original_raw_streams = _stream_map(path)
    differing_parsed_record_streams = _parsed_record_stream_diffs(original_doc)

    original_doc.save_copy(temp_path, preserve_original_bytes=(writer_mode == "preserve"))

    roundtrip_doc = HwpBinaryDocument.open(temp_path)
    roundtrip_logical_streams = _logical_stream_map(roundtrip_doc)
    roundtrip_raw_streams = _stream_map(temp_path)

    differing_raw_streams = _diff_stream_maps(original_raw_streams, roundtrip_raw_streams)
    differing_logical_streams = _diff_stream_maps(original_logical_streams, roundtrip_logical_streams)

    return RoundtripResult(
        path=str(path),
        writer_mode=writer_mode,
        byte_identical=path.read_bytes() == temp_path.read_bytes(),
        raw_streams_identical=not differing_raw_streams,
        logical_streams_identical=not differing_logical_streams,
        parsed_record_streams_identical=not differing_parsed_record_streams,
        original_sha256=_sha256(path),
        roundtrip_sha256=_sha256(temp_path),
        differing_raw_streams=differing_raw_streams,
        differing_logical_streams=differing_logical_streams,
        differing_parsed_record_streams=differing_parsed_record_streams,
    )


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Audit no-op HWP binary roundtrips for exact byte preservation.")
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=DEFAULT_INPUT_DIR,
        help="Directory containing .hwp files to audit.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="Only audit the first N files after sorting. Use 0 for all files.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Optional JSON report path.",
    )
    parser.add_argument(
        "--fail-on-mismatch",
        action="store_true",
        help="Exit with code 1 when any audit result is not fully identical.",
    )
    parser.add_argument(
        "--writer-mode",
        choices=("preserve", "reencode"),
        default="preserve",
        help="Choose whether no-op saves preserve original file bytes or force CFB re-encoding.",
    )
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    input_dir = args.input_dir.expanduser().resolve()
    paths = sorted(input_dir.glob("*.hwp"))
    if args.limit > 0:
        paths = paths[: args.limit]
    if not paths:
        raise SystemExit(f"No .hwp files found under {input_dir}")

    results: list[RoundtripResult] = []
    for path in paths:
        try:
            result = _roundtrip_once(path, writer_mode=args.writer_mode)
        except Exception as exc:  # pragma: no cover - audit harness should keep going.
            result = RoundtripResult(
                path=str(path),
                writer_mode=args.writer_mode,
                byte_identical=False,
                raw_streams_identical=False,
                logical_streams_identical=False,
                parsed_record_streams_identical=False,
                original_sha256=_sha256(path),
                roundtrip_sha256="",
                differing_raw_streams=[],
                differing_logical_streams=[],
                differing_parsed_record_streams=[],
                error=str(exc),
            )
        results.append(result)
        status = (
            "OK"
            if result.byte_identical
            and result.raw_streams_identical
            and result.logical_streams_identical
            and result.parsed_record_streams_identical
            else "DIFF"
        )
        print(
            f"[{status}] {Path(result.path).name} "
            f"bytes={result.byte_identical} raw={result.raw_streams_identical} "
            f"logical={result.logical_streams_identical} parsed={result.parsed_record_streams_identical}"
        )
        if result.error:
            print(f"  error: {result.error}")
        elif result.differing_raw_streams:
            print(f"  raw-stream-diff: {', '.join(result.differing_raw_streams[:8])}")
        elif result.differing_logical_streams:
            print(f"  logical-stream-diff: {', '.join(result.differing_logical_streams[:8])}")
        elif result.differing_parsed_record_streams:
            print(f"  parsed-record-diff: {', '.join(result.differing_parsed_record_streams[:8])}")

    summary = {
        "input_dir": str(input_dir),
        "writer_mode": args.writer_mode,
        "case_count": len(results),
        "byte_identical_count": sum(1 for result in results if result.byte_identical),
        "raw_stream_identical_count": sum(1 for result in results if result.raw_streams_identical),
        "logical_stream_identical_count": sum(1 for result in results if result.logical_streams_identical),
        "parsed_record_stream_identical_count": sum(1 for result in results if result.parsed_record_streams_identical),
        "error_count": sum(1 for result in results if result.error is not None),
        "results": [asdict(result) for result in results],
    }

    print(
        "[summary] "
        f"cases={summary['case_count']} "
        f"bytes={summary['byte_identical_count']} "
        f"raw={summary['raw_stream_identical_count']} "
        f"logical={summary['logical_stream_identical_count']} "
        f"parsed={summary['parsed_record_stream_identical_count']} "
        f"errors={summary['error_count']}"
    )

    if args.output is not None:
        output_path = args.output.expanduser().resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    if args.fail_on_mismatch and any(
        not (
            result.byte_identical
            and result.raw_streams_identical
            and result.logical_streams_identical
            and result.parsed_record_streams_identical
        )
        for result in results
    ):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
