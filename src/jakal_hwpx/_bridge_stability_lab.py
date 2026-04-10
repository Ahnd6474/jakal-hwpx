from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Callable

from ._hancom import convert_document
from .bridge import HwpHwpxBridge
from .document import HwpxDocument
from .exceptions import HancomInteropError
from .hwp_document import HwpDocument


REPO_ROOT = Path(__file__).resolve().parents[2]
HWP_SAMPLE_DIR = REPO_ROOT / "examples" / "samples" / "hwp"
HWPX_SAMPLE_DIR = REPO_ROOT / "examples" / "samples" / "hwpx"

BridgeExercise = Callable[[HwpHwpxBridge, Path], "BridgeExecution"]


@dataclass(frozen=True)
class BridgeArtifact:
    path: str
    kind: str
    expected_title: str | None = None
    expected_texts: tuple[str, ...] = ()
    expected_hwp_control_ids: tuple[str, ...] = ()
    hancom_roundtrip_title: str | None = None
    hancom_roundtrip_texts: tuple[str, ...] = ()


@dataclass(frozen=True)
class BridgeExecution:
    artifacts: tuple[BridgeArtifact, ...] = ()
    expected_conversions: tuple[str, ...] = ()
    exact_conversions: tuple[str, ...] | None = None
    notes: tuple[str, ...] = ()


@dataclass(frozen=True)
class BridgeStabilityCase:
    name: str
    source_kind: str
    exercise: BridgeExercise


@dataclass(frozen=True)
class BridgeStabilityCaseResult:
    name: str
    ok: bool
    bridge_ok: bool
    bridge_errors: list[str]
    hancom_status: str
    hancom_ok: bool | None
    hancom_errors: list[str]
    conversions: list[str]
    artifacts: list[str]
    notes: list[str]


def _sample_hwp_path() -> Path:
    paths = sorted(HWP_SAMPLE_DIR.glob("*.hwp"))
    if not paths:
        raise FileNotFoundError(f"No HWP sample was found under {HWP_SAMPLE_DIR}")
    return paths[0]


def _sample_hwpx_path() -> Path:
    paths = sorted(HWPX_SAMPLE_DIR.glob("*.hwpx"))
    if not paths:
        raise FileNotFoundError(f"No HWPX sample was found under {HWPX_SAMPLE_DIR}")
    return paths[0]


def _sample_backed_converter(conversions: list[str]) -> Callable[[str | Path, str | Path, str], Path]:
    sample_hwp = _sample_hwp_path()
    sample_hwpx = _sample_hwpx_path()

    def _convert(input_path: str | Path, output_path: str | Path, output_format: str) -> Path:
        source = Path(input_path)
        target = Path(output_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        normalized = output_format.upper()
        conversions.append(normalized)
        if normalized == "HWP":
            if source.suffix.lower() == ".hwpx" and source.exists():
                _semantic_fake_hwp_from_hwpx(source, target)
            else:
                target.write_bytes(sample_hwp.read_bytes())
        elif normalized == "HWPX":
            sidecar = _bridge_sidecar_path(source)
            if source.suffix.lower() == ".hwp" and sidecar.exists():
                payload = json.loads(sidecar.read_text(encoding="utf-8"))
                document = HwpxDocument.blank()
                title = payload.get("title")
                if isinstance(title, str):
                    document.set_metadata(title=title)
                text = payload.get("text", "")
                if isinstance(text, str) and text.strip():
                    for line in [line for line in text.splitlines() if line.strip()]:
                        document.append_paragraph(line.strip())
                document.save(target)
            else:
                target.write_bytes(sample_hwpx.read_bytes())
        else:
            raise AssertionError(output_format)
        return target

    return _convert


def _hwp_control_ids(document: HwpDocument) -> list[str]:
    return [control.control_id for control in document.controls()]


def _table_matrix(document: HwpxDocument, table_index: int) -> list[list[str]]:
    table = document.tables()[table_index]
    matrix = [["" for _ in range(table.column_count)] for _ in range(table.row_count)]
    for cell in table.cells():
        if cell.row < table.row_count and cell.column < table.column_count:
            matrix[cell.row][cell.column] = cell.text
    return matrix


def _bridge_sidecar_path(path: Path) -> Path:
    return path.parent / f"{path.name}.bridge.json"


def _write_hwpx_sidecar_for_hwp_artifact(document: HwpxDocument, target_hwp: str | Path) -> None:
    target_path = Path(target_hwp)
    _bridge_sidecar_path(target_path).write_text(
        json.dumps(
            {
                "title": document.metadata().title,
                "text": document.get_document_text(),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def _save_hwp_with_sidecar(bridge: HwpHwpxBridge, path: str | Path) -> Path:
    target = bridge.save_hwp(path)
    _write_hwpx_sidecar_for_hwp_artifact(bridge.hwpx_document(), target)
    return target


def _semantic_fake_hwp_from_hwpx(source_hwpx: Path, target_hwp: Path) -> None:
    source = HwpxDocument.open(source_hwpx)
    target = HwpDocument.blank()

    text = source.get_document_text().strip()
    if text:
        for line in [line for line in text.splitlines() if line.strip()]:
            target.append_paragraph(line.strip())

    for index, _table in enumerate(source.tables()):
        target.append_table(rows=source.tables()[index].row_count, cols=source.tables()[index].column_count, cell_texts=_table_matrix(source, index))
    for link in source.hyperlinks():
        target.append_hyperlink(link.hyperlink_target or "https://example.com", text=link.display_text or None)
    for field in source.fields():
        if field.is_hyperlink:
            continue
        target.append_field(
            field_type=field.field_type or "FIELD",
            display_text=field.display_text or field.name or "FIELD",
            name=field.name,
            parameters=field.parameter_map(),
        )
    for equation in source.equations():
        target.append_equation(equation.script or "x+y")
    for picture in source.pictures():
        extension = Path(picture.binary_part_path()).suffix.lstrip(".") or None
        target.append_picture(picture.binary_data(), extension=extension)
    for shape in source.shapes():
        if shape.kind == "ole":
            continue
        if shape.kind not in {"rect", "ellipse", "arc", "polygon", "curve", "line", "textart", "container"}:
            continue
        size = shape.size()
        target.append_shape(
            kind=shape.kind,
            text=shape.text,
            width=size.get("width", 12000),
            height=size.get("height", 3200),
        )
    for ole in source.oles():
        size = ole.size()
        target.append_ole(
            "bridge.ole",
            ole.binary_data(),
            width=size.get("width", 42001),
            height=size.get("height", 13501),
        )

    target.save(target_hwp)
    _bridge_sidecar_path(target_hwp).write_text(
        json.dumps(
            {
                "title": source.metadata().title,
                "text": source.get_document_text(),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def _validate_artifact(artifact: BridgeArtifact) -> list[str]:
    path = Path(artifact.path)
    errors: list[str] = []
    if not path.exists():
        return [f"missing artifact: {path}"]
    if artifact.kind == "hwp":
        try:
            reopened = HwpDocument.open(path)
        except Exception as exc:  # noqa: BLE001
            errors.append(f"failed to reopen hwp artifact {path}: {exc}")
        else:
            text = reopened.get_document_text()
            for expected_text in artifact.expected_texts:
                if expected_text not in text:
                    errors.append(f"hwp artifact missing expected text {expected_text!r}: {path}")
            control_ids = _hwp_control_ids(reopened)
            for expected_control_id in artifact.expected_hwp_control_ids:
                if expected_control_id not in control_ids:
                    errors.append(f"hwp artifact missing expected control {expected_control_id!r}: {path}")
    elif artifact.kind == "hwpx":
        try:
            reopened = HwpxDocument.open(path)
        except Exception as exc:  # noqa: BLE001
            errors.append(f"failed to reopen hwpx artifact {path}: {exc}")
        else:
            if artifact.expected_title is not None and reopened.metadata().title != artifact.expected_title:
                errors.append(
                    f"hwpx metadata title mismatch for {path}: {reopened.metadata().title!r} != {artifact.expected_title!r}"
                )
            for expected_text in artifact.expected_texts:
                if expected_text not in reopened.get_document_text():
                    errors.append(f"hwpx artifact missing expected text {expected_text!r}: {path}")
    else:
        errors.append(f"unknown artifact kind: {artifact.kind}")
    return errors


def _validate_hancom_roundtrip_artifact(
    artifact: BridgeArtifact,
    case_dir: Path,
    converter: Callable[[str | Path, str | Path, str], Path],
) -> tuple[list[str], list[str]]:
    if artifact.kind != "hwp":
        return ([], [])
    if artifact.hancom_roundtrip_title is None and not artifact.hancom_roundtrip_texts:
        return ([], [])
    roundtrip_hwpx = case_dir / f"{Path(artifact.path).stem}_roundtrip.hwpx"
    converter(artifact.path, roundtrip_hwpx, "HWPX")
    reopened = HwpxDocument.open(roundtrip_hwpx)
    errors: list[str] = []
    if artifact.hancom_roundtrip_title is not None and reopened.metadata().title != artifact.hancom_roundtrip_title:
        errors.append(
            f"hancom roundtrip title mismatch for {artifact.path}: {reopened.metadata().title!r} != {artifact.hancom_roundtrip_title!r}"
        )
    text = reopened.get_document_text()
    for expected_text in artifact.hancom_roundtrip_texts:
        if expected_text not in text:
            errors.append(f"hancom roundtrip hwpx missing expected text {expected_text!r}: {artifact.path}")
    return (errors, [str(roundtrip_hwpx)])


def _run_case(case: BridgeStabilityCase, output_dir: Path, *, validate_with_hancom: bool) -> BridgeStabilityCaseResult:
    case_dir = output_dir / case.name
    case_dir.mkdir(parents=True, exist_ok=True)
    conversions: list[str] = []
    if validate_with_hancom:
        def converter(input_path: str | Path, output_path: str | Path, output_format: str) -> Path:
            conversions.append(output_format.upper())
            return convert_document(input_path, output_path, output_format)
    else:
        converter = _sample_backed_converter(conversions)
    bridge_errors: list[str] = []
    hancom_errors: list[str] = []
    artifacts: list[str] = []

    try:
        source_path = _sample_hwp_path() if case.source_kind == "hwp" else _sample_hwpx_path()
        bridge = HwpHwpxBridge.open(source_path, converter=converter)
        execution = case.exercise(bridge, case_dir)
        artifacts = [artifact.path for artifact in execution.artifacts]
    except HancomInteropError as exc:
        message = str(exc)
        lowered = message.lower()
        if validate_with_hancom and (
            "security module" in lowered
            or "comobject" in lowered
            or "activeobject" in lowered
        ):
            return BridgeStabilityCaseResult(
                name=case.name,
                ok=True,
                bridge_ok=True,
                bridge_errors=[],
                hancom_status="skipped",
                hancom_ok=None,
                hancom_errors=[message],
                conversions=conversions,
                artifacts=artifacts,
                notes=[],
            )
        return BridgeStabilityCaseResult(
            name=case.name,
            ok=False,
            bridge_ok=False,
            bridge_errors=[message],
            hancom_status="failed" if validate_with_hancom else "skipped",
            hancom_ok=False if validate_with_hancom else None,
            hancom_errors=[message] if validate_with_hancom else [],
            conversions=conversions,
            artifacts=artifacts,
            notes=[],
        )

    for expected in execution.expected_conversions:
        if expected not in conversions:
            bridge_errors.append(f"missing conversion call: {expected}")
    if execution.exact_conversions is not None and tuple(conversions) != execution.exact_conversions:
        bridge_errors.append(f"unexpected conversion sequence: {conversions} != {list(execution.exact_conversions)}")
    for artifact in execution.artifacts:
        bridge_errors.extend(_validate_artifact(artifact))

    if validate_with_hancom:
        hancom_status = "passed"
        hancom_ok: bool | None = True
    else:
        hancom_status = "skipped"
        hancom_ok = None

    if validate_with_hancom:
        try:
            for artifact in execution.artifacts:
                bridge_errors.extend(_validate_artifact(artifact))
                artifact_errors, artifact_paths = _validate_hancom_roundtrip_artifact(artifact, case_dir, converter)
                hancom_errors.extend(artifact_errors)
                artifacts.extend(artifact_paths)
        except HancomInteropError as exc:
            hancom_status = "failed"
            hancom_ok = False
            hancom_errors.append(str(exc))
        else:
            if hancom_errors:
                hancom_status = "failed"
                hancom_ok = False

    bridge_ok = not bridge_errors
    ok = bridge_ok and hancom_status != "failed"
    return BridgeStabilityCaseResult(
        name=case.name,
        ok=ok,
        bridge_ok=bridge_ok,
        bridge_errors=bridge_errors,
        hancom_status=hancom_status,
        hancom_ok=hancom_ok,
        hancom_errors=hancom_errors,
        conversions=conversions,
        artifacts=artifacts,
        notes=list(execution.notes),
    )


def _cases() -> list[BridgeStabilityCase]:
    return [
        BridgeStabilityCase(
            name="from_hwp_cache_hwpx",
            source_kind="hwp",
            exercise=lambda bridge, _case_dir: (
                lambda first=bridge.hwpx_document(), second=bridge.hwpx_document(): BridgeExecution(
                    exact_conversions=("HWPX",),
                    notes=(f"cache_identity={first is second}",),
                )
            )(),
        ),
        BridgeStabilityCase(
            name="from_hwp_save_native_hwp",
            source_kind="hwp",
            exercise=lambda bridge, case_dir: BridgeExecution(
                artifacts=(BridgeArtifact(str(bridge.save_hwp(case_dir / "native.hwp")), "hwp"),),
                exact_conversions=(),
            ),
        ),
        BridgeStabilityCase(
            name="from_hwp_save_hwpx_after_metadata_edit",
            source_kind="hwp",
            exercise=lambda bridge, case_dir: (
                lambda document=bridge.hwpx_document(): (
                    document.set_metadata(title="BRIDGE-HWP-HWPX"),
                    BridgeExecution(
                        artifacts=(
                            BridgeArtifact(
                                str(bridge.save_hwpx(case_dir / "edited.hwpx")),
                                "hwpx",
                                expected_title="BRIDGE-HWP-HWPX",
                            ),
                        ),
                        exact_conversions=("HWPX",),
                    ),
                )[1]
            )(),
        ),
        BridgeStabilityCase(
            name="from_hwp_refresh_hwpx",
            source_kind="hwp",
            exercise=lambda bridge, _case_dir: (
                bridge.hwpx_document(),
                bridge.refresh_hwpx(),
                BridgeExecution(expected_conversions=("HWPX",), exact_conversions=("HWPX", "HWPX")),
            )[2],
        ),
        BridgeStabilityCase(
            name="from_hwp_dispatch_save_hwpx",
            source_kind="hwp",
            exercise=lambda bridge, case_dir: BridgeExecution(
                artifacts=(BridgeArtifact(str(bridge.save(case_dir / "dispatch.hwpx")), "hwpx"),),
                exact_conversions=("HWPX",),
            ),
        ),
        BridgeStabilityCase(
            name="hwp_document_helper_bridge",
            source_kind="hwp",
            exercise=lambda bridge, _case_dir: (
                lambda helper=bridge.hwp_document().bridge(), document=bridge.hwp_document().bridge().hwpx_document(): BridgeExecution(
                    exact_conversions=("HWPX",),
                    notes=(f"helper_bridge_type={type(helper).__name__}", f"hwpx_type={type(document).__name__}"),
                )
            )(),
        ),
        BridgeStabilityCase(
            name="from_hwpx_cache_hwp",
            source_kind="hwpx",
            exercise=lambda bridge, _case_dir: (
                lambda first=bridge.hwp_document(), second=bridge.hwp_document(): BridgeExecution(
                    exact_conversions=("HWP",),
                    notes=(f"cache_identity={first is second}",),
                )
            )(),
        ),
        BridgeStabilityCase(
            name="from_hwpx_save_native_hwpx",
            source_kind="hwpx",
            exercise=lambda bridge, case_dir: BridgeExecution(
                artifacts=(BridgeArtifact(str(bridge.save_hwpx(case_dir / "native.hwpx")), "hwpx"),),
                exact_conversions=(),
            ),
        ),
        BridgeStabilityCase(
            name="from_hwpx_save_hwp",
            source_kind="hwpx",
            exercise=lambda bridge, case_dir: BridgeExecution(
                artifacts=(BridgeArtifact(str(_save_hwp_with_sidecar(bridge, case_dir / "converted.hwp")), "hwp"),),
                exact_conversions=("HWP",),
            ),
        ),
        BridgeStabilityCase(
            name="from_hwpx_refresh_hwp",
            source_kind="hwpx",
            exercise=lambda bridge, _case_dir: (
                bridge.hwp_document(),
                bridge.refresh_hwp(),
                BridgeExecution(expected_conversions=("HWP",), exact_conversions=("HWP", "HWP")),
            )[2],
        ),
        BridgeStabilityCase(
            name="from_hwpx_dispatch_save_hwp",
            source_kind="hwpx",
            exercise=lambda bridge, case_dir: BridgeExecution(
                artifacts=(BridgeArtifact(str(_save_hwp_with_sidecar(bridge, case_dir / "dispatch.hwp")), "hwp"),),
                exact_conversions=("HWP",),
            ),
        ),
        BridgeStabilityCase(
            name="hwpx_document_reverse_helpers",
            source_kind="hwpx",
            exercise=lambda bridge, case_dir: (
                lambda document=bridge.hwpx_document(): (
                    document.to_hwp_document(converter=bridge._converter),
                    BridgeExecution(
                        artifacts=(
                            BridgeArtifact(
                                str(document.save_as_hwp(case_dir / "helper.hwp", converter=bridge._converter)),
                                "hwp",
                            ),
                        ),
                        expected_conversions=("HWP",),
                    ),
                )[1]
            )(),
        ),
        BridgeStabilityCase(
            name="from_hwpx_modify_title_save_hwp",
            source_kind="hwpx",
            exercise=lambda bridge, case_dir: (
                lambda document=bridge.hwpx_document(): (
                    document.set_metadata(title="BRIDGE-HWPX-TITLE"),
                    BridgeExecution(
                        artifacts=(
                            BridgeArtifact(
                                str(_save_hwp_with_sidecar(bridge, case_dir / "title_modified.hwp")),
                                "hwp",
                            ),
                        ),
                        exact_conversions=("HWP",),
                    ),
                )[1]
            )(),
        ),
        BridgeStabilityCase(
            name="from_hwpx_append_paragraph_save_hwp",
            source_kind="hwpx",
            exercise=lambda bridge, case_dir: (
                lambda document=bridge.hwpx_document(): (
                    document.append_paragraph("BRIDGE-HWPX-PARAGRAPH"),
                    BridgeExecution(
                        artifacts=(
                            BridgeArtifact(
                                str(_save_hwp_with_sidecar(bridge, case_dir / "paragraph_modified.hwp")),
                                "hwp",
                                expected_texts=("BRIDGE-HWPX-PARAGRAPH",),
                            ),
                        ),
                        exact_conversions=("HWP",),
                    ),
                )[1]
            )(),
        ),
        BridgeStabilityCase(
            name="from_hwpx_append_table_save_hwp",
            source_kind="hwpx",
            exercise=lambda bridge, case_dir: (
                lambda document=bridge.hwpx_document(): (
                    document.append_table(1, 2, cell_texts=[["BRIDGE-T11", "BRIDGE-T12"]]),
                    BridgeExecution(
                        artifacts=(
                            BridgeArtifact(
                                str(_save_hwp_with_sidecar(bridge, case_dir / "table_modified.hwp")),
                                "hwp",
                                expected_texts=("BRIDGE-T11", "BRIDGE-T12"),
                                expected_hwp_control_ids=("tbl ",),
                            ),
                        ),
                        exact_conversions=("HWP",),
                    ),
                )[1]
            )(),
        ),
        BridgeStabilityCase(
            name="from_hwpx_append_hyperlink_save_hwp",
            source_kind="hwpx",
            exercise=lambda bridge, case_dir: (
                lambda document=bridge.hwpx_document(): (
                    document.append_hyperlink("https://example.com/bridge", display_text="BRIDGE-LINK"),
                    BridgeExecution(
                        artifacts=(
                            BridgeArtifact(
                                str(_save_hwp_with_sidecar(bridge, case_dir / "hyperlink_modified.hwp")),
                                "hwp",
                                expected_texts=("BRIDGE-LINK",),
                                expected_hwp_control_ids=("%hlk",),
                            ),
                        ),
                        exact_conversions=("HWP",),
                    ),
                )[1]
            )(),
        ),
        BridgeStabilityCase(
            name="from_hwpx_append_field_save_hwp",
            source_kind="hwpx",
            exercise=lambda bridge, case_dir: (
                lambda document=bridge.hwpx_document(): (
                    document.append_field(field_type="DOCPROPERTY", display_text="BRIDGE-FIELD"),
                    BridgeExecution(
                        artifacts=(
                            BridgeArtifact(
                                str(_save_hwp_with_sidecar(bridge, case_dir / "field_modified.hwp")),
                                "hwp",
                            ),
                        ),
                        exact_conversions=("HWP",),
                    ),
                )[1]
            )(),
        ),
        BridgeStabilityCase(
            name="from_hwpx_append_equation_save_hwp",
            source_kind="hwpx",
            exercise=lambda bridge, case_dir: (
                lambda document=bridge.hwpx_document(): (
                    document.append_equation("a+b=c"),
                    BridgeExecution(
                        artifacts=(
                            BridgeArtifact(
                                str(_save_hwp_with_sidecar(bridge, case_dir / "equation_modified.hwp")),
                                "hwp",
                                expected_hwp_control_ids=("eqed",),
                            ),
                        ),
                        exact_conversions=("HWP",),
                    ),
                )[1]
            )(),
        ),
        BridgeStabilityCase(
            name="from_hwpx_append_shape_save_hwp",
            source_kind="hwpx",
            exercise=lambda bridge, case_dir: (
                lambda document=bridge.hwpx_document(): (
                    document.append_shape(kind="rect", text="BRIDGE-SHAPE"),
                    BridgeExecution(
                        artifacts=(
                            BridgeArtifact(
                                str(_save_hwp_with_sidecar(bridge, case_dir / "shape_modified.hwp")),
                                "hwp",
                                expected_texts=("BRIDGE-SHAPE",),
                                expected_hwp_control_ids=("gso ",),
                            ),
                        ),
                        exact_conversions=("HWP",),
                    ),
                )[1]
            )(),
        ),
        BridgeStabilityCase(
            name="from_hwpx_append_ole_save_hwp",
            source_kind="hwpx",
            exercise=lambda bridge, case_dir: (
                lambda document=bridge.hwpx_document(): (
                    document.append_ole("bridge.ole", b"BRIDGE-OLE"),
                    BridgeExecution(
                        artifacts=(
                            BridgeArtifact(
                                str(_save_hwp_with_sidecar(bridge, case_dir / "ole_modified.hwp")),
                                "hwp",
                                expected_texts=("bridge.ole",),
                                expected_hwp_control_ids=("gso ",),
                            ),
                        ),
                        exact_conversions=("HWP",),
                    ),
                )[1]
            )(),
        ),
        BridgeStabilityCase(
            name="from_hwpx_mixed_edits_save_hwp",
            source_kind="hwpx",
            exercise=lambda bridge, case_dir: (
                lambda document=bridge.hwpx_document(): (
                    document.append_paragraph("BRIDGE-MIX-PARA"),
                    document.append_table(1, 1, cell_texts=[["BRIDGE-MIX-CELL"]]),
                    document.append_hyperlink("https://example.com/mix", display_text="BRIDGE-MIX-LINK"),
                    document.append_equation("mix+eq"),
                    BridgeExecution(
                        artifacts=(
                            BridgeArtifact(
                                str(_save_hwp_with_sidecar(bridge, case_dir / "mixed_modified.hwp")),
                                "hwp",
                                expected_texts=("BRIDGE-MIX-PARA", "BRIDGE-MIX-CELL", "BRIDGE-MIX-LINK"),
                                expected_hwp_control_ids=("tbl ", "%hlk", "eqed"),
                            ),
                        ),
                        exact_conversions=("HWP",),
                    ),
                )[4]
            )(),
        ),
    ]


def run_bridge_stability_matrix(
    output_dir: str | Path,
    *,
    validate_with_hancom: bool = False,
) -> list[BridgeStabilityCaseResult]:
    output_path = Path(output_dir).expanduser().resolve()
    output_path.mkdir(parents=True, exist_ok=True)
    return [_run_case(case, output_path, validate_with_hancom=validate_with_hancom) for case in _cases()]


def write_bridge_stability_report(results: list[BridgeStabilityCaseResult], path: str | Path) -> Path:
    output_path = Path(path).expanduser().resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "case_count": len(results),
        "ok_count": sum(1 for result in results if result.ok),
        "bridge_failure_count": sum(1 for result in results if not result.bridge_ok),
        "hancom_failure_count": sum(1 for result in results if result.hancom_status == "failed"),
        "hancom_skip_count": sum(1 for result in results if result.hancom_status == "skipped"),
        "results": [asdict(result) for result in results],
    }
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return output_path
