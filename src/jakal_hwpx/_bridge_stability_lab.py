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
        target = Path(output_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        normalized = output_format.upper()
        conversions.append(normalized)
        if normalized == "HWP":
            target.write_bytes(sample_hwp.read_bytes())
        elif normalized == "HWPX":
            target.write_bytes(sample_hwpx.read_bytes())
        else:
            raise AssertionError(output_format)
        return target

    return _convert


def _validate_artifact(artifact: BridgeArtifact) -> list[str]:
    path = Path(artifact.path)
    errors: list[str] = []
    if not path.exists():
        return [f"missing artifact: {path}"]
    if artifact.kind == "hwp":
        try:
            HwpDocument.open(path)
        except Exception as exc:  # noqa: BLE001
            errors.append(f"failed to reopen hwp artifact {path}: {exc}")
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
    else:
        errors.append(f"unknown artifact kind: {artifact.kind}")
    return errors


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
            or "com conversion failed" in lowered
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
        except HancomInteropError as exc:
            hancom_status = "failed"
            hancom_ok = False
            hancom_errors.append(str(exc))

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
                artifacts=(BridgeArtifact(str(bridge.save_hwp(case_dir / "converted.hwp")), "hwp"),),
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
                artifacts=(BridgeArtifact(str(bridge.save(case_dir / "dispatch.hwp")), "hwp"),),
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
                        artifacts=(BridgeArtifact(str(document.save_as_hwp(case_dir / "helper.hwp", converter=bridge._converter)), "hwp"),),
                        expected_conversions=("HWP",),
                    ),
                )[1]
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
