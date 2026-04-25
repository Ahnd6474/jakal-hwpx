from __future__ import annotations

import json
import platform
import subprocess
import sys
import tomllib
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from .document import HwpxDocument
from .hwp_document import HwpDocument


@dataclass(frozen=True)
class CorpusSampleResult:
    category: str
    path: str
    format: str
    ok: bool
    errors: list[str]


@dataclass(frozen=True)
class CorpusValidationReport:
    sample_count: int
    failure_count: int
    covered_categories: list[str]
    missing_categories: list[str]
    results: list[CorpusSampleResult]

    @property
    def ok(self) -> bool:
        return self.failure_count == 0 and not self.missing_categories


@dataclass(frozen=True)
class HancomCorpusSmokeResult:
    available: bool
    ok: bool
    case_count: int
    failure_count: int
    dialog_count: int
    report_path: str | None
    errors: list[str]


def repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def contract_path(root: str | Path | None = None) -> Path:
    base = Path(root).expanduser().resolve() if root is not None else repo_root()
    return base / "stability_contract.json"


def pyproject_path(root: str | Path | None = None) -> Path:
    base = Path(root).expanduser().resolve() if root is not None else repo_root()
    return base / "pyproject.toml"


def load_release_contract(root: str | Path | None = None) -> dict[str, Any]:
    path = contract_path(root)
    return json.loads(path.read_text(encoding="utf-8"))


def _supported_python_versions_from_pyproject(path: Path) -> list[str]:
    project = tomllib.loads(path.read_text(encoding="utf-8"))["project"]
    classifiers = project.get("classifiers", [])
    versions = sorted(
        {
            classifier.rsplit("::", 1)[-1].strip()
            for classifier in classifiers
            if classifier.startswith("Programming Language :: Python :: 3.")
        }
    )
    return versions


def validate_contract_definition(contract: dict[str, Any], root: str | Path | None = None) -> list[str]:
    errors: list[str] = []

    if not contract.get("unsupported_scopes"):
        errors.append("stability_contract.json must list at least one unsupported scope.")

    python_versions = list(contract.get("python_versions", []))
    if not python_versions:
        errors.append("stability_contract.json must list supported python_versions.")
    else:
        pyproject_versions = _supported_python_versions_from_pyproject(pyproject_path(root))
        if sorted(python_versions) != pyproject_versions:
            errors.append(
                f"contract python_versions {sorted(python_versions)!r} do not match pyproject classifiers {pyproject_versions!r}."
            )

    supported_formats = {entry.get("format") for entry in contract.get("supported_scopes", [])}
    for required_format in ("HWPX", "HWP", "HWP<->HWPX"):
        if required_format not in supported_formats:
            errors.append(f"supported_scopes is missing format {required_format!r}.")

    gate_profiles = contract.get("gate_profiles", {})
    for required_profile in ("ci", "release"):
        if required_profile not in gate_profiles:
            errors.append(f"gate_profiles is missing profile {required_profile!r}.")

    minimum_case_counts = contract.get("minimum_case_counts", {})
    for key in ("hwpx_stability", "hwp_stability", "bridge_stability"):
        value = minimum_case_counts.get(key)
        if not isinstance(value, int) or value <= 0:
            errors.append(f"minimum_case_counts[{key!r}] must be a positive integer.")

    required_categories = set(contract.get("required_corpus_categories", []))
    sample_categories = {sample.get("category") for sample in contract.get("corpus_samples", [])}
    missing_categories = sorted(category for category in required_categories if category not in sample_categories)
    if missing_categories:
        errors.append(f"corpus_samples is missing required categories: {missing_categories!r}.")

    sample_paths = [sample.get("path") for sample in contract.get("corpus_samples", [])]
    duplicate_paths = sorted({path for path in sample_paths if sample_paths.count(path) > 1})
    if duplicate_paths:
        errors.append(f"corpus_samples contains duplicate paths: {duplicate_paths!r}.")

    return errors


def _hwp_control_ids(document: HwpDocument) -> list[str]:
    return [control.control_id for control in document.controls()]


def _hwpx_has_feature(document: HwpxDocument, feature: str) -> bool:
    if feature == "header":
        return bool(document.headers())
    if feature == "footer":
        return bool(document.footers())
    if feature == "footnote":
        return any(note.kind == "footNote" for note in document.notes())
    if feature == "endnote":
        return any(note.kind == "endNote" for note in document.notes())
    if feature == "bookmark":
        return bool(document.bookmarks())
    if feature == "field":
        return bool(document.fields())
    if feature == "hyperlink":
        return bool(document.hyperlinks())
    if feature == "table":
        return bool(document.tables())
    if feature == "picture":
        return bool(document.pictures())
    if feature == "equation":
        return bool(document.equations())
    if feature == "shape":
        return bool(document.shapes())
    if feature == "ole":
        return bool(document.oles())
    raise ValueError(f"Unsupported HWPX feature label: {feature}")


def _validate_hwpx_sample(root: Path, sample: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    sample_path = root / sample["path"]
    document = HwpxDocument.open(sample_path)

    expected_title = sample.get("expected_title")
    if expected_title is not None and document.metadata().title != expected_title:
        errors.append(f"title {document.metadata().title!r} != {expected_title!r}")

    min_sections = int(sample.get("min_sections", 1))
    if len(document.sections) < min_sections:
        errors.append(f"section_count {len(document.sections)} < {min_sections}")

    document_text = document.get_document_text()
    for expected_text in sample.get("expected_texts", []):
        if expected_text not in document_text:
            errors.append(f"missing expected text {expected_text!r}")

    for feature in sample.get("required_features", []):
        if not _hwpx_has_feature(document, feature):
            errors.append(f"missing required feature {feature!r}")

    return errors


def _validate_hwp_sample(root: Path, sample: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    sample_path = root / sample["path"]
    document = HwpDocument.open(sample_path)

    min_sections = int(sample.get("min_sections", 1))
    if len(document.sections()) < min_sections:
        errors.append(f"section_count {len(document.sections())} < {min_sections}")

    preview_text = document.preview_text()
    for expected_text in sample.get("expected_preview_texts", []):
        if expected_text not in preview_text:
            errors.append(f"missing expected preview text {expected_text!r}")

    document_text = document.get_document_text()
    for expected_text in sample.get("expected_texts", []):
        if expected_text not in document_text:
            errors.append(f"missing expected text {expected_text!r}")

    control_ids = _hwp_control_ids(document)
    for control_id in sample.get("required_control_ids", []):
        if control_id not in control_ids:
            errors.append(f"missing required control id {control_id!r}")

    return errors


def validate_corpus_samples(contract: dict[str, Any], root: str | Path | None = None) -> CorpusValidationReport:
    base = Path(root).expanduser().resolve() if root is not None else repo_root()
    results: list[CorpusSampleResult] = []

    for sample in contract.get("corpus_samples", []):
        sample_path = base / sample["path"]
        errors: list[str] = []
        if not sample_path.exists():
            errors.append(f"missing file: {sample_path}")
        else:
            try:
                if sample["format"] == "hwpx":
                    errors.extend(_validate_hwpx_sample(base, sample))
                elif sample["format"] == "hwp":
                    errors.extend(_validate_hwp_sample(base, sample))
                else:
                    errors.append(f"unsupported sample format {sample['format']!r}")
            except Exception as exc:  # noqa: BLE001
                errors.append(str(exc))
        results.append(
            CorpusSampleResult(
                category=sample["category"],
                path=sample["path"],
                format=sample["format"],
                ok=not errors,
                errors=errors,
            )
        )

    covered_categories = sorted({result.category for result in results})
    required_categories = sorted(set(contract.get("required_corpus_categories", [])))
    missing_categories = [category for category in required_categories if category not in covered_categories]
    failure_count = sum(1 for result in results if not result.ok)
    return CorpusValidationReport(
        sample_count=len(results),
        failure_count=failure_count,
        covered_categories=covered_categories,
        missing_categories=missing_categories,
        results=results,
    )


def write_corpus_validation_report(report: CorpusValidationReport, output_path: str | Path) -> Path:
    path = Path(output_path).expanduser().resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "sample_count": report.sample_count,
        "failure_count": report.failure_count,
        "covered_categories": report.covered_categories,
        "missing_categories": report.missing_categories,
        "results": [asdict(result) for result in report.results],
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def run_hancom_corpus_smoke(output_root: str | Path, root: str | Path | None = None) -> HancomCorpusSmokeResult:
    base = Path(root).expanduser().resolve() if root is not None else repo_root()
    smoke_output_root = Path(output_root).expanduser().resolve()
    report_path = smoke_output_root / "corpus-smoke-report.json"

    if platform.system() != "Windows":
        return HancomCorpusSmokeResult(
            available=False,
            ok=False,
            case_count=0,
            failure_count=0,
            dialog_count=0,
            report_path=str(report_path),
            errors=["Hancom corpus smoke requires Windows."],
        )

    script_path = base / "scripts" / "run_hancom_corpus_smoke_validation.ps1"
    if not script_path.exists():
        return HancomCorpusSmokeResult(
            available=False,
            ok=False,
            case_count=0,
            failure_count=0,
            dialog_count=0,
            report_path=str(report_path),
            errors=[f"Hancom corpus smoke script was not found: {script_path}"],
        )

    smoke_output_root.mkdir(parents=True, exist_ok=True)
    command = [
        "powershell",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script_path),
        "-OutputRoot",
        str(smoke_output_root),
        "-AllowExistingHwpProcesses",
    ]
    completed = subprocess.run(
        command,
        cwd=base,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )

    errors: list[str] = []
    if completed.returncode != 0:
        detail = completed.stderr.strip() or completed.stdout.strip() or f"exit code {completed.returncode}"
        errors.append(detail)

    entries: list[dict[str, Any]] = []
    if report_path.exists():
        decoded_report = json.loads(report_path.read_text(encoding="utf-8-sig"))
        if isinstance(decoded_report, list):
            entries = [entry for entry in decoded_report if isinstance(entry, dict)]
        elif isinstance(decoded_report, dict):
            entries = [decoded_report]
        else:
            errors.append(f"Hancom corpus smoke report has unsupported shape: {type(decoded_report).__name__}")
    else:
        errors.append(f"Hancom corpus smoke did not produce report {report_path}")

    failure_count = 0
    dialog_count = 0
    for entry in entries:
        if entry.get("Status") != "ok":
            failure_count += 1
            errors.append(f"{entry.get('Name')}: {entry.get('Message') or entry.get('Status')}")
        raw_log_path = entry.get("LogPath")
        if not raw_log_path:
            continue
        log_path = Path(raw_log_path)
        if not log_path.exists():
            errors.append(f"missing Hancom run log: {log_path}")
            continue
        try:
            log_entries = json.loads(log_path.read_text(encoding="utf-8-sig"))
        except Exception as exc:  # noqa: BLE001
            errors.append(f"failed to read Hancom run log {log_path}: {exc}")
            continue
        dialog_count += sum(1 for log_entry in log_entries if log_entry.get("kind") == "dialog")

    if dialog_count:
        errors.append(f"Hancom corpus smoke recorded {dialog_count} dialog events.")

    return HancomCorpusSmokeResult(
        available=True,
        ok=not errors,
        case_count=len(entries),
        failure_count=failure_count,
        dialog_count=dialog_count,
        report_path=str(report_path),
        errors=errors,
    )


def python_runtime_summary() -> dict[str, str]:
    return {
        "executable": sys.executable,
        "version": platform.python_version(),
        "implementation": platform.python_implementation(),
        "platform": platform.platform(),
    }
