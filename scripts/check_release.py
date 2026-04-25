from __future__ import annotations

import argparse
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import tarfile
import zipfile
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"
DIST_DIR = REPO_ROOT / "dist"
PYPROJECT = REPO_ROOT / "pyproject.toml"
INIT_FILE = REPO_ROOT / "src" / "jakal_hwpx" / "__init__.py"

if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from jakal_hwpx._bridge_stability_lab import run_bridge_stability_matrix, write_bridge_stability_report  # noqa: E402
from jakal_hwpx._hwp_stability_lab import run_hwp_stability_matrix, write_hwp_stability_report  # noqa: E402
from jakal_hwpx._release_gate import (  # noqa: E402
    load_release_contract,
    python_runtime_summary,
    run_hancom_corpus_smoke,
    validate_contract_definition,
    validate_corpus_samples,
    write_corpus_validation_report,
)
from jakal_hwpx._stability_lab import run_stability_matrix, write_stability_report  # noqa: E402


REQUIRED_WHEEL_PATTERNS = (
    "jakal_hwpx/__init__.py",
    "jakal_hwpx/document.py",
    "jakal_hwpx/elements.py",
    "jakal_hwpx/parts.py",
    "jakal_hwpx/py.typed",
)
REQUIRED_SDIST_PATTERNS = tuple(f"src/{path}" for path in REQUIRED_WHEEL_PATTERNS)
OPTIONAL_TOP_LEVEL_FILES = ("LICENSE", "THIRD_PARTY_NOTICES.md")


def extract_version(pattern: str, text: str, source: Path) -> str:
    match = re.search(pattern, text, re.MULTILINE)
    if match is None:
        raise RuntimeError(f"Could not find version in {source}")
    return match.group(1)


def read_versions() -> str:
    pyproject_version = extract_version(r'^version = "([^"]+)"$', PYPROJECT.read_text(encoding="utf-8"), PYPROJECT)
    init_version = extract_version(r'^__version__ = "([^"]+)"$', INIT_FILE.read_text(encoding="utf-8"), INIT_FILE)
    if pyproject_version != init_version:
        raise RuntimeError(
            f"Version mismatch: pyproject.toml has {pyproject_version}, src/jakal_hwpx/__init__.py has {init_version}"
        )
    return pyproject_version


def run(*args: str) -> None:
    print("+", " ".join(args))
    subprocess.run(args, cwd=REPO_ROOT, check=True)


def _workspace_temp_env(temp_root: Path) -> dict[str, str]:
    temp_root.mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    env["TMP"] = str(temp_root)
    env["TEMP"] = str(temp_root)
    env["TMPDIR"] = str(temp_root)
    env["JAKAL_HWPX_TEMP_ROOT"] = str(temp_root)
    return env


def rebuild_dist(*, temp_root: Path) -> list[Path]:
    shutil.rmtree(DIST_DIR, ignore_errors=True)
    DIST_DIR.mkdir(parents=True, exist_ok=True)
    env = _workspace_temp_env(temp_root)
    print("+", " ".join((sys.executable, "-m", "build", "--sdist", "--wheel", "--outdir", str(DIST_DIR), "--no-isolation")))
    subprocess.run(
        (sys.executable, "-m", "build", "--sdist", "--wheel", "--outdir", str(DIST_DIR), "--no-isolation"),
        cwd=REPO_ROOT,
        check=True,
        env=env,
    )
    artifacts = sorted(DIST_DIR.iterdir())
    if not artifacts:
        raise RuntimeError("No artifacts were produced in dist/")
    print("+", " ".join((sys.executable, "-m", "twine", "check", *(str(path) for path in artifacts))))
    subprocess.run(
        (sys.executable, "-m", "twine", "check", *(str(path) for path in artifacts)),
        cwd=REPO_ROOT,
        check=True,
        env=env,
    )
    return artifacts


def ensure_required_entries(artifacts: list[Path], version: str) -> list[str]:
    warnings: list[str] = []
    wheel_name = f"jakal_hwpx-{version}-py3-none-any.whl"
    sdist_name = f"jakal_hwpx-{version}.tar.gz"
    wheel = next((path for path in artifacts if path.name == wheel_name), None)
    sdist = next((path for path in artifacts if path.name == sdist_name), None)
    if wheel is None or sdist is None:
        raise RuntimeError(f"Expected {wheel_name} and {sdist_name} in dist/, found {[path.name for path in artifacts]}")

    with zipfile.ZipFile(wheel) as archive:
        names = set(archive.namelist())
        for required in REQUIRED_WHEEL_PATTERNS:
            if required not in names:
                raise RuntimeError(f"Wheel is missing required file: {required}")

    prefix = f"jakal_hwpx-{version}/"
    with tarfile.open(sdist, "r:gz") as archive:
        names = set(archive.getnames())
        for required in REQUIRED_SDIST_PATTERNS:
            candidate = prefix + required
            if candidate not in names:
                raise RuntimeError(f"sdist is missing required file: {candidate}")

    missing_optional = [name for name in OPTIONAL_TOP_LEVEL_FILES if not (REPO_ROOT / name).exists()]
    if missing_optional:
        warnings.append(f"Missing top-level release files: {', '.join(missing_optional)}")
    return warnings


def _write_text(path: Path, value: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(value, encoding="utf-8")


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _run_pytest(output_dir: Path, *, temp_root: Path) -> dict[str, Any]:
    output_dir.mkdir(parents=True, exist_ok=True)
    basetemp = output_dir / "basetemp"
    basetemp.mkdir(parents=True, exist_ok=True)
    shutil.rmtree(REPO_ROOT / ".pytest_cache", ignore_errors=True)
    for stale in (REPO_ROOT / "tests").glob("pytest-cache-files-*"):
        shutil.rmtree(stale, ignore_errors=True)
    command = [
        sys.executable,
        "-m",
        "pytest",
        "-q",
        "-p",
        "no:cacheprovider",
        "--ignore-glob=tests/pytest-cache-files-*",
        f"--basetemp={basetemp}",
    ]
    completed = subprocess.run(
        command,
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
        check=False,
        env=_workspace_temp_env(temp_root),
    )
    _write_text(output_dir / "pytest.stdout.log", completed.stdout)
    _write_text(output_dir / "pytest.stderr.log", completed.stderr)
    detail = completed.stderr.strip() or completed.stdout.strip()
    return {
        "command": command,
        "returncode": completed.returncode,
        "ok": completed.returncode == 0,
        "error": detail if completed.returncode != 0 else "",
        "stdout_log": str((output_dir / "pytest.stdout.log").resolve()),
        "stderr_log": str((output_dir / "pytest.stderr.log").resolve()),
    }


def _summarize_hwpx_stability(results: list[Any], minimum_case_count: int) -> dict[str, Any]:
    failures = [result for result in results if not result.ok]
    signature_failures = sum(1 for result in results if result.signature_mismatch)
    control_failures = sum(1 for result in results if result.control_errors)
    return {
        "case_count": len(results),
        "minimum_case_count": minimum_case_count,
        "failure_count": len(failures),
        "signature_failure_count": signature_failures,
        "control_failure_count": control_failures,
        "ok": not failures and len(results) >= minimum_case_count,
    }


def _summarize_hwp_stability(results: list[Any], minimum_case_count: int) -> dict[str, Any]:
    failures = [result for result in results if not result.ok]
    binary_failures = [result for result in results if not result.binary_ok]
    hancom_failures = [result for result in results if result.hancom_status == "failed"]
    hancom_skips = [result for result in results if result.hancom_status == "skipped"]
    return {
        "case_count": len(results),
        "minimum_case_count": minimum_case_count,
        "failure_count": len(failures),
        "binary_failure_count": len(binary_failures),
        "hancom_failure_count": len(hancom_failures),
        "hancom_skip_count": len(hancom_skips),
        "ok": not failures and len(results) >= minimum_case_count,
    }


def _summarize_bridge_stability(results: list[Any], minimum_case_count: int) -> dict[str, Any]:
    failures = [result for result in results if not result.ok]
    bridge_failures = [result for result in results if not result.bridge_ok]
    hancom_failures = [result for result in results if result.hancom_status == "failed"]
    hancom_skips = [result for result in results if result.hancom_status == "skipped"]
    return {
        "case_count": len(results),
        "minimum_case_count": minimum_case_count,
        "failure_count": len(failures),
        "bridge_failure_count": len(bridge_failures),
        "hancom_failure_count": len(hancom_failures),
        "hancom_skip_count": len(hancom_skips),
        "ok": not failures and len(results) >= minimum_case_count,
    }


def _validate_gate_thresholds(
    *,
    contract_errors: list[str],
    profile_name: str,
    profile_config: dict[str, Any],
    thresholds: dict[str, Any],
    corpus_report: dict[str, Any],
    pytest_report: dict[str, Any],
    hwpx_summary: dict[str, Any],
    hwp_summary: dict[str, Any],
    bridge_summary: dict[str, Any],
    hancom_corpus_summary: dict[str, Any] | None,
    precondition_errors: list[str],
) -> list[str]:
    failures = list(contract_errors)
    failures.extend(precondition_errors)

    if corpus_report["failure_count"] > thresholds["corpus_failures_max"]:
        failures.append(
            f"Corpus validation failures {corpus_report['failure_count']} exceed {thresholds['corpus_failures_max']}."
        )
    if corpus_report["missing_categories"]:
        failures.append(f"Corpus validation is missing categories: {corpus_report['missing_categories']!r}.")

    if pytest_report["returncode"] > thresholds["pytest_failures_max"]:
        failures.append(f"pytest failed with return code {pytest_report['returncode']}.")

    if hwpx_summary["case_count"] < hwpx_summary["minimum_case_count"]:
        failures.append(
            f"HWPX stability case_count {hwpx_summary['case_count']} is below {hwpx_summary['minimum_case_count']}."
        )
    if hwpx_summary["failure_count"] > thresholds["hwpx_stability_failures_max"]:
        failures.append(
            f"HWPX stability failures {hwpx_summary['failure_count']} exceed {thresholds['hwpx_stability_failures_max']}."
        )

    if hwp_summary["case_count"] < hwp_summary["minimum_case_count"]:
        failures.append(f"HWP stability case_count {hwp_summary['case_count']} is below {hwp_summary['minimum_case_count']}.")
    if hwp_summary["failure_count"] > thresholds["hwp_stability_failures_max"]:
        failures.append(f"HWP stability failures {hwp_summary['failure_count']} exceed {thresholds['hwp_stability_failures_max']}.")

    if bridge_summary["case_count"] < bridge_summary["minimum_case_count"]:
        failures.append(
            f"Bridge stability case_count {bridge_summary['case_count']} is below {bridge_summary['minimum_case_count']}."
        )
    if bridge_summary["failure_count"] > thresholds["bridge_stability_failures_max"]:
        failures.append(
            f"Bridge stability failures {bridge_summary['failure_count']} exceed {thresholds['bridge_stability_failures_max']}."
        )

    if profile_config["require_hancom"]:
        if hwp_summary["hancom_skip_count"]:
            failures.append(f"Release profile does not allow skipped Hancom HWP cases: {hwp_summary['hancom_skip_count']}.")
        if bridge_summary["hancom_skip_count"]:
            failures.append(
                f"Release profile does not allow skipped Hancom bridge cases: {bridge_summary['hancom_skip_count']}."
            )
        if hancom_corpus_summary is None:
            failures.append("Release profile requires Hancom corpus smoke results.")
        else:
            if not hancom_corpus_summary["available"]:
                failures.append(f"Hancom corpus smoke is unavailable: {hancom_corpus_summary['errors']!r}.")
            if hancom_corpus_summary["failure_count"] > thresholds["hancom_corpus_failures_max"]:
                failures.append(
                    "Hancom corpus smoke failures "
                    f"{hancom_corpus_summary['failure_count']} exceed {thresholds['hancom_corpus_failures_max']}."
                )
            if hancom_corpus_summary["dialog_count"] > thresholds["hancom_dialogs_max"]:
                failures.append(
                    f"Hancom warning/recovery dialog count {hancom_corpus_summary['dialog_count']} exceeds "
                    f"{thresholds['hancom_dialogs_max']}."
                )

    return failures


def main() -> int:
    parser = argparse.ArgumentParser(description="Run the repository release gate.")
    parser.add_argument(
        "--profile",
        choices=("ci", "release"),
        default="ci",
        help="ci runs the portable gate; release requires a Windows + Hancom machine.",
    )
    parser.add_argument(
        "--output-dir",
        default=str(REPO_ROOT / "build" / "release_gate"),
        help="Directory where gate reports are written.",
    )
    args = parser.parse_args()

    output_root = Path(args.output_dir).expanduser().resolve() / args.profile
    output_root.mkdir(parents=True, exist_ok=True)
    configured_temp_root = os.environ.get("JAKAL_RELEASE_TEMP_ROOT")
    temp_root_base = Path(configured_temp_root).expanduser().resolve() if configured_temp_root else REPO_ROOT / "tests" / "_release_tmp"
    temp_root = temp_root_base / args.profile
    shutil.rmtree(temp_root, ignore_errors=True)
    temp_root.mkdir(parents=True, exist_ok=True)

    contract = load_release_contract(REPO_ROOT)
    contract_errors = validate_contract_definition(contract, REPO_ROOT)
    profile_config = contract["gate_profiles"][args.profile]
    minimum_case_counts = contract["minimum_case_counts"]
    thresholds = contract["thresholds"]

    precondition_errors: list[str] = []
    windows_available = platform.system() == "Windows"
    hancom_enabled = bool(profile_config["require_hancom"] and windows_available)
    if profile_config["require_windows"] and not windows_available:
        precondition_errors.append(f"Profile {args.profile!r} requires Windows, current platform is {platform.system()!r}.")

    corpus_validation = validate_corpus_samples(contract, REPO_ROOT)
    corpus_report_path = write_corpus_validation_report(corpus_validation, output_root / "corpus_validation.json")
    print(
        f"[contract] errors={len(contract_errors)} corpus_samples={corpus_validation.sample_count} "
        f"corpus_failures={corpus_validation.failure_count}"
    )

    pytest_report = (
        _run_pytest(output_root / "pytest", temp_root=temp_root / "pytest")
        if profile_config["run_pytest"]
        else {"returncode": 0, "ok": True}
    )
    print(f"[pytest] ok={pytest_report['ok']} returncode={pytest_report['returncode']}")

    hwpx_results = run_stability_matrix(output_root / "hwpx_stability")
    hwpx_report_path = write_stability_report(hwpx_results, output_root / "hwpx_stability" / "stability_report.json")
    hwpx_summary = _summarize_hwpx_stability(hwpx_results, minimum_case_counts["hwpx_stability"])
    print(
        f"[hwpx] cases={hwpx_summary['case_count']} failures={hwpx_summary['failure_count']} "
        f"signature_failures={hwpx_summary['signature_failure_count']}"
    )

    hwp_results = run_hwp_stability_matrix(output_root / "hwp_stability", validate_with_hancom=hancom_enabled)
    hwp_report_path = write_hwp_stability_report(hwp_results, output_root / "hwp_stability" / "stability_report.json")
    hwp_summary = _summarize_hwp_stability(hwp_results, minimum_case_counts["hwp_stability"])
    print(
        f"[hwp] cases={hwp_summary['case_count']} failures={hwp_summary['failure_count']} "
        f"hancom_failures={hwp_summary['hancom_failure_count']} hancom_skips={hwp_summary['hancom_skip_count']}"
    )

    bridge_results = run_bridge_stability_matrix(output_root / "bridge_stability", validate_with_hancom=hancom_enabled)
    bridge_report_path = write_bridge_stability_report(
        bridge_results,
        output_root / "bridge_stability" / "stability_report.json",
    )
    bridge_summary = _summarize_bridge_stability(bridge_results, minimum_case_counts["bridge_stability"])
    print(
        f"[bridge] cases={bridge_summary['case_count']} failures={bridge_summary['failure_count']} "
        f"hancom_failures={bridge_summary['hancom_failure_count']} hancom_skips={bridge_summary['hancom_skip_count']}"
    )

    hancom_corpus_result = None
    if profile_config["run_hancom_corpus_smoke"]:
        hancom_corpus_result = run_hancom_corpus_smoke(output_root / "hancom_corpus_smoke", REPO_ROOT)
        print(
            f"[hancom-corpus] available={hancom_corpus_result.available} case_count={hancom_corpus_result.case_count} "
            f"failures={hancom_corpus_result.failure_count} dialogs={hancom_corpus_result.dialog_count}"
        )

    packaging_report: dict[str, Any] = {"ok": True, "warnings": [], "artifacts": []}
    try:
        version = read_versions()
        artifacts = rebuild_dist(temp_root=temp_root)
        packaging_warnings = ensure_required_entries(artifacts, version)
        packaging_report = {
            "ok": True,
            "version": version,
            "warnings": packaging_warnings,
            "artifacts": [str(path) for path in artifacts],
        }
        print(f"[packaging] version={version} artifacts={[path.name for path in artifacts]}")
        for warning in packaging_warnings:
            print(f"[packaging-warn] {warning}")
    except Exception as exc:  # noqa: BLE001
        packaging_report = {"ok": False, "error": str(exc), "warnings": [], "artifacts": []}
        print(f"[packaging] failed: {exc}")

    hancom_corpus_summary = None
    if hancom_corpus_result is not None:
        hancom_corpus_summary = {
            "available": hancom_corpus_result.available,
            "ok": hancom_corpus_result.ok,
            "case_count": hancom_corpus_result.case_count,
            "failure_count": hancom_corpus_result.failure_count,
            "dialog_count": hancom_corpus_result.dialog_count,
            "report_path": hancom_corpus_result.report_path,
            "errors": hancom_corpus_result.errors,
        }

    failures = _validate_gate_thresholds(
        contract_errors=contract_errors,
        profile_name=args.profile,
        profile_config=profile_config,
        thresholds=thresholds,
        corpus_report={
            "sample_count": corpus_validation.sample_count,
            "failure_count": corpus_validation.failure_count,
            "missing_categories": corpus_validation.missing_categories,
        },
        pytest_report=pytest_report,
        hwpx_summary=hwpx_summary,
        hwp_summary=hwp_summary,
        bridge_summary=bridge_summary,
        hancom_corpus_summary=hancom_corpus_summary,
        precondition_errors=precondition_errors,
    )
    if not packaging_report["ok"]:
        failures.append(str(packaging_report["error"]))

    overall_report = {
        "profile": args.profile,
        "runtime": python_runtime_summary(),
        "contract_errors": contract_errors,
        "precondition_errors": precondition_errors,
        "corpus_validation": {
            "sample_count": corpus_validation.sample_count,
            "failure_count": corpus_validation.failure_count,
            "covered_categories": corpus_validation.covered_categories,
            "missing_categories": corpus_validation.missing_categories,
            "report_path": str(corpus_report_path),
        },
        "pytest": pytest_report,
        "hwpx_stability": {
            **hwpx_summary,
            "report_path": str(hwpx_report_path),
        },
        "hwp_stability": {
            **hwp_summary,
            "report_path": str(hwp_report_path),
            "hancom_enabled": hancom_enabled,
        },
        "bridge_stability": {
            **bridge_summary,
            "report_path": str(bridge_report_path),
            "hancom_enabled": hancom_enabled,
        },
        "hancom_corpus_smoke": hancom_corpus_summary,
        "packaging": packaging_report,
        "pdf_visual_diff": {
            "status": "not_configured",
            "threshold": thresholds["pdf_visual_diff_max_changed_ratio"],
            "blocking": False,
        },
        "failures": failures,
        "ok": not failures,
    }
    report_path = output_root / "release_report.json"
    _write_json(report_path, overall_report)
    print(f"[report] {report_path}")

    if failures:
        for failure in failures:
            print(f"[FAIL] {failure}")
        return 1

    print(f"[ok] release gate passed for profile={args.profile}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
