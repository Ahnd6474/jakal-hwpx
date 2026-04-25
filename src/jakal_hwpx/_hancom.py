from __future__ import annotations

import os
import subprocess
from pathlib import Path

from .exceptions import HancomInteropError


DEFAULT_SECURITY_MODULE_NAME = "FilePathCheckerModuleExample"
DEFAULT_SECURITY_REGISTRY_ROOT = r"HKCU:\SOFTWARE\HNC\HwpAutomation\Modules"


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _default_security_module_install_root() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "jakal-hwpx" / "hancom-security"
    return _repo_root() / ".codex-temp" / "hancom-security"


def _smoke_script_path() -> Path:
    return _repo_root() / "scripts" / "run_hancom_smoke_validation.ps1"


def convert_document(
    input_path: str | Path,
    output_path: str | Path,
    output_format: str,
    *,
    timeout_seconds: int = 120,
    retry_count: int = 1,
    allow_existing_hwp_processes: bool = True,
    security_module_name: str = DEFAULT_SECURITY_MODULE_NAME,
    security_module_path: str = "",
    security_module_install_root: str | Path | None = None,
    security_registry_root: str = DEFAULT_SECURITY_REGISTRY_ROOT,
    skip_security_module_registration: bool = False,
) -> Path:
    if skip_security_module_registration:
        raise HancomInteropError(
            "Skipping Hancom security module registration is disabled. "
            "Remove skip_security_module_registration=True."
        )

    resolved_input = Path(input_path).expanduser().resolve()
    resolved_output = Path(output_path).expanduser().resolve()
    resolved_output.parent.mkdir(parents=True, exist_ok=True)
    resolved_install_root = (
        Path(security_module_install_root).expanduser().resolve()
        if security_module_install_root
        else _default_security_module_install_root().resolve()
    )

    script_path = _smoke_script_path()
    if not script_path.exists():
        raise HancomInteropError(f"Hancom smoke script was not found: {script_path}")

    command = [
        "powershell",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script_path),
        "-InputPath",
        str(resolved_input),
        "-OutputPath",
        str(resolved_output),
        "-OutputFormat",
        output_format,
        "-TimeoutSeconds",
        str(timeout_seconds),
        "-SecurityModuleName",
        security_module_name,
        "-SecurityModuleInstallRoot",
        str(resolved_install_root),
        "-SecurityRegistryRoot",
        security_registry_root,
    ]
    if allow_existing_hwp_processes:
        command.append("-AllowExistingHwpProcesses")
    if security_module_path:
        command.extend(["-SecurityModulePath", security_module_path])

    attempts = max(1, int(retry_count) + 1)
    last_detail = ""
    for attempt in range(1, attempts + 1):
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
        if completed.returncode == 0:
            if resolved_output.exists():
                return resolved_output
            last_detail = f"Hancom conversion did not produce the expected output file: {resolved_output}"
        else:
            stderr = completed.stderr.strip()
            stdout = completed.stdout.strip()
            last_detail = stderr or stdout or f"exit code {completed.returncode}"

        if attempt < attempts and _is_transient_hancom_failure(last_detail):
            continue
        break

    if last_detail.startswith("Hancom conversion did not produce"):
        raise HancomInteropError(last_detail)
    raise HancomInteropError(f"Hancom conversion failed for {resolved_input} -> {resolved_output}: {last_detail}")


def _is_transient_hancom_failure(detail: str) -> bool:
    lowered = detail.lower()
    return any(
        marker in lowered
        for marker in (
            "remote procedure call failed",
            "rpc server is unavailable",
            "0x800706be",
            "0x800706ba",
        )
    )
