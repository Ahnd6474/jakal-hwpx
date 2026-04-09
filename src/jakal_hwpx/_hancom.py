from __future__ import annotations

import subprocess
from pathlib import Path

from .exceptions import HancomInteropError


DEFAULT_SECURITY_MODULE_NAME = "FilePathCheckerModuleExample"
DEFAULT_SECURITY_REGISTRY_ROOT = r"HKCU:\SOFTWARE\HNC\HwpAutomation\Modules"


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _smoke_script_path() -> Path:
    return _repo_root() / "scripts" / "run_hancom_smoke_validation.ps1"


def convert_document(
    input_path: str | Path,
    output_path: str | Path,
    output_format: str,
    *,
    timeout_seconds: int = 120,
    allow_existing_hwp_processes: bool = True,
    security_module_name: str = DEFAULT_SECURITY_MODULE_NAME,
    security_module_path: str = "",
    security_registry_root: str = DEFAULT_SECURITY_REGISTRY_ROOT,
    skip_security_module_registration: bool = False,
) -> Path:
    resolved_input = Path(input_path).expanduser().resolve()
    resolved_output = Path(output_path).expanduser().resolve()
    resolved_output.parent.mkdir(parents=True, exist_ok=True)

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
        "-SecurityRegistryRoot",
        security_registry_root,
    ]
    if allow_existing_hwp_processes:
        command.append("-AllowExistingHwpProcesses")
    if security_module_path:
        command.extend(["-SecurityModulePath", security_module_path])
    if skip_security_module_registration:
        command.append("-SkipSecurityModuleRegistration")

    completed = subprocess.run(
        command,
        capture_output=True,
        text=True,
        check=False,
    )
    if completed.returncode != 0:
        stderr = completed.stderr.strip()
        stdout = completed.stdout.strip()
        detail = stderr or stdout or f"exit code {completed.returncode}"
        raise HancomInteropError(f"Hancom conversion failed for {resolved_input} -> {resolved_output}: {detail}")

    if not resolved_output.exists():
        raise HancomInteropError(f"Hancom conversion did not produce the expected output file: {resolved_output}")
    return resolved_output
