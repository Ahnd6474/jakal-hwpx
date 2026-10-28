"""Optional local Hancom conversion without rebuilding document objects."""
from __future__ import annotations

import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile

from .exceptions import HancomInteropError


_SCRIPT = r'''
param([string]$ConfigPath)
$ErrorActionPreference = 'Stop'
$config = [IO.File]::ReadAllText($ConfigPath) | ConvertFrom-Json
if (@(Get-Process Hwp -ErrorAction SilentlyContinue).Count -gt 0) {
    throw 'An existing Hancom process is running. Close it before isolated conversion.'
}
$hwp = $null
try {
    $hwp = New-Object -ComObject HWPFrame.HwpObject
    $owned = @(Get-Process Hwp -ErrorAction SilentlyContinue | ForEach-Object {
        @{ id = $_.Id; started = $_.StartTime.ToUniversalTime().ToString('o') }
    })
    $owned | ConvertTo-Json | Set-Content -LiteralPath $config.process_file -Encoding UTF8
    $hwp.XHwpWindows.Item(0).Visible = $false
    $hwp.SetMessageBoxMode(0x00010000) | Out-Null
    if (-not $hwp.RegisterModule('FilePathCheckDLL', $config.security_module)) {
        throw 'Hancom file access security module is not registered.'
    }
    if (-not $hwp.Open($config.input_path, '', '')) { throw 'Hancom Open failed.' }
    if (-not $hwp.SaveAs($config.output_path, $config.format, '')) { throw 'Hancom SaveAs failed.' }
} finally {
    if ($hwp) {
        try { $hwp.Quit() } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($hwp) }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    foreach ($entry in $owned) {
        $process = Get-Process -Id $entry.id -ErrorAction SilentlyContinue
        if ($process -and $process.ProcessName -eq 'Hwp' -and $process.StartTime.ToUniversalTime().ToString('o') -eq $entry.started) {
            if (-not $process.WaitForExit(3000)) { Stop-Process -Id $process.Id }
        }
    }
}
'''


class HancomOfficeConverter:
    """Isolated Windows Hancom automation using an already registered security module.

    This requires an installed Hancom Office and refuses to reuse an open user's
    application. No software, registry entry, or security module is installed.
    """

    def __init__(self, *, security_module: str = "FilePathCheckerModuleExample", timeout: float = 120) -> None:
        self.security_module = security_module
        self.timeout = timeout

    def __call__(self, source: str | Path, destination: str | Path, output_format: str) -> Path:
        if os.name != "nt":
            raise HancomInteropError("Native formatting preservation requires Windows Hancom Office or a supplied converter.")
        source, destination = Path(source).resolve(), Path(destination).resolve()
        output_format = output_format.upper()
        if output_format not in {"HWP", "HWPX", "DOCX"}:
            raise ValueError("HancomOfficeConverter supports HWP, HWPX and DOCX output.")
        if not source.is_file():
            raise FileNotFoundError(source)
        with tempfile.TemporaryDirectory(prefix="jakal_native_convert_") as directory:
            root = Path(directory)
            script = root / "convert.ps1"
            config = root / "config.json"
            output = root / ("converted." + output_format.lower())
            process_file = root / "processes.json"
            script.write_text(_SCRIPT, encoding="utf-8-sig")
            config.write_text(json.dumps({
                # Hancom names its DOCX filter OOXML (not DOCX).
                "input_path": str(source), "output_path": str(output), "format": "OOXML" if output_format == "DOCX" else output_format,
                "security_module": self.security_module, "process_file": str(process_file),
            }), encoding="utf-8")
            try:
                result = subprocess.run(
                    ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(script), "-ConfigPath", str(config)],
                    capture_output=True, timeout=self.timeout, creationflags=0x08000000,
                )
            except subprocess.TimeoutExpired as error:
                # Only stop the exact process started by this isolated operation.
                if process_file.exists():
                    cleanup = r'''param([string]$Path)
$owned = @([IO.File]::ReadAllText($Path) | ConvertFrom-Json)
foreach ($entry in $owned) {
    $process = Get-Process -Id $entry.id -ErrorAction SilentlyContinue
    if ($process -and $process.ProcessName -eq 'Hwp' -and $process.StartTime.ToUniversalTime().ToString('o') -eq $entry.started) {
        Stop-Process -Id $process.Id
    }
}
'''
                    cleanup_path = root / "cleanup.ps1"
                    cleanup_path.write_text(cleanup, encoding="utf-8-sig")
                    subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(cleanup_path), str(process_file)], capture_output=True, timeout=15, creationflags=0x08000000)
                raise HancomInteropError("Hancom conversion timed out; no output was written.") from error
            if result.returncode or not output.is_file():
                message = result.stderr.decode("utf-8", errors="replace").strip()
                raise HancomInteropError("Native Hancom conversion failed: " + message)
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(output, destination)
        return destination
