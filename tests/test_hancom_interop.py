from __future__ import annotations

import subprocess
from pathlib import Path

import pytest

import jakal_hwpx._hancom as hancom
from jakal_hwpx.exceptions import HancomInteropError


REPO_ROOT = Path(__file__).resolve().parents[1]


def test_default_security_module_install_root_prefers_local_app_data(monkeypatch) -> None:
    monkeypatch.setenv("LOCALAPPDATA", r"C:\Users\TestUser\AppData\Local")

    resolved = hancom._default_security_module_install_root()

    assert resolved == Path(r"C:\Users\TestUser\AppData\Local") / "jakal-hwpx" / "hancom-security"


def test_default_security_module_install_root_falls_back_to_repo_temp(monkeypatch) -> None:
    monkeypatch.delenv("LOCALAPPDATA", raising=False)

    resolved = hancom._default_security_module_install_root()

    assert resolved == REPO_ROOT / ".codex-temp" / "hancom-security"


def test_convert_document_always_passes_security_module_install_root(monkeypatch, tmp_path: Path) -> None:
    input_path = tmp_path / "input.hwp"
    output_path = tmp_path / "output.hwpx"
    script_path = tmp_path / "run_hancom_smoke_validation.ps1"
    input_path.write_bytes(b"fake-hwp")
    script_path.write_text("exit 0", encoding="utf-8")

    commands: list[list[str]] = []

    def fake_run(command: list[str], *, capture_output: bool, text: bool, check: bool):
        assert capture_output is True
        assert text is True
        assert check is False
        commands.append(command)
        output_path.write_text("<hwpx />", encoding="utf-8")
        return subprocess.CompletedProcess(command, 0, stdout="", stderr="")

    monkeypatch.setenv("LOCALAPPDATA", r"C:\Users\TestUser\AppData\Local")
    monkeypatch.setattr(hancom, "_smoke_script_path", lambda: script_path)
    monkeypatch.setattr(hancom.subprocess, "run", fake_run)

    result = hancom.convert_document(input_path, output_path, "HWPX", allow_existing_hwp_processes=False)

    assert result == output_path.resolve()
    assert len(commands) == 1
    assert "-SecurityModuleInstallRoot" in commands[0]
    assert str(Path(r"C:\Users\TestUser\AppData\Local") / "jakal-hwpx" / "hancom-security") in commands[0]
    assert "-SkipSecurityModuleRegistration" not in commands[0]


def test_convert_document_rejects_skip_security_module_registration(tmp_path: Path) -> None:
    input_path = tmp_path / "input.hwp"
    output_path = tmp_path / "output.hwpx"
    input_path.write_bytes(b"fake-hwp")

    with pytest.raises(HancomInteropError, match="disabled"):
        hancom.convert_document(
            input_path,
            output_path,
            "HWPX",
            skip_security_module_registration=True,
        )


def test_hancom_smoke_scripts_force_security_module_setup() -> None:
    smoke_script = (REPO_ROOT / "scripts" / "run_hancom_smoke_validation.ps1").read_text(encoding="utf-8")
    corpus_script = (REPO_ROOT / "scripts" / "run_hancom_corpus_smoke_validation.ps1").read_text(encoding="utf-8")
    setup_script = (REPO_ROOT / "scripts" / "setup_hancom_security_module.ps1").read_text(encoding="utf-8")

    assert "Invoke-SecurityModuleSetup" in smoke_script
    assert "Skipping Hancom security module registration is disabled." in smoke_script
    assert "[string]$SecurityModuleInstallRoot = \"\"" in smoke_script
    assert "Hancom RegisterModule(FilePathCheckDLL, $SecurityModuleName) failed." in smoke_script

    assert "Skipping Hancom security module setup is disabled." in corpus_script
    assert "-SecurityModuleInstallRoot $SecurityModuleInstallRoot" in corpus_script

    assert '[Environment]::GetFolderPath("LocalApplicationData")' in setup_script
    assert 'return Join-Path $localAppData "jakal-hwpx\\hancom-security"' in setup_script


def test_hancom_textart_probe_scripts_cover_plain_text_selection() -> None:
    action_probe = (REPO_ROOT / "scripts" / "probe_hancom_textart_actions.ps1").read_text(encoding="utf-8")
    run_probe = (REPO_ROOT / "scripts" / "probe_hancom_textart_run_commands.ps1").read_text(encoding="utf-8")

    assert '[string]$SeedText = "TEXTART-SEED"' in action_probe
    assert "function Create-PlainTextDocument" in action_probe
    assert '$result.plain_text += Get-ParameterSnapshot' in action_probe
    assert '$result.plain_text_execute_attempts += Try-ExecutePlainTextAction' in action_probe

    assert '[string[]]$Modes = @("donor", "textbox", "plain_text")' in run_probe
    assert '[string]$SeedText = "TEXTART-SEED"' in run_probe
    assert "function Initialize-PlainTextSelection" in run_probe
    assert 'elseif ($SelectedMode -eq "plain_text")' in run_probe


def test_hancom_connectline_probe_script_targets_hwpx_to_hwp_conversion() -> None:
    script = (REPO_ROOT / "scripts" / "probe_hancom_connectline_mapping.py").read_text(encoding="utf-8")

    assert '"connectLine"' in script
    assert 'convert_document(connect_hwpx' in script or '"probe_connectline_from_hancom.hwp"' in script
    assert '"HWP"' in script
    assert "PROBE-CONNECTLINE" in script
    assert "comparisons" in script


def test_hancom_picture_ole_probe_script_targets_hwpx_to_hwp_conversion() -> None:
    script = (REPO_ROOT / "scripts" / "probe_hancom_picture_ole_mapping.py").read_text(encoding="utf-8")

    assert "set_crop(" in script
    assert "set_extent(" in script
    assert '"HWP"' in script
    assert "PROBE-PICTURE-MUTATED" in script
    assert "PROBE-OLE-MUTATED" in script
    assert "comparisons" in script
