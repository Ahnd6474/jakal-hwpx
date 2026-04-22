param(
    [string]$DonorPath = "",
    [string]$OutputRoot = "",
    [string[]]$Commands = @("InsertTextArtDlg", "InsertTextArtButton", "TextArtToolBoxGallery", "ModifyTextArtDlg", "TextArtPropertyDlg"),
    [string[]]$Modes = @("donor", "textbox", "plain_text"),
    [string]$SeedText = "TEXTART-SEED",
    [int]$TimeoutSeconds = 8,
    [switch]$ChildMode,
    [string]$Mode = "",
    [string]$CommandName = "",
    [string]$StatusPath = "",
    [string]$OutputPath = "",
    [string]$SecurityModuleName = "FilePathCheckerModuleExample",
    [string]$SecurityModulePath = "",
    [string]$SecurityModuleInstallRoot = "",
    [string]$SecurityRegistryRoot = "HKCU:\SOFTWARE\HNC\HwpAutomation\Modules"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ($Commands.Count -eq 1 -and $Commands[0] -like "*,*") {
    $Commands = @(
        $Commands[0].Split(",") |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

if ($Modes.Count -eq 1 -and $Modes[0] -like "*,*") {
    $Modes = @(
        $Modes[0].Split(",") |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function Get-RepoRoot {
    return Split-Path -Parent $PSScriptRoot
}

function Resolve-DonorPath {
    param([string]$RequestedPath)

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        return [System.IO.Path]::GetFullPath($RequestedPath)
    }
    return Join-Path (Get-RepoRoot) ".codex-temp\reverse_shapes\com_textart_try.hwp"
}

function Resolve-OutputRoot {
    param([string]$RequestedPath)

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        return [System.IO.Path]::GetFullPath($RequestedPath)
    }
    return Join-Path (Get-RepoRoot) ".codex-temp\hancom-textart-run-probe"
}

function Ensure-SecurityModule {
    param(
        [string]$ModuleName,
        [string]$ModulePath,
        [string]$InstallRoot,
        [string]$RegistryRoot
    )

    $setupScript = Join-Path $PSScriptRoot "setup_hancom_security_module.ps1"
    $setupArgs = @{
        ModuleName = $ModuleName
        RegistryRoot = $RegistryRoot
    }
    if (-not [string]::IsNullOrWhiteSpace($InstallRoot)) {
        $setupArgs["InstallRoot"] = $InstallRoot
    }
    if (-not [string]::IsNullOrWhiteSpace($ModulePath)) {
        $setupArgs["ModulePath"] = $ModulePath
    }
    else {
        $setupArgs["DownloadIfMissing"] = $true
    }
    & $setupScript @setupArgs | Out-Null
}

function New-HwpObject {
    param([string]$ModuleName)

    $hwp = New-Object -ComObject HWPFrame.HwpObject
    $registerResult = $hwp.RegisterModule("FilePathCheckDLL", $ModuleName)
    if (-not $registerResult) {
        throw "RegisterModule(FilePathCheckDLL, $ModuleName) failed."
    }
    return $hwp
}

function Select-GraphicControl {
    param(
        [object]$Hwp,
        [object]$Control
    )

    $anchor = $Control.GetAnchorPos(0)
    $Hwp.SetPosBySet($anchor) | Out-Null
    $Hwp.FindCtrl() | Out-Null
}

function Select-FirstGso {
    param([object]$Hwp)

    $control = $Hwp.HeadCtrl
    while ($null -ne $control) {
        if ($control.CtrlID -eq "gso") {
            Select-GraphicControl -Hwp $Hwp -Control $control
            return
        }
        $control = $control.Next
    }
    throw "Could not find a gso control."
}

function Select-LastGso {
    param([object]$Hwp)

    $control = $Hwp.LastCtrl
    while ($null -ne $control) {
        if ($control.CtrlID -eq "gso") {
            Select-GraphicControl -Hwp $Hwp -Control $control
            return
        }
        $control = $control.Prev
    }
    throw "Could not find the last gso control."
}

function Initialize-TextboxSelection {
    param([object]$Hwp)

    $Hwp.Run("FileNew") | Out-Null
    $action = $Hwp.CreateAction("DrawObjCreatorTextBox")
    if ($null -eq $action) {
        throw "DrawObjCreatorTextBox action is unavailable."
    }
    $shapeObject = $Hwp.HParameterSet.HShapeObject
    $action.GetDefault($shapeObject.HSet) | Out-Null
    $shapeObject.Width = 22500
    $shapeObject.Height = 5000
    $shapeObject.TreatAsChar = 1
    $shapeObject.ShapeDrawLineAttr.Color = 0x00665544
    $shapeObject.ShapeDrawFillAttr.type = 1
    $shapeObject.ShapeDrawFillAttr.WinBrushFaceColor = 0x00FFEEDD
    $action.Execute($shapeObject.HSet) | Out-Null
    Select-LastGso -Hwp $Hwp
}

function Initialize-PlainTextSelection {
    param(
        [object]$Hwp,
        [string]$Text
    )

    $Hwp.Run("FileNew") | Out-Null
    $action = $Hwp.CreateAction("InsertText")
    if ($null -eq $action) {
        throw "InsertText action is unavailable."
    }

    $insertText = $Hwp.HParameterSet.HInsertText
    $action.GetDefault($insertText.HSet) | Out-Null
    $insertText.Text = $Text
    $action.Execute($insertText.HSet) | Out-Null
    $Hwp.Run("SelectAll") | Out-Null
}

function Invoke-ChildRunProbe {
    param(
        [string]$SelectedMode,
        [string]$SelectedCommand,
        [string]$SelectedDonorPath,
        [string]$SelectedOutputPath,
        [string]$SelectedStatusPath,
        [string]$ModuleName,
        [string]$SelectedSeedText
    )

    $status = [ordered]@{
        mode = $SelectedMode
        command = $SelectedCommand
        returned = $false
        run_result = $null
        save = $false
        error = $null
        output_path = $SelectedOutputPath
    }

    $hwp = $null
    try {
        $hwp = New-HwpObject -ModuleName $ModuleName
        if ($SelectedMode -eq "donor") {
            $hwp.Open($SelectedDonorPath, "", "") | Out-Null
            Select-FirstGso -Hwp $hwp
        }
        elseif ($SelectedMode -eq "textbox") {
            Initialize-TextboxSelection -Hwp $hwp
        }
        elseif ($SelectedMode -eq "plain_text") {
            Initialize-PlainTextSelection -Hwp $hwp -Text $SelectedSeedText
        }
        else {
            throw "Unsupported mode: $SelectedMode"
        }

        $status.run_result = $hwp.Run($SelectedCommand)
        $status.returned = $true
    }
    catch {
        $status.error = $_.Exception.Message
    }
    finally {
        if ($null -ne $hwp) {
            try {
                if (-not [string]::IsNullOrWhiteSpace($SelectedOutputPath)) {
                    $status.save = [bool]$hwp.SaveAs($SelectedOutputPath, "HWP", "")
                }
            }
            catch {
            }
            try {
                $hwp.Quit()
            }
            catch {
            }
        }
    }

    $json = $status | ConvertTo-Json -Depth 4
    $json | Set-Content -Encoding UTF8 $SelectedStatusPath
}

$resolvedDonorPath = Resolve-DonorPath -RequestedPath $DonorPath
$resolvedOutputRoot = Resolve-OutputRoot -RequestedPath $OutputRoot
New-Item -ItemType Directory -Force $resolvedOutputRoot | Out-Null

Ensure-SecurityModule `
    -ModuleName $SecurityModuleName `
    -ModulePath $SecurityModulePath `
    -InstallRoot $SecurityModuleInstallRoot `
    -RegistryRoot $SecurityRegistryRoot

if ($ChildMode) {
    Invoke-ChildRunProbe `
        -SelectedMode $Mode `
        -SelectedCommand $CommandName `
        -SelectedDonorPath $resolvedDonorPath `
        -SelectedOutputPath $OutputPath `
        -SelectedStatusPath $StatusPath `
        -ModuleName $SecurityModuleName `
        -SelectedSeedText $SeedText
    exit 0
}

$results = New-Object System.Collections.Generic.List[object]
foreach ($selectedMode in $Modes) {
    foreach ($selectedCommand in $Commands) {
        $baseName = "{0}_{1}" -f $selectedMode, $selectedCommand
        $statusPath = Join-Path $resolvedOutputRoot ("{0}.json" -f $baseName)
        $outputPath = Join-Path $resolvedOutputRoot ("{0}.hwp" -f $baseName)
        $child = Start-Process `
            -FilePath "powershell.exe" `
            -ArgumentList @(
                "-ExecutionPolicy", "Bypass",
                "-File", $PSCommandPath,
                "-ChildMode",
                "-Mode", $selectedMode,
                "-CommandName", $selectedCommand,
                "-DonorPath", $resolvedDonorPath,
                "-OutputRoot", $resolvedOutputRoot,
                "-StatusPath", $statusPath,
                "-OutputPath", $outputPath,
                "-SecurityModuleName", $SecurityModuleName,
                "-SecurityRegistryRoot", $SecurityRegistryRoot,
                "-SeedText", $SeedText
            ) `
            -WindowStyle Hidden `
            -PassThru

        $finished = $child.WaitForExit($TimeoutSeconds * 1000)
        if (-not $finished) {
            try {
                Stop-Process -Id $child.Id -Force
            }
            catch {
            }
            $results.Add([pscustomobject]@{
                mode = $selectedMode
                command = $selectedCommand
                timed_out = $true
                status_path = $statusPath
                output_path = $outputPath
            }) | Out-Null
            continue
        }

        $payload = if (Test-Path $statusPath) {
            Get-Content -Raw -Encoding UTF8 $statusPath | ConvertFrom-Json
        }
        else {
            [pscustomobject]@{
                mode = $selectedMode
                command = $selectedCommand
                returned = $false
                run_result = $null
                save = $false
                error = "status file missing"
                output_path = $outputPath
            }
        }
        $results.Add($payload) | Out-Null
    }
}

$jsonPath = Join-Path $resolvedOutputRoot "summary.json"
$results | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 $jsonPath
$results
