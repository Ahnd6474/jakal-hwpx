param(
    [string]$CorpusDir = "",
    [string]$OutputRoot = "",
    [string]$SecurityModuleName = "FilePathCheckerModuleExample",
    [string]$SecurityModulePath = "",
    [string]$SecurityModuleInstallRoot = "",
    [string]$SecurityRegistryRoot = "HKCU:\SOFTWARE\HNC\HwpAutomation\Modules",
    [int]$TimeoutSeconds = 90,
    [switch]$SkipSecurityModuleSetup,
    [switch]$AllowExistingHwpProcesses
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-RepoRoot {
    return Split-Path -Parent $PSScriptRoot
}

function Resolve-CorpusDir {
    param([string]$RequestedPath)

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        return [System.IO.Path]::GetFullPath($RequestedPath)
    }
    return Join-Path (Get-RepoRoot) "examples\samples\hwpx"
}

function Resolve-OutputRoot {
    param([string]$RequestedPath)

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        return [System.IO.Path]::GetFullPath($RequestedPath)
    }
    return Join-Path (Get-RepoRoot) ".codex-temp\hancom-corpus-smoke"
}

$resolvedCorpusDir = Resolve-CorpusDir -RequestedPath $CorpusDir
$resolvedOutputRoot = Resolve-OutputRoot -RequestedPath $OutputRoot
$setupScript = Join-Path $PSScriptRoot "setup_hancom_security_module.ps1"
$smokeScript = Join-Path $PSScriptRoot "run_hancom_smoke_validation.ps1"

if ($SkipSecurityModuleSetup) {
    throw "Skipping Hancom security module setup is disabled. Remove -SkipSecurityModuleSetup."
}

if (-not (Test-Path $resolvedCorpusDir)) {
    throw "Corpus directory does not exist: $resolvedCorpusDir"
}

New-Item -ItemType Directory -Force $resolvedOutputRoot | Out-Null

$setupArgs = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", $setupScript,
    "-ModuleName", $SecurityModuleName,
    "-RegistryRoot", $SecurityRegistryRoot
)
if (-not [string]::IsNullOrWhiteSpace($SecurityModuleInstallRoot)) {
    $setupArgs += @("-InstallRoot", $SecurityModuleInstallRoot)
}
if (-not [string]::IsNullOrWhiteSpace($SecurityModulePath)) {
    $setupArgs += @("-ModulePath", $SecurityModulePath)
}
else {
    $setupArgs += "-DownloadIfMissing"
}
& powershell.exe @setupArgs | Out-Null

$results = New-Object System.Collections.Generic.List[object]
$files = Get-ChildItem -LiteralPath $resolvedCorpusDir -Filter *.hwpx | Sort-Object Name

foreach ($file in $files) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    $safeName = ($baseName -replace '[^0-9A-Za-z._-]', '_')
    $outputDir = Join-Path $resolvedOutputRoot $safeName
    $roundtripPath = Join-Path $outputDir ($safeName + ".roundtrip.hwpx")
    $logPath = Join-Path $outputDir "hancom-dialog-log.json"

    New-Item -ItemType Directory -Force $outputDir | Out-Null
    $status = "ok"
    $message = ""

    try {
        & $smokeScript `
            -InputPath $file.FullName `
            -OutputPath $roundtripPath `
            -TimeoutSeconds $TimeoutSeconds `
            -LogPath $logPath `
            -SecurityModuleName $SecurityModuleName `
            -SecurityModulePath $SecurityModulePath `
            -SecurityModuleInstallRoot $SecurityModuleInstallRoot `
            -SecurityRegistryRoot $SecurityRegistryRoot `
            -AllowExistingHwpProcesses:$AllowExistingHwpProcesses
    }
    catch {
        $status = "error"
        $message = $_.Exception.Message
    }

    $results.Add([pscustomobject]@{
        Name = $file.Name
        InputPath = $file.FullName
        OutputPath = $roundtripPath
        LogPath = [System.IO.Path]::ChangeExtension($logPath, ".run.json")
        Status = $status
        Message = $message
    }) | Out-Null

    if ($status -ne "ok") {
        break
    }
}

$reportPath = Join-Path $resolvedOutputRoot "corpus-smoke-report.json"
$results | ConvertTo-Json -Depth 8 | Set-Content -Path $reportPath -Encoding UTF8
$results
