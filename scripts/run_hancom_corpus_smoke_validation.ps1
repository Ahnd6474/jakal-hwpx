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

function Get-HancomProcessIds {
    return @(
        Get-Process -Name Hwp -ErrorAction SilentlyContinue |
            ForEach-Object { [int]$_.Id }
    )
}

function Stop-NewHancomProcesses {
    param([int[]]$ExistingProcessIds)

    $newProcesses = @(
        Get-Process -Name Hwp -ErrorAction SilentlyContinue |
            Where-Object { $ExistingProcessIds -notcontains [int]$_.Id }
    )
    foreach ($process in $newProcesses) {
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-IsolatedHancomSmoke {
    param(
        [string]$ScriptPath,
        [string]$SourcePath,
        [string]$DestinationPath,
        [string]$DialogLogPath,
        [int]$OperationTimeoutSeconds
    )

    $existingProcessIds = @(Get-HancomProcessIds)
    $childCommand = @'
$ErrorActionPreference = "Stop"
& $env:JAKAL_HANCOM_SMOKE_SCRIPT `
    -InputPath $env:JAKAL_HANCOM_SMOKE_INPUT `
    -OutputPath $env:JAKAL_HANCOM_SMOKE_OUTPUT `
    -TimeoutSeconds ([int]$env:JAKAL_HANCOM_SMOKE_TIMEOUT) `
    -LogPath $env:JAKAL_HANCOM_SMOKE_LOG `
    -SecurityModuleName $env:JAKAL_HANCOM_SECURITY_NAME `
    -SecurityModulePath $env:JAKAL_HANCOM_SECURITY_PATH `
    -SecurityModuleInstallRoot $env:JAKAL_HANCOM_SECURITY_ROOT `
    -SecurityRegistryRoot $env:JAKAL_HANCOM_SECURITY_REGISTRY `
    -AllowExistingHwpProcesses:([bool]::Parse($env:JAKAL_HANCOM_ALLOW_EXISTING))
'@
    $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($childCommand))
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = "powershell.exe"
    $startInfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $encodedCommand"
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.EnvironmentVariables["JAKAL_HANCOM_SMOKE_SCRIPT"] = $ScriptPath
    $startInfo.EnvironmentVariables["JAKAL_HANCOM_SMOKE_INPUT"] = $SourcePath
    $startInfo.EnvironmentVariables["JAKAL_HANCOM_SMOKE_OUTPUT"] = $DestinationPath
    $startInfo.EnvironmentVariables["JAKAL_HANCOM_SMOKE_TIMEOUT"] = [string]$OperationTimeoutSeconds
    $startInfo.EnvironmentVariables["JAKAL_HANCOM_SMOKE_LOG"] = $DialogLogPath
    $startInfo.EnvironmentVariables["JAKAL_HANCOM_SECURITY_NAME"] = $SecurityModuleName
    $startInfo.EnvironmentVariables["JAKAL_HANCOM_SECURITY_PATH"] = $SecurityModulePath
    $startInfo.EnvironmentVariables["JAKAL_HANCOM_SECURITY_ROOT"] = $SecurityModuleInstallRoot
    $startInfo.EnvironmentVariables["JAKAL_HANCOM_SECURITY_REGISTRY"] = $SecurityRegistryRoot
    $startInfo.EnvironmentVariables["JAKAL_HANCOM_ALLOW_EXISTING"] = ([bool]$AllowExistingHwpProcesses).ToString()
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo

    try {
        if (-not $process.Start()) {
            throw "Failed to start isolated Hancom smoke validation for $SourcePath."
        }
        $finished = $process.WaitForExit($OperationTimeoutSeconds * 1000)
        if (-not $finished) {
            try { $process.Kill() } catch { }
            try { $process.WaitForExit() } catch { }
            throw "Hancom smoke validation timed out after $OperationTimeoutSeconds seconds for $SourcePath."
        }

        $standardOutput = $process.StandardOutput.ReadToEnd().Trim()
        $standardError = $process.StandardError.ReadToEnd().Trim()
        if ($process.ExitCode -ne 0) {
            $detail = if (-not [string]::IsNullOrWhiteSpace($standardError)) {
                $standardError
            } elseif (-not [string]::IsNullOrWhiteSpace($standardOutput)) {
                $standardOutput
            } else {
                "exit code $($process.ExitCode)"
            }
            throw "Hancom smoke validation failed for ${SourcePath}: $detail"
        }
    }
    finally {
        $process.Dispose()
        Stop-NewHancomProcesses -ExistingProcessIds $existingProcessIds
    }
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
        Invoke-IsolatedHancomSmoke `
            -ScriptPath $smokeScript `
            -SourcePath $file.FullName `
            -DestinationPath $roundtripPath `
            -DialogLogPath $logPath `
            -OperationTimeoutSeconds $TimeoutSeconds
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
