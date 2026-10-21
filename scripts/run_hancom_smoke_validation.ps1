param(
    [string]$InputPath,
    [string]$OutputPath,
    [string]$OutputFormat = "HWPX",
    [int]$TimeoutSeconds = 90,
    [string]$LogPath = "",
    [string]$SecurityModuleName = "FilePathCheckerModuleExample",
    [string]$SecurityModulePath = "",
    [string]$SecurityModuleInstallRoot = "",
    [string]$SecurityRegistryRoot = "HKCU:\SOFTWARE\HNC\HwpAutomation\Modules",
    [switch]$SkipSecurityModuleRegistration,
    [switch]$AllowExistingHwpProcesses,
    [switch]$ReuseExistingHwpObject
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function New-LogEntry {
    param(
        [string]$Kind,
        [hashtable]$Data
    )

    $entry = [ordered]@{
        timestamp = (Get-Date).ToString("o")
        kind = $Kind
    }
    foreach ($key in $Data.Keys) {
        $entry[$key] = $Data[$key]
    }
    return [pscustomobject]$entry
}

function Write-LogEntries {
    param(
        [string]$Path,
        [System.Collections.Generic.List[object]]$Entries
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }
    $directory = Split-Path -Parent $Path
    if ($directory) {
        New-Item -ItemType Directory -Force $directory | Out-Null
    }
    $Entries | ConvertTo-Json -Depth 8 | Set-Content -Path $Path -Encoding UTF8
}

function Get-RepoRoot {
    return Split-Path -Parent $PSScriptRoot
}

function Get-DefaultSecurityModuleInstallRoot {
    $userProfile = [Environment]::GetFolderPath("UserProfile")
    if (-not [string]::IsNullOrWhiteSpace($userProfile)) {
        return Join-Path $userProfile "jakal-hwpx\hancom-security"
    }
    return Join-Path (Get-RepoRoot) ".codex-temp\hancom-security"
}

function Resolve-SecurityModuleInstallRoot {
    param([string]$RequestedRoot)

    if (-not [string]::IsNullOrWhiteSpace($RequestedRoot)) {
        return [System.IO.Path]::GetFullPath($RequestedRoot)
    }
    return Get-DefaultSecurityModuleInstallRoot
}

function Get-ExpandedSecurityModulePath {
    param(
        [string]$Root,
        [string]$ModuleName
    )

    return Join-Path $Root ("expanded\{0}.dll" -f $ModuleName)
}

function Resolve-SecurityModulePath {
    param(
        [string]$RequestedPath,
        [string]$ModuleName,
        [string]$InstallRoot
    )

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        return [System.IO.Path]::GetFullPath($RequestedPath)
    }

    $resolvedInstallRoot = Resolve-SecurityModuleInstallRoot -RequestedRoot $InstallRoot
    $defaultInstallRoot = Get-DefaultSecurityModuleInstallRoot
    $repoFallback = Join-Path (Get-RepoRoot) (".codex-temp\hancom-security\expanded\{0}.dll" -f $ModuleName)
    $candidates = @(
        (Get-ExpandedSecurityModulePath -Root $resolvedInstallRoot -ModuleName $ModuleName),
        (Get-ExpandedSecurityModulePath -Root $defaultInstallRoot -ModuleName $ModuleName),
        $repoFallback
    ) | Select-Object -Unique
    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return [System.IO.Path]::GetFullPath($candidate)
        }
    }
    return ""
}

function Invoke-SecurityModuleSetup {
    param(
        [string]$ModuleName,
        [string]$ModulePath,
        [string]$InstallRoot,
        [string]$RegistryRoot
    )

    $setupScript = Join-Path $PSScriptRoot "setup_hancom_security_module.ps1"
    if (-not (Test-Path $setupScript)) {
        throw "Hancom security module setup script was not found: $setupScript"
    }

    $resolvedInstallRoot = Resolve-SecurityModuleInstallRoot -RequestedRoot $InstallRoot
    $resolvedExistingPath = Resolve-SecurityModulePath -RequestedPath $ModulePath -ModuleName $ModuleName -InstallRoot $resolvedInstallRoot
    $setupArgs = @{
        ModuleName = $ModuleName
        RegistryRoot = $RegistryRoot
        InstallRoot = $resolvedInstallRoot
    }
    if (-not [string]::IsNullOrWhiteSpace($ModulePath)) {
        $setupArgs["ModulePath"] = [System.IO.Path]::GetFullPath($ModulePath)
    }
    elseif (-not [string]::IsNullOrWhiteSpace($resolvedExistingPath)) {
        $setupArgs["ModulePath"] = $resolvedExistingPath
    }
    else {
        $setupArgs["DownloadIfMissing"] = $true
    }

    $setupResult = & $setupScript @setupArgs
    if ($setupResult -is [System.Array]) {
        return $setupResult[-1]
    }
    return $setupResult
}

function Get-HancomComObject {
    param([switch]$ReuseExisting)

    if ($ReuseExisting) {
        try {
            return [pscustomobject]@{
                Object = [System.Runtime.InteropServices.Marshal]::GetActiveObject("HWPFrame.HwpObject")
                Reused = $true
            }
        }
        catch {
        }
    }
    return [pscustomobject]@{
        Object = New-Object -ComObject HWPFrame.HwpObject
        Reused = $false
    }
}

function Test-TruthyResult {
    param([object]$Value)

    if ($null -eq $Value) {
        return $false
    }
    if ($Value -is [bool]) {
        return [bool]$Value
    }
    $stringValue = [string]$Value
    if ([string]::IsNullOrWhiteSpace($stringValue)) {
        return $false
    }
    return ($stringValue -ne "0" -and $stringValue.ToLowerInvariant() -ne "false")
}

function Get-HancomProcessSnapshot {
    return @(
        Get-Process -Name Hwp -ErrorAction SilentlyContinue |
            ForEach-Object {
                $pathValue = $null
                $startTimeValue = $null
                $titleValue = $null
                $respondingValue = $null
                try { $pathValue = $_.Path } catch { }
                try { $startTimeValue = $_.StartTime } catch { }
                try { $titleValue = $_.MainWindowTitle } catch { }
                try { $respondingValue = $_.Responding } catch { }
                [pscustomobject]@{
                    Id = $_.Id
                    ProcessName = $_.ProcessName
                    MainWindowTitle = $titleValue
                    Responding = $respondingValue
                    StartTime = $startTimeValue
                    Path = $pathValue
                }
            }
    )
}

function Invoke-HancomSmokeValidation {
    if ($SkipSecurityModuleRegistration) {
        throw "Skipping Hancom security module registration is disabled. Remove -SkipSecurityModuleRegistration."
    }

    $resolvedInput = (Resolve-Path -LiteralPath $InputPath).Path
    $resolvedOutput = [System.IO.Path]::GetFullPath($OutputPath)
    $logTarget = if ([string]::IsNullOrWhiteSpace($LogPath)) {
        Join-Path ([System.IO.Path]::GetDirectoryName($resolvedOutput)) "hancom-run-log.json"
    } else {
        [System.IO.Path]::GetFullPath($LogPath)
    }
    New-Item -ItemType Directory -Force ([System.IO.Path]::GetDirectoryName($resolvedOutput)) | Out-Null

    $entries = New-Object 'System.Collections.Generic.List[object]'
    $existingProcesses = @(Get-HancomProcessSnapshot)
    $existingProcessIds = @($existingProcesses | ForEach-Object { [int]$_.Id })
    if ($existingProcesses.Count -gt 0 -and -not $AllowExistingHwpProcesses) {
        $processList = ($existingProcesses | ForEach-Object { "{0} ({1})" -f $_.Id, $_.Path }) -join ", "
        throw "Existing Hwp.exe processes detected before smoke validation: $processList. Close those instances or rerun with -AllowExistingHwpProcesses."
    }

    $hwp = $null
    $reusedHwpObject = $false
    try {
        $setupResult = Invoke-SecurityModuleSetup -ModuleName $SecurityModuleName -ModulePath $SecurityModulePath -InstallRoot $SecurityModuleInstallRoot -RegistryRoot $SecurityRegistryRoot
        $resolvedSecurityModulePath = [string]$setupResult.ModulePath
        $entries.Add((New-LogEntry -Kind "security_module_setup" -Data @{
            moduleName = $SecurityModuleName
            modulePath = $resolvedSecurityModulePath
            installRoot = [string]$setupResult.InstallRoot
            registryRoot = $SecurityRegistryRoot
            registryValue = [string]$setupResult.RegistryValue
            downloaded = [bool]$setupResult.Downloaded
        })) | Out-Null

        $comSession = Get-HancomComObject -ReuseExisting:$ReuseExistingHwpObject
        $hwp = $comSession.Object
        $reusedHwpObject = [bool]$comSession.Reused
        $entries.Add((New-LogEntry -Kind "com" -Data @{ message = "Created HWPFrame.HwpObject"; reused = $reusedHwpObject })) | Out-Null

        try {
            $hwp.SetMessageBoxMode(0x00010000) | Out-Null
        }
        catch {
            $entries.Add((New-LogEntry -Kind "messagebox_mode_error" -Data @{ message = $_.Exception.Message })) | Out-Null
        }

        $registerResult = $hwp.RegisterModule("FilePathCheckDLL", $SecurityModuleName)
        $entries.Add((New-LogEntry -Kind "register_module" -Data @{
            moduleType = "FilePathCheckDLL"
            moduleName = $SecurityModuleName
            result = [string]$registerResult
        })) | Out-Null
        if (-not (Test-TruthyResult -Value $registerResult)) {
            throw "Hancom RegisterModule(FilePathCheckDLL, $SecurityModuleName) failed. No dialog-click fallback is used."
        }

        $entries.Add((New-LogEntry -Kind "open_start" -Data @{ path = $resolvedInput })) | Out-Null
        Write-LogEntries -Path $logTarget -Entries $entries
        $openResult = $hwp.Open($resolvedInput, "", "")
        $entries.Add((New-LogEntry -Kind "open" -Data @{ path = $resolvedInput; result = [string]$openResult })) | Out-Null
        Write-LogEntries -Path $logTarget -Entries $entries
        if (-not (Test-TruthyResult -Value $openResult)) {
            throw "Hancom Open failed for $resolvedInput."
        }

        $entries.Add((New-LogEntry -Kind "saveas_start" -Data @{ path = $resolvedOutput; format = $OutputFormat })) | Out-Null
        Write-LogEntries -Path $logTarget -Entries $entries
        $saveResult = $hwp.SaveAs($resolvedOutput, $OutputFormat, "")
        $entries.Add((New-LogEntry -Kind "saveas" -Data @{ path = $resolvedOutput; format = $OutputFormat; result = [string]$saveResult })) | Out-Null
        Write-LogEntries -Path $logTarget -Entries $entries
        if (-not (Test-TruthyResult -Value $saveResult)) {
            throw "Hancom SaveAs($OutputFormat) failed for $resolvedOutput."
        }
    }
    catch {
        $entries.Add((New-LogEntry -Kind "error" -Data @{ message = $_.Exception.Message })) | Out-Null
        throw
    }
    finally {
        if ($hwp) {
            if (-not $reusedHwpObject) {
                try { $hwp.Quit() } catch { }
            }
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($hwp) } catch { }
        }
        if (-not $reusedHwpObject) {
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            $spawnedProcesses = @()
            for ($attempt = 0; $attempt -lt 20; $attempt++) {
                $spawnedProcesses = @(
                    Get-HancomProcessSnapshot |
                        Where-Object { $existingProcessIds -notcontains [int]$_.Id }
                )
                if ($spawnedProcesses.Count -eq 0) {
                    break
                }
                Start-Sleep -Milliseconds 250
            }
            foreach ($process in $spawnedProcesses) {
                Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
                $entries.Add((New-LogEntry -Kind "process_cleanup" -Data @{
                    processId = [int]$process.Id
                    message = "Stopped the Hwp.exe process created by this smoke-validation run after COM Quit did not exit it."
                })) | Out-Null
            }
        }
        Write-LogEntries -Path $logTarget -Entries $entries
    }
}

if ([string]::IsNullOrWhiteSpace($InputPath)) {
    throw "InputPath is required."
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    throw "OutputPath is required."
}

Invoke-HancomSmokeValidation
