param(
    [string]$ModuleName = "FilePathCheckerModuleExample",
    [string]$ModulePath = "",
    [string]$RegistryRoot = "HKCU:\SOFTWARE\HNC\HwpAutomation\Modules",
    [string]$InstallRoot = "",
    [switch]$DownloadIfMissing
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$OfficialSecurityModuleZipUrl = "https://github.com/hancom-io/devcenter-archive/raw/main/hwp-automation/%EB%B3%B4%EC%95%88%EB%AA%A8%EB%93%88%28Automation%29.zip"

function Get-RepoRoot {
    return Split-Path -Parent $PSScriptRoot
}

function Get-DefaultInstallRoot {
    # Packaged hosts can virtualize LocalApplicationData. Build the ordinary
    # user path from UserProfile so Hwp.exe can load the same DLL.
    $userProfile = [Environment]::GetFolderPath("UserProfile")
    if (-not [string]::IsNullOrWhiteSpace($userProfile)) {
        return Join-Path $userProfile "jakal-hwpx\hancom-security"
    }
    return Join-Path (Get-RepoRoot) ".codex-temp\hancom-security"
}

function Resolve-InstallRoot {
    param([string]$RequestedRoot)

    if (-not [string]::IsNullOrWhiteSpace($RequestedRoot)) {
        return [System.IO.Path]::GetFullPath($RequestedRoot)
    }
    return Get-DefaultInstallRoot
}

function Get-ExpandedModulePath {
    param(
        [string]$Root,
        [string]$Name
    )

    return Join-Path $Root ("expanded\{0}.dll" -f $Name)
}

function Install-OfficialSecurityModule {
    param(
        [string]$Root,
        [string]$Name
    )

    $zipPath = Join-Path $Root "hancom-automation-security.zip"
    $expandedDir = Join-Path $Root "expanded"
    New-Item -ItemType Directory -Force $Root | Out-Null

    if (-not (Test-Path $zipPath)) {
        $ProgressPreference = "SilentlyContinue"
        Invoke-WebRequest -Uri $OfficialSecurityModuleZipUrl -OutFile $zipPath
    }

    if (-not (Test-Path (Get-ExpandedModulePath -Root $Root -Name $Name))) {
        Expand-Archive -LiteralPath $zipPath -DestinationPath $expandedDir -Force
    }

    $modulePath = Get-ExpandedModulePath -Root $Root -Name $Name
    if (-not (Test-Path $modulePath)) {
        throw "Expanded security module DLL was not found: $modulePath"
    }
    return [System.IO.Path]::GetFullPath($modulePath)
}

function Resolve-ModulePath {
    param(
        [string]$RequestedPath,
        [string]$Root,
        [string]$Name,
        [switch]$DownloadIfMissing
    )

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        $resolved = [System.IO.Path]::GetFullPath($RequestedPath)
        if (-not (Test-Path $resolved)) {
            throw "Security module DLL does not exist: $resolved"
        }
        return $resolved
    }

    $defaultPath = Get-ExpandedModulePath -Root $Root -Name $Name
    if (Test-Path $defaultPath) {
        return [System.IO.Path]::GetFullPath($defaultPath)
    }

    if ($DownloadIfMissing) {
        return Install-OfficialSecurityModule -Root $Root -Name $Name
    }

    throw "Security module DLL was not found. Pass -ModulePath or rerun with -DownloadIfMissing."
}

function Set-SecurityModuleRegistry {
    param(
        [string]$Root,
        [string]$Name,
        [string]$Path
    )

    if (-not (Test-Path $Root)) {
        New-Item -Path $Root -Force | Out-Null
    }
    New-ItemProperty -Path $Root -Name $Name -Value $Path -PropertyType String -Force | Out-Null

    # Codex Desktop is an MSIX-packaged host. Registry writes from its child
    # processes can be redirected to the package hive, which is invisible to
    # the unpackaged Hancom process. Run one short-lived, current-user task so
    # Windows writes the value to the ordinary HKCU hive. No UI interaction or
    # elevation is involved.
    if (-not [string]::IsNullOrWhiteSpace($env:CODEX_WINDOWS_SANDBOX_PACKAGE_FAMILY)) {
        if ($Root -notmatch '^HKCU:\\') {
            throw "Packaged-host registry bridging supports HKCU roots only: $Root"
        }
        if ($Name -notmatch '^[A-Za-z0-9_.-]+$') {
            throw "Security module name contains unsupported characters: $Name"
        }

        $nativeRoot = $Root -replace '^HKCU:\\', 'HKCU\'
        $taskName = "JakalHwpxSecurityModule-{0}-{1}" -f $PID, ([guid]::NewGuid().ToString("N").Substring(0, 8))
        $taskCommand = 'reg.exe add "{0}" /v "{1}" /t REG_SZ /d "{2}" /f' -f $nativeRoot, $Name, $Path
        $startTime = (Get-Date).AddMinutes(5).ToString("HH:mm")
        try {
            & schtasks.exe /Create /TN $taskName /TR $taskCommand /SC ONCE /ST $startTime /RL LIMITED /F | Out-Null
            if ($LASTEXITCODE -ne 0) {
                throw "Unable to create the temporary registry bridge task (exit $LASTEXITCODE)."
            }

            $scheduler = New-Object -ComObject "Schedule.Service"
            $scheduler.Connect()
            $task = $scheduler.GetFolder("\").GetTask($taskName)
            $previousRunTime = $task.LastRunTime
            & schtasks.exe /Run /TN $taskName | Out-Null
            if ($LASTEXITCODE -ne 0) {
                throw "Unable to run the temporary registry bridge task (exit $LASTEXITCODE)."
            }

            $deadline = (Get-Date).AddSeconds(20)
            do {
                Start-Sleep -Milliseconds 100
                $task = $scheduler.GetFolder("\").GetTask($taskName)
                $hasRun = $task.LastRunTime -gt $previousRunTime
                $hasFinished = $hasRun -and $task.State -ne 4
            } while (-not $hasFinished -and (Get-Date) -lt $deadline)
            if (-not $hasFinished) {
                throw "The temporary registry bridge task did not finish within 20 seconds."
            }
            if ($task.LastTaskResult -ne 0) {
                throw "The temporary registry bridge task failed with result $($task.LastTaskResult)."
            }
        }
        finally {
            & schtasks.exe /Delete /TN $taskName /F 2>$null | Out-Null
        }
    }
    return Get-ItemProperty -Path $Root -Name $Name
}

$resolvedInstallRoot = Resolve-InstallRoot -RequestedRoot $InstallRoot
$resolvedModulePath = Resolve-ModulePath -RequestedPath $ModulePath -Root $resolvedInstallRoot -Name $ModuleName -DownloadIfMissing:$DownloadIfMissing
$registryEntry = Set-SecurityModuleRegistry -Root $RegistryRoot -Name $ModuleName -Path $resolvedModulePath

[pscustomobject]@{
    ModuleName = $ModuleName
    ModulePath = $resolvedModulePath
    RegistryRoot = $RegistryRoot
    RegistryValue = $registryEntry.$ModuleName
    InstallRoot = $resolvedInstallRoot
    Downloaded = [bool](Test-Path (Join-Path $resolvedInstallRoot "hancom-automation-security.zip"))
}
