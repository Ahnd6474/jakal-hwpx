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

function Resolve-InstallRoot {
    param([string]$RequestedRoot)

    if (-not [string]::IsNullOrWhiteSpace($RequestedRoot)) {
        return [System.IO.Path]::GetFullPath($RequestedRoot)
    }
    return Join-Path (Get-RepoRoot) ".codex-temp\hancom-security"
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
