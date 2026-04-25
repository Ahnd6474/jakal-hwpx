param(
    [string]$InputPath,
    [string]$OutputPath,
    [string]$OutputFormat = "HWPX",
    [int]$TimeoutSeconds = 90,
    [int]$PollMilliseconds = 250,
    [switch]$ClickerOnly,
    [int]$ParentPid = 0,
    [string]$LogPath = "",
    [string[]]$AllowButtons = @(),
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

if ($AllowButtons.Count -eq 0) {
    $AllowButtons = @(
        [string]::Concat([char]0xBAA8, [char]0xB450, " ", [char]0xD5C8, [char]0xC6A9),
        [string]::Concat([char]0xD5C8, [char]0xC6A9),
        [string]([char]0xC608),
        [string]::Concat([char]0xD655, [char]0xC778),
        "Yes",
        "OK",
        "Allow",
        "Open",
        "Save"
    )
}

function Add-WindowInterop {
    if ("Jakal.Hancom.Win32" -as [type]) {
        return
    }

    Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Runtime.InteropServices;

namespace Jakal.Hancom {
    public static class Win32 {
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows(IntPtr hWnd, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
    }
}
"@
}

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

function Resolve-AllowButtons {
    param([string[]]$Buttons)

    if ($Buttons.Count -eq 1 -and $Buttons[0] -like "*|*") {
        return @($Buttons[0].Split("|") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    return $Buttons
}

function Get-RepoRoot {
    return Split-Path -Parent $PSScriptRoot
}

function Get-DefaultSecurityModuleInstallRoot {
    $localAppData = [Environment]::GetFolderPath("LocalApplicationData")
    if (-not [string]::IsNullOrWhiteSpace($localAppData)) {
        return Join-Path $localAppData "jakal-hwpx\hancom-security"
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

function Ensure-SecurityModuleRegistry {
    param(
        [string]$ModuleName,
        [string]$ModulePath,
        [string]$RegistryRoot
    )

    if ([string]::IsNullOrWhiteSpace($ModulePath)) {
        throw "Security module DLL could not be resolved. Run scripts/setup_hancom_security_module.ps1 first or pass -SecurityModulePath."
    }
    if (-not (Test-Path $ModulePath)) {
        throw "Security module DLL does not exist: $ModulePath"
    }
    if (-not (Test-Path $RegistryRoot)) {
        New-Item -Path $RegistryRoot -Force | Out-Null
    }
    New-ItemProperty -Path $RegistryRoot -Name $ModuleName -Value $ModulePath -PropertyType String -Force | Out-Null
    return Get-ItemProperty -Path $RegistryRoot -Name $ModuleName
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

function Get-OptionalValue {
    param(
        [object]$Object,
        [string]$Name
    )

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }
    return $property.Value
}

function Get-WindowTextValue {
    param([IntPtr]$Handle)

    $buffer = New-Object System.Text.StringBuilder 512
    [Jakal.Hancom.Win32]::GetWindowText($Handle, $buffer, $buffer.Capacity) | Out-Null
    return $buffer.ToString()
}

function Get-WindowClassValue {
    param([IntPtr]$Handle)

    $buffer = New-Object System.Text.StringBuilder 256
    [Jakal.Hancom.Win32]::GetClassName($Handle, $buffer, $buffer.Capacity) | Out-Null
    return $buffer.ToString()
}

function Get-WindowProcessIdValue {
    param([IntPtr]$Handle)

    [uint32]$processId = 0
    [Jakal.Hancom.Win32]::GetWindowThreadProcessId($Handle, [ref]$processId) | Out-Null
    return [int]$processId
}

function Get-ChildWindowObjects {
    param([IntPtr]$ParentHandle)

    $children = New-Object System.Collections.Generic.List[object]
    $callback = [Jakal.Hancom.Win32+EnumWindowsProc]{
        param([IntPtr]$ChildHandle, [IntPtr]$LParam)

        $children.Add([pscustomobject]@{
            Handle = $ChildHandle
            Title = Get-WindowTextValue -Handle $ChildHandle
            ClassName = Get-WindowClassValue -Handle $ChildHandle
            ProcessId = Get-WindowProcessIdValue -Handle $ChildHandle
        }) | Out-Null
        return $true
    }
    [Jakal.Hancom.Win32]::EnumChildWindows($ParentHandle, $callback, [IntPtr]::Zero) | Out-Null
    return $children
}

function Get-HancomDialogCandidates {
    param(
        [string[]]$ProcessNames,
        [string[]]$AllowButtons
    )

    $processMap = @{}
    foreach ($name in $ProcessNames) {
        foreach ($process in @(Get-Process -Name $name -ErrorAction SilentlyContinue)) {
            $processMap[$process.Id] = $process.ProcessName
        }
    }

    $results = New-Object System.Collections.Generic.List[object]
    $callback = [Jakal.Hancom.Win32+EnumWindowsProc]{
        param([IntPtr]$Handle, [IntPtr]$LParam)

        if (-not [Jakal.Hancom.Win32]::IsWindowVisible($Handle)) {
            return $true
        }

        $processId = Get-WindowProcessIdValue -Handle $Handle
        if (-not $processMap.ContainsKey($processId)) {
            return $true
        }

        $title = Get-WindowTextValue -Handle $Handle
        $className = Get-WindowClassValue -Handle $Handle
        $children = Get-ChildWindowObjects -ParentHandle $Handle
        $buttons = @($children | Where-Object { $_.ClassName -eq "Button" -and -not [string]::IsNullOrWhiteSpace($_.Title) })
        $matchedButtons = @($buttons | Where-Object { $AllowButtons -contains $_.Title })
        if ($matchedButtons.Count -eq 0) {
            return $true
        }

        $bodyTexts = @(
            $children |
                Where-Object { $_.ClassName -in @("Static", "RICHEDIT50W") -and -not [string]::IsNullOrWhiteSpace($_.Title) } |
                Select-Object -ExpandProperty Title
        )

        $results.Add([pscustomobject]@{
            Handle = $Handle
            ProcessId = $processId
            ProcessName = $processMap[$processId]
            Title = $title
            ClassName = $className
            BodyTexts = $bodyTexts
            Buttons = $buttons
            MatchedButtons = $matchedButtons
        }) | Out-Null
        return $true
    }

    [Jakal.Hancom.Win32]::EnumWindows($callback, [IntPtr]::Zero) | Out-Null
    return $results
}

function Invoke-HancomDialogAutoClicker {
    param(
        [int]$ParentPid,
        [int]$TimeoutSeconds,
        [int]$PollMilliseconds,
        [string]$LogPath,
        [string[]]$AllowButtons
    )

    Add-WindowInterop
    $entries = New-Object 'System.Collections.Generic.List[object]'
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $clickedHandles = New-Object 'System.Collections.Generic.HashSet[string]'

    while ((Get-Date) -lt $deadline) {
        if ($ParentPid -gt 0 -and -not (Get-Process -Id $ParentPid -ErrorAction SilentlyContinue)) {
            break
        }

        $dialogs = Get-HancomDialogCandidates -ProcessNames @("Hwp") -AllowButtons $AllowButtons
        foreach ($dialog in $dialogs) {
            $button = $null
            foreach ($label in $AllowButtons) {
                $button = $dialog.MatchedButtons | Where-Object { $_.Title -eq $label } | Select-Object -First 1
                if ($button) {
                    break
                }
            }
            if (-not $button) {
                continue
            }

            $key = ("{0}:{1}" -f $dialog.Handle.ToInt64(), $button.Handle.ToInt64())
            if ($clickedHandles.Contains($key)) {
                continue
            }

            $entries.Add((New-LogEntry -Kind "dialog" -Data @{
                processId = $dialog.ProcessId
                processName = $dialog.ProcessName
                title = $dialog.Title
                className = $dialog.ClassName
                bodyTexts = $dialog.BodyTexts
                button = $button.Title
            })) | Out-Null

            [Jakal.Hancom.Win32]::SendMessage($button.Handle, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
            $clickedHandles.Add($key) | Out-Null
            Start-Sleep -Milliseconds 300
        }

        Start-Sleep -Milliseconds $PollMilliseconds
    }

    Write-LogEntries -Path $LogPath -Entries $entries
}

function Invoke-HancomSmokeValidation {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [int]$TimeoutSeconds,
        [int]$PollMilliseconds,
        [string]$LogPath,
        [string[]]$AllowButtons,
        [string]$SecurityModuleName,
        [string]$SecurityModulePath,
        [string]$SecurityModuleInstallRoot,
        [string]$SecurityRegistryRoot,
        [switch]$SkipSecurityModuleRegistration,
        [switch]$AllowExistingHwpProcesses
    )

    Add-WindowInterop

    if ($SkipSecurityModuleRegistration) {
        throw "Skipping Hancom security module registration is disabled. Remove -SkipSecurityModuleRegistration."
    }

    $resolvedInput = (Resolve-Path -LiteralPath $InputPath).Path
    $resolvedOutput = [System.IO.Path]::GetFullPath($OutputPath)
    $logTarget = if ([string]::IsNullOrWhiteSpace($LogPath)) {
        Join-Path ([System.IO.Path]::GetDirectoryName($resolvedOutput)) "hancom-dialog-log.json"
    } else {
        [System.IO.Path]::GetFullPath($LogPath)
    }

    New-Item -ItemType Directory -Force ([System.IO.Path]::GetDirectoryName($resolvedOutput)) | Out-Null

    $clickerArgs = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $PSCommandPath,
        "-ClickerOnly",
        "-ParentPid", $PID,
        "-TimeoutSeconds", $TimeoutSeconds,
        "-PollMilliseconds", $PollMilliseconds,
        "-LogPath", $logTarget,
        "-AllowButtons", ($AllowButtons -join "|")
    )

    $clicker = Start-Process -FilePath "powershell.exe" -ArgumentList $clickerArgs -PassThru -WindowStyle Hidden
    $entries = New-Object 'System.Collections.Generic.List[object]'
    $existingProcesses = @(Get-HancomProcessSnapshot)

    if ($existingProcesses.Count -gt 0) {
        $processSnapshots = @(
            $existingProcesses | ForEach-Object {
                [ordered]@{
                    Id = Get-OptionalValue -Object $_ -Name "Id"
                    ProcessName = Get-OptionalValue -Object $_ -Name "ProcessName"
                    MainWindowTitle = Get-OptionalValue -Object $_ -Name "MainWindowTitle"
                    Responding = Get-OptionalValue -Object $_ -Name "Responding"
                    StartTime = [string](Get-OptionalValue -Object $_ -Name "StartTime")
                    Path = Get-OptionalValue -Object $_ -Name "Path"
                }
            }
        )
        $entries.Add((New-LogEntry -Kind "existing_hwp_processes" -Data @{
            count = $existingProcesses.Count
            processes = $processSnapshots
        })) | Out-Null
        if (-not $AllowExistingHwpProcesses) {
            $processList = ($existingProcesses | ForEach-Object {
                $pathValue = Get-OptionalValue -Object $_ -Name "Path"
                if ([string]::IsNullOrWhiteSpace([string]$pathValue)) {
                    $pathValue = "unknown-path"
                }
                "{0} ({1})" -f $_.Id, $pathValue
            }) -join ", "
            $runLog = [System.IO.Path]::ChangeExtension($logTarget, ".run.json")
            Write-LogEntries -Path $runLog -Entries $entries
            throw "Existing Hwp.exe processes detected before smoke validation: $processList. Close those instances or rerun with -AllowExistingHwpProcesses."
        }
    }

    $hwp = $null
    $reusedHwpObject = $false

    try {
        $setupResult = Invoke-SecurityModuleSetup -ModuleName $SecurityModuleName -ModulePath $SecurityModulePath -InstallRoot $SecurityModuleInstallRoot -RegistryRoot $SecurityRegistryRoot
        $resolvedSecurityModulePath = [string]$setupResult.ModulePath
        if ([string]::IsNullOrWhiteSpace($resolvedSecurityModulePath)) {
            $resolvedSecurityModulePath = Resolve-SecurityModulePath -RequestedPath $SecurityModulePath -ModuleName $SecurityModuleName -InstallRoot $SecurityModuleInstallRoot
        }
        $entries.Add((New-LogEntry -Kind "security_module_setup" -Data @{
            moduleName = $SecurityModuleName
            modulePath = $resolvedSecurityModulePath
            installRoot = [string]$setupResult.InstallRoot
            registryRoot = $SecurityRegistryRoot
            registryValue = [string]$setupResult.RegistryValue
            downloaded = [bool]$setupResult.Downloaded
        })) | Out-Null

        $registryEntry = Ensure-SecurityModuleRegistry -ModuleName $SecurityModuleName -ModulePath $resolvedSecurityModulePath -RegistryRoot $SecurityRegistryRoot
        $entries.Add((New-LogEntry -Kind "security_module_registry" -Data @{
            moduleName = $SecurityModuleName
            modulePath = $resolvedSecurityModulePath
            registryRoot = $SecurityRegistryRoot
            registryValue = $registryEntry.$SecurityModuleName
        })) | Out-Null

        $comSession = Get-HancomComObject -ReuseExisting:$ReuseExistingHwpObject
        $hwp = $comSession.Object
        $reusedHwpObject = [bool]$comSession.Reused
        $entries.Add((New-LogEntry -Kind "com" -Data @{
            message = "Created HWPFrame.HwpObject"
            reused = $reusedHwpObject
        })) | Out-Null

        try {
            $hwp.SetMessageBoxMode(0x00010000) | Out-Null
            $entries.Add((New-LogEntry -Kind "messagebox_mode" -Data @{ value = "0x00010000" })) | Out-Null
        } catch {
            $entries.Add((New-LogEntry -Kind "messagebox_mode_error" -Data @{ message = $_.Exception.Message })) | Out-Null
        }

        $registerResult = $hwp.RegisterModule("FilePathCheckDLL", $SecurityModuleName)
        $entries.Add((New-LogEntry -Kind "register_module" -Data @{
            moduleType = "FilePathCheckDLL"
            moduleName = $SecurityModuleName
            result = [string]$registerResult
        })) | Out-Null
        if (-not (Test-TruthyResult -Value $registerResult)) {
            throw "Hancom RegisterModule(FilePathCheckDLL, $SecurityModuleName) failed."
        }

        $openResult = $hwp.Open($resolvedInput, "", "")
        $entries.Add((New-LogEntry -Kind "open" -Data @{ path = $resolvedInput; result = [string]$openResult })) | Out-Null

        $saveResult = $hwp.SaveAs($resolvedOutput, $OutputFormat, "")
        $entries.Add((New-LogEntry -Kind "saveas" -Data @{ path = $resolvedOutput; format = $OutputFormat; result = [string]$saveResult })) | Out-Null

    }
    catch {
        $entries.Add((New-LogEntry -Kind "error" -Data @{ message = $_.Exception.Message })) | Out-Null
        throw
    }
    finally {
        if ($hwp) {
            if (-not $reusedHwpObject) {
                try {
                    $hwp.Quit()
                    $entries.Add((New-LogEntry -Kind "quit" -Data @{ result = "ok" })) | Out-Null
                }
                catch {
                    $entries.Add((New-LogEntry -Kind "quit_error" -Data @{ message = $_.Exception.Message })) | Out-Null
                }
            }
            try {
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($hwp)
            }
            catch {
            }
        }

        if ($clicker -and -not $clicker.HasExited) {
            Stop-Process -Id $clicker.Id -Force -ErrorAction SilentlyContinue
        }

        if (Test-Path $logTarget) {
            $dialogEntries = Get-Content $logTarget -Raw | ConvertFrom-Json
            foreach ($entry in @($dialogEntries)) {
                $entries.Add($entry) | Out-Null
            }
        }

        $runLog = [System.IO.Path]::ChangeExtension($logTarget, ".run.json")
        Write-LogEntries -Path $runLog -Entries $entries
    }
}

if ($ClickerOnly) {
    $AllowButtons = Resolve-AllowButtons -Buttons $AllowButtons
    Invoke-HancomDialogAutoClicker -ParentPid $ParentPid -TimeoutSeconds $TimeoutSeconds -PollMilliseconds $PollMilliseconds -LogPath $LogPath -AllowButtons $AllowButtons
    exit 0
}

if ([string]::IsNullOrWhiteSpace($InputPath)) {
    throw "InputPath is required unless -ClickerOnly is used."
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    throw "OutputPath is required unless -ClickerOnly is used."
}

 $AllowButtons = Resolve-AllowButtons -Buttons $AllowButtons
Invoke-HancomSmokeValidation -InputPath $InputPath -OutputPath $OutputPath -TimeoutSeconds $TimeoutSeconds -PollMilliseconds $PollMilliseconds -LogPath $LogPath -AllowButtons $AllowButtons -SecurityModuleName $SecurityModuleName -SecurityModulePath $SecurityModulePath -SecurityModuleInstallRoot $SecurityModuleInstallRoot -SecurityRegistryRoot $SecurityRegistryRoot -SkipSecurityModuleRegistration:$SkipSecurityModuleRegistration -AllowExistingHwpProcesses:$AllowExistingHwpProcesses
