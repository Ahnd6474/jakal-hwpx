param(
    [string]$OutputRoot = "",
    [string[]]$Kinds = @("rect", "ellipse", "arc", "polygon"),
    [int]$Width = 3800,
    [int]$Height = 2000,
    [int]$LineWidth = 120,
    [string]$LineColor = "#302010",
    [string]$FillColor = "#A1B2C3",
    [string]$SecurityModuleName = "FilePathCheckerModuleExample",
    [string]$SecurityModulePath = "",
    [string]$SecurityModuleInstallRoot = "",
    [string]$SecurityRegistryRoot = "HKCU:\SOFTWARE\HNC\HwpAutomation\Modules"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-RepoRoot {
    return Split-Path -Parent $PSScriptRoot
}

function Resolve-OutputRoot {
    param([string]$RequestedPath)

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        return [System.IO.Path]::GetFullPath($RequestedPath)
    }
    return Join-Path (Get-RepoRoot) ".codex-temp\hancom-shape-donors"
}

function Convert-HexColorToColorRef {
    param([string]$Color)

    $normalized = $Color.Trim()
    if ($normalized.StartsWith("#")) {
        $normalized = $normalized.Substring(1)
    }
    if ($normalized.Length -ne 6) {
        throw "Expected a 6-digit RGB color, got: $Color"
    }

    $red = [Convert]::ToInt32($normalized.Substring(0, 2), 16)
    $green = [Convert]::ToInt32($normalized.Substring(2, 2), 16)
    $blue = [Convert]::ToInt32($normalized.Substring(4, 2), 16)
    return ($blue -bor ($green -shl 8) -bor ($red -shl 16))
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

function Get-ShapeSpec {
    param([string]$Kind)

    switch ($Kind.ToLowerInvariant()) {
        "rect" {
            return @{
                Action = "DrawObjCreatorRectangle"
                Configure = {
                    param($shapeObject)
                }
            }
        }
        "ellipse" {
            return @{
                Action = "DrawObjCreatorEllipse"
                Configure = {
                    param($shapeObject)
                }
            }
        }
        "arc" {
            return @{
                Action = "DrawObjCreatorArc"
                Configure = {
                    param($shapeObject)
                    $shapeObject.ShapeDrawArcType.type = 0
                }
            }
        }
        "polygon" {
            return @{
                Action = "DrawObjCreatorPolygon"
                Configure = {
                    param($shapeObject)
                }
            }
        }
        default {
            throw "Unsupported kind: $Kind. Supported values: rect, ellipse, arc, polygon."
        }
    }
}

$resolvedOutputRoot = Resolve-OutputRoot -RequestedPath $OutputRoot
New-Item -ItemType Directory -Force $resolvedOutputRoot | Out-Null

Ensure-SecurityModule `
    -ModuleName $SecurityModuleName `
    -ModulePath $SecurityModulePath `
    -InstallRoot $SecurityModuleInstallRoot `
    -RegistryRoot $SecurityRegistryRoot

$lineColorRef = Convert-HexColorToColorRef -Color $LineColor
$fillColorRef = Convert-HexColorToColorRef -Color $FillColor
$results = New-Object System.Collections.Generic.List[object]

foreach ($kind in $Kinds) {
    $spec = Get-ShapeSpec -Kind $kind
    $outputPath = Join-Path $resolvedOutputRoot ("{0}.hwp" -f $kind.ToLowerInvariant())
    $hwp = $null

    try {
        $hwp = New-Object -ComObject HWPFrame.HwpObject
        $registerResult = $hwp.RegisterModule("FilePathCheckDLL", $SecurityModuleName)
        if (-not $registerResult) {
            throw "RegisterModule(FilePathCheckDLL, $SecurityModuleName) failed."
        }
        $hwp.Run("FileNew") | Out-Null

        $action = $hwp.CreateAction($spec.Action)
        if ($null -eq $action) {
            throw "CreateAction($($spec.Action)) returned null."
        }

        $shapeObject = $hwp.HParameterSet.HShapeObject
        $set = $shapeObject.HSet
        $action.GetDefault($set) | Out-Null
        $shapeObject.Width = $Width
        $shapeObject.Height = $Height
        $shapeObject.TreatAsChar = 1
        $shapeObject.ShapeDrawLineAttr.Color = $lineColorRef
        $shapeObject.ShapeDrawLineAttr.Width = $LineWidth
        $shapeObject.ShapeDrawFillAttr.type = 1
        $shapeObject.ShapeDrawFillAttr.WinBrushFaceColor = $fillColorRef
        & $spec.Configure $shapeObject

        $executeResult = $action.Execute($set)
        $saveResult = $hwp.SaveAs($outputPath, "HWP", "")
        $results.Add([pscustomobject]@{
            Kind = $kind.ToLowerInvariant()
            Action = $spec.Action
            Execute = [bool]$executeResult
            Save = [bool]$saveResult
            OutputPath = $outputPath
            Status = if ($executeResult -and $saveResult) { "ok" } else { "failed" }
        }) | Out-Null
    }
    finally {
        if ($null -ne $hwp) {
            try {
                $hwp.Quit()
            }
            catch {
            }
        }
    }
}

$results
