param(
    [string]$DonorPath = "",
    [string]$OutputRoot = "",
    [string[]]$CandidateActions = @("ShapeObjDialog", "TextartAttr", "TextArtM", "TEXTART", "TextShape", "TextEffects"),
    [string]$SeedText = "TEXTART-SEED",
    [string]$SecurityModuleName = "FilePathCheckerModuleExample",
    [string]$SecurityModulePath = "",
    [string]$SecurityModuleInstallRoot = "",
    [string]$SecurityRegistryRoot = "HKCU:\SOFTWARE\HNC\HwpAutomation\Modules"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ($CandidateActions.Count -eq 1 -and $CandidateActions[0] -like "*,*") {
    $CandidateActions = @(
        $CandidateActions[0].Split(",") |
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
    return Join-Path (Get-RepoRoot) ".codex-temp\hancom-textart-probe"
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

function Get-ComPropertyValue {
    param(
        [object]$Object,
        [string]$Name
    )

    try {
        return $Object.$Name
    }
    catch {
        return $null
    }
}

function Set-ComPropertyValue {
    param(
        [object]$Object,
        [string]$Name,
        [object]$Value
    )

    try {
        $Object.$Name = $Value
        return $true
    }
    catch {
        return $false
    }
}

function Get-ComPropertyNames {
    param([object]$Object)

    if ($null -eq $Object) {
        return @()
    }
    try {
        return @($Object | Get-Member -MemberType Property | Select-Object -ExpandProperty Name)
    }
    catch {
        return @()
    }
}

function Convert-ComSnapshotValue {
    param(
        [object]$Value,
        [int]$Depth = 1
    )

    if ($null -eq $Value) {
        return $null
    }
    if ($Value -is [string] -or $Value -is [ValueType]) {
        return $Value
    }
    if ($Depth -le 0) {
        return "$Value"
    }

    $names = Get-ComPropertyNames -Object $Value
    if (-not $names) {
        return "$Value"
    }

    $snapshot = [ordered]@{}
    foreach ($name in $names) {
        if ($name -eq "HSet") {
            continue
        }
        $child = Get-ComPropertyValue -Object $Value -Name $name
        if ($null -eq $child) {
            continue
        }
        $snapshot[$name] = Convert-ComSnapshotValue -Value $child -Depth ($Depth - 1)
    }
    return $snapshot
}

function Get-TextartAssignments {
    param(
        [string]$TextValue
    )

    return [ordered]@{
        Apply = 1
        string = $TextValue
        FontName = "Malgun Gothic"
        FontStyle = "Bold"
        FontType = 1
        Shape = 0
        AlignType = 0
        CharSpacing = 100
        LineSpacing = 120
        NumberOfLines = 1
        ShadowColor = 0
        ShadowOffsetX = 0
        ShadowOffsetY = 0
        ShadowType = 0
    }
}

function Apply-TextartAssignments {
    param(
        [string]$ParameterSetName,
        [object]$ParameterSet,
        [string]$TextValue,
        [switch]$IncludeShapeSizing
    )

    $target = if ($ParameterSetName -eq "HShapeObject") { $ParameterSet.ShapeDrawTextart } else { $ParameterSet }
    if ($null -eq $target) {
        return
    }

    if ($IncludeShapeSizing -and $ParameterSetName -eq "HShapeObject") {
        Set-ComPropertyValue -Object $ParameterSet -Name "Width" -Value 22500 | Out-Null
        Set-ComPropertyValue -Object $ParameterSet -Name "Height" -Value 5000 | Out-Null
        Set-ComPropertyValue -Object $ParameterSet -Name "TreatAsChar" -Value 1 | Out-Null
    }

    foreach ($entry in (Get-TextartAssignments -TextValue $TextValue).GetEnumerator()) {
        Set-ComPropertyValue -Object $target -Name $entry.Key -Value $entry.Value | Out-Null
    }

    if ($ParameterSetName -eq "HShapeObject") {
        Set-ComPropertyValue -Object $ParameterSet.ShapeDrawLineAttr -Name "Color" -Value 0x00665544 | Out-Null
        Set-ComPropertyValue -Object $ParameterSet.ShapeDrawFillAttr -Name "type" -Value 1 | Out-Null
        Set-ComPropertyValue -Object $ParameterSet.ShapeDrawFillAttr -Name "WinBrushFaceColor" -Value 0x00FFEEDD | Out-Null
    }
}

function Select-GraphicControl {
    param(
        [object]$Hwp,
        [object]$Control
    )

    if ($null -eq $Control) {
        throw "Control is null."
    }
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
            return $control
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
            return $control
        }
        $control = $control.Prev
    }
    throw "Could not find the last gso control."
}

function Get-ParameterSnapshot {
    param(
        [object]$Hwp,
        [string]$ActionName,
        [string]$ParameterSetName
    )

    $action = $Hwp.CreateAction($ActionName)
    if ($null -eq $action) {
        return [pscustomobject]@{
            action = $ActionName
            parameter_set = $ParameterSetName
            create_action = $false
            get_default = $false
            values = @{}
        }
    }

    $parameterSet = Get-ComPropertyValue -Object $Hwp.HParameterSet -Name $ParameterSetName
    if ($null -eq $parameterSet) {
        return [pscustomobject]@{
            action = $ActionName
            parameter_set = $ParameterSetName
            create_action = $true
            get_default = $false
            values = @{}
        }
    }

    $result = $false
    try {
        $result = [bool]$action.GetDefault($parameterSet.HSet)
    }
    catch {
        $result = $false
    }

    $values = [ordered]@{}
    $scalarNames = if ($ParameterSetName -eq "HDrawTextart") {
        @(
            "Apply",
            "string",
            "FontName",
            "FontStyle",
            "FontType",
            "Shape",
            "AlignType",
            "CharSpacing",
            "LineSpacing",
            "NumberOfLines",
            "ShadowColor",
            "ShadowOffsetX",
            "ShadowOffsetY",
            "ShadowType"
        )
    }
    else {
        @(
            "Width",
            "Height",
            "TreatAsChar",
            "ShapeType",
            "ShapeComment",
            "VisualString"
        )
    }
    foreach ($name in $scalarNames) {
        $value = Get-ComPropertyValue -Object $parameterSet -Name $name
        if ($null -ne $value) {
            $values[$name] = $value
        }
    }

    $nestedValues = [ordered]@{}
    if ($ParameterSetName -eq "HShapeObject") {
        foreach ($name in @("ShapeDrawTextart", "ShapeDrawLineAttr", "ShapeDrawFillAttr", "ShapeDrawShadow")) {
            $value = Get-ComPropertyValue -Object $parameterSet -Name $name
            if ($null -ne $value) {
                $nestedValues[$name] = Convert-ComSnapshotValue -Value $value -Depth 2
            }
        }
    }
    else {
        $nestedValues["self"] = Convert-ComSnapshotValue -Value $parameterSet -Depth 1
    }

    return [pscustomobject]@{
        action = $ActionName
        parameter_set = $ParameterSetName
        create_action = $true
        get_default = $result
        values = $values
        nested_values = $nestedValues
    }
}

function New-HwpObject {
    param(
        [string]$ModuleName
    )

    $hwp = New-Object -ComObject HWPFrame.HwpObject
    $registerResult = $hwp.RegisterModule("FilePathCheckDLL", $ModuleName)
    if (-not $registerResult) {
        throw "RegisterModule(FilePathCheckDLL, $ModuleName) failed."
    }
    return $hwp
}

function Create-TextboxDocument {
    param(
        [string]$ModuleName,
        [string]$OutputPath
    )

    $hwp = New-HwpObject -ModuleName $ModuleName
    try {
        $hwp.Run("FileNew") | Out-Null
        $action = $hwp.CreateAction("DrawObjCreatorTextBox")
        if ($null -eq $action) {
            throw "DrawObjCreatorTextBox action is unavailable."
        }
        $shapeObject = $hwp.HParameterSet.HShapeObject
        $action.GetDefault($shapeObject.HSet) | Out-Null
        $shapeObject.Width = 22500
        $shapeObject.Height = 5000
        $shapeObject.TreatAsChar = 1
        $shapeObject.ShapeDrawLineAttr.Color = 0x00665544
        $shapeObject.ShapeDrawFillAttr.type = 1
        $shapeObject.ShapeDrawFillAttr.WinBrushFaceColor = 0x00FFEEDD
        $action.Execute($shapeObject.HSet) | Out-Null
        Select-LastGso -Hwp $hwp | Out-Null
        $hwp.SaveAs($OutputPath, "HWP", "") | Out-Null
        return $hwp
    }
    catch {
        try {
            $hwp.Quit()
        }
        catch {
        }
        throw
    }
}

function Create-PlainTextDocument {
    param(
        [string]$ModuleName,
        [string]$OutputPath,
        [string]$Text
    )

    $hwp = New-HwpObject -ModuleName $ModuleName
    try {
        $hwp.Run("FileNew") | Out-Null
        $action = $hwp.CreateAction("InsertText")
        if ($null -eq $action) {
            throw "InsertText action is unavailable."
        }

        $insertText = $hwp.HParameterSet.HInsertText
        $action.GetDefault($insertText.HSet) | Out-Null
        $insertText.Text = $Text
        $action.Execute($insertText.HSet) | Out-Null
        $hwp.Run("SelectAll") | Out-Null
        $hwp.SaveAs($OutputPath, "HWP", "") | Out-Null
        return $hwp
    }
    catch {
        try {
            $hwp.Quit()
        }
        catch {
        }
        throw
    }
}

function Try-ExecuteTextartAction {
    param(
        [string]$ModuleName,
        [string]$ActionName,
        [string]$ParameterSetName,
        [string]$OutputPath
    )

    $hwp = Create-TextboxDocument -ModuleName $ModuleName -OutputPath (Join-Path ([System.IO.Path]::GetDirectoryName($OutputPath)) "textbox_seed.hwp")
    try {
        $action = $hwp.CreateAction($ActionName)
        if ($null -eq $action) {
            return [pscustomobject]@{
                action = $ActionName
                created = $false
                get_default = $false
                execute = $false
                parameter_set = $ParameterSetName
                output_path = $OutputPath
            }
        }

        $parameterSet = Get-ComPropertyValue -Object $hwp.HParameterSet -Name $parameterSetName
        if ($null -eq $parameterSet) {
            return [pscustomobject]@{
                action = $ActionName
                created = $true
                get_default = $false
                execute = $false
                parameter_set = $ParameterSetName
                output_path = $OutputPath
            }
        }

        $getDefault = $false
        try {
            $getDefault = [bool]$action.GetDefault($parameterSet.HSet)
        }
        catch {
            $getDefault = $false
        }

        Apply-TextartAssignments `
            -ParameterSetName $ParameterSetName `
            -ParameterSet $parameterSet `
            -TextValue "PROBE-TEXTART" `
            -IncludeShapeSizing

        $execute = $false
        try {
            $execute = [bool]$action.Execute($parameterSet.HSet)
        }
        catch {
            $execute = $false
        }
        $save = [bool]$hwp.SaveAs($OutputPath, "HWP", "")
        return [pscustomobject]@{
            action = $ActionName
            created = $true
            get_default = $getDefault
            execute = $execute
            save = $save
            parameter_set = $ParameterSetName
            output_path = $OutputPath
        }
    }
    finally {
        try {
            $hwp.Quit()
        }
        catch {
        }
    }
}

function Try-ExecuteDonorTextartAction {
    param(
        [string]$ModuleName,
        [string]$DonorPath,
        [string]$ActionName,
        [string]$ParameterSetName,
        [string]$OutputPath
    )

    $hwp = New-HwpObject -ModuleName $ModuleName
    try {
        $hwp.Open($DonorPath, "", "") | Out-Null
        Select-FirstGso -Hwp $hwp | Out-Null
        $action = $hwp.CreateAction($ActionName)
        if ($null -eq $action) {
            return [pscustomobject]@{
                action = $ActionName
                created = $false
                get_default = $false
                execute = $false
                parameter_set = $ParameterSetName
                output_path = $OutputPath
            }
        }

        $parameterSet = Get-ComPropertyValue -Object $hwp.HParameterSet -Name $ParameterSetName
        if ($null -eq $parameterSet) {
            return [pscustomobject]@{
                action = $ActionName
                created = $true
                get_default = $false
                execute = $false
                parameter_set = $ParameterSetName
                output_path = $OutputPath
            }
        }

        $getDefault = $false
        try {
            $getDefault = [bool]$action.GetDefault($parameterSet.HSet)
        }
        catch {
            $getDefault = $false
        }

        Apply-TextartAssignments `
            -ParameterSetName $ParameterSetName `
            -ParameterSet $parameterSet `
            -TextValue "DONOR-PROBE-TEXTART"

        $execute = $false
        try {
            $execute = [bool]$action.Execute($parameterSet.HSet)
        }
        catch {
            $execute = $false
        }
        $save = [bool]$hwp.SaveAs($OutputPath, "HWP", "")
        return [pscustomobject]@{
            action = $ActionName
            created = $true
            get_default = $getDefault
            execute = $execute
            save = $save
            parameter_set = $ParameterSetName
            output_path = $OutputPath
        }
    }
    finally {
        try {
            $hwp.Quit()
        }
        catch {
        }
    }
}

function Try-ExecutePlainTextAction {
    param(
        [string]$ModuleName,
        [string]$ActionName,
        [string]$ParameterSetName,
        [string]$OutputPath,
        [string]$TextValue
    )

    $hwp = Create-PlainTextDocument `
        -ModuleName $ModuleName `
        -OutputPath (Join-Path ([System.IO.Path]::GetDirectoryName($OutputPath)) "plain_text_seed.hwp") `
        -Text $TextValue
    try {
        $action = $hwp.CreateAction($ActionName)
        if ($null -eq $action) {
            return [pscustomobject]@{
                action = $ActionName
                created = $false
                get_default = $false
                execute = $false
                parameter_set = $ParameterSetName
                output_path = $OutputPath
            }
        }

        $parameterSet = Get-ComPropertyValue -Object $hwp.HParameterSet -Name $ParameterSetName
        if ($null -eq $parameterSet) {
            return [pscustomobject]@{
                action = $ActionName
                created = $true
                get_default = $false
                execute = $false
                parameter_set = $ParameterSetName
                output_path = $OutputPath
            }
        }

        $getDefault = $false
        try {
            $getDefault = [bool]$action.GetDefault($parameterSet.HSet)
        }
        catch {
            $getDefault = $false
        }

        Apply-TextartAssignments `
            -ParameterSetName $ParameterSetName `
            -ParameterSet $parameterSet `
            -TextValue "PLAIN-PROBE-TEXTART"

        $execute = $false
        try {
            $execute = [bool]$action.Execute($parameterSet.HSet)
        }
        catch {
            $execute = $false
        }
        $save = [bool]$hwp.SaveAs($OutputPath, "HWP", "")
        return [pscustomobject]@{
            action = $ActionName
            created = $true
            get_default = $getDefault
            execute = $execute
            save = $save
            parameter_set = $ParameterSetName
            output_path = $OutputPath
        }
    }
    finally {
        try {
            $hwp.Quit()
        }
        catch {
        }
    }
}

$resolvedDonorPath = Resolve-DonorPath -RequestedPath $DonorPath
$resolvedOutputRoot = Resolve-OutputRoot -RequestedPath $OutputRoot
New-Item -ItemType Directory -Force $resolvedOutputRoot | Out-Null

Ensure-SecurityModule `
    -ModuleName $SecurityModuleName `
    -ModulePath $SecurityModulePath `
    -InstallRoot $SecurityModuleInstallRoot `
    -RegistryRoot $SecurityRegistryRoot

$result = [ordered]@{
    donor_path = $resolvedDonorPath
    output_root = $resolvedOutputRoot
    donor = @()
    textbox = @()
    plain_text = @()
    donor_execute_attempts = @()
    execute_attempts = @()
    plain_text_execute_attempts = @()
}

$donorHwp = $null
try {
    $donorHwp = New-HwpObject -ModuleName $SecurityModuleName
    $donorHwp.Open($resolvedDonorPath, "", "") | Out-Null
    Select-FirstGso -Hwp $donorHwp | Out-Null
    foreach ($actionName in $CandidateActions) {
        $result.donor += Get-ParameterSnapshot -Hwp $donorHwp -ActionName $actionName -ParameterSetName "HShapeObject"
        $result.donor += Get-ParameterSnapshot -Hwp $donorHwp -ActionName $actionName -ParameterSetName "HDrawTextart"
    }
}
finally {
    if ($null -ne $donorHwp) {
        try {
            $donorHwp.Quit()
        }
        catch {
        }
    }
}

$textboxSeedPath = Join-Path $resolvedOutputRoot "textbox_seed.hwp"
$textboxHwp = $null
try {
    $textboxHwp = Create-TextboxDocument -ModuleName $SecurityModuleName -OutputPath $textboxSeedPath
    foreach ($actionName in $CandidateActions) {
        $result.textbox += Get-ParameterSnapshot -Hwp $textboxHwp -ActionName $actionName -ParameterSetName "HShapeObject"
        $result.textbox += Get-ParameterSnapshot -Hwp $textboxHwp -ActionName $actionName -ParameterSetName "HDrawTextart"
    }
}
finally {
    if ($null -ne $textboxHwp) {
        try {
            $textboxHwp.Quit()
        }
        catch {
        }
    }
}

$plainTextSeedPath = Join-Path $resolvedOutputRoot "plain_text_seed.hwp"
$plainTextHwp = $null
try {
    $plainTextHwp = Create-PlainTextDocument `
        -ModuleName $SecurityModuleName `
        -OutputPath $plainTextSeedPath `
        -Text $SeedText
    foreach ($actionName in $CandidateActions) {
        $result.plain_text += Get-ParameterSnapshot -Hwp $plainTextHwp -ActionName $actionName -ParameterSetName "HShapeObject"
        $result.plain_text += Get-ParameterSnapshot -Hwp $plainTextHwp -ActionName $actionName -ParameterSetName "HDrawTextart"
    }
}
finally {
    if ($null -ne $plainTextHwp) {
        try {
            $plainTextHwp.Quit()
        }
        catch {
        }
    }
}

foreach ($attempt in @(
    @{ action = "ShapeObjDialog"; parameter_set = "HShapeObject" },
    @{ action = "TextArtModify"; parameter_set = "HShapeObject" },
    @{ action = "TextArtModify"; parameter_set = "HDrawTextart" },
    @{ action = "ChangeObject"; parameter_set = "HShapeObject" },
    @{ action = "ChangeObject"; parameter_set = "HDrawTextart" }
)) {
    $outputPath = Join-Path $resolvedOutputRoot ("textbox_{0}_{1}.hwp" -f $attempt.action, $attempt.parameter_set)
    $result.execute_attempts += Try-ExecuteTextartAction `
        -ModuleName $SecurityModuleName `
        -ActionName $attempt.action `
        -ParameterSetName $attempt.parameter_set `
        -OutputPath $outputPath
}

foreach ($attempt in @(
    @{ action = "TextArtModify"; parameter_set = "HShapeObject" },
    @{ action = "TextArtModify"; parameter_set = "HDrawTextart" }
)) {
    $outputPath = Join-Path $resolvedOutputRoot ("donor_{0}_{1}.hwp" -f $attempt.action, $attempt.parameter_set)
    $result.donor_execute_attempts += Try-ExecuteDonorTextartAction `
        -ModuleName $SecurityModuleName `
        -DonorPath $resolvedDonorPath `
        -ActionName $attempt.action `
        -ParameterSetName $attempt.parameter_set `
        -OutputPath $outputPath
}

foreach ($attempt in @(
    @{ action = "TextArtModify"; parameter_set = "HShapeObject" },
    @{ action = "TextArtModify"; parameter_set = "HDrawTextart" },
    @{ action = "ChangeObject"; parameter_set = "HShapeObject" },
    @{ action = "ChangeObject"; parameter_set = "HDrawTextart" }
)) {
    $outputPath = Join-Path $resolvedOutputRoot ("plain_text_{0}_{1}.hwp" -f $attempt.action, $attempt.parameter_set)
    $result.plain_text_execute_attempts += Try-ExecutePlainTextAction `
        -ModuleName $SecurityModuleName `
        -ActionName $attempt.action `
        -ParameterSetName $attempt.parameter_set `
        -OutputPath $outputPath `
        -TextValue $SeedText
}

$jsonPath = Join-Path $resolvedOutputRoot "probe.json"
$json = $result | ConvertTo-Json -Depth 8
$json | Set-Content -Encoding UTF8 $jsonPath
$result
