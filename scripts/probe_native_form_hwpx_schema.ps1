param(
    [string]$OutputDir = ".codex-temp/native_form_hwpx_probe",
    [string[]]$Kinds = @("CHECKBUTTON", "RADIOBUTTON", "BUTTON", "EDIT", "COMBOBOX", "LISTBOX", "SCROLLBAR"),
    [string]$ModuleName = "FilePathCheckerModuleExample"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Get-RepoRoot {
    return Split-Path -Parent $PSScriptRoot
}

function Get-HwpObject {
    param([string]$RegisteredModuleName)

    $hwp = New-Object -ComObject HWPFrame.HwpObject
    $hwp.SetMessageBoxMode(0x00010000) | Out-Null
    $registered = $hwp.RegisterModule("FilePathCheckDLL", $RegisteredModuleName)
    if (-not $registered) {
        throw "RegisterModule(FilePathCheckDLL, $RegisteredModuleName) failed."
    }
    return $hwp
}

function Get-BlankHwpml {
    param([object]$Hwp)

    $Hwp.Run("FileNew") | Out-Null
    $xml = $Hwp.GetTextFile("HWPML2X", "")
    if ([string]::IsNullOrWhiteSpace($xml)) {
        throw "GetTextFile(HWPML2X) returned empty content."
    }
    return $xml
}

function Get-CommonShapeObject {
    param(
        [int]$InstId,
        [int]$Width = 2400,
        [int]$Height = 1200
    )

    return @"
<SHAPEOBJECT InstId="$InstId" ZOrder="0" NumberingType="None" TextWrap="TopAndBottom" TextFlow="BothSides" Lock="false">
  <SIZE Width="$Width" WidthRelTo="Absolute" Height="$Height" HeightRelTo="Absolute" Protect="false"/>
  <POSITION TreatAsChar="true" AffectLSpacing="false" FlowWithText="true" AllowOverlap="false" HoldAnchorAndSO="false" VertRelTo="Para" HorzRelTo="Column" VertAlign="Top" HorzAlign="Left" VertOffset="0" HorzOffset="0"/>
  <OUTSIDEMARGIN Left="0" Right="0" Top="0" Bottom="0"/>
</SHAPEOBJECT>
"@
}

function Get-CommonFormObject {
    param(
        [string]$Name,
        [string]$ExtraChildren = ""
    )

    $children = @"
<FORMCHARSHAPE CharShape="0" FollowContext="false" AutoSize="false" WordWrap="false"/>
$ExtraChildren
"@
    return @"
<FORMOBJECT Name="$Name" ForeColor="0" BackColor="4294967295" GroupName="" TabStop="true" TabOrder="1" Enabled="true" BorderType="0" DrawFrame="true" Printable="true">
$children
</FORMOBJECT>
"@
}

function Get-FormSnippets {
    $snippets = [ordered]@{}

    $snippets["CHECKBUTTON"] = @"
<CHECKBUTTON>
$(Get-CommonShapeObject -InstId 4101)
$(Get-CommonFormObject -Name "check1" -ExtraChildren '<BUTTONSET Caption="Check" Value="UNCHECKED" RadioGroupName="" TriState="false" BackStyle="Transparent"/>')
</CHECKBUTTON>
"@

    $snippets["RADIOBUTTON"] = @"
<RADIOBUTTON>
$(Get-CommonShapeObject -InstId 4102)
$(Get-CommonFormObject -Name "radio1" -ExtraChildren '<BUTTONSET Caption="Radio" Value="SELECTED" RadioGroupName="group1" TriState="false" BackStyle="Transparent"/>')
</RADIOBUTTON>
"@

    $snippets["BUTTON"] = @"
<BUTTON>
$(Get-CommonShapeObject -InstId 4103 -Width 3200 -Height 1400)
$(Get-CommonFormObject -Name "button1" -ExtraChildren '<BUTTONSET Caption="Run" Value="" RadioGroupName="" TriState="false" BackStyle="Transparent"/>')
</BUTTON>
"@

    $snippets["EDIT"] = @"
<EDIT MultiLine="false" PasswordChar="" MaxLength="32" ScrollBars="0" TabKeyBehavior="false" Number="false" ReadOnly="false" AlignText="Left">
$(Get-CommonShapeObject -InstId 4104 -Width 4200 -Height 1400)
$(Get-CommonFormObject -Name "edit1")
<EDITTEXT>Draft</EDITTEXT>
</EDIT>
"@

    $snippets["COMBOBOX"] = @"
<COMBOBOX ListBoxRows="4" ListBoxWidth="2400" Text="Pending" EditEnable="true">
$(Get-CommonShapeObject -InstId 4105 -Width 4200 -Height 1400)
$(Get-CommonFormObject -Name "combo1")
</COMBOBOX>
"@

    $snippets["LISTBOX"] = @"
<LISTBOX Text="Pending" ItemHeight="300" TopIndex="0">
$(Get-CommonShapeObject -InstId 4106 -Width 4200 -Height 1800)
$(Get-CommonFormObject -Name "list1")
</LISTBOX>
"@

    $snippets["SCROLLBAR"] = @"
<SCROLLBAR Delay="0" LargeChange="10" SmallChange="1" Min="0" Max="100" Page="10" Value="25" Type="Horizontal">
$(Get-CommonShapeObject -InstId 4107 -Width 4200 -Height 900)
$(Get-CommonFormObject -Name "scroll1")
</SCROLLBAR>
"@

    return $snippets
}

function Build-HwpmlWithForm {
    param(
        [string]$BlankHwpml,
        [string]$Snippet
    )

    $marker = "<CHAR/></TEXT>"
    if (-not $BlankHwpml.Contains($marker)) {
        throw "Blank HWPML marker was not found: $marker"
    }
    return $BlankHwpml.Replace($marker, "$Snippet<CHAR/></TEXT>")
}

function Get-SectionXmlFromHwpx {
    param([string]$HwpxPath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($HwpxPath)
    try {
        $entry = $archive.GetEntry("Contents/section0.xml")
        if ($null -eq $entry) {
            throw "Contents/section0.xml was not found in $HwpxPath"
        }
        $stream = $entry.Open()
        try {
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
            try {
                return $reader.ReadToEnd()
            }
            finally {
                $reader.Dispose()
            }
        }
        finally {
            $stream.Dispose()
        }
    }
    finally {
        $archive.Dispose()
    }
}

function Get-DetectedFormTags {
    param([string]$SectionXml)

    $matches = [regex]::Matches($SectionXml, "<hp:(checkBtn|radioBtn|btn|edit|comboBox|listBox|scrollBar)\b")
    $values = New-Object System.Collections.Generic.List[string]
    foreach ($match in $matches) {
        $values.Add($match.Groups[1].Value)
    }
    return $values | Select-Object -Unique
}

function Get-FormSchemaReport {
    param([string]$SectionXml)

    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($SectionXml)
    $nsmgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsmgr.AddNamespace("hp", "http://www.hancom.co.kr/hwpml/2011/paragraph")

    $formNodes = $doc.SelectNodes("//hp:checkBtn | //hp:radioBtn | //hp:btn | //hp:edit | //hp:comboBox | //hp:listBox | //hp:scrollBar", $nsmgr)
    $forms = @()
    foreach ($node in $formNodes) {
        $attributes = [ordered]@{}
        foreach ($attribute in $node.Attributes) {
            $attributes[$attribute.Name] = $attribute.Value
        }
        $childTags = @()
        foreach ($child in $node.ChildNodes) {
            if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                $childTags += $child.LocalName
            }
        }
        $outerXml = $node.OuterXml
        $forms += [ordered]@{
            tag = $node.LocalName
            attributes = $attributes
            child_tags = @($childTags | Select-Object -Unique)
            has_caption = ($outerXml -match "caption")
            has_placeholder_like_token = ($outerXml -match "placeholder|placeHolder|hint|cue|prompt|watermark|emptyText")
            command = if ($node.Attributes["command"] -ne $null) { $node.Attributes["command"].Value } else { $null }
        }
    }

    return [ordered]@{
        form_count = $formNodes.Count
        package_has_placeholder_like_token = ($SectionXml -match "placeholder|placeHolder|hint|cue|prompt|watermark|emptyText")
        forms = @($forms)
    }
}

$setupScript = Join-Path $PSScriptRoot "setup_hancom_security_module.ps1"
$setupResult = & $setupScript -ModuleName $ModuleName -DownloadIfMissing
$registeredModuleName = $setupResult.ModuleName

$resolvedOutputDir = Join-Path (Get-RepoRoot) $OutputDir
New-Item -ItemType Directory -Force $resolvedOutputDir | Out-Null

$snippets = Get-FormSnippets
$report = [ordered]@{}

foreach ($kind in $Kinds) {
    if (-not $snippets.Contains($kind)) {
        Write-Warning "Skipping unsupported kind: $kind"
        continue
    }

    $key = $kind.ToLowerInvariant()
    $hwpxPath = Join-Path $resolvedOutputDir "${key}_probe.hwpx"
    $sectionDumpPath = Join-Path $resolvedOutputDir "${key}_section0.xml"

    $hwp = Get-HwpObject -RegisteredModuleName $registeredModuleName
    try {
        $blank = Get-BlankHwpml -Hwp $hwp
        $patched = Build-HwpmlWithForm -BlankHwpml $blank -Snippet $snippets[$kind]
        $setResult = $hwp.SetTextFile($patched, "HWPML2X", "")
        $saveResult = $hwp.SaveAs($hwpxPath, "HWPX", "")
        $sectionXml = Get-SectionXmlFromHwpx -HwpxPath $hwpxPath
        [System.IO.File]::WriteAllText($sectionDumpPath, $sectionXml, [System.Text.UTF8Encoding]::new($false))

        $report[$kind] = [ordered]@{
            set_result = $setResult
            save_result = $saveResult
            detected_tags = @(Get-DetectedFormTags -SectionXml $sectionXml)
            schema = Get-FormSchemaReport -SectionXml $sectionXml
            hwpx_path = $hwpxPath
            section_path = $sectionDumpPath
        }
    }
    catch {
        $report[$kind] = [ordered]@{
            error = $_.Exception.Message
        }
    }
    finally {
        try { $hwp.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($hwp) | Out-Null } catch {}
    }
}

$reportPath = Join-Path $resolvedOutputDir "report.json"
$report | ConvertTo-Json -Depth 8 | Set-Content -Path $reportPath -Encoding UTF8
$report | ConvertTo-Json -Depth 8
