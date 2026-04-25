param(
    [string]$OutputDir = ".codex-temp/native_memo_hwpx_probe",
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

function Add-Memo {
    param(
        [object]$Hwp,
        [string]$Text
    )

    $action = $Hwp.CreateAction("Comment")
    if ($null -eq $action) {
        throw "CreateAction(Comment) returned null."
    }

    $memoShape = $Hwp.HParameterSet.HMemoShape
    $getDefault = [bool]$action.GetDefault($memoShape.HSet)
    $execute = [bool]$action.Execute($memoShape.HSet)

    $insertResult = $null
    if (-not [string]::IsNullOrEmpty($Text)) {
        $insertText = $Hwp.CreateAction("InsertText")
        if ($null -eq $insertText) {
            throw "CreateAction(InsertText) returned null."
        }
        $insertSet = $Hwp.HParameterSet.HInsertText
        $insertText.GetDefault($insertSet.HSet) | Out-Null
        $insertSet.Text = $Text
        $insertResult = [bool]$insertText.Execute($insertSet.HSet)
    }

    return [ordered]@{
        get_default = $getDefault
        execute = $execute
        insert_text = $insertResult
    }
}

function Get-HwpmlText {
    param([object]$Hwp)

    $xml = $Hwp.GetTextFile("HWPML2X", "")
    if ([string]::IsNullOrWhiteSpace($xml)) {
        return ""
    }
    return [string]$xml
}

function Expand-HwpxPackage {
    param(
        [string]$HwpxPath,
        [string]$Destination
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    if (Test-Path $Destination) {
        Remove-Item -LiteralPath $Destination -Recurse -Force
    }
    [System.IO.Compression.ZipFile]::ExtractToDirectory($HwpxPath, $Destination)
}

function Get-XmlText {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        return ""
    }
    return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
}

function Get-PackageFiles {
    param([string]$Root)
    return @(
        Get-ChildItem -LiteralPath $Root -Recurse -File |
            ForEach-Object {
                $_.FullName.Substring($Root.Length).TrimStart('\').Replace('\', '/')
            } |
            Sort-Object
    )
}

function Get-MemoReport {
    param(
        [string]$HwpxPath,
        [string]$ExpandedRoot
    )

    $headerPath = Join-Path $ExpandedRoot "Contents\header.xml"
    $sectionPath = Join-Path $ExpandedRoot "Contents\section0.xml"
    $contentPath = Join-Path $ExpandedRoot "Contents\content.hpf"

    $headerXml = Get-XmlText -Path $headerPath
    $sectionXml = Get-XmlText -Path $sectionPath
    $contentXml = Get-XmlText -Path $contentPath

    $sectionDoc = New-Object System.Xml.XmlDocument
    $sectionDoc.LoadXml($sectionXml)
    $nsmgr = New-Object System.Xml.XmlNamespaceManager($sectionDoc.NameTable)
    $nsmgr.AddNamespace("hp", "http://www.hancom.co.kr/hwpml/2011/paragraph")

    $hiddenComment = $sectionDoc.SelectSingleNode("//hp:hiddenComment", $nsmgr)
    $secPr = $sectionDoc.SelectSingleNode("//hp:secPr", $nsmgr)
    $memoText = ""
    if ($hiddenComment -ne $null) {
        $textNodes = $hiddenComment.SelectNodes(".//hp:t", $nsmgr)
        if ($textNodes -ne $null) {
            $memoText = (($textNodes | ForEach-Object { $_.InnerText }) -join "")
        }
    }

    return [ordered]@{
        hwpx_path = $HwpxPath
        files = @(Get-PackageFiles -Root $ExpandedRoot)
        header_has_memo_properties = $headerXml.Contains("<hh:memoProperties")
        header_has_track_change_authors = $headerXml.Contains("trackChangeAuthor")
        header_has_memo_style = $headerXml.Contains('engName="Memo"')
        content_has_comments_part = $contentXml.Contains("comments.xml")
        content_has_memo_extended_part = $contentXml.Contains("memoExtended.xml")
        section_has_hidden_comment = ($hiddenComment -ne $null)
        section_memo_shape_id = if ($secPr -ne $null) { $secPr.Attributes["memoShapeIDRef"].Value } else { $null }
        section_memo_para_style_id = if ($hiddenComment -ne $null) {
            $node = $hiddenComment.SelectSingleNode(".//hp:subList/hp:p", $nsmgr)
            if ($node -ne $null) { $node.Attributes["styleIDRef"].Value } else { $null }
        } else { $null }
        section_memo_char_pr_id = if ($hiddenComment -ne $null) {
            $node = $hiddenComment.SelectSingleNode(".//hp:subList/hp:p/hp:run", $nsmgr)
            if ($node -ne $null) { $node.Attributes["charPrIDRef"].Value } else { $null }
        } else { $null }
        section_memo_text = $memoText
    }
}

function Get-HwpmlMemoReport {
    param([string]$Xml)

    if ([string]::IsNullOrWhiteSpace($Xml)) {
        return [ordered]@{
            has_hidden_comment = $false
            has_memo_properties = $false
            has_comments_part = $false
            has_memo_extended_part = $false
            has_author = $false
            has_memo_id = $false
            has_anchor_id = $false
            has_visible = $false
            has_order = $false
            memo_text = ""
        }
    }

    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($Xml)
    $nsmgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsmgr.AddNamespace("h", "http://www.hancom.co.kr/hwpml/2008/7.0")

    $hiddenComment = $doc.SelectSingleNode("//*[local-name()='HIDDENCOMMENT']", $nsmgr)
    $memoText = ""
    $hiddenCommentXml = ""
    if ($hiddenComment -ne $null) {
        $memoText = $hiddenComment.InnerText
        $hiddenCommentXml = $hiddenComment.OuterXml
    }

    return [ordered]@{
        has_hidden_comment = ($hiddenComment -ne $null)
        has_memo_properties = ($Xml -match "memoProperties")
        has_comments_part = ($Xml -match "comments\\.xml")
        has_memo_extended_part = ($Xml -match "memoExtended\\.xml")
        has_author = ($hiddenCommentXml -match "Author")
        has_memo_id = ($hiddenCommentXml -match "MemoId")
        has_anchor_id = ($hiddenCommentXml -match "AnchorId")
        has_visible = ($hiddenCommentXml -match "Visible")
        has_order = ($hiddenCommentXml -match "Order")
        memo_text = $memoText
    }
}

$setupScript = Join-Path $PSScriptRoot "setup_hancom_security_module.ps1"
$setupResult = & $setupScript -ModuleName $ModuleName -DownloadIfMissing
$registeredModuleName = $setupResult.ModuleName

$resolvedOutputDir = Join-Path (Get-RepoRoot) $OutputDir
New-Item -ItemType Directory -Force $resolvedOutputDir | Out-Null

$cases = @(
    [ordered]@{ name = "empty"; text = "" },
    [ordered]@{ name = "with_text"; text = "comment-body" }
)

$report = [ordered]@{}

foreach ($case in $cases) {
    $caseName = [string]$case.name
    $hwpxPath = Join-Path $resolvedOutputDir ("memo_{0}.hwpx" -f $caseName)
    $expandedDir = Join-Path $resolvedOutputDir ("{0}_unzipped" -f $caseName)

    $hwp = Get-HwpObject -RegisteredModuleName $registeredModuleName
    try {
        $hwp.Run("FileNew") | Out-Null
        $actionStatus = Add-Memo -Hwp $hwp -Text ([string]$case.text)
        $hwpml = Get-HwpmlText -Hwp $hwp
        $saveResult = [bool]$hwp.SaveAs($hwpxPath, "HWPX", "")
        Expand-HwpxPackage -HwpxPath $hwpxPath -Destination $expandedDir
        $report[$caseName] = [ordered]@{
            action = $actionStatus
            hwpml = Get-HwpmlMemoReport -Xml $hwpml
            save_result = $saveResult
            package = Get-MemoReport -HwpxPath $hwpxPath -ExpandedRoot $expandedDir
        }
    }
    catch {
        $report[$caseName] = [ordered]@{
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
