param(
    [Parameter(Mandatory = $true)]
    [string]$ChunkFile,
    [Parameter(Mandatory = $true)]
    [string]$RepoRoot
)

$ErrorActionPreference = 'Stop'

function Get-RelativeDocPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FullPath
    )

    $absolute = [System.IO.Path]::GetFullPath($FullPath)
    $root = [System.IO.Path]::GetPathRoot($absolute)
    $driveName = $root.TrimEnd('\').TrimEnd(':')
    $relative = $absolute.Substring($root.Length).TrimStart('\')

    if ([string]::IsNullOrWhiteSpace($relative)) {
        return $driveName
    }

    return Join-Path $driveName $relative
}

function Ensure-ParentDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $parent = Split-Path -Parent $Path
    if ($parent) {
        New-Item -ItemType Directory -Force -Path $parent | Out-Null
    }
}

$chunkName = [System.IO.Path]::GetFileNameWithoutExtension($ChunkFile)
$javaExe = Join-Path $RepoRoot 'tools\jdk-21.0.10+7\bin\java.exe'
$converterJar = Join-Path $RepoRoot 'tools\hwp-batch-converter.jar'
$manifestDir = Join-Path $RepoRoot 'manifests'
$copyRoot = Join-Path $RepoRoot 'collected_files'
$copyHwpRoot = Join-Path $copyRoot 'hwp'
$copyHwpxRoot = Join-Path $copyRoot 'hwpx'
$convertedRoot = Join-Path $RepoRoot 'hwpx_files'
$convertManifest = Join-Path $manifestDir ("convert_{0}.tsv" -f $chunkName)
$convertLog = Join-Path $manifestDir ("convert_{0}.log.tsv" -f $chunkName)
$copyLog = Join-Path $manifestDir ("copy_{0}.log.tsv" -f $chunkName)
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

New-Item -ItemType Directory -Force -Path $manifestDir | Out-Null
New-Item -ItemType Directory -Force -Path $copyHwpRoot | Out-Null
New-Item -ItemType Directory -Force -Path $copyHwpxRoot | Out-Null
New-Item -ItemType Directory -Force -Path $convertedRoot | Out-Null

$copyWriter = [System.IO.StreamWriter]::new($copyLog, $false, $utf8NoBom)
$convertWriter = [System.IO.StreamWriter]::new($convertManifest, $false, $utf8NoBom)
$hwpCount = 0
$hwpxCount = 0
$copyErrors = 0

try {
    $copyWriter.WriteLine("source`ttarget`tstatus`tmessage")

    foreach ($source in Get-Content -LiteralPath $ChunkFile -Encoding UTF8) {
        if ([string]::IsNullOrWhiteSpace($source)) {
            continue
        }

        $relative = Get-RelativeDocPath -FullPath $source
        $extension = [System.IO.Path]::GetExtension($source)

        if ($extension -ieq '.hwp') {
            $copyTarget = Join-Path $copyHwpRoot $relative
            $convertTarget = Join-Path $convertedRoot ([System.IO.Path]::ChangeExtension($relative, '.hwpx'))
            $hwpCount++
        } elseif ($extension -ieq '.hwpx') {
            $copyTarget = Join-Path $copyHwpxRoot $relative
            $convertTarget = $null
            $hwpxCount++
        } else {
            continue
        }

        try {
            Ensure-ParentDirectory -Path $copyTarget
            Copy-Item -LiteralPath $source -Destination $copyTarget -Force
            $copyWriter.WriteLine(("{0}`t{1}`tOK`t" -f $source, $copyTarget))
        } catch {
            $copyErrors++
            $message = $_.Exception.Message.Replace("`t", ' ').Replace("`r", ' ').Replace("`n", ' ')
            $copyWriter.WriteLine(("{0}`t{1}`tERROR`t{2}" -f $source, $copyTarget, $message))
            continue
        }

        if ($convertTarget) {
            Ensure-ParentDirectory -Path $convertTarget
            $convertWriter.WriteLine(("{0}`t{1}" -f $source, $convertTarget))
        }
    }
} finally {
    $copyWriter.Dispose()
    $convertWriter.Dispose()
}

$convertFailures = 0
if ((Get-Item -LiteralPath $convertManifest).Length -gt 0) {
    & $javaExe -jar $converterJar --manifest $convertManifest --log $convertLog
    if (Test-Path -LiteralPath $convertLog) {
        $convertFailures = @(
            Get-Content -LiteralPath $convertLog -Encoding UTF8 |
                Select-Object -Skip 1 |
                Where-Object { $_ -match "`tERROR`t" }
        ).Count
    }
}

[pscustomobject]@{
    Chunk = $chunkName
    HwpCount = $hwpCount
    HwpxCount = $hwpxCount
    CopyErrors = $copyErrors
    ConvertFailures = $convertFailures
}
