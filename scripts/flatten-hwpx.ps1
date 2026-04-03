param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path,
    [string]$TargetFolderName = 'all_hwpx_flat'
)

$ErrorActionPreference = 'Stop'

$target = Join-Path $RepoRoot $TargetFolderName
$manifestDir = Join-Path $RepoRoot 'manifests'
$mapPath = Join-Path $manifestDir 'all_hwpx_flat_map.tsv'
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

New-Item -ItemType Directory -Force -Path $manifestDir | Out-Null
New-Item -ItemType Directory -Force -Path $target | Out-Null

Get-ChildItem -LiteralPath $target -File -ErrorAction SilentlyContinue | Remove-Item -Force

$files = Get-ChildItem -LiteralPath $RepoRoot -Recurse -File |
    Where-Object {
        $_.Extension -ieq '.hwpx' -and
        $_.FullName -notlike "$RepoRoot\.git\*" -and
        $_.FullName -notlike "$target\*"
    } |
    Sort-Object FullName

$sha1 = [System.Security.Cryptography.SHA1]::Create()
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("source`tflattened_name`ttarget")

foreach ($file in $files) {
    $name = $file.Name
    $destination = Join-Path $target $name

    if (Test-Path -LiteralPath $destination) {
        $hashBytes = $sha1.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($file.FullName))
        $hash = [System.BitConverter]::ToString($hashBytes).Replace('-', '').Substring(0, 10).ToLowerInvariant()
        $base = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $name = '{0}__{1}{2}' -f $base, $hash, $file.Extension
        $destination = Join-Path $target $name
    }

    Copy-Item -LiteralPath $file.FullName -Destination $destination -Force
    $lines.Add(('{0}`t{1}`t{2}' -f $file.FullName, $name, $destination))
}

$sha1.Dispose()
[System.IO.File]::WriteAllLines($mapPath, $lines, $utf8NoBom)

[pscustomobject]@{
    Target = $target
    TotalCopied = $files.Count
    FlattenedCount = (Get-ChildItem -LiteralPath $target -File | Measure-Object).Count
    Mapping = $mapPath
}
