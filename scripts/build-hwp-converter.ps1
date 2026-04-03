param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
)

$ErrorActionPreference = 'Stop'

$javaHome = Get-ChildItem (Join-Path $RepoRoot 'tools') -Directory |
    Where-Object { $_.Name -like 'jdk-*' } |
    Select-Object -First 1 -ExpandProperty FullName

if (-not $javaHome) {
    throw 'JDK was not found under tools\.'
}

$javac = Join-Path $javaHome 'bin\javac.exe'
$jarExe = Join-Path $javaHome 'bin\jar.exe'
$buildRoot = Join-Path $RepoRoot 'tools\build'
$classesRoot = Join-Path $buildRoot 'classes'
$manifestPath = Join-Path $buildRoot 'manifest.mf'
$sourceListPath = Join-Path $buildRoot 'sources.txt'
$outputJar = Join-Path $RepoRoot 'tools\hwp-batch-converter.jar'

New-Item -ItemType Directory -Force -Path $buildRoot | Out-Null
if (Test-Path $classesRoot) {
    Remove-Item -Recurse -Force $classesRoot
}
New-Item -ItemType Directory -Force -Path $classesRoot | Out-Null

$classpath = @(
    (Join-Path $RepoRoot 'tools\lib\hwplib-1.1.10.jar')
    (Join-Path $RepoRoot 'tools\lib\hwpxlib-1.0.8.jar')
) -join ';'

$sources = @(
    Get-ChildItem (Join-Path $RepoRoot 'vendor\hwp2hwpx-src\src\main\java') -Recurse -Filter *.java |
        Select-Object -ExpandProperty FullName
    Get-ChildItem (Join-Path $RepoRoot 'tools\java-src') -Recurse -Filter *.java |
        Select-Object -ExpandProperty FullName
)

if ($sources.Count -eq 0) {
    throw 'No Java sources were found.'
}

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllLines($sourceListPath, $sources, $utf8NoBom)

& $javac -encoding UTF-8 -cp $classpath -d $classesRoot "@$sourceListPath"
if ($LASTEXITCODE -ne 0) {
    throw "javac failed with exit code $LASTEXITCODE"
}

$manifestLines = @(
    'Main-Class: local.jakaldocs.HwpBatchConverterMain'
    'Class-Path: lib/hwplib-1.1.10.jar lib/hwpxlib-1.0.8.jar'
    ''
)
[System.IO.File]::WriteAllLines($manifestPath, $manifestLines, $utf8NoBom)

if (Test-Path $outputJar) {
    Remove-Item -Force $outputJar
}

& $jarExe cfm $outputJar $manifestPath -C $classesRoot .
if ($LASTEXITCODE -ne 0) {
    throw "jar failed with exit code $LASTEXITCODE"
}

Write-Output $outputJar
