param(
    [string[]]$DriveRoots,
    [int]$WorkerCount = [Math]::Min([Math]::Max([Environment]::ProcessorCount, 2), 6)
)

$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$manifestDir = Join-Path $repoRoot 'manifests'
$scanDir = Join-Path $manifestDir 'scan'
$chunkDir = Join-Path $manifestDir 'chunks'
$buildScript = Join-Path $repoRoot 'scripts\build-hwp-converter.ps1'
$workerScript = Join-Path $repoRoot 'scripts\process-hwp-chunk.ps1'
$copyRoot = Join-Path $repoRoot 'collected_files'
$convertedRoot = Join-Path $repoRoot 'hwpx_files'
$allFilesManifest = Join-Path $manifestDir 'all_files.txt'
$summaryPath = Join-Path $manifestDir 'summary.json'
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

if (-not $DriveRoots -or $DriveRoots.Count -eq 0) {
    $DriveRoots = Get-PSDrive -PSProvider FileSystem | Select-Object -ExpandProperty Root
}

New-Item -ItemType Directory -Force -Path $manifestDir | Out-Null
New-Item -ItemType Directory -Force -Path $scanDir | Out-Null
New-Item -ItemType Directory -Force -Path $chunkDir | Out-Null
New-Item -ItemType Directory -Force -Path $copyRoot | Out-Null
New-Item -ItemType Directory -Force -Path $convertedRoot | Out-Null

& $buildScript -RepoRoot $repoRoot | Out-Null

$scanJobScript = {
    param(
        [string]$DriveRoot,
        [string]$ExcludeRoot,
        [string]$OutputPath
    )

    $utf8NoBomInner = New-Object System.Text.UTF8Encoding($false)
    $writer = [System.IO.StreamWriter]::new($OutputPath, $false, $utf8NoBomInner)
    $count = 0
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($DriveRoot)

    try {
        while ($stack.Count -gt 0) {
            $current = $stack.Pop()

            if ($ExcludeRoot -and $current.StartsWith($ExcludeRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
                continue
            }

            try {
                foreach ($file in [System.IO.Directory]::EnumerateFiles($current)) {
                    $extension = [System.IO.Path]::GetExtension($file)
                    if ($extension -ieq '.hwp' -or $extension -ieq '.hwpx') {
                        $writer.WriteLine($file)
                        $count++
                    }
                }
            } catch {
            }

            try {
                foreach ($directory in [System.IO.Directory]::EnumerateDirectories($current)) {
                    if ($ExcludeRoot -and $directory.StartsWith($ExcludeRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
                        continue
                    }

                    try {
                        $attributes = [System.IO.File]::GetAttributes($directory)
                        if (($attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                            continue
                        }
                    } catch {
                        continue
                    }

                    $stack.Push($directory)
                }
            } catch {
            }
        }
    } finally {
        $writer.Dispose()
    }

    [pscustomobject]@{
        DriveRoot = $DriveRoot
        Count = $count
        OutputPath = $OutputPath
    }
}

$scanJobs = @()
foreach ($driveRoot in $DriveRoots) {
    $driveName = $driveRoot.TrimEnd('\').TrimEnd(':')
    $scanPath = Join-Path $scanDir ("{0}.txt" -f $driveName)
    $excludeRoot = ''

    if ($repoRoot.StartsWith($driveRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        $excludeRoot = $repoRoot
    }

    $scanJobs += Start-Job -ScriptBlock $scanJobScript -ArgumentList $driveRoot, $excludeRoot, $scanPath
}

$null = Wait-Job -Job $scanJobs
$scanResults = $scanJobs | Receive-Job
$scanJobs | Remove-Job

$allFiles = New-Object System.Collections.Generic.List[string]
foreach ($scanResult in $scanResults) {
    if (Test-Path -LiteralPath $scanResult.OutputPath) {
        foreach ($line in Get-Content -LiteralPath $scanResult.OutputPath -Encoding UTF8) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                $allFiles.Add($line)
            }
        }
    }
}

[System.IO.File]::WriteAllLines($allFilesManifest, $allFiles, $utf8NoBom)

if ($allFiles.Count -eq 0) {
    $summary = [pscustomobject]@{
        scannedDrives = $DriveRoots
        totalFound = 0
        copiedHwp = 0
        copiedHwpx = 0
        copyErrors = 0
        convertFailureChunks = 0
    }
    $summary | ConvertTo-Json | Set-Content -LiteralPath $summaryPath -Encoding UTF8
    $summary
    return
}

$effectiveWorkers = [Math]::Min($WorkerCount, $allFiles.Count)
$chunkFiles = @()
for ($index = 0; $index -lt $effectiveWorkers; $index++) {
    $chunkPath = Join-Path $chunkDir ("chunk_{0:D2}.txt" -f $index)
    if (Test-Path $chunkPath) {
        Remove-Item -Force $chunkPath
    }
    $chunkFiles += $chunkPath
}

for ($index = 0; $index -lt $allFiles.Count; $index++) {
    $chunkIndex = $index % $effectiveWorkers
    [System.IO.File]::AppendAllText(
        $chunkFiles[$chunkIndex],
        ($allFiles[$index] + [Environment]::NewLine),
        $utf8NoBom
    )
}

$workerJobs = @()
foreach ($chunkFile in $chunkFiles) {
    $workerJobs += Start-Job -FilePath $workerScript -ArgumentList $chunkFile, $repoRoot
}

$null = Wait-Job -Job $workerJobs
$workerResults = $workerJobs | Receive-Job
$workerJobs | Remove-Job

$summary = [pscustomobject]@{
    scannedDrives = $DriveRoots
    totalFound = $allFiles.Count
    copiedHwp = ($workerResults | Measure-Object -Property HwpCount -Sum).Sum
    copiedHwpx = ($workerResults | Measure-Object -Property HwpxCount -Sum).Sum
    copyErrors = ($workerResults | Measure-Object -Property CopyErrors -Sum).Sum
    convertFailureChunks = ($workerResults | Measure-Object -Property ConvertFailures -Sum).Sum
}

$summary | ConvertTo-Json | Set-Content -LiteralPath $summaryPath -Encoding UTF8
$summary
