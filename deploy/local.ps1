param([ValidateSet('start','stop','status','logs')][string]$Action = 'start')
$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
if (-not (Test-Path -LiteralPath 'compose.local.yaml') -and (Test-Path -LiteralPath '../compose.local.yaml')) { Set-Location -LiteralPath (Split-Path -Parent $PSScriptRoot) }
if (-not (Get-Command docker -ErrorAction SilentlyContinue)) { throw 'Docker Engine 28+ with Compose v2 is required. Start Docker first.' }
& docker compose version
if ($LASTEXITCODE -ne 0) { throw 'Docker Compose v2 is required.' }
$DockerVersion = & docker version --format '{{.Server.Version}}'
if ($LASTEXITCODE -ne 0 -or $DockerVersion -notmatch '^(\d+)\.' -or [int]$Matches[1] -lt 28) { throw 'Docker Engine 28 or newer must be running.' }
$LocalPort = if ($env:JAKAL_LOCAL_PORT) { $env:JAKAL_LOCAL_PORT } else { '18988' }
if ($LocalPort -notmatch '^\d{4,5}$' -or [int]$LocalPort -lt 1024 -or [int]$LocalPort -gt 65535) { throw 'JAKAL_LOCAL_PORT must be from 1024 to 65535.' }
$LocalProject = if ($env:JAKAL_LOCAL_PROJECT) { $env:JAKAL_LOCAL_PROJECT } else { 'jakal-local' }
if ($LocalProject -notmatch '^[a-z0-9][a-z0-9_-]{0,59}$') { throw 'Invalid local project name.' }
$ComposeArguments = @('compose', '--project-name', $LocalProject, '--file', 'compose.local.yaml')
& docker @ComposeArguments config --quiet
if ($LASTEXITCODE -ne 0) { throw 'Local configuration is invalid.' }
switch ($Action) {
    'start' { & docker @ComposeArguments up -d --build --wait --wait-timeout 240 }
    'stop' { & docker @ComposeArguments stop }
    'status' { & docker @ComposeArguments ps }
    'logs' { & docker @ComposeArguments logs --tail 100 -f }
}
if ($LASTEXITCODE -ne 0) { throw 'Docker did not complete the local operation. See guide.html.' }
if ($Action -eq 'start') { Write-Host "Open http://127.0.0.1:$LocalPort/ in your browser." }
