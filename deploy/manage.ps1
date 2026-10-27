param(
    [ValidateSet('start','stop','status','logs','update')][string]$Action = 'start',
    [string]$Domain
)
$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
if (-not (Get-Command docker -ErrorAction SilentlyContinue)) { throw 'Docker with Compose v2 is required.' }
& docker compose version
if ($LASTEXITCODE -ne 0) { throw 'Docker Compose v2 is required.' }
if (-not (Test-Path -LiteralPath '.env.production')) {
    if ($Action -ne 'start') { throw 'Run manage.ps1 start first.' }
    if (-not $Domain) { $Domain = Read-Host 'Public domain connected to this server (example: editor.example.com)' }
    $Domain = $Domain.ToLowerInvariant()
    if ($Domain.Length -gt 253 -or $Domain -notmatch '^[a-z0-9]([a-z0-9-]{0,61}[a-z0-9])?(\.[a-z0-9]([a-z0-9-]{0,61}[a-z0-9])?)+$') { throw 'Enter a hostname without https://, a port, or a path.' }
    $Config = "JAKAL_DOMAIN=$Domain`nJAKAL_MEMORY_LIMIT=3g`nJAKAL_PROXY_SUBNET=172.30.88.0/29`nJAKAL_PROXY_ADDRESS=172.30.88.2`n"
    $Stream = [IO.File]::Open((Join-Path $PSScriptRoot '.env.production'), [IO.FileMode]::CreateNew)
    try { $Bytes = [Text.Encoding]::UTF8.GetBytes($Config); $Stream.Write($Bytes, 0, $Bytes.Length) } finally { $Stream.Dispose() }
}
& docker compose --env-file .env.production config --quiet
if ($LASTEXITCODE -ne 0) { throw 'Deployment configuration is invalid.' }
switch ($Action) {
    { $_ -in 'start','update' } { & docker compose --env-file .env.production up -d --build --wait --wait-timeout 180 }
    'stop' { & docker compose --env-file .env.production stop }
    'status' { & docker compose --env-file .env.production ps }
    'logs' { & docker compose --env-file .env.production logs --tail 100 -f }
}
if ($LASTEXITCODE -ne 0) { throw 'Docker did not complete the requested operation.' }
