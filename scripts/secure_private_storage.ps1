param([Parameter(Mandatory = $true)][string]$PrivateRoot)
$ErrorActionPreference = 'Stop'
$TaskPrivateRoot = (Resolve-Path -LiteralPath $PrivateRoot).ProviderPath
if ($TaskPrivateRoot.TrimEnd('\') -eq [IO.Path]::GetPathRoot($TaskPrivateRoot).TrimEnd('\')) {
    throw 'A drive root cannot be used for private document storage.'
}
$TaskRootItem = Get-Item -LiteralPath $TaskPrivateRoot -Force
if (-not $TaskRootItem.PSIsContainer -or ($TaskRootItem.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
    throw 'Private storage must be a regular directory, not a junction or symbolic link.'
}
$TaskChildren = @(Get-ChildItem -LiteralPath $TaskPrivateRoot -Force -Recurse)
foreach ($TaskItem in $TaskChildren) {
    if (-not $TaskItem.FullName.StartsWith($TaskPrivateRoot.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase) -or
        ($TaskItem.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
        throw 'A linked path was found inside private document storage; no permissions were changed.'
    }
}
$TaskIdentity = [Security.Principal.WindowsIdentity]::GetCurrent().User
$TaskAcl = New-Object Security.AccessControl.DirectorySecurity
$TaskAcl.SetOwner($TaskIdentity)
$TaskAcl.SetAccessRuleProtection($true, $false)
foreach ($TaskSid in @($TaskIdentity.Value, 'S-1-5-18', 'S-1-5-32-544')) {
    $TaskPrincipal = New-Object Security.Principal.SecurityIdentifier($TaskSid)
    $TaskRule = New-Object Security.AccessControl.FileSystemAccessRule($TaskPrincipal, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
    $TaskAcl.AddAccessRule($TaskRule)
}
Set-Acl -LiteralPath $TaskPrivateRoot -AclObject $TaskAcl
# Reset descendants only, after the protected parent exists. This removes old
# explicit grants without ever reopening the root's parent permissions.
foreach ($TaskItem in ($TaskChildren | Sort-Object { $_.FullName.Length })) {
    & icacls.exe $TaskItem.FullName /reset /Q | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'Could not restrict one or more private storage items.' }
}
Write-Output 'Private storage is restricted to this service account, SYSTEM and Administrators.'
