param([string]$ConfigPath = $env:RESUME_MANAGER_CONFIG)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $env:LOCALAPPDATA 'ResumeManager\settings.local.json'
}
if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    throw "Local settings not found: $ConfigPath. Follow README.md to create settings outside the program repository."
}
$ConfigPath = (Resolve-Path -LiteralPath $ConfigPath).Path
$Settings = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$CodeRoot = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot)).TrimEnd('\') + '\'
if ($ConfigPath.StartsWith($CodeRoot, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Personal settings must be stored outside the program repository.'
}
foreach ($Key in @('DataRoot','PrivateContact','PublicContact','ContactAnchor')) {
    if ([string]::IsNullOrWhiteSpace($Settings.$Key)) { throw "Missing local setting: $Key" }
}
if (-not $Settings.PrivateOnly -or $Settings.PrivateOnly -is [string]) { throw 'PrivateOnly must be a nonempty array of private contact values.' }
$Root = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($Settings.DataRoot))
if (($Root.TrimEnd('\') + '\').StartsWith($CodeRoot, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Resume data must be stored outside the program repository.'
}
$CurrentDir = Join-Path $Root 'Current'
$UpdatesDir = Join-Path $Root 'Updates'
$PrivateName = if ($Settings.ArchivePrivateName) { [string]$Settings.ArchivePrivateName } else { 'resume-private.docx' }
$PublicName = if ($Settings.ArchivePublicName) { [string]$Settings.ArchivePublicName } else { 'resume-public.docx' }
$CurrentPrivateName = if ($Settings.CurrentPrivateName) { [string]$Settings.CurrentPrivateName } else { 'resume-private.docx' }
$CurrentPublicName = if ($Settings.CurrentPublicName) { [string]$Settings.CurrentPublicName } else { 'resume-public.docx' }
foreach ($Name in @($PrivateName,$PublicName,$CurrentPrivateName,$CurrentPublicName)) {
    if ([IO.Path]::GetFileName($Name) -ne $Name -or $Name -notmatch '(?i)\.docx$') { throw 'Resume filenames must be plain DOCX filenames.' }
}
if ($PrivateName -eq $PublicName -or $CurrentPrivateName -eq $CurrentPublicName) { throw 'Private and public filenames must differ.' }
$CurrentPrivate = Join-Path $CurrentDir $CurrentPrivateName
$CurrentPublic = Join-Path $CurrentDir $CurrentPublicName
$WebsitePublic = [Environment]::ExpandEnvironmentVariables([string]$Settings.WebsitePublicPath)
if ($WebsitePublic -and ([IO.Path]::GetFullPath($WebsitePublic)).StartsWith($CodeRoot, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Website document destination must be outside the program repository.'
}
$PythonExe = if ($Settings.PythonExe) { [Environment]::ExpandEnvironmentVariables([string]$Settings.PythonExe) } else { (Get-Command python.exe -ErrorAction Stop).Source }
$VariantScript = Join-Path $PSScriptRoot 'update_resume_projects.py'
