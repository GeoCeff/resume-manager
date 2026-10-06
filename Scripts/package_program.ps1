param([Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
$Root = Split-Path -Parent $PSScriptRoot
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
if(($OutputDirectory.TrimEnd('\')+'\').StartsWith(($Root.TrimEnd('\')+'\'),[StringComparison]::OrdinalIgnoreCase)){throw 'Distribution must be written outside the checkout.'}
Set-Location $Root
& (Join-Path $PSScriptRoot 'check_repository.ps1') -WorkingTree
if($LASTEXITCODE -ne 0){throw 'Source privacy audit failed.'}
& (Join-Path $PSScriptRoot 'check_repository.ps1') -History
if($LASTEXITCODE -ne 0){throw 'History privacy audit failed.'}
$Files = @(& (Join-Path $PSScriptRoot 'check_repository.ps1') -ListFiles)
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
[void][IO.Directory]::CreateDirectory($OutputDirectory)
$Version = [regex]::Match((Get-Content -LiteralPath (Join-Path $PSScriptRoot 'manager_dialogs.ps1') -Raw), "ProgramVersion = '([^']+)'").Groups[1].Value
if(-not $Version){throw 'Program version is missing.'}
$Target = Join-Path $OutputDirectory ('resume-manager-' + $Version + '-source-' + (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss') + '-' + [guid]::NewGuid().ToString('N').Substring(0,8) + '.zip')
$Stream = [IO.File]::Open($Target,'CreateNew','Write','None')
$Archive = New-Object IO.Compression.ZipArchive($Stream,[IO.Compression.ZipArchiveMode]::Create)
try{foreach($Name in $Files){if(-not (Test-Path -LiteralPath $Name -PathType Leaf)){throw "Missing allowlisted file: $Name"}; [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($Archive,(Join-Path $Root $Name),$Name)}}finally{$Archive.Dispose();$Stream.Dispose()}
$Archive=[IO.Compression.ZipFile]::OpenRead($Target)
try{
    $Names=@($Archive.Entries | ForEach-Object {$_.FullName})
    if(@(Compare-Object $Files $Names).Count){throw 'Distribution ZIP differs from the source allowlist.'}
    foreach($Entry in $Archive.Entries){if($Entry.Length -eq 0){throw "Empty distribution entry: $($Entry.FullName)"}}
}finally{$Archive.Dispose()}
& (Join-Path $PSScriptRoot 'check_repository.ps1') -ZipPath $Target
if($LASTEXITCODE -ne 0){throw 'Generated ZIP privacy audit failed. Do not distribute it.'}
Write-Output "Program-only ZIP verified: $Target"
Write-Output 'Not published. MIT license included; review the release checklist before public distribution.'
