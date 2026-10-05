param([switch]$Staged, [switch]$History, [switch]$WorkingTree, [switch]$ListFiles, [string]$ZipPath)
$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)
$Allowed = @('.gitignore','.gitattributes','README.md','Resume Manager.cmd','requirements.txt','settings.example.json',
    'Scripts/Resume Manager.ps1','Scripts/resume_settings.ps1','Scripts/resume_update_workflow.ps1',
    'Scripts/update_resume_projects.py','Scripts/check_repository.ps1','Scripts/test_resume_program.py',
    '.githooks/pre-commit','.githooks/pre-push')
$Allowed += @('Scripts/manager_dialogs.ps1','Scripts/package_program.ps1','LICENSE')
if($ListFiles){$Allowed;return}
function Check-Content([string]$Name,[string]$Content) {
    if ($Allowed -cnotcontains $Name) { throw "Blocked non-program file: $Name" }
    if ($Content -match '(?i)(?<![a-z0-9])[a-z]:[\\/]') { throw "Blocked machine-specific path in $Name" }
    foreach ($Email in [regex]::Matches($Content, '(?i)[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}')) {
        if ($Email.Value -notmatch '(?i)@example\.(com|org|net)$') { throw "Blocked non-example email in $Name" }
    }
    foreach ($Number in [regex]::Matches($Content, '(?<![-\d])\+?\d[\d .()-]{5,}\d(?!\d)')) {
        $Digits = $Number.Value -replace '\D',''
        if ($Digits.Length -ge 7 -and $Digits -notmatch '^0+$') { throw "Blocked possible personal phone number in $Name" }
    }
}
function Check-Tree([string]$Revision) {
    $Names = if($Revision -eq 'WORKTREE'){@($Allowed | Where-Object {Test-Path -LiteralPath $_ -PathType Leaf})} elseif ($Revision -eq '') { @(git ls-files) } else { @(git ls-tree -r --name-only $Revision) }
    if ($Revision -ne 'WORKTREE' -and $LASTEXITCODE -ne 0) { throw 'Cannot list Git files.' }
    foreach ($Name in $Names) {
        if($Revision -eq 'WORKTREE'){$Content=Get-Content -LiteralPath $Name -Raw}
        else{$Content = (git show "${Revision}:$Name") -join "`n"; if ($LASTEXITCODE -ne 0) { throw "Cannot inspect Git blob: $Name" }}
        Check-Content $Name $Content
    }
}
try {
    if($ZipPath){
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $Archive=[IO.Compression.ZipFile]::OpenRead([IO.Path]::GetFullPath($ZipPath))
        try{
            $Names=@($Archive.Entries | ForEach-Object {$_.FullName})
            if($Names.Count -ne $Allowed.Count -or @(Compare-Object $Allowed $Names).Count){throw 'ZIP does not match the exact program allowlist.'}
            foreach($Entry in $Archive.Entries){
                $Reader=New-Object IO.StreamReader($Entry.Open())
                try{Check-Content $Entry.FullName $Reader.ReadToEnd()}finally{$Reader.Dispose()}
            }
        }finally{$Archive.Dispose()}
    }
    elseif($WorkingTree){
        Check-Tree 'WORKTREE'
    }
    elseif ($Staged) { Check-Tree '' }
    elseif ($History) {
        $Commits = @(git rev-list --all)
        if ($LASTEXITCODE -ne 0) { throw 'Cannot inspect Git history.' }
        foreach ($Commit in $Commits) { Check-Tree $Commit }
    }
    else { Check-Tree 'HEAD' }
    $global:LASTEXITCODE = 0
    Write-Output 'Program-only privacy check passed.'
}
catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1 }
