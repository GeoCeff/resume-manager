param([switch]$Staged, [switch]$History)
$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)
$Allowed = @('.gitignore','.gitattributes','README.md','Resume Manager.cmd','requirements.txt','settings.example.json',
    'Scripts/Resume Manager.ps1','Scripts/resume_settings.ps1','Scripts/resume_update_workflow.ps1',
    'Scripts/update_resume_projects.py','Scripts/check_repository.ps1','Scripts/test_resume_program.py',
    '.githooks/pre-commit','.githooks/pre-push')
function Check-Tree([string]$Revision) {
    $Names = if ($Revision -eq '') { @(git ls-files) } else { @(git ls-tree -r --name-only $Revision) }
    if ($LASTEXITCODE -ne 0) { throw 'Cannot list Git files.' }
    foreach ($Name in $Names) {
        if ($Allowed -cnotcontains $Name) { throw "Blocked non-program file: $Name" }
        $Content = (git show "${Revision}:$Name") -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Cannot inspect Git blob: $Name" }
        if ($Content -match '(?i)(?<![a-z0-9])[a-z]:[\\/]') { throw "Blocked machine-specific path in $Name" }
        $Emails = [regex]::Matches($Content, '(?i)[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}')
        foreach ($Email in $Emails) {
            if ($Email.Value -notmatch '(?i)@example\.(com|org|net)$') { throw "Blocked non-example email in $Name" }
        }
        foreach ($Number in [regex]::Matches($Content, '(?<![-\d])\+?\d[\d .()-]{5,}\d(?!\d)')) {
            $Digits = $Number.Value -replace '\D',''
            if ($Digits.Length -ge 7 -and $Digits -notmatch '^0+$') { throw "Blocked possible personal phone number in $Name" }
        }
    }
}
try {
    if ($Staged) { Check-Tree '' }
    elseif ($History) {
        $Commits = @(git rev-list --all)
        if ($LASTEXITCODE -ne 0) { throw 'Cannot inspect Git history.' }
        foreach ($Commit in $Commits) { Check-Tree $Commit }
    }
    else { Check-Tree 'HEAD' }
    Write-Output 'Program-only privacy check passed.'
}
catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1 }
