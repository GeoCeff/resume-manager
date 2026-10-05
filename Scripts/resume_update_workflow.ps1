param(
    [ValidatePattern('^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}$')]
    [string]$Timestamp = (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'),
    [string]$SourcePath,
    [string]$WebsitePublicPath,
    [string]$ConfigPath = $env:RESUME_MANAGER_CONFIG,
    [string]$ProgressPath,
    [switch]$SkipProjectRewrite,
    [switch]$NoPdf
)

$ErrorActionPreference = "Stop"

function Write-Phase([string]$Phase) {
    if (-not $ProgressPath) { return }
    $Target = [IO.Path]::GetFullPath($ProgressPath)
    $Code = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot)).TrimEnd('\') + '\'
    if ($Target.StartsWith($Code,[StringComparison]::OrdinalIgnoreCase)) { throw 'Progress files must be outside the repository.' }
    $Temporary = $Target + '.tmp'
    [IO.File]::WriteAllText($Temporary, (@{Phase=$Phase;Time=[datetime]::UtcNow.ToString('o')} | ConvertTo-Json))
    if (Test-Path -LiteralPath $Target) { [IO.File]::Replace($Temporary,$Target,[NullString]::Value) } else { [IO.File]::Move($Temporary,$Target) }
}

. (Join-Path $PSScriptRoot 'resume_settings.ps1') -ConfigPath $ConfigPath
if (-not $PSBoundParameters.ContainsKey('WebsitePublicPath')) { $WebsitePublicPath = $WebsitePublic }
if ($WebsitePublicPath -and ([IO.Path]::GetFullPath($WebsitePublicPath)).StartsWith($CodeRoot, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Website document destination must be outside the program repository.'
}

function Publish-ResumeFiles {
    param([array]$Files)
    $Prepared = @()
    $Applied = @()
    $OriginalError = $null
    try {
        foreach ($File in $Files) {
            $Dir = Split-Path -Parent $File.Destination
            if (-not (Test-Path -LiteralPath $Dir)) { New-Item -ItemType Directory -Path $Dir | Out-Null }
            $Id = [guid]::NewGuid().ToString('N')
            $Entry = @{ Destination=$File.Destination; Temp=(Join-Path $Dir ".$Id.tmp"); Backup=(Join-Path $Dir ".$Id.bak"); Existed=(Test-Path -LiteralPath $File.Destination) }
            $Prepared += $Entry
            Copy-Item -LiteralPath $File.Source -Destination $Entry.Temp
            if ($Entry.Existed) {
                $Handle = [IO.File]::Open($Entry.Destination, 'Open', 'ReadWrite', 'None')
                $Handle.Dispose()
            }
        }
        foreach ($Entry in $Prepared) {
            if ($Entry.Existed) { [IO.File]::Replace($Entry.Temp, $Entry.Destination, $Entry.Backup, $true) }
            else { [IO.File]::Move($Entry.Temp, $Entry.Destination) }
            $Applied += $Entry
        }
    }
    catch {
        $OriginalError = $_.Exception.Message
        [array]::Reverse($Applied)
        foreach ($Entry in $Applied) {
            try {
                if ($Entry.Existed) {
                    $Discard = $Entry.Temp
                    [IO.File]::Replace($Entry.Backup, $Entry.Destination, $Discard, $true)
                }
                else { [IO.File]::Delete($Entry.Destination) }
            }
            catch { $OriginalError += " Rollback failed: $($_.Exception.Message). Backup retained: $($Entry.Backup)" }
        }
        throw $OriginalError
    }
    finally {
        foreach ($Entry in $Prepared) {
            if (Test-Path -LiteralPath $Entry.Temp) { [IO.File]::Delete($Entry.Temp) }
            # ponytail: completed archives are the recovery source; never delete rollback backups on failure.
            if ($Applied.Count -eq $Prepared.Count -and -not $OriginalError -and (Test-Path -LiteralPath $Entry.Backup)) { [IO.File]::Delete($Entry.Backup) }
        }
    }
}

$Lock = New-Object Threading.Mutex($false, 'ResumeManager.Workflow')
if (-not $Lock.WaitOne(0)) { $Lock.Dispose(); throw 'Another resume update is running. Try again when it finishes.' }

try {
if (-not (Test-Path -LiteralPath $UpdatesDir)) {
    New-Item -ItemType Directory -Path $UpdatesDir | Out-Null
}

if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    if (Test-Path -LiteralPath $CurrentPublic) {
        $SourcePath = $CurrentPublic
    }
    else {
        $Latest = Get-ChildItem -LiteralPath $UpdatesDir -Directory |
            Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}(-\d+)?$' } |
            Sort-Object @{Expression={$_.Name.Substring(0,19)};Descending=$true}, @{Expression={if($_.Name.Length -gt 19){[int]$_.Name.Substring(20)}else{1}};Descending=$true} |
            Select-Object -First 1

        if ($null -eq $Latest) {
            throw "No current or archived public resume was found."
        }

        $SourcePath = Join-Path $Latest.FullName $PublicName
    }
}

if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
    throw "Resume source not found: $SourcePath"
}

$SourcePath = (Resolve-Path -LiteralPath $SourcePath).Path
$ArchiveName = $Timestamp
$Counter = 2
while (Test-Path -LiteralPath (Join-Path $UpdatesDir $ArchiveName)) {
    $ArchiveName = "$Timestamp-$Counter"
    $Counter++
}

$NewDir = Join-Path $UpdatesDir $ArchiveName
$StageDir = Join-Path $UpdatesDir (".staging-{0}-{1}" -f $ArchiveName, ([guid]::NewGuid().ToString("N")))
$FailedDir = $null
$Word = $null

try {
    Write-Phase 'Preparing variants'
    New-Item -ItemType Directory -Path $StageDir | Out-Null
    Copy-Item -LiteralPath $SourcePath -Destination (Join-Path $StageDir $PrivateName)
    Copy-Item -LiteralPath (Join-Path $StageDir $PrivateName) -Destination (Join-Path $StageDir $PublicName)

    if ($SkipProjectRewrite) {
        Write-Warning "-SkipProjectRewrite is retained for compatibility; project content is no longer rewritten."
    }

    & $PythonExe $VariantScript $StageDir --config $ConfigPath
    if ($LASTEXITCODE -ne 0) {
        throw "Resume variant generation failed with exit code $LASTEXITCODE."
    }

    if (-not $NoPdf) {
        Write-Phase 'Exporting PDFs'
        $Word = New-Object -ComObject Word.Application
        $Word.Visible = $false
        $Word.DisplayAlerts = 0

        foreach ($DocxName in @($PrivateName, $PublicName)) {
            $Doc = $null
            try {
                $DocxPath = Join-Path $StageDir $DocxName
                $PdfPath = [System.IO.Path]::ChangeExtension($DocxPath, ".pdf")
                $Doc = Open-ResumeWordDocument $Word $DocxPath $true
                $Doc.ExportAsFixedFormat($PdfPath, 17)
            }
            finally {
                if ($null -ne $Doc) {
                    $Doc.Close($false)
                    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Doc)
                }
            }
        }
    }

    $ExpectedFiles = @($PrivateName, $PublicName)
    if (-not $NoPdf) {
        Write-Phase 'Checking privacy'
        $ExpectedFiles += @(
            [System.IO.Path]::ChangeExtension($PrivateName, ".pdf"),
            [System.IO.Path]::ChangeExtension($PublicName, ".pdf")
        )
    }

    foreach ($Name in $ExpectedFiles) {
        $Path = Join-Path $StageDir $Name
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf) -or (Get-Item -LiteralPath $Path).Length -eq 0) {
            throw "Expected output was not created: $Path"
        }
    }

    if (-not $NoPdf) {
        & $PythonExe $VariantScript --validate-pdfs $StageDir --config $ConfigPath
        if ($LASTEXITCODE -ne 0) { throw "PDF privacy validation failed with exit code $LASTEXITCODE." }
    }

    Move-Item -LiteralPath $StageDir -Destination $NewDir
    Write-Phase 'Updating Current'

    $PublishFiles = @(
        @{Source=(Join-Path $NewDir $PrivateName); Destination=$CurrentPrivate},
        @{Source=(Join-Path $NewDir $PublicName); Destination=$CurrentPublic}
    )

    $WebsiteResumeDir = if ($WebsitePublicPath) { Split-Path -Parent $WebsitePublicPath } else { $null }
    if ($WebsiteResumeDir -and (Test-Path -LiteralPath $WebsiteResumeDir -PathType Container)) {
        $PublishFiles += @{Source=(Join-Path $NewDir $PublicName); Destination=$WebsitePublicPath}
    }
    Publish-ResumeFiles -Files $PublishFiles
    Write-Phase 'Version saved'

    Write-Output "Created update folder: $NewDir"
    Write-Output "Updated current private resume: $CurrentPrivate"
    Write-Output "Updated current public resume: $CurrentPublic"
    if ($WebsiteResumeDir -and (Test-Path -LiteralPath $WebsiteResumeDir -PathType Container)) {
        Write-Output "Updated website public resume: $WebsitePublicPath"
    }
    elseif ($WebsitePublicPath) { Write-Output 'Website copy skipped: destination folder does not exist.' }
    else { Write-Output 'Website copy disabled.' }
}
catch {
    $OriginalFailure = $_.Exception.Message
    if (Test-Path -LiteralPath $StageDir) {
        $FailedDir = Join-Path $UpdatesDir (".failed-{0}-{1}" -f $ArchiveName, ([guid]::NewGuid().ToString("N")))
        try {
            Move-Item -LiteralPath $StageDir -Destination $FailedDir
        }
        catch {
            $FailedDir = $StageDir
        }
    }

    $FailureMessage = "Resume update failed: $OriginalFailure Source preserved at: $SourcePath"
    if (Test-Path -LiteralPath $NewDir) { $FailureMessage += " Completed archive retained at: $NewDir" }
    if ($null -ne $FailedDir) {
        $FailureMessage += " Staging preserved at: $FailedDir"
    }
    throw $FailureMessage
}
finally {
    if ($null -ne $Word) {
        try { $Word.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Word)
    }
}
}
finally { $Lock.ReleaseMutex(); $Lock.Dispose() }
