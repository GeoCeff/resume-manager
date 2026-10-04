param([switch]$SelfTest, [switch]$LoadOnly, [string]$ConfigPath = $env:RESUME_MANAGER_CONFIG)

$ErrorActionPreference = "Stop"

try {
    . (Join-Path $PSScriptRoot 'resume_settings.ps1') -ConfigPath $ConfigPath
}
catch {
    if ($SelfTest -or $LoadOnly) { throw }
    Add-Type -AssemblyName System.Windows.Forms
    [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Resume Manager - Setup')
    exit 1
}
$WorkflowPath = Join-Path $PSScriptRoot "resume_update_workflow.ps1"
$PowerShellExe = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
$SessionRoot = Join-Path $env:TEMP 'ResumeManagerSessions'

function Get-ResumeHash {
    param([Parameter(Mandatory)][string]$Path)
    $Result = & $PythonExe $VariantScript --fingerprint $Path 2>&1
    if ($LASTEXITCODE -ne 0) { throw ($Result -join "`n") }
    return [string]($Result | Select-Object -Last 1)
}

function Test-ArchiveNeeded {
    param([string]$CurrentHash, [string]$ArchivedHash)
    return -not [string]::Equals($CurrentHash, $ArchivedHash, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-LatestArchive {
    $Latest = Get-ChildItem -LiteralPath $UpdatesDir -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}(-\d+)?$' } |
        Sort-Object @{Expression={$_.Name.Substring(0,19)};Descending=$true}, @{Expression={if($_.Name.Length -gt 19){[int]$_.Name.Substring(20)}else{1}};Descending=$true} |
        Select-Object -First 1
    if ($null -eq $Latest) { return "None yet" }
    return $Latest.Name
}

if ($SelfTest) {
    foreach ($Path in @($WorkflowPath, $CurrentPrivate, $CurrentPublic)) {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
            throw "Required file missing: $Path"
        }
    }
    if (Test-ArchiveNeeded -CurrentHash "same" -ArchivedHash "same") {
        throw "Unchanged files must not create a new archive"
    }
    if (-not (Test-ArchiveNeeded -CurrentHash "new" -ArchivedHash "old")) {
        throw "Changed files must create a new archive"
    }
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $TestForm = New-Object System.Windows.Forms.Form
    $TestForm.Dispose()
    $TestWatcher = New-Object System.IO.FileSystemWatcher($CurrentDir, "*.docx")
    $TestWatcher.Dispose()
    $TestWord = New-Object -ComObject Word.Application
    try { $TestWord.Quit() } finally { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($TestWord) }
    Write-Output "Resume Manager self-test passed. Latest archive: $(Get-LatestArchive)"
    exit 0
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$CreatedNew = $false
$Mutex = New-Object System.Threading.Mutex($true, 'ResumeManager.Application', [ref]$CreatedNew)
if (-not $CreatedNew) {
    [System.Windows.Forms.MessageBox]::Show(
        "Resume Manager is already running.",
        "Resume Manager",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
    exit 0
}

$script:SessionActive = $false
$script:WorkingPath = $null
$script:WorkingDir = $null
$script:Word = $null
$script:WordDocument = $null
$script:Watcher = $null
$script:PendingChange = $false
$script:LastFileEvent = [datetime]::MinValue
$script:LastSignature = $null
$script:LastArchivedHash = $null
$script:HashBeingArchived = $null
$script:ArchiveProcess = $null
$script:ArchiveStdout = $null
$script:ArchiveStderr = $null
$script:WordClosed = $false
$script:LastFailedHash = $null

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Resume Manager"
$Form.StartPosition = "CenterScreen"
$Form.ClientSize = New-Object System.Drawing.Size(520, 330)
$Form.FormBorderStyle = "FixedDialog"
$Form.MaximizeBox = $false
$Form.BackColor = [System.Drawing.Color]::White

$Title = New-Object System.Windows.Forms.Label
$Title.Text = "Resume Manager"
$Title.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$Title.AutoSize = $true
$Title.Location = New-Object System.Drawing.Point(28, 22)
$Form.Controls.Add($Title)

$Intro = New-Object System.Windows.Forms.Label
$Intro.Text = "Edit in Word. Saved changes become dated private and public versions."
$Intro.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$Intro.AutoSize = $true
$Intro.Location = New-Object System.Drawing.Point(31, 61)
$Form.Controls.Add($Intro)

function New-ManagerButton {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width = 220)
    $Button = New-Object System.Windows.Forms.Button
    $Button.Text = $Text
    $Button.Location = New-Object System.Drawing.Point($X, $Y)
    $Button.Size = New-Object System.Drawing.Size($Width, 42)
    $Button.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $Button.FlatStyle = "System"
    $Form.Controls.Add($Button)
    return $Button
}

$EditPrivateButton = New-ManagerButton -Text "Edit Private Resume" -X 28 -Y 92
$EditPublicButton = New-ManagerButton -Text "Edit Public Resume" -X 270 -Y 92
$OpenCurrentButton = New-ManagerButton -Text "Open Current Folder" -X 28 -Y 148
$OpenHistoryButton = New-ManagerButton -Text "Open Version History" -X 270 -Y 148

$LatestCaption = New-Object System.Windows.Forms.Label
$LatestCaption.Text = "Latest archived version"
$LatestCaption.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$LatestCaption.AutoSize = $true
$LatestCaption.Location = New-Object System.Drawing.Point(28, 214)
$Form.Controls.Add($LatestCaption)

$LatestValue = New-Object System.Windows.Forms.Label
$LatestValue.Text = Get-LatestArchive
$LatestValue.Font = New-Object System.Drawing.Font("Consolas", 10)
$LatestValue.AutoSize = $true
$LatestValue.Location = New-Object System.Drawing.Point(28, 236)
$Form.Controls.Add($LatestValue)

$StatusCaption = New-Object System.Windows.Forms.Label
$StatusCaption.Text = "Status"
$StatusCaption.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$StatusCaption.AutoSize = $true
$StatusCaption.Location = New-Object System.Drawing.Point(270, 214)
$Form.Controls.Add($StatusCaption)

$StatusValue = New-Object System.Windows.Forms.Label
$StatusValue.Text = "Ready"
$StatusValue.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$StatusValue.AutoEllipsis = $true
$StatusValue.Location = New-Object System.Drawing.Point(270, 236)
$StatusValue.Size = New-Object System.Drawing.Size(220, 24)
$Form.Controls.Add($StatusValue)

$Hint = New-Object System.Windows.Forms.Label
$Hint.Text = "Save normally in Word. The manager handles versioning automatically."
$Hint.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$Hint.ForeColor = [System.Drawing.Color]::DimGray
$Hint.AutoSize = $true
$Hint.Location = New-Object System.Drawing.Point(28, 298)
$Form.Controls.Add($Hint)

$RetryButton = New-ManagerButton -Text 'Retry save' -X 270 -Y 266 -Width 110
$RetryButton.Height = 26
$RetryButton.Visible = $false

function Set-Status {
    param([string]$Text, [switch]$Error)
    $StatusValue.Text = $Text
    $StatusValue.ForeColor = if ($Error) { [System.Drawing.Color]::Firebrick } else { [System.Drawing.Color]::Black }
}

function Set-EditButtonsEnabled {
    param([bool]$Enabled)
    $EditPrivateButton.Enabled = $Enabled
    $EditPublicButton.Enabled = $Enabled
}

function Remove-WorkingDirectory {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return }
    $ResolvedRoot = [System.IO.Path]::GetFullPath($SessionRoot).TrimEnd('\') + '\'
    $ResolvedPath = [System.IO.Path]::GetFullPath($Path).TrimEnd('\') + '\'
    if (-not $ResolvedPath.StartsWith($ResolvedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove unexpected working directory: $ResolvedPath"
    }
    [System.IO.Directory]::Delete($Path, $true)
}

function Release-WordObjects {
    if ($null -ne $script:WordDocument) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:WordDocument) } catch {}
        $script:WordDocument = $null
    }
    if ($null -ne $script:Word) {
        try {
            if ($script:Word.Documents.Count -eq 0) { $script:Word.Quit() }
        }
        catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:Word) } catch {}
        $script:Word = $null
    }
}

function End-EditSession {
    param([switch]$PreserveWorkingCopy, [string]$Message = "Ready")

    if ($null -ne $script:Watcher) {
        try { $script:Watcher.EnableRaisingEvents = $false; $script:Watcher.Dispose() } catch {}
        $script:Watcher = $null
    }
    Release-WordObjects

    $SavedWorkingPath = $script:WorkingPath
    if (-not $PreserveWorkingCopy) {
        try { Remove-WorkingDirectory -Path $script:WorkingDir } catch {}
    }

    $script:SessionActive = $false
    $script:WorkingPath = $null
    $script:WorkingDir = $null
    $script:PendingChange = $false
    $script:ArchiveProcess = $null
    $script:WordClosed = $false
    $RetryButton.Visible = $false
    Set-EditButtonsEnabled -Enabled $true
    Set-Status -Text $Message

    if ($PreserveWorkingCopy) {
        [System.Windows.Forms.MessageBox]::Show(
            "The edited working copy was preserved at:`n$SavedWorkingPath",
            "Resume Manager",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
    }
}

function Start-ArchiveProcess {
    $SnapshotPath = Join-Path $script:WorkingDir 'saved-snapshot.docx'
    $Before = Get-Item -LiteralPath $script:WorkingPath
    $BeforeSignature = "$($Before.Length)|$($Before.LastWriteTimeUtc.Ticks)"
    # Word keeps a writer reservation after Save. Read the saved package without blocking it.
    $InputFile = [IO.File]::Open($script:WorkingPath, 'Open', 'Read', 'ReadWrite')
    try {
        $OutputFile = [IO.File]::Create($SnapshotPath)
        try { $InputFile.CopyTo($OutputFile) } finally { $OutputFile.Dispose() }
    }
    finally { $InputFile.Dispose() }
    $Hash = Get-ResumeHash -Path $SnapshotPath
    $After = Get-Item -LiteralPath $script:WorkingPath
    $AfterSignature = "$($After.Length)|$($After.LastWriteTimeUtc.Ticks)"
    if ($BeforeSignature -ne $AfterSignature -or $Hash -ne (Get-ResumeHash -Path $script:WorkingPath)) {
        $script:LastFileEvent = [datetime]::UtcNow
        throw 'Word is still saving; waiting for a stable saved package.'
    }
    if (-not (Test-ArchiveNeeded -CurrentHash $Hash -ArchivedHash $script:LastArchivedHash)) {
        $script:PendingChange = $false
        if ($script:WordClosed) { End-EditSession }
        return
    }
    if ($Hash -eq $script:LastFailedHash) {
        $script:PendingChange = $false
        if ($script:WordClosed) { End-EditSession -PreserveWorkingCopy -Message 'Error' }
        return
    }

    Set-Status -Text "Saving new version"
    $script:HashBeingArchived = $Hash
    $script:PendingChange = $false
    $script:ArchiveStdout = Join-Path $script:WorkingDir "workflow.stdout.txt"
    $script:ArchiveStderr = Join-Path $script:WorkingDir "workflow.stderr.txt"

    $EscapedWorkflow = $WorkflowPath.Replace("'", "''")
    $EscapedSource = $SnapshotPath.Replace("'", "''")
    $EscapedConfig = $ConfigPath.Replace("'", "''")
    $Command = "try { & '$EscapedWorkflow' -SourcePath '$EscapedSource' -ConfigPath '$EscapedConfig'; exit 0 } catch { [Console]::Error.WriteLine(`$_); exit 1 }"
    $Encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Command))

    $script:ArchiveProcess = Start-Process -FilePath $PowerShellExe `
        -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", $Encoded) `
        -WindowStyle Hidden `
        -RedirectStandardOutput $script:ArchiveStdout `
        -RedirectStandardError $script:ArchiveStderr `
        -PassThru
    # Keep the process handle so Windows PowerShell can retrieve ExitCode after Refresh.
    $null = $script:ArchiveProcess.Handle
}

function Complete-ArchiveProcess {
    $script:ArchiveProcess.Refresh()
    if (-not $script:ArchiveProcess.HasExited) { return }
    $script:ArchiveProcess.WaitForExit()
    $ExitCode = $script:ArchiveProcess.ExitCode
    $Output = if (Test-Path -LiteralPath $script:ArchiveStdout) { Get-Content -LiteralPath $script:ArchiveStdout -Raw } else { "" }
    $Errors = if (Test-Path -LiteralPath $script:ArchiveStderr) { Get-Content -LiteralPath $script:ArchiveStderr -Raw } else { "" }
    $script:ArchiveProcess.Dispose()
    $script:ArchiveProcess = $null

    if ($ExitCode -ne 0) {
        $Detail = ($Errors + "`n" + $Output).Trim()
        if ([string]::IsNullOrWhiteSpace($Detail)) { $Detail = "The resume workflow exited with code $ExitCode." }
        Set-Status -Text "Error" -Error
        $script:LastFailedHash = $script:HashBeingArchived
        $RetryButton.Visible = $true
        [System.Windows.Forms.MessageBox]::Show(
            $Detail,
            "Resume Manager Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        if ($script:WordClosed) { End-EditSession -PreserveWorkingCopy -Message "Error" }
        return
    }

    $script:LastArchivedHash = $script:HashBeingArchived
    $script:LastFailedHash = $null
    $RetryButton.Visible = $false
    $LatestValue.Text = Get-LatestArchive
    Set-Status -Text "Version saved"

    try {
        $CurrentHash = Get-ResumeHash -Path $script:WorkingPath
        if (Test-ArchiveNeeded -CurrentHash $CurrentHash -ArchivedHash $script:LastArchivedHash) {
            $script:PendingChange = $true
            $script:LastFileEvent = [datetime]::UtcNow
        }
        elseif ($script:WordClosed) {
            End-EditSession -Message "Version saved"
        }
    }
    catch {
        $script:PendingChange = $true
        $script:LastFileEvent = [datetime]::UtcNow
    }
}

function Start-EditSession {
    param([ValidateSet("Private", "Public")][string]$Kind)

    if ($script:SessionActive) { return }
    $Source = if ($Kind -eq "Private") { $CurrentPrivate } else { $CurrentPublic }
    if (-not (Test-Path -LiteralPath $Source -PathType Leaf)) {
        [System.Windows.Forms.MessageBox]::Show("Resume not found:`n$Source", "Resume Manager") | Out-Null
        return
    }

    $script:WorkingDir = Join-Path $SessionRoot ([guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $script:WorkingDir -Force | Out-Null
    $script:WorkingPath = Join-Path $script:WorkingDir ([System.IO.Path]::GetFileName($Source))
    Copy-Item -LiteralPath $Source -Destination $script:WorkingPath

    try {
        $script:LastArchivedHash = Get-ResumeHash -Path $script:WorkingPath
        $Item = Get-Item -LiteralPath $script:WorkingPath
        $script:LastSignature = "$($Item.Length)|$($Item.LastWriteTimeUtc.Ticks)"
        $script:PendingChange = $false
        $script:WordClosed = $false
        $script:LastFailedHash = $null
        $RetryButton.Visible = $false

        $script:Watcher = New-Object System.IO.FileSystemWatcher($script:WorkingDir, ([System.IO.Path]::GetFileName($script:WorkingPath)))
        $script:Watcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite -bor [System.IO.NotifyFilters]::Size -bor [System.IO.NotifyFilters]::FileName
        $script:Watcher.SynchronizingObject = $Form
        $script:Watcher.add_Changed({ $script:PendingChange = $true; $script:LastFileEvent = [datetime]::UtcNow })
        $script:Watcher.add_Created({ $script:PendingChange = $true; $script:LastFileEvent = [datetime]::UtcNow })
        $script:Watcher.EnableRaisingEvents = $true

        $script:Word = New-Object -ComObject Word.Application
        $script:Word.Visible = $true
        $script:WordDocument = $script:Word.Documents.Open($script:WorkingPath)

        $script:SessionActive = $true
        Set-EditButtonsEnabled -Enabled $false
        Set-Status -Text "Editing $($Kind.ToLower()) resume"
    }
    catch {
        if ($null -ne $script:Watcher) { try { $script:Watcher.Dispose() } catch {} }
        Release-WordObjects
        try { Remove-WorkingDirectory -Path $script:WorkingDir } catch {}
        $script:WorkingPath = $null
        $script:WorkingDir = $null
        Set-Status -Text "Error" -Error
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Resume Manager Error") | Out-Null
    }
}

$EditPrivateButton.Add_Click({ Start-EditSession -Kind "Private" })
$EditPublicButton.Add_Click({ Start-EditSession -Kind "Public" })
$OpenCurrentButton.Add_Click({ Start-Process -FilePath "explorer.exe" -ArgumentList ('"{0}"' -f $CurrentDir) })
$OpenHistoryButton.Add_Click({ Start-Process -FilePath "explorer.exe" -ArgumentList ('"{0}"' -f $UpdatesDir) })
$RetryButton.Add_Click({
    $script:LastFailedHash = $null
    $script:PendingChange = $true
    $script:LastFileEvent = [datetime]::UtcNow.AddSeconds(-3)
    $RetryButton.Visible = $false
})

$Timer = New-Object System.Windows.Forms.Timer
$Timer.Interval = 750
$Timer.Add_Tick({
    if (-not $script:SessionActive) { return }

    if ($null -ne $script:ArchiveProcess) {
        Complete-ArchiveProcess
        return
    }

    if (-not $script:WordClosed) {
        try {
            if ($script:WordDocument.Windows.Count -eq 0) { $script:WordClosed = $true }
            elseif ($script:WordDocument.FullName -ne $script:WorkingPath) {
                # Follow Word Save As without deleting the user's chosen destination.
                $script:WorkingPath = $script:WordDocument.FullName
                $script:Watcher.EnableRaisingEvents = $false
                $script:Watcher.Path = Split-Path -Parent $script:WorkingPath
                $script:Watcher.Filter = [IO.Path]::GetFileName($script:WorkingPath)
                $script:Watcher.EnableRaisingEvents = $true
                $script:LastSignature = $null
            }
        }
        catch {
            $Code = $_.Exception.HResult
            if ($Code -eq -2147418111 -or $Code -eq -2147417846) { return }
            $script:WordClosed = $true
        }
        if ($script:WordClosed) {
            Set-Status -Text "Waiting for Word"
            $script:PendingChange = $true
            $script:LastFileEvent = [datetime]::UtcNow.AddSeconds(-3)
        }
    }

    try {
        $Item = Get-Item -LiteralPath $script:WorkingPath
        $Signature = "$($Item.Length)|$($Item.LastWriteTimeUtc.Ticks)"
        if ($Signature -ne $script:LastSignature) {
            $script:LastSignature = $Signature
            $script:PendingChange = $true
            $script:LastFileEvent = [datetime]::UtcNow
        }
    }
    catch { return }

    if (-not $script:PendingChange) { return }
    if (([datetime]::UtcNow - $script:LastFileEvent).TotalSeconds -lt 1.5) { return }

    try {
        Start-ArchiveProcess
    }
    catch {
        Set-Status -Text "Waiting for Word"
    }
})
$Timer.Start()

$Form.Add_FormClosing({
    param($Sender, $EventArgs)
    if ($script:SessionActive -or $null -ne $script:ArchiveProcess) {
        $EventArgs.Cancel = $true
        [System.Windows.Forms.MessageBox]::Show(
            "Close the Word document and wait for Resume Manager to finish saving first.",
            "Resume Manager",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
})

if ($LoadOnly) { return }

try {
    [void]$Form.ShowDialog()
}
finally {
    $Timer.Stop()
    $Timer.Dispose()
    Release-WordObjects
    if ($CreatedNew) { try { $Mutex.ReleaseMutex() } catch {} }
    $Mutex.Dispose()
}
