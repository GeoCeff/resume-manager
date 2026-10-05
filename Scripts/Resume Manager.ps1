param([switch]$SelfTest, [switch]$LoadOnly, [string]$ConfigPath = $env:RESUME_MANAGER_CONFIG)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot 'manager_dialogs.ps1')
if ([string]::IsNullOrWhiteSpace($ConfigPath)) { $ConfigPath = Join-Path $env:LOCALAPPDATA 'ResumeManager\settings.local.json' }
$CreatedNew = $false
$Mutex = New-Object System.Threading.Mutex($true, 'ResumeManager.Application', [ref]$CreatedNew)
if (-not $CreatedNew) {
    [void][Windows.Forms.MessageBox]::Show('Resume Manager is already running.','Resume Manager')
    $Mutex.Dispose(); exit 0
}

while ($true) {
    $StartupIssue = Get-StartupIssue $ConfigPath
    if($StartupIssue.Kind -eq 'Ready'){break}
    if($SelfTest -or $LoadOnly){$Mutex.ReleaseMutex();$Mutex.Dispose();throw ($StartupIssue.Message + "`n" + $StartupIssue.Details)}
    $Choice = if($StartupIssue.Kind -eq 'Setup'){'Edit settings'}else{Show-StartupRecovery $StartupIssue}
    if($Choice -eq 'Retry'){continue}
    if($Choice -eq 'Cancel' -or -not (Show-SettingsDialog $ConfigPath $null)){$Mutex.ReleaseMutex();$Mutex.Dispose();exit 0}
}
. (Join-Path $PSScriptRoot 'resume_settings.ps1') -ConfigPath $ConfigPath
$PythonExe = $StartupIssue.Python
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
    $Policy = New-ContactPolicy @('public@example.com','personal@example.com','Example City') @('public@example.com','Example City') @('personal@example.com')
    if($Policy.PublicContact -match 'personal@' -or (Test-PublicContactField '000-0000-0000' @('00000000000'))){throw 'Public contact policy failed'}
    if((Get-StartupIssue (Join-Path $env:TEMP ([guid]::NewGuid().ToString('N')+'.json'))).Kind -ne 'Setup'){throw 'Missing settings must be first-run setup'}
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $TestForm = New-Object System.Windows.Forms.Form
    $TestForm.Dispose()
    $TestWatcher = New-Object System.IO.FileSystemWatcher($CurrentDir, "*.docx")
    $TestWatcher.Dispose()
    $TestWord = New-Object -ComObject Word.Application
    try { if($TestWord.Documents.Count -eq 0){$TestWord.Quit()} } finally { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($TestWord) }
    Write-Output "Resume Manager self-test passed. Latest archive: $(Get-LatestArchive)"
    $Mutex.ReleaseMutex(); $Mutex.Dispose()
    exit 0
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

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
$script:LastErrorDetail = ''
$script:RecoveryPath = $Root
$script:ArchiveStarted = [datetime]::MinValue
$script:ArchiveProgress = $null
$script:StallShown = $false
$script:Restoring = $false

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Resume Manager"
$Form.StartPosition = "CenterScreen"
$Form.ClientSize = New-Object System.Drawing.Size(720, 540)
$Form.MinimumSize = New-Object System.Drawing.Size(680,560)
$Form.Font = New-Object Drawing.Font('Segoe UI',10)
$Form.AutoScaleMode = 'Font'
$Form.BackColor = [Drawing.SystemColors]::Window
$Form.ForeColor = [Drawing.SystemColors]::WindowText
$Form.KeyPreview = $true
$MainLayout = New-Object Windows.Forms.TableLayoutPanel
$MainLayout.Dock = 'Fill'; $MainLayout.Padding = New-Object Windows.Forms.Padding(20)
$MainLayout.ColumnCount = 1; $MainLayout.AutoScroll = $true
$Form.Controls.Add($MainLayout)

$Title = New-Object System.Windows.Forms.Label
$Title.Text = "Resume Manager"
$Title.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$Title.AutoSize = $true
$MainLayout.Controls.Add($Title)

$Intro = New-Object System.Windows.Forms.Label
$Intro.Text = "Edit either version in Word. Saved content updates both; contacts stay separate."
$Intro.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$Intro.AutoSize = $true
$Intro.MaximumSize = New-Object Drawing.Size(640,0)
$Intro.Margin = New-Object Windows.Forms.Padding(0,8,0,18)
$MainLayout.Controls.Add($Intro)

$ResumePanels = New-Object Windows.Forms.TableLayoutPanel
$ResumePanels.AutoSize = $true; $ResumePanels.Dock = 'Fill'; $ResumePanels.ColumnCount = 2
foreach($i in 1..2){[void]$ResumePanels.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent',50)))}
$MainLayout.Controls.Add($ResumePanels)
$PrivatePanel = New-Object Windows.Forms.GroupBox; $PrivatePanel.Text = 'Private Resume'; $PrivatePanel.Dock = 'Fill'; $PrivatePanel.AutoSize = $true
$PublicPanel = New-Object Windows.Forms.GroupBox; $PublicPanel.Text = 'Public Resume'; $PublicPanel.Dock = 'Fill'; $PublicPanel.AutoSize = $true
$ResumePanels.Controls.Add($PrivatePanel,0,0); $ResumePanels.Controls.Add($PublicPanel,1,0)
$PrivateActions = New-Object Windows.Forms.FlowLayoutPanel; $PrivateActions.FlowDirection = 'TopDown'; $PrivateActions.Dock = 'Fill'; $PrivateActions.AutoSize = $true; $PrivateActions.Padding = New-Object Windows.Forms.Padding(8,12,8,8)
$PublicActions = New-Object Windows.Forms.FlowLayoutPanel; $PublicActions.FlowDirection = 'TopDown'; $PublicActions.Dock = 'Fill'; $PublicActions.AutoSize = $true; $PublicActions.Padding = $PrivateActions.Padding
$PrivatePanel.Controls.Add($PrivateActions); $PublicPanel.Controls.Add($PublicActions)
$EditPrivateButton = New-UiButton 'Edit &Private Resume'; $EditPublicButton = New-UiButton 'Edit P&ublic Resume'
$PreviewPrivateButton = New-UiButton 'Open Private PDF'; $PreviewPublicButton = New-UiButton 'Open Public PDF'
$PrivateActions.Controls.AddRange(@($EditPrivateButton,$PreviewPrivateButton)); $PublicActions.Controls.AddRange(@($EditPublicButton,$PreviewPublicButton))
$PdfPreviewDate = New-UiLabel 'Archived PDFs only - not unsaved Word edits.'; $PdfPreviewDate.MaximumSize = New-Object Drawing.Size(640,0); $MainLayout.Controls.Add($PdfPreviewDate)
$Secondary = New-Object Windows.Forms.FlowLayoutPanel; $Secondary.AutoSize = $true; $Secondary.Dock = 'Fill'; $Secondary.Margin = New-Object Windows.Forms.Padding(0,12,0,12)
$OpenCurrentButton = New-UiButton 'Open &Current Folder'; $OpenHistoryButton = New-UiButton 'Open &Version History'; $SettingsButton = New-UiButton '&Settings'
$Secondary.Controls.AddRange(@($OpenCurrentButton,$OpenHistoryButton,$SettingsButton)); $MainLayout.Controls.Add($Secondary)

$LatestCaption = New-Object System.Windows.Forms.Label
$LatestCaption.Text = "Latest archived version"
$LatestCaption.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$LatestCaption.AutoSize = $true
$MainLayout.Controls.Add($LatestCaption)

$LatestValue = New-Object System.Windows.Forms.Label
$LatestValue.Text = Format-ArchiveDate (Get-LatestArchive)
$LatestValue.Font = New-Object System.Drawing.Font("Consolas", 10)
$LatestValue.AutoSize = $true
$MainLayout.Controls.Add($LatestValue)

$StatusCaption = New-Object System.Windows.Forms.Label
$StatusCaption.Text = "Status"
$StatusCaption.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$StatusCaption.AutoSize = $true
$MainLayout.Controls.Add($StatusCaption)

$StatusValue = New-Object System.Windows.Forms.Label
$StatusValue.Text = "Ready"
$StatusValue.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$StatusValue.AutoSize = $true; $StatusValue.MaximumSize = New-Object Drawing.Size(640,0); $StatusValue.AccessibleName = 'Save status'
$MainLayout.Controls.Add($StatusValue)
$ProgressBar = New-Object Windows.Forms.ProgressBar; $ProgressBar.Dock = 'Fill'; $ProgressBar.Height = 8; $ProgressBar.Style = 'Marquee'; $ProgressBar.Visible = $false
$MainLayout.Controls.Add($ProgressBar)
$ReviewSummary = New-UiLabel 'Review the public PDF before sharing.'; $ReviewSummary.MaximumSize = New-Object Drawing.Size(640,0); $MainLayout.Controls.Add($ReviewSummary)
$DetailsToggle = New-Object Windows.Forms.CheckBox; $DetailsToggle.Text = 'Show &details'; $DetailsToggle.AutoSize = $true; $MainLayout.Controls.Add($DetailsToggle)
$DetailsPanel = New-Object Windows.Forms.FlowLayoutPanel; $DetailsPanel.AutoSize = $true; $DetailsPanel.Dock = 'Fill'; $DetailsPanel.FlowDirection = 'TopDown'; $DetailsPanel.Visible = $false
$StatusDetail = New-UiLabel 'Ready'; $StatusDetail.MaximumSize = New-Object Drawing.Size(640,0); $DetailsPanel.Controls.Add($StatusDetail)
$ReviewValue = New-UiLabel 'Open a public PDF to review its layout before sharing.'; $ReviewValue.MaximumSize = New-Object Drawing.Size(640,0); $DetailsPanel.Controls.Add($ReviewValue)
$ErrorDetail = New-UiLabel ''; $ErrorDetail.MaximumSize = New-Object Drawing.Size(640,0); $DetailsPanel.Controls.Add($ErrorDetail); $MainLayout.Controls.Add($DetailsPanel)
$DetailsToggle.Add_CheckedChanged({$DetailsPanel.Visible=$DetailsToggle.Checked})
$ErrorPanel = New-Object Windows.Forms.FlowLayoutPanel; $ErrorPanel.AutoSize = $true; $ErrorPanel.FlowDirection = 'TopDown'; $ErrorPanel.Dock = 'Fill'; $ErrorPanel.Visible = $false
$ErrorValue = New-UiLabel ''; $ErrorValue.MaximumSize = New-Object Drawing.Size(640,150)
$ErrorPanel.Controls.Add($ErrorValue)
$RecoveryActions = New-Object Windows.Forms.FlowLayoutPanel; $RecoveryActions.AutoSize = $true
$RetryButton = New-UiButton '&Retry save'; $OpenRecoveryButton = New-UiButton 'Open recovery folder'; $CopyDiagnosticsButton = New-UiButton 'Copy redacted diagnostics'
$RecoveryActions.Controls.AddRange(@($RetryButton,$OpenRecoveryButton,$CopyDiagnosticsButton)); $ErrorPanel.Controls.Add($RecoveryActions); $MainLayout.Controls.Add($ErrorPanel)

$Hint = New-Object System.Windows.Forms.Label
$Hint.Text = "Save normally in Word. The manager handles versioning automatically."
$Hint.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$Hint.ForeColor = [Drawing.SystemColors]::WindowText
$Hint.AutoSize = $true
$Hint.MaximumSize = New-Object Drawing.Size(640,0); $MainLayout.Controls.Add($Hint)
$RetryButton.Visible = $false

function Set-Status {
    param([string]$Text, [switch]$Error)
    $StatusDetail.Text = $Text
    $StatusValue.Text = if($Text -match '^Editing'){'Editing in Word'}elseif($Text -match '^(Preparing|Exporting|Checking|Updating|Saving)'){'Saving version'}else{$Text}
    $StatusValue.ForeColor = if ($Error) { [Drawing.SystemColors]::HotTrack } else { [Drawing.SystemColors]::WindowText }
    $ProgressBar.Visible = $null -ne $script:ArchiveProcess
}

function Set-EditButtonsEnabled {
    param([bool]$Enabled)
    $EditPrivateButton.Enabled = $Enabled
    $EditPublicButton.Enabled = $Enabled
    $SettingsButton.Enabled = $Enabled
}

function Show-ManagerError([string]$Detail) {
    $script:LastErrorDetail = $Detail
    $ErrorDetail.Text = $Detail
    $ErrorValue.Text = if($Detail -match 'Private contact details|private-only|PublicContact|privacy validation|Configured private details'){'Privacy check failed. Do not share this output; review details to correct it.'}else{'The operation failed. Your source is retained. Retry or open recovery; the exact error is under Show details.'}
    $ErrorPanel.Visible = $true
    $RetryButton.Visible = $script:SessionActive -and $null -eq $script:ArchiveProcess
    Set-Status 'Error' -Error
}

function Update-PublicReview {
    $Latest = @(Get-ArchiveDirectories $UpdatesDir) | Select-Object -First 1
    $PreviewPrivateButton.Enabled = $Latest -and (Test-Path -LiteralPath (Join-Path $Latest.FullName ([IO.Path]::ChangeExtension($PrivateName,'.pdf'))))
    $PreviewPublicButton.Enabled = $Latest -and (Test-Path -LiteralPath (Join-Path $Latest.FullName ([IO.Path]::ChangeExtension($PublicName,'.pdf'))))
    if(-not $Latest){$PdfPreviewDate.Text='No archived PDF yet. Save an edit to create the first version.';return}
    $PdfPreviewDate.Text = 'PDF previews saved: ' + (Format-ArchiveDate $Latest.Name) + '. Not unsaved Word edits.'
    try {
        $Result = & $PythonExe $VariantScript --review $Latest.FullName --config $ConfigPath 2>&1
        if($LASTEXITCODE -ne 0){throw ($Result -join "`n")}
        $Review = ($Result -join "`n") | ConvertFrom-Json
        $ReviewValue.Text = "Private PDF: $($Review.PrivatePages) page(s) | Public PDF: $($Review.PublicPages) page(s)`n$($Review.Checks)`n$($Review.Warnings -join '; ')"
        $ReviewSummary.Text = "Private PDF: $($Review.PrivatePages) page(s) | Public PDF: $($Review.PublicPages) page(s). Review public output before sharing."
        if($Review.PrivatePages -gt 1 -or $Review.PublicPages -gt 1){$ReviewSummary.Text += ' More than one page: check layout.'}
        $ReviewSummary.ForeColor = [Drawing.SystemColors]::WindowText
    }catch{$ReviewValue.Text = 'Latest output review unavailable: ' + $_.Exception.Message; $ReviewSummary.Text='Latest output is incomplete or failed review. Do not share it before checking details.'; $ReviewSummary.ForeColor=[Drawing.SystemColors]::HotTrack}
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
    $script:RecoveryPath = $script:WorkingDir
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

    if ($PreserveWorkingCopy) { Show-ManagerError ($script:LastErrorDetail + "`nEdited copy preserved: $SavedWorkingPath") }
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
    $script:ArchiveProgress = Join-Path $script:WorkingDir 'progress.json'
    $script:ArchiveStarted = [datetime]::UtcNow
    $script:StallShown = $false
    $ErrorPanel.Visible = $false

    $EscapedWorkflow = $WorkflowPath.Replace("'", "''")
    $EscapedSource = $SnapshotPath.Replace("'", "''")
    $EscapedConfig = $ConfigPath.Replace("'", "''")
    $Command = "try { & '$EscapedWorkflow' -SourcePath '$EscapedSource' -ConfigPath '$EscapedConfig' -ProgressPath '$($script:ArchiveProgress.Replace("'","''"))'; exit 0 } catch { [Console]::Error.WriteLine(`$_); exit 1 }"
    $Encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Command))

    $script:ArchiveProcess = Start-Process -FilePath $PowerShellExe `
        -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", $Encoded) `
        -WindowStyle Hidden `
        -RedirectStandardOutput $script:ArchiveStdout `
        -RedirectStandardError $script:ArchiveStderr `
        -PassThru
    # Keep the process handle so Windows PowerShell can retrieve ExitCode after Refresh.
    $null = $script:ArchiveProcess.Handle
    $ProgressBar.Visible = $true
}

function Complete-ArchiveProcess {
    $script:ArchiveProcess.Refresh()
    if (-not $script:ArchiveProcess.HasExited) {
        $Phase = 'Saving new version'
        if(Test-Path -LiteralPath $script:ArchiveProgress){try{$Phase=(Get-Content -LiteralPath $script:ArchiveProgress -Raw | ConvertFrom-Json).Phase}catch{}}
        $Queued = if($script:PendingChange){' - another save queued'}else{''}
        $Elapsed = [int]([datetime]::UtcNow-$script:ArchiveStarted).TotalSeconds
        Set-Status "Saving version - $Phase ($Elapsed s)$Queued"
        if($Elapsed -gt 120 -and -not $script:StallShown){
            $script:StallShown=$true; $script:RecoveryPath=$script:WorkingDir
            $ErrorValue.Text = 'Export is taking longer than expected. Check Word for a dialog. The manager is still monitoring; recovery files are retained. It will not terminate Word automatically.'
            $script:LastErrorDetail=$ErrorValue.Text; $ErrorDetail.Text=$ErrorValue.Text
            $ErrorPanel.Visible=$true; $RetryButton.Visible=$false
        }
        return
    }
    $script:ArchiveProcess.WaitForExit()
    $ExitCode = $script:ArchiveProcess.ExitCode
    $Output = if (Test-Path -LiteralPath $script:ArchiveStdout) { Get-Content -LiteralPath $script:ArchiveStdout -Raw } else { "" }
    $Errors = if (Test-Path -LiteralPath $script:ArchiveStderr) { Get-Content -LiteralPath $script:ArchiveStderr -Raw } else { "" }
    $script:ArchiveProcess.Dispose()
    $script:ArchiveProcess = $null
    $ProgressBar.Visible = $false

    if ($ExitCode -ne 0) {
        $Detail = ($Errors + "`n" + $Output).Trim()
        if ([string]::IsNullOrWhiteSpace($Detail)) { $Detail = "The resume workflow exited with code $ExitCode." }
        $script:LastFailedHash = $script:HashBeingArchived
        $script:RecoveryPath = $script:WorkingDir
        Show-ManagerError $Detail
        if ($script:WordClosed) { End-EditSession -PreserveWorkingCopy -Message "Error" }
        return
    }

    $script:LastArchivedHash = $script:HashBeingArchived
    $script:LastFailedHash = $null
    $RetryButton.Visible = $false
    $ErrorPanel.Visible = $false
    $LatestValue.Text = Format-ArchiveDate (Get-LatestArchive)
    Update-PublicReview
    Set-Status -Text "Version saved"
    $WebsiteOutcome = @($Output -split "`n" | Where-Object {$_ -match 'Website copy|Updated website public resume'}) | Select-Object -Last 1
    if($WebsiteOutcome){$ReviewValue.Text += "`n" + $(if($WebsiteOutcome -match '^Updated website'){'Website copy updated.'}else{$WebsiteOutcome.Trim()})}

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
        $script:SessionKind = $Kind.ToLower()
        $ErrorPanel.Visible = $false
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
        $Check = & $PythonExe $VariantScript --check-source $script:WorkingPath --config $ConfigPath 2>&1
        if($LASTEXITCODE -ne 0){throw ($Check -join "`n")}
        $script:WordDocument = Open-ResumeWordDocument $script:Word $script:WorkingPath

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
        Show-ManagerError $_.Exception.Message
    }
}

$EditPrivateButton.Add_Click({ Start-EditSession -Kind "Private" })
$EditPublicButton.Add_Click({ Start-EditSession -Kind "Public" })
$OpenCurrentButton.Add_Click({ Start-Process -FilePath "explorer.exe" -ArgumentList ('"{0}"' -f $CurrentDir) })
$OpenHistoryButton.Add_Click({ Show-HistoryDialog $Form })
$SettingsButton.Add_Click({
    if($script:SessionActive -or $null -ne $script:ArchiveProcess){return}
    if(Show-SettingsDialog $ConfigPath $Form){
        . (Join-Path $PSScriptRoot 'resume_settings.ps1') -ConfigPath $ConfigPath
        foreach($Name in @('Root','CurrentDir','UpdatesDir','PrivateName','PublicName','CurrentPrivate','CurrentPublic','CurrentPrivateName','CurrentPublicName','PythonExe','VariantScript','Settings','WebsitePublic')){Set-Variable -Name $Name -Scope Script -Value (Get-Variable -Name $Name -ValueOnly)}
        $ErrorPanel.Visible=$false; $ErrorDetail.Text=''; $script:LastErrorDetail=''
        $LatestValue.Text = Format-ArchiveDate (Get-LatestArchive); Update-PublicReview; Set-Status 'Ready'
    }
})
$PreviewPrivateButton.Add_Click({try{$Latest=@(Get-ArchiveDirectories $UpdatesDir)|Select-Object -First 1; if(-not $Latest){throw 'No archived PDFs yet.'}; Open-LocalFile (Join-Path $Latest.FullName ([IO.Path]::ChangeExtension($PrivateName,'.pdf')))}catch{Show-ManagerError $_.Exception.Message}})
$PreviewPublicButton.Add_Click({try{$Latest=@(Get-ArchiveDirectories $UpdatesDir)|Select-Object -First 1; if(-not $Latest){throw 'No archived PDFs yet.'}; Open-LocalFile (Join-Path $Latest.FullName ([IO.Path]::ChangeExtension($PublicName,'.pdf')))}catch{Show-ManagerError $_.Exception.Message}})
$OpenRecoveryButton.Add_Click({if(Test-Path -LiteralPath $script:RecoveryPath){Start-Process explorer.exe -ArgumentList ('"{0}"' -f $script:RecoveryPath)}})
$CopyDiagnosticsButton.Add_Click({$Text = Get-RedactedDiagnostics $script:LastErrorDetail $Settings; if ($Text) {[Windows.Forms.Clipboard]::SetText($Text)}})
$RetryButton.Add_Click({
    $script:LastFailedHash = $null
    $script:PendingChange = $true
    $script:LastFileEvent = [datetime]::UtcNow.AddSeconds(-3)
    $RetryButton.Visible = $false
    $ErrorPanel.Visible = $false
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
    catch {
        $script:PendingChange = $false
        Show-ManagerError $_.Exception.Message
        if($script:WordClosed){End-EditSession -PreserveWorkingCopy -Message 'Error'}
        return
    }

    if (-not $script:PendingChange) {
        if(-not $script:WordClosed -and -not $ErrorPanel.Visible){try{if(-not $script:WordDocument.Saved){Set-Status "Editing $script:SessionKind - unsaved changes in Word"}}catch{}}
        return
    }
    if (([datetime]::UtcNow - $script:LastFileEvent).TotalSeconds -lt 1.5) { return }

    try {
        Start-ArchiveProcess
    }
    catch {
        $Message = $_.Exception.Message
        $Inner = $_.Exception
        while($Inner.InnerException){$Inner=$Inner.InnerException}
        $Sharing = $Inner -is [IO.IOException] -and (($Inner.HResult -band 65535) -in @(32,33))
        $RacingPackage = $Message -match 'BadZipFile|not a zip file' -and ([datetime]::UtcNow-$script:LastFileEvent).TotalSeconds -lt 10
        if($Sharing -or $RacingPackage -or $Message -eq 'Word is still saving; waiting for a stable saved package.') { Set-Status 'Waiting for Word' }
        else {
            $script:PendingChange=$false; $script:RecoveryPath=$script:WorkingDir
            Show-ManagerError $Message
            if($script:WordClosed){End-EditSession -PreserveWorkingCopy -Message 'Error'}
        }
    }
})
$Timer.Start()
Update-PublicReview

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
$Form.Add_KeyDown({
    param($Sender,$Event)
    if($Event.Control -and $Event.KeyCode -eq 'P' -and $EditPrivateButton.Enabled){$EditPrivateButton.PerformClick();$Event.SuppressKeyPress=$true}
    elseif($Event.Control -and $Event.KeyCode -eq 'U' -and $EditPublicButton.Enabled){$EditPublicButton.PerformClick();$Event.SuppressKeyPress=$true}
    elseif($Event.Control -and $Event.KeyCode -eq 'H'){$OpenHistoryButton.PerformClick();$Event.SuppressKeyPress=$true}
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
