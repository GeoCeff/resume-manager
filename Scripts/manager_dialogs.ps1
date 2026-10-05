# Native dialogs share the existing workflow; personal state never belongs in the checkout.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function New-UiLabel([string]$Text) {
    $Label = New-Object Windows.Forms.Label
    $Label.Text = $Text; $Label.AutoSize = $true
    $Label.Margin = New-Object Windows.Forms.Padding(4,8,4,8)
    return $Label
}

function New-UiButton([string]$Text) {
    $Button = New-Object Windows.Forms.Button
    $Button.Text = $Text; $Button.AutoSize = $true; $Button.MinimumSize = New-Object Drawing.Size(120,34)
    $Button.Margin = New-Object Windows.Forms.Padding(4)
    $Button.AccessibleName = $Text.Replace('&','')
    return $Button
}

function Get-ArchiveDirectories([string]$Directory) {
    Get-ChildItem -LiteralPath $Directory -Directory -ErrorAction SilentlyContinue |
        Where-Object {$_.Name -match '^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}(-\d+)?$'} |
        Sort-Object @{Expression={$_.Name.Substring(0,19)};Descending=$true},@{Expression={if($_.Name.Length -gt 19){[int]$_.Name.Substring(20)}else{1}};Descending=$true}
}

function Format-ArchiveDate([string]$Name) {
    if ($Name -notmatch '^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}') { return $Name }
    $Date = [datetime]::ParseExact($Name.Substring(0,19),'yyyy-MM-dd_HH-mm-ss',[Globalization.CultureInfo]::InvariantCulture)
    $Suffix = if($Name.Length -gt 19){' (version ' + $Name.Substring(20) + ')'}else{''}
    return $Date.ToString('MMM d, yyyy - h:mm:ss tt') + $Suffix
}

function Open-LocalFile([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw 'The selected file is missing. Open the version folder to inspect its contents.' }
    Start-Process -FilePath $Path
}

function Confirm-ResumeReplacement($Owner,[string]$Message,[string]$Title) {
    return [Windows.Forms.MessageBox]::Show($Owner,$Message,$Title,'YesNo','Question') -eq 'Yes'
}

function Get-RedactedDiagnostics([string]$Detail,$Configuration) {
    foreach ($Value in @($Configuration.PrivateOnly) + @($Configuration.PrivateContact,$Configuration.PublicContact,$Configuration.ContactAnchor,$env:USERNAME)) {
        if ($Value) { $Detail = [regex]::Replace($Detail,[regex]::Escape([string]$Value),'[redacted]','IgnoreCase') }
    }
    $Detail = [regex]::Replace($Detail,'(?i)[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}','[email]')
    $Detail = [regex]::Replace($Detail,'(?i)(?<![a-z0-9])[a-z]:[\\/][^\r\n"<>]*','[local path]')
    $Detail = [regex]::Replace($Detail,'\\\\[^\r\n"<>]+','[network path]')
    return $Detail
}

function Show-HistoryDialog($Owner) {
    $Dialog = New-Object Windows.Forms.Form; $Dialog.Text = 'Resume Manager - Version History'
    $Dialog.Size = New-Object Drawing.Size(780,460); $Dialog.MinimumSize = New-Object Drawing.Size(700,400)
    $Dialog.Font = $Owner.Font; $Dialog.StartPosition = 'CenterParent'; $Dialog.AutoScaleMode = 'Font'
    $Layout = New-Object Windows.Forms.TableLayoutPanel; $Layout.Dock = 'Fill'; $Layout.Padding = New-Object Windows.Forms.Padding(14); $Layout.ColumnCount = 1; $Layout.RowCount = 3
    [void]$Layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('AutoSize')))
    [void]$Layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent',100)))
    [void]$Layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('AutoSize')))
    $Dialog.Controls.Add($Layout)
    $Message = New-UiLabel 'Restore creates a new version using today''s contact settings. Archives are never modified.'
    $Message.MaximumSize = New-Object Drawing.Size(700,0); $Layout.Controls.Add($Message,0,0)
    $List = New-Object Windows.Forms.ListView; $List.Dock = 'Fill'; $List.View = 'Details'; $List.FullRowSelect = $true; $List.MultiSelect = $false; $List.AccessibleName = 'Archived resume versions'
    [void]$List.Columns.Add('Saved',310); [void]$List.Columns.Add('Files',150); [void]$List.Columns.Add('Folder',240)
    foreach($Archive in @(Get-ArchiveDirectories $UpdatesDir)) {
        $Expected = @($PrivateName,$PublicName,[IO.Path]::ChangeExtension($PrivateName,'.pdf'),[IO.Path]::ChangeExtension($PublicName,'.pdf'))
        $Present = @($Expected | Where-Object {Test-Path -LiteralPath (Join-Path $Archive.FullName $_) -PathType Leaf}).Count
        $Item = New-Object Windows.Forms.ListViewItem((Format-ArchiveDate $Archive.Name))
        [void]$Item.SubItems.Add($(if($Present -eq 4){'Complete'}else{"Incomplete ($Present/4)"})); [void]$Item.SubItems.Add($Archive.Name)
        $Item.Tag = $Archive.FullName; [void]$List.Items.Add($Item)
        if($Present -ne 4){$Item.ForeColor=[Drawing.SystemColors]::HotTrack}
    }
    $Layout.Controls.Add($List,0,1)
    $Actions = New-Object Windows.Forms.FlowLayoutPanel; $Actions.AutoSize = $true; $Actions.Dock = 'Fill'
    $Public = New-UiButton 'Open Public PDF'; $Private = New-UiButton 'Open Private PDF'; $Folder = New-UiButton 'Open Folder'; $Restore = New-UiButton 'Restore as New Version'
    $Restore.Enabled = -not $script:SessionActive -and $null -eq $script:ArchiveProcess
    $Actions.Controls.AddRange(@($Public,$Private,$Folder,$Restore)); $Layout.Controls.Add($Actions,0,2)
    $Action = {
        param([string]$Kind)
        try {
            if ($List.SelectedItems.Count -ne 1) { throw 'Select a version first.' }
            $Selected = [string]$List.SelectedItems[0].Tag
            switch ($Kind) {
                'Public' { Open-LocalFile (Join-Path $Selected ([IO.Path]::ChangeExtension($PublicName,'.pdf'))) }
                'Private' { Open-LocalFile (Join-Path $Selected ([IO.Path]::ChangeExtension($PrivateName,'.pdf'))) }
                'Folder' { Start-Process explorer.exe -ArgumentList ('"{0}"' -f $Selected) }
                'Restore' {
                    if($script:SessionActive -or $null -ne $script:ArchiveProcess){throw 'Close the editing session before restoring.'}
                    $Source = Join-Path $Selected $PublicName
                    if(-not (Test-Path -LiteralPath $Source)){throw 'Public DOCX missing; this version cannot be restored.'}
                    if(-not (Confirm-ResumeReplacement $Dialog ('Restore ' + $List.SelectedItems[0].Text + ' as a new version? Current files will be preserved in Recovery first.') 'Confirm restore')){return}
                    [void](Backup-CurrentPair $Root)
                    Invoke-WorkflowInteractive $Source $ConfigPath $Dialog
                    $LatestValue.Text = Format-ArchiveDate (Get-LatestArchive)
                    Update-PublicReview
                    $ErrorPanel.Visible = $false
                    Set-Status 'Version saved'
                    $Dialog.Close()
                }
            }
        } catch { $Message.Text = $_.Exception.Message }
    }
    $Public.Add_Click({ & $Action 'Public' }); $Private.Add_Click({ & $Action 'Private' }); $Folder.Add_Click({ & $Action 'Folder' }); $Restore.Add_Click({ & $Action 'Restore' })
    try {[void]$Dialog.ShowDialog($Owner)}finally{$Dialog.Dispose()}
}

function Invoke-ResumePython([string]$Python, [string[]]$Arguments) {
    $Previous = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        $Result = & $Python @Arguments 2>&1
        if ($LASTEXITCODE -ne 0) { throw ($Result -join "`n") }
        return $Result
    } finally { $ErrorActionPreference = $Previous }
}

function Test-Prerequisites([string]$Python) {
    try { $Resolved = (Get-Command ([Environment]::ExpandEnvironmentVariables($Python)) -ErrorAction Stop).Source }
    catch { throw 'Python is not available. Select a working Python executable in Advanced settings.' }
    try { [void](Invoke-ResumePython $Resolved @('-c','import sys; assert sys.version_info >= (3,10), "Python 3.10 or newer required"; import lxml, pypdf')) }
    catch { throw "Python requirements are unavailable. Install requirements manually using python -m pip install -r requirements.txt.`n$($_.Exception.Message)" }
    if (-not [Type]::GetTypeFromProgID('Word.Application')) { throw 'Microsoft Word desktop is required. Install or repair Word; web Word is not supported.' }
    return $Resolved
}

function Get-StartupIssue([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return [pscustomobject]@{Kind='Setup';Message='Choose a resume to get started.';Details=''} }
    try {
        $Loaded = & { . (Join-Path $PSScriptRoot 'resume_settings.ps1') -ConfigPath $Path; [pscustomobject]@{Private=$CurrentPrivate;Public=$CurrentPublic;Python=$PythonExe;Configuration=$Settings} }
    } catch { return [pscustomobject]@{Kind='Settings';Message='Your saved settings need attention. Existing resume files have not been changed.';Details=$_.Exception.Message} }
    try { $Python = Test-Prerequisites $Loaded.Python }
    catch { return [pscustomobject]@{Kind='Prerequisites';Message='Word or Python is not ready. Fix the requirement and retry; your saved settings are retained.';Details=$_.Exception.Message} }
    try { [void](Invoke-ResumePython $Python @((Join-Path $PSScriptRoot 'update_resume_projects.py'),'--validate-settings','--config',$Path)) }
    catch { return [pscustomobject]@{Kind='Settings';Message='Review the saved contact and privacy settings. Nothing has been reset.';Details=$_.Exception.Message} }
    if (-not (Test-Path -LiteralPath $Loaded.Private -PathType Leaf) -or -not (Test-Path -LiteralPath $Loaded.Public -PathType Leaf)) {
        return [pscustomobject]@{Kind='Files';Message='Current resume files are missing. Locate the original data folder in Settings, or explicitly choose a resume to recover from.';Details='The private/public Current DOCX pair was not found.'}
    }
    return [pscustomobject]@{Kind='Ready';Python=$Python;Message='Ready';Details=''}
}

function Show-StartupRecovery($Issue) {
    $Dialog = New-Object Windows.Forms.Form; $Dialog.Text='Resume Manager - Startup help'; $Dialog.Size=New-Object Drawing.Size(660,370)
    $Dialog.StartPosition='CenterScreen'; $Dialog.Font=New-Object Drawing.Font('Segoe UI',10)
    $Layout=New-Object Windows.Forms.FlowLayoutPanel; $Layout.Dock='Fill'; $Layout.Padding=New-Object Windows.Forms.Padding(18); $Layout.FlowDirection='TopDown'; $Layout.AutoScroll=$true; $Dialog.Controls.Add($Layout)
    $Message=New-UiLabel $Issue.Message; $Message.MaximumSize=New-Object Drawing.Size(580,0); $Layout.Controls.Add($Message)
    $Details=New-Object Windows.Forms.TextBox; $Details.Multiline=$true; $Details.ReadOnly=$true; $Details.ScrollBars='Vertical'; $Details.Size=New-Object Drawing.Size(580,90); $Details.Text=$Issue.Details; $Details.Visible=$false
    $More=New-Object Windows.Forms.CheckBox; $More.Text='Show technical details'; $More.AutoSize=$true; $More.Add_CheckedChanged({$Details.Visible=$More.Checked}); $Layout.Controls.Add($More); $Layout.Controls.Add($Details)
    $Actions=New-Object Windows.Forms.FlowLayoutPanel; $Actions.AutoSize=$true; $Layout.Controls.Add($Actions)
    $Choice=@{Value='Cancel'}
    foreach($Name in @('Retry','Edit settings','Cancel')){$Button=New-UiButton $Name; $Button.Tag=$Name; $Button.Add_Click({$Choice.Value=$this.Tag; $Dialog.Close()}); $Actions.Controls.Add($Button)}
    try{[void]$Dialog.ShowDialog();return $Choice.Value}finally{$Dialog.Dispose()}
}

function Test-PublicContactField([string]$Value,[array]$NeverPublic) {
    $Decoded=[uri]::UnescapeDataString($Value).ToLowerInvariant()
    foreach($Private in $NeverPublic){
        if(-not $Private){continue}; $Digits=[regex]::Replace($Private,'\D','')
        if($Decoded.Contains($Private.ToLowerInvariant()) -or ($Digits.Length -ge 7 -and $Private -notmatch '[a-zA-Z]' -and ([regex]::Replace($Decoded,'\D','')).Contains($Digits))){return $false}
    }
    return $true
}

function New-ContactPolicy([string[]]$Contacts,[string[]]$PublicFields,[string[]]$NeverPublic) {
    if(-not $PublicFields){throw 'Choose at least one contact detail for the public resume.'}
    foreach($Value in $PublicFields){if($Contacts -notcontains $Value -or -not (Test-PublicContactField $Value $NeverPublic)){throw 'A never-public contact detail cannot be included in the public resume.'}}
    $PrivateValues=@(@($NeverPublic)+@($Contacts | Where-Object {$PublicFields -notcontains $_}) | Where-Object {$_} | Select-Object -Unique)
    if(-not $PrivateValues){throw 'Keep at least one contact detail private, or add a never-public value in Advanced settings.'}
    return [pscustomobject]@{PrivateContact=($Contacts -join ' | ');PublicContact=($PublicFields -join ' | ');ContactAnchor=$PublicFields[0];PrivateOnly=$PrivateValues}
}

function Save-LocalSettings([string]$Path, $Value) {
    $Code = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot)).TrimEnd('\') + '\'
    $Path = [IO.Path]::GetFullPath($Path)
    if ($Path.StartsWith($Code,[StringComparison]::OrdinalIgnoreCase)) { throw 'Settings must be outside the program repository.' }
    $Directory = Split-Path -Parent $Path
    [void][IO.Directory]::CreateDirectory($Directory)
    $Temporary = Join-Path $Directory ([guid]::NewGuid().ToString('N') + '.tmp')
    try {
        [IO.File]::WriteAllText($Temporary, ($Value | ConvertTo-Json -Depth 100))
        if (Test-Path -LiteralPath $Path) { [IO.File]::Replace($Temporary,$Path,[NullString]::Value) }
        else { [IO.File]::Move($Temporary,$Path) }
    }
    finally { if (Test-Path -LiteralPath $Temporary) { [IO.File]::Delete($Temporary) } }
}

function Invoke-WorkflowInteractive([string]$Source, [string]$Configuration, $Owner) {
    $JobDir = Join-Path $env:TEMP ('ResumeManagerSetup-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($JobDir)
    $ProgressFile = Join-Path $JobDir 'progress.json'
    $Workflow = (Join-Path $PSScriptRoot 'resume_update_workflow.ps1').Replace("'","''")
    $Command = "try { & '$Workflow' -SourcePath '$($Source.Replace("'","''"))' -ConfigPath '$($Configuration.Replace("'","''"))' -ProgressPath '$($ProgressFile.Replace("'","''"))'; exit 0 } catch { [Console]::Error.WriteLine(`$_); exit 1 }"
    $Process = Start-Process -FilePath "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-EncodedCommand',[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Command))) -WindowStyle Hidden -RedirectStandardOutput (Join-Path $JobDir 'output.txt') -RedirectStandardError (Join-Path $JobDir 'error.txt') -PassThru
    $null = $Process.Handle
    $OriginalTitle = $Owner.Text
    try {
        $Owner.Enabled = $false
        while (-not $Process.HasExited) {
            if (Test-Path -LiteralPath $ProgressFile) {
                try { $Owner.Text = $OriginalTitle + ' - ' + ((Get-Content -LiteralPath $ProgressFile -Raw | ConvertFrom-Json).Phase) } catch {}
            }
            [Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 100; $Process.Refresh()
        }
        $Process.WaitForExit()
        if ($Process.ExitCode -ne 0) { throw ((Get-Content -LiteralPath (Join-Path $JobDir 'error.txt') -Raw) + "`nRecovery details: $JobDir") }
    }
    finally { $Owner.Text = $OriginalTitle; $Owner.Enabled = $true; $Process.Dispose() }
}

function Show-SettingsDialog([string]$Path, $Owner) {
    $Existing=$null; $ReadError=''
    if(Test-Path -LiteralPath $Path){try{$Existing=Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json}catch{$ReadError='Saved settings could not be read. Review these values before replacing them.'}}
    $SetupState=@{Saved=$false;Updating=$false;Contacts=@();Protected=@();HardPrivate=@();ManualPrivateOnlyDirty=$false;Details=''}
    if($Existing){
        $PublicFields=@($Existing.PublicContact -split '\s*\|\s*')
        $SetupState.HardPrivate=@(@($Existing.PrivateOnly)+@($Existing.PrivateContact -split '\s*\|\s*' | Where-Object {$PublicFields -notcontains $_}) | Where-Object {$_} | Select-Object -Unique)
        $SetupState.Protected=$SetupState.HardPrivate
    }
    $Dialog=New-Object Windows.Forms.Form; $Dialog.Text='Resume Manager - Local Setup'
    $Dialog.Size=New-Object Drawing.Size(740,740); $Dialog.MinimumSize=New-Object Drawing.Size(680,600)
    $Dialog.StartPosition='CenterScreen'; $Dialog.Font=New-Object Drawing.Font('Segoe UI',10); $Dialog.AutoScaleMode='Font'
    $Layout=New-Object Windows.Forms.TableLayoutPanel; $Layout.Dock='Fill'; $Layout.Padding=New-Object Windows.Forms.Padding(18); $Layout.ColumnCount=1; $Layout.AutoScroll=$true
    $HostLayout=New-Object Windows.Forms.TableLayoutPanel; $HostLayout.Dock='Fill'; $HostLayout.ColumnCount=1; $HostLayout.RowCount=2
    [void]$HostLayout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent',100))); [void]$HostLayout.RowStyles.Add((New-Object Windows.Forms.RowStyle('AutoSize')))
    $Dialog.Controls.Add($HostLayout); $HostLayout.Controls.Add($Layout,0,0)
    $Intro=New-UiLabel $(if($Existing){'Your saved settings are loaded. Choose a Word resume only to import different content.'}else{'Choose a resume, decide what stays public, then review both contact lines.'})
    $Intro.MaximumSize=New-Object Drawing.Size(640,0); $Layout.Controls.Add($Intro)
    $SourceRow=New-Object Windows.Forms.TableLayoutPanel; $SourceRow.AutoSize=$true; $SourceRow.Dock='Fill'; $SourceRow.ColumnCount=3
    [void]$SourceRow.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('AutoSize')))
    [void]$SourceRow.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent',100)))
    [void]$SourceRow.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('AutoSize')))
    $SourceBox=New-Object Windows.Forms.TextBox; $SourceBox.Dock='Fill'; $SourceBox.AccessibleName='Starting DOCX (optional for existing data)'
    $SourceBrowse=New-UiButton 'Choose Word resume...'
    $SourceRow.Controls.Add((New-UiLabel 'Word resume'),0,0); $SourceRow.Controls.Add($SourceBox,1,0); $SourceRow.Controls.Add($SourceBrowse,2,0); $Layout.Controls.Add($SourceRow)
    $ReadContacts=New-UiButton 'Read contact details'; $Layout.Controls.Add($ReadContacts)
    $ContactHelp=New-UiLabel 'Check only the details allowed in public. Unchecked details stay private.'
    $ContactHelp.MaximumSize=New-Object Drawing.Size(640,0); $Layout.Controls.Add($ContactHelp)
    $ContactList=New-Object Windows.Forms.CheckedListBox; $ContactList.CheckOnClick=$true; $ContactList.Dock='Fill'; $ContactList.Height=110; $ContactList.HorizontalScrollbar=$true; $ContactList.AccessibleName='Contact details allowed in public'
    $Layout.Controls.Add($ContactList)
    $Storage=New-UiLabel ''; $Storage.MaximumSize=New-Object Drawing.Size(640,0); $Layout.Controls.Add($Storage)
    $AdvancedToggle=New-UiButton '&Advanced settings'; $Layout.Controls.Add($AdvancedToggle)
    $Advanced=New-Object Windows.Forms.TableLayoutPanel; $Advanced.AutoSize=$true; $Advanced.Dock='Fill'; $Advanced.ColumnCount=3; $Advanced.Visible=$false
    [void]$Advanced.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('AutoSize')))
    [void]$Advanced.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent',100)))
    [void]$Advanced.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('AutoSize')))
    $Layout.Controls.Add($Advanced)
    $Fields=@{Source=$SourceBox}; $Row=0
    $Definitions=@(
        @('DataRoot','Resume data folder',(Join-Path $env:LOCALAPPDATA 'ResumeManager\Data')),
        @('PrivateContact','Private contact line',''),@('PublicContact','Public contact line',''),
        @('ContactAnchor','Shared contact anchor',''),@('PrivateOnly','Never-public values (one per line)',''),
        @('PythonExe','Python executable','python.exe'),@('WebsitePublicPath','Website public DOCX (optional)',''),
        @('CurrentPrivateName','Current private filename','resume-private.docx'),@('CurrentPublicName','Current public filename','resume-public.docx'),
        @('ArchivePrivateName','Archived private filename','resume-private.docx'),@('ArchivePublicName','Archived public filename','resume-public.docx'))
    foreach($Definition in $Definitions){
        $Key=$Definition[0]; $Input=New-Object Windows.Forms.TextBox; $Input.Dock='Fill'; $Input.AccessibleName=$Definition[1]
        $Input.Text=if($Existing -and $Existing.$Key){if($Key -eq 'PrivateOnly'){$Existing.$Key -join [Environment]::NewLine}else{[string]$Existing.$Key}}else{$Definition[2]}
        if($Key -eq 'PrivateOnly'){$Input.Multiline=$true;$Input.Height=60;$Input.ScrollBars='Vertical'}
        $Fields[$Key]=$Input; $Advanced.Controls.Add((New-UiLabel $Definition[1]),0,$Row); $Advanced.Controls.Add($Input,1,$Row)
        if($Key -in @('DataRoot','PythonExe','WebsitePublicPath')){
            $Browse=New-UiButton 'Browse...'; $Browse.Tag=$Key
            $Browse.Add_Click({
                $Key=$this.Tag
                if($Key -eq 'DataRoot'){$Picker=New-Object Windows.Forms.FolderBrowserDialog}
                elseif($Key -eq 'WebsitePublicPath'){$Picker=New-Object Windows.Forms.SaveFileDialog;$Picker.Filter='Word document (*.docx)|*.docx'}
                else{$Picker=New-Object Windows.Forms.OpenFileDialog;$Picker.Filter='Python executable (*.exe)|*.exe'}
                try{if($Picker.ShowDialog($Dialog) -eq 'OK'){$Fields[$Key].Text=if($Key -eq 'DataRoot'){$Picker.SelectedPath}else{$Picker.FileName}}}finally{$Picker.Dispose()}
            })
            $Advanced.Controls.Add($Browse,2,$Row)
        }
        $Row++
    }
    $Preview=New-UiLabel ''; $Preview.MaximumSize=New-Object Drawing.Size(640,0); $Preview.AccessibleName='Private and public contact previews'; $Layout.Controls.Add($Preview)
    $Reviewed=New-Object Windows.Forms.CheckBox; $Reviewed.Text='I reviewed the public contact details'; $Reviewed.AutoSize=$true; $Layout.Controls.Add($Reviewed)
    $ErrorLabel=New-UiLabel $ReadError; $ErrorLabel.MaximumSize=New-Object Drawing.Size(640,0); $ErrorLabel.ForeColor=[Drawing.SystemColors]::HotTrack; $ErrorLabel.AccessibleName='Setup validation message'; $Layout.Controls.Add($ErrorLabel)
    $ErrorProvider=New-Object Windows.Forms.ErrorProvider; $ErrorProvider.ContainerControl=$Dialog; $ErrorProvider.BlinkStyle='NeverBlink'
    $More=New-Object Windows.Forms.CheckBox; $More.Text='Show technical details'; $More.AutoSize=$true; $More.Visible=$false; $Layout.Controls.Add($More)
    $Details=New-Object Windows.Forms.TextBox; $Details.Multiline=$true; $Details.ReadOnly=$true; $Details.ScrollBars='Vertical'; $Details.Dock='Fill'; $Details.Height=85; $Details.Visible=$false; $Layout.Controls.Add($Details)
    $Copy=New-UiButton 'Copy redacted details'; $Copy.Visible=$false; $Layout.Controls.Add($Copy)
    $Actions=New-Object Windows.Forms.FlowLayoutPanel; $Actions.AutoSize=$true; $Actions.Dock='Fill'; $Actions.Padding=New-Object Windows.Forms.Padding(18,4,18,10); $HostLayout.Controls.Add($Actions,0,1)
    $Save=New-UiButton '&Save and finish'; $Cancel=New-UiButton 'Cancel'; $Cancel.DialogResult='Cancel'; $Actions.Controls.AddRange(@($Save,$Cancel))
    $Dialog.CancelButton=$Cancel; $Dialog.AcceptButton=$Save

    $ShowError={
        param([string]$Message,$Control,[string]$Raw='')
        $ErrorLabel.Text=$Message; $SetupState.Details=$Raw; $Details.Text=$Raw; $More.Visible=[bool]$Raw
        if($Control){$ErrorProvider.SetError($Control,$Message);if($Advanced.Contains($Control)){$Advanced.Visible=$true;$AdvancedToggle.Text='Hide advanced settings'};$Dialog.ScrollControlIntoView($Layout);$Control.Focus() | Out-Null}
    }
    $UpdatePreview={
        $Preview.Text=if($Fields.PrivateContact.Text){'Private: '+$Fields.PrivateContact.Text+[Environment]::NewLine+'Public: '+$Fields.PublicContact.Text}else{'Choose a resume to see the private and public contact previews.'}
        $Storage.Text='Versions saved locally in: '+$Fields.DataRoot.Text+[Environment]::NewLine+'Change the folder under Advanced settings.'
    }
    $FillContacts={
        param([string[]]$Values,[string[]]$PublicValues)
        $SetupState.Updating=$true
        try{
            $SetupState.Contacts=@($Values | Where-Object {$_} | Select-Object -Unique); $ContactList.Items.Clear()
            foreach($Value in $SetupState.Contacts){
                $Allowed=Test-PublicContactField $Value $SetupState.Protected
                $Label=if($Allowed){$Value}else{$Value+' (always private)'}
                [void]$ContactList.Items.Add($Label,($Allowed -and $PublicValues -contains $Value))
            }
        }finally{$SetupState.Updating=$false}
    }
    $SyncPolicy={
        param($Change)
        $Chosen=@()
        for($Index=0;$Index -lt $ContactList.Items.Count;$Index++){
            $Checked=if($Change -and $Change.Index -eq $Index){$Change.NewValue -eq 'Checked'}else{$ContactList.GetItemChecked($Index)}
            if($Checked){$Chosen+=$SetupState.Contacts[$Index]}
        }
        $SetupState.Updating=$true
        try{
            $Fields.PrivateContact.Text=$SetupState.Contacts -join ' | '
            $Fields.PublicContact.Text=$Chosen -join ' | '
            $Fields.ContactAnchor.Text=if($Chosen){$Chosen[0]}else{''}
            $Fields.PrivateOnly.Text=(@(@($SetupState.Protected)+@($SetupState.Contacts | Where-Object {$Chosen -notcontains $_}) | Where-Object {$_} | Select-Object -Unique) -join [Environment]::NewLine)
            $Reviewed.Checked=$false
            & $UpdatePreview
            if($Chosen){try{[void](New-ContactPolicy $SetupState.Contacts $Chosen $SetupState.Protected);$ErrorLabel.Text=''}catch{$ErrorLabel.Text=$_.Exception.Message}}
        }finally{$SetupState.Updating=$false}
    }
    $ContactList.Add_ItemCheck({
        param($Sender,$Change)
        if($SetupState.Updating){return}
        if($Change.NewValue -eq 'Checked' -and -not (Test-PublicContactField $SetupState.Contacts[$Change.Index] $SetupState.Protected)){
            $Change.NewValue='Unchecked'; & $ShowError 'This detail is configured as always private.' $ContactList; return
        }
        & $SyncPolicy $Change
    })
    $ReadContacts.Add_Click({
        $ErrorProvider.Clear(); $ErrorLabel.Text=''
        try{
            if(-not (Test-Path -LiteralPath $SourceBox.Text -PathType Leaf)){& $ShowError 'Choose an existing Word resume first.' $SourceBox;return}
            $Python=Test-Prerequisites $Fields.PythonExe.Text
            $Arguments=@((Join-Path $PSScriptRoot 'update_resume_projects.py'),'--inspect-contact',$SourceBox.Text)
            if($Fields.ContactAnchor.Text){$Arguments+=@('--anchor',$Fields.ContactAnchor.Text)}
            $Inspection=((Invoke-ResumePython $Python $Arguments) -join [Environment]::NewLine) | ConvertFrom-Json
            $Values=@($Inspection.Fields)
            if($Existing){
                $Values=@(@($Existing.PrivateContact -split '\s*\|\s*' | Where-Object {$Values -contains $_ -or $SetupState.Protected -contains $_})+@($Values | Where-Object {($Existing.PrivateContact -split '\s*\|\s*') -notcontains $_}))
            }
            & $FillContacts $Values @($Fields.PublicContact.Text -split '\s*\|\s*')
            & $SyncPolicy $null
            $ContactHelp.Text='Review the detected details. Check only those allowed in public.'
        }catch{& $ShowError 'Contact details could not be read. Check Word/Python or use Advanced settings for manual contact configuration.' $SourceBox $_.Exception.Message}
    })
    $SourceBrowse.Add_Click({
        $Picker=New-Object Windows.Forms.OpenFileDialog;$Picker.Filter='Word document (*.docx)|*.docx'
        try{if($Picker.ShowDialog($Dialog) -eq 'OK'){$SourceBox.Text=$Picker.FileName;$ReadContacts.PerformClick()}}finally{$Picker.Dispose()}
    })
    $SourceBox.Add_TextChanged({$Reviewed.Checked=$false})
    $AdvancedToggle.Add_Click({$Advanced.Visible=-not $Advanced.Visible;$AdvancedToggle.Text=if($Advanced.Visible){'Hide advanced settings'}else{'Advanced settings'};$Layout.PerformLayout()})
    foreach($Key in $Fields.Keys){
        if($Key -eq 'Source'){continue}
        $Fields[$Key].Add_TextChanged({if(-not $SetupState.Updating){$Reviewed.Checked=$false;& $UpdatePreview}})
    }
    $RefreshManual={
        if($SetupState.ManualPrivateOnlyDirty){$SetupState.Protected=@(@($SetupState.HardPrivate)+@($Fields.PrivateOnly.Lines) | Where-Object {$_} | Select-Object -Unique);$SetupState.ManualPrivateOnlyDirty=$false}
        & $FillContacts @($Fields.PrivateContact.Text -split '\s*\|\s*') @($Fields.PublicContact.Text -split '\s*\|\s*')
    }
    $Fields.PrivateOnly.Add_TextChanged({if(-not $SetupState.Updating){$SetupState.ManualPrivateOnlyDirty=$true}})
    $Fields.PrivateContact.Add_Leave({& $RefreshManual}); $Fields.PublicContact.Add_Leave({& $RefreshManual}); $Fields.PrivateOnly.Add_Leave({& $RefreshManual})
    $More.Add_CheckedChanged({$Details.Visible=$More.Checked;$Copy.Visible=$More.Checked})
    $Copy.Add_Click({$Redacted=Get-RedactedDiagnostics $SetupState.Details ([pscustomobject]@{PrivateOnly=$Fields.PrivateOnly.Lines;PrivateContact=$Fields.PrivateContact.Text;PublicContact=$Fields.PublicContact.Text;ContactAnchor=$Fields.ContactAnchor.Text});if($Redacted){[Windows.Forms.Clipboard]::SetText($Redacted)}})
    $Dialog.Add_FormClosing({param($Sender,$EventArgs) if(-not $Dialog.Enabled){$EventArgs.Cancel=$true}})
    $Save.Add_Click({
        $CandidatePath=$null; $FocusControl=$SourceBox; $WorkflowCompleted=$false
        $ErrorProvider.Clear(); $ErrorLabel.Text=''; $More.Checked=$false; $More.Visible=$false
        try{
            if(-not $Fields.PrivateContact.Text){& $ShowError 'Choose a resume and read its contacts, or enter the private contact line in Advanced settings.' $SourceBox;return}
            if(-not $Fields.PublicContact.Text){& $ShowError 'Choose at least one contact detail for the public resume.' $ContactList;return}
            if(-not $Reviewed.Checked){& $ShowError 'Review both contact previews and tick the confirmation before saving.' $Reviewed;return}
            $Candidate=if($Existing){$Existing | ConvertTo-Json -Depth 100 | ConvertFrom-Json}else{[pscustomobject]@{}}
            foreach($Key in $Fields.Keys){
                if($Key -eq 'Source'){continue}
                $Value=if($Key -eq 'PrivateOnly'){@(@($SetupState.HardPrivate)+@($Fields[$Key].Lines | ForEach-Object {$_.Trim()}) | Where-Object {$_} | Select-Object -Unique)}else{$Fields[$Key].Text.Trim()}
                $Candidate | Add-Member -NotePropertyName $Key -NotePropertyValue $Value -Force
            }
            $FocusControl=$Fields.PythonExe; $Candidate.PythonExe=Test-Prerequisites $Candidate.PythonExe
            $CandidatePath=Join-Path (Split-Path -Parent $Path) ('candidate-'+[guid]::NewGuid().ToString('N')+'.json')
            Save-LocalSettings $CandidatePath $Candidate
            $FocusControl=$Fields.PublicContact
            [void](Invoke-ResumePython $Candidate.PythonExe @((Join-Path $PSScriptRoot 'update_resume_projects.py'),'--validate-settings','--config',$CandidatePath))
            $FocusControl=$Fields.DataRoot
            $Validated=& { . (Join-Path $PSScriptRoot 'resume_settings.ps1') -ConfigPath $CandidatePath; [pscustomobject]@{Private=$CurrentPrivate;Public=$CurrentPublic;Root=$Root} }
            $Source=$SourceBox.Text.Trim()
            $PrivacyChanged=$Existing -and @('PrivateContact','PublicContact','ContactAnchor','PrivateOnly' | Where-Object {($Candidate.$_ | ConvertTo-Json -Compress) -ne ($Existing.$_ | ConvertTo-Json -Compress)}).Count -gt 0
            $Initializing=$Source -or $PrivacyChanged -or -not (Test-Path -LiteralPath $Validated.Private) -or -not (Test-Path -LiteralPath $Validated.Public)
            if($Initializing -and -not $Source){
                if(Test-Path -LiteralPath $Validated.Public){$Source=$Validated.Public}
                else{& $ShowError 'Choose a Word resume to create or recover the missing files.' $SourceBox;return}
            }
            $CheckSource=if($Source){$Source}else{$Validated.Public}
            $FocusControl=$Fields.ContactAnchor
            [void](Invoke-ResumePython $Candidate.PythonExe @((Join-Path $PSScriptRoot 'update_resume_projects.py'),'--check-source',$CheckSource,'--config',$CandidatePath))
            if($Initializing){
                if((Test-Path -LiteralPath $Validated.Private) -or (Test-Path -LiteralPath $Validated.Public)){
                    if(-not (Confirm-ResumeReplacement $Dialog 'Create a new version using these contact settings? Current will be backed up first. Existing archives remain unchanged.' 'Confirm new version')){return}
                    [void](Backup-CurrentPair $Validated.Root)
                }
                $FocusControl=$SourceBox; Invoke-WorkflowInteractive $Source $CandidatePath $Dialog; $WorkflowCompleted=$true
            }
            Save-LocalSettings $Path $Candidate
            $SetupState.Saved=$true; $Dialog.DialogResult='OK'; $Dialog.Close()
        }catch{
            $Raw=$_.Exception.Message
            if($WorkflowCompleted){& $ShowError 'A new resume version was saved, but settings could not be saved. Cancelling does not undo that version. Recovery settings are retained; see details.' $Save ($Raw+[Environment]::NewLine+'Recovery settings: '+$CandidatePath);return}
            $Message=if($Raw -match 'PrivateOnly'){'Keep at least one detail private, or add a never-public value in Advanced settings.'}
                elseif($Raw -match 'PublicContact|private-only|Private contact details|never-public'){'A private contact detail is included in public output. Review the public contact line.'}
                elseif($Raw -match 'Python|lxml|pypdf|requirements'){'Python is not ready. Select its executable or repair the required libraries in Advanced settings.'}
                elseif($Raw -match 'Microsoft Word'){'Microsoft Word desktop is required. Install or repair Word, then retry.'}
                elseif($Raw -match 'ContactAnchor|Contact line|contact text|first six'){'Contact text could not be located in this resume. Check the shared contact anchor in Advanced settings.'}
                elseif($Raw -match 'outside|filenames|filename|Missing local setting'){'Review the folder and filename settings. Personal data must stay outside the program folder.'}
                else{'The operation could not finish. Your entered values and source are retained; open details to investigate.'}
            if($Raw -match 'ContactAnchor|Contact line|first six'){$FocusControl=$Fields.ContactAnchor}
            elseif($Raw -match 'PrivateOnly'){$FocusControl=$ContactList}
            elseif($Raw -match 'Website|website'){$FocusControl=$Fields.WebsitePublicPath}
            elseif($Raw -match 'filename'){
                $BadName=@('CurrentPrivateName','CurrentPublicName','ArchivePrivateName','ArchivePublicName' | Where-Object {$Fields[$_].Text -and ([IO.Path]::GetFileName($Fields[$_].Text) -ne $Fields[$_].Text -or $Fields[$_].Text -notmatch '(?i)\.docx$')}) | Select-Object -First 1
                $FocusControl=if($BadName){$Fields[$BadName]}elseif($Candidate.CurrentPrivateName -eq $Candidate.CurrentPublicName){$Fields.CurrentPublicName}else{$Fields.ArchivePublicName}
            }
            & $ShowError $Message $FocusControl $Raw
        }finally{if($CandidatePath -and (-not $WorkflowCompleted -or $SetupState.Saved) -and (Test-Path -LiteralPath $CandidatePath)){[IO.File]::Delete($CandidatePath)}}
    })
    if($Existing){& $FillContacts @($Existing.PrivateContact -split '\s*\|\s*') @($Existing.PublicContact -split '\s*\|\s*')}
    & $UpdatePreview
    try{if($Owner){[void]$Dialog.ShowDialog($Owner)}else{[void]$Dialog.ShowDialog()};return $SetupState.Saved}
    finally{$ErrorProvider.Dispose();$Dialog.Dispose()}
}

function Backup-CurrentPair([string]$DataRoot) {
    $Directory = Join-Path $DataRoot ('Recovery\before-replace-' + (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss') + '-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($Directory)
    Get-ChildItem -LiteralPath (Join-Path $DataRoot 'Current') -Filter '*.docx' -File | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $Directory }
    return $Directory
}
