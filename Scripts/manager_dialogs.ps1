# Native dialogs share the existing workflow; personal state never belongs in the checkout.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$script:ProgramVersion = '1.1.0'

function Test-ManagerBusy {
    return $script:SessionActive -or $null -ne $script:ArchiveProcess -or $script:OperationBusy
}

function Test-PathWithin([string]$Path,[string]$Parent) {
    $Full = [IO.Path]::GetFullPath($Path).TrimEnd('\')
    $Base = [IO.Path]::GetFullPath($Parent).TrimEnd('\')
    return $Full.Equals($Base,[StringComparison]::OrdinalIgnoreCase) -or $Full.StartsWith($Base+'\',[StringComparison]::OrdinalIgnoreCase)
}

function Assert-PlainPath([string]$Path) {
    $Full = [IO.Path]::GetFullPath($Path)
    for($Cursor=$Full;$Cursor;$Cursor=Split-Path -Parent $Cursor){
        if(Test-Path -LiteralPath $Cursor){
            if((Get-Item -LiteralPath $Cursor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Choose a plain local folder, not a junction or symbolic link.'}
        }
    }
    return $Full
}

function Assert-ProfileRoot([string]$DataRoot,[string]$SettingsPath) {
    $Full = Assert-PlainPath ([Environment]::ExpandEnvironmentVariables($DataRoot))
    $Code = Split-Path -Parent $PSScriptRoot
    if($Full.TrimEnd('\') -eq [IO.Path]::GetPathRoot($Full).TrimEnd('\') -or
       (Test-PathWithin $Full $Code) -or (Test-PathWithin $Code $Full) -or
       (Test-PathWithin $Full $env:WINDIR) -or $Full.TrimEnd('\') -eq $env:USERPROFILE.TrimEnd('\') -or
       $Full.TrimEnd('\') -eq $env:LOCALAPPDATA.TrimEnd('\')){throw 'Choose a dedicated data folder outside the program and Windows folders.'}
    foreach($Entry in @($script:ProfileEntries | Where-Object {$_})){
        if([IO.Path]::GetFullPath($Entry.SettingsPath) -eq [IO.Path]::GetFullPath($SettingsPath)){continue}
        if(-not (Test-Path -LiteralPath $Entry.SettingsPath)){throw 'A registered profile settings file is missing. Locate it before changing profile folders.'}
        $Other = Get-Content -LiteralPath $Entry.SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $OtherRoot = Assert-PlainPath ([Environment]::ExpandEnvironmentVariables($Other.DataRoot))
        if((Test-PathWithin $Full $OtherRoot) -or (Test-PathWithin $OtherRoot $Full)){throw 'Profile data folders must not overlap. Choose a separate folder.'}
    }
}

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
    $Message = New-UiLabel 'Select one version to label/restore, or two to compare body text. Archives are never modified.'
    $Message.MaximumSize = New-Object Drawing.Size(700,0); $Layout.Controls.Add($Message,0,0)
    $List = New-Object Windows.Forms.ListView; $List.Dock = 'Fill'; $List.View = 'Details'; $List.FullRowSelect = $true; $List.MultiSelect = $true; $List.AccessibleName = 'Archived resume versions'
    [void]$List.Columns.Add('Saved',260); [void]$List.Columns.Add('Label',200); [void]$List.Columns.Add('Files',110); [void]$List.Columns.Add('Folder',230)
    $Notes=Get-VersionNotes $Root
    foreach($Archive in @(Get-ArchiveDirectories $UpdatesDir)) {
        $Expected = @($PrivateName,$PublicName,[IO.Path]::ChangeExtension($PrivateName,'.pdf'),[IO.Path]::ChangeExtension($PublicName,'.pdf'))
        $Present = @($Expected | Where-Object {Test-Path -LiteralPath (Join-Path $Archive.FullName $_) -PathType Leaf}).Count
        $Item = New-Object Windows.Forms.ListViewItem((Format-ArchiveDate $Archive.Name))
        [void]$Item.SubItems.Add([string]$Notes.($Archive.Name)); [void]$Item.SubItems.Add($(if($Present -eq 4){'Complete'}else{"Incomplete ($Present/4)"})); [void]$Item.SubItems.Add($Archive.Name)
        $Item.Tag = $Archive.FullName; [void]$List.Items.Add($Item)
        if($Present -ne 4){$Item.ForeColor=[Drawing.SystemColors]::HotTrack}
    }
    $Layout.Controls.Add($List,0,1)
    $Actions = New-Object Windows.Forms.FlowLayoutPanel; $Actions.AutoSize = $true; $Actions.Dock = 'Fill'
    $Public = New-UiButton 'Open Public PDF'; $Private = New-UiButton 'Open Private PDF'; $Folder = New-UiButton 'Open Folder'; $Restore = New-UiButton 'Restore as New Version'
    $Restore.Enabled = -not (Test-ManagerBusy)
    $LabelButton=New-UiButton 'Label Version'; $CompareButton=New-UiButton 'Compare Two Versions'
    $Variant=New-Object Windows.Forms.ComboBox; $Variant.DropDownStyle='DropDownList'; $Variant.AccessibleName='Variant to compare'; [void]$Variant.Items.AddRange(@('Public','Private')); $Variant.SelectedIndex=0
    $Actions.Controls.AddRange(@($Public,$Private,$Folder,$Restore,$LabelButton,$Variant,$CompareButton)); $Layout.Controls.Add($Actions,0,2)
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
                    if(Test-ManagerBusy){throw 'Finish the current operation before restoring.'}
                    $Source = Join-Path $Selected $PublicName
                    if(-not (Test-Path -LiteralPath $Source)){throw 'Public DOCX missing; this version cannot be restored.'}
                    if(-not (Confirm-ResumeReplacement $Dialog ('Restore ' + $List.SelectedItems[0].Text + ' as a new version? Current files will be preserved in Recovery first.') 'Confirm restore')){return}
                    $script:OperationBusy=$true
                    try{[void](Backup-CurrentPair $Root); Invoke-WorkflowInteractive $Source $ConfigPath $Dialog}finally{$script:OperationBusy=$false}
                    $LatestValue.Text = Format-ArchiveDate (Get-LatestArchive)
                    Update-PublicReview
                    $ErrorPanel.Visible = $false
                    Set-Status ('Both versions saved at ' + (Get-Date -Format 'h:mm:ss tt'))
                    $Dialog.Close()
                }
            }
        } catch { $Message.Text = $_.Exception.Message }
    }
    $Public.Add_Click({ & $Action 'Public' }); $Private.Add_Click({ & $Action 'Private' }); $Folder.Add_Click({ & $Action 'Folder' }); $Restore.Add_Click({ & $Action 'Restore' })
    $LabelButton.Add_Click({
        try{
            if($List.SelectedItems.Count -ne 1){throw 'Select one version to label.'}
            $Item=$List.SelectedItems[0]; $Name=Split-Path -Leaf $Item.Tag
            $CurrentNotes=Get-VersionNotes $Root
            $Text=Show-TextEntry $Dialog 'Label Version' 'A short local note, such as Added internship. Blank removes the label.' ([string]$CurrentNotes.$Name)
            if($null -eq $Text){return}
            Set-VersionNote $Root $Name $Text
            $Notes=Get-VersionNotes $Root; $Item.SubItems[1].Text=$Text
        }catch{$Message.Text=$_.Exception.Message}
    })
    $CompareButton.Add_Click({
        try{
            if($List.SelectedItems.Count -ne 2){throw 'Select two versions (hold Ctrl while clicking), then choose Public or Private.'}
            $Name=if($Variant.SelectedItem -eq 'Private'){$PrivateName}else{$PublicName}
            $Paths=@($List.SelectedItems | ForEach-Object {Join-Path ([string]$_.Tag) $Name})
            $Text=((Invoke-ResumePython $PythonExe @($VariantScript,'--compare-text',$Paths[1],$Paths[0])) -join "`n") | ConvertFrom-Json
            Show-ReadOnlyText $Dialog 'Resume Manager - Text Comparison' ('Body text only, including tables. Not images, formatting, headers, or layout. Use PDF previews for visual comparison.' + "`r`n`r`n" + $Text)
        }catch{$Message.Text=$_.Exception.Message}
    })
    try {[void]$Dialog.ShowDialog($Owner)}finally{$Dialog.Dispose()}
}

function Show-TextEntry($Owner,[string]$Title,[string]$Prompt,[string]$Initial='') {
    $Dialog=New-Object Windows.Forms.Form; $Dialog.Text=$Title; $Dialog.Size=New-Object Drawing.Size(600,220); $Dialog.StartPosition='CenterParent'; $Dialog.Font=$Owner.Font
    $Layout=New-Object Windows.Forms.FlowLayoutPanel; $Layout.Dock='Fill'; $Layout.Padding=New-Object Windows.Forms.Padding(14); $Layout.FlowDirection='TopDown'; $Dialog.Controls.Add($Layout)
    $Help=New-UiLabel $Prompt; $Help.MaximumSize=New-Object Drawing.Size(540,0); $Layout.Controls.Add($Help)
    $Input=New-Object Windows.Forms.TextBox; $Input.Width=520; $Input.MaxLength=160; $Input.Text=$Initial; $Input.AccessibleName=$Title; $Layout.Controls.Add($Input)
    $Actions=New-Object Windows.Forms.FlowLayoutPanel; $Actions.AutoSize=$true; $Layout.Controls.Add($Actions)
    $Save=New-UiButton 'Save'; $Save.DialogResult='OK'; $Cancel=New-UiButton 'Cancel'; $Cancel.DialogResult='Cancel'; $Actions.Controls.AddRange(@($Save,$Cancel)); $Dialog.AcceptButton=$Save; $Dialog.CancelButton=$Cancel
    try{if($Dialog.ShowDialog($Owner) -eq 'OK'){return $Input.Text.Trim()};return $null}finally{$Dialog.Dispose()}
}

function Show-ReadOnlyText($Owner,[string]$Title,[string]$Text) {
    $Dialog=New-Object Windows.Forms.Form; $Dialog.Text=$Title; $Dialog.Size=New-Object Drawing.Size(820,540); $Dialog.MinimumSize=New-Object Drawing.Size(620,360); $Dialog.StartPosition='CenterParent'
    $Input=New-Object Windows.Forms.TextBox; $Input.Multiline=$true; $Input.ReadOnly=$true; $Input.ScrollBars='Both'; $Input.WordWrap=$false; $Input.Dock='Fill'; $Input.Font=New-Object Drawing.Font('Consolas',10); $Input.Text=$Text.Replace("`n","`r`n").Replace("`r`r","`r"); $Input.AccessibleName=$Title
    $Dialog.Controls.Add($Input); try{[void]$Dialog.ShowDialog($Owner)}finally{$Dialog.Dispose()}
}

function Get-VersionNotes([string]$DataRoot) {
    $Path=Join-Path $DataRoot 'version-notes.local.json'
    if(Test-Path -LiteralPath $Path){return Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json}
    return [pscustomobject]@{}
}

function Set-VersionNote([string]$DataRoot,[string]$ArchiveName,[string]$Text) {
    if($ArchiveName -notmatch '^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}(-\d+)?$' -or -not (Test-Path -LiteralPath (Join-Path $DataRoot ('Updates\'+$ArchiveName)) -PathType Container)){throw 'Choose an existing dated version.'}
    if($Text.Length -gt 160 -or $Text -match '[\r\n]'){throw 'Use a single-line label of at most 160 characters.'}
    $Notes=Get-VersionNotes $DataRoot
    $Notes | Add-Member -NotePropertyName $ArchiveName -NotePropertyValue $Text -Force
    Save-LocalSettings (Join-Path $DataRoot 'version-notes.local.json') $Notes
}

function Export-ResumeCopy([string]$Archive,[ValidateSet('Public','Private')][string]$Kind,[ValidateSet('PDF','DOCX')][string]$Format,[string]$Destination,[switch]$Overwrite) {
    $Archive=Assert-PlainPath $Archive
    if((Split-Path -Parent $Archive) -ne [IO.Path]::GetFullPath($UpdatesDir) -or (Split-Path -Leaf $Archive) -notmatch '^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}(-\d+)?$'){throw 'Select a dated version from the active profile history.'}
    $Destination=Assert-PlainPath $Destination
    $Filename=Split-Path -Leaf $Destination
    if($Filename -match '[<>:"/\\|?*\x00-\x1F]' -or $Filename -match '[. ]$' -or $Filename -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)' -or [IO.Path]::GetExtension($Filename) -ne ('.'+$Format.ToLower())){throw 'Choose a valid Windows filename with the selected extension.'}
    $Protected=@((Split-Path -Parent $PSScriptRoot),$Root,$SessionRoot,(Split-Path -Parent $ConfigPath))
    foreach($Entry in @($script:ProfileEntries)){
        if(Test-Path -LiteralPath $Entry.SettingsPath){
            $Other=Get-Content -LiteralPath $Entry.SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
            $Protected+= [Environment]::ExpandEnvironmentVariables($Other.DataRoot)
            if($Other.WebsitePublicPath -and $Destination -eq [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($Other.WebsitePublicPath))){throw 'Export must not replace another profile website asset.'}
        }
    }
    foreach($Folder in $Protected){if($Folder -and (Test-PathWithin $Destination $Folder)){throw 'Export outside program, settings, resume data, and recovery folders.'}}
    if($WebsitePublic -and $Destination -eq [IO.Path]::GetFullPath($WebsitePublic)){throw 'Export must not replace the website asset.'}
    if((Test-Path -LiteralPath $Destination) -and -not $Overwrite){throw 'Export already exists. Confirm replacement first.'}
    $Review=((Invoke-ResumePython $PythonExe @($VariantScript,'--review',$Archive,'--config',$ConfigPath)) -join "`n") | ConvertFrom-Json
    $Name=if($Kind -eq 'Private'){$PrivateName}else{$PublicName}; $Name=[IO.Path]::ChangeExtension($Name,$Format.ToLower())
    $Temporary=Join-Path (Split-Path -Parent $Destination) ('.export-'+[guid]::NewGuid().ToString('N')+'.'+$Format.ToLower())
    try{
        [IO.File]::Copy((Join-Path $Archive $Name),$Temporary,$false)
        [void](Invoke-ResumePython $PythonExe @($VariantScript,'--validate-export',$Temporary,'--variant',$Kind,'--config',$ConfigPath))
        if(Test-Path -LiteralPath $Destination){[IO.File]::Replace($Temporary,$Destination,[NullString]::Value)}else{[IO.File]::Move($Temporary,$Destination)}
        return $Review
    }finally{if(Test-Path -LiteralPath $Temporary){[IO.File]::Delete($Temporary)}}
}

function Show-ExportDialog($Owner) {
    $Dialog=New-Object Windows.Forms.Form; $Dialog.Text='Resume Manager - Export Resume'; $Dialog.Size=New-Object Drawing.Size(650,540); $Dialog.MinimumSize=$Dialog.Size; $Dialog.Font=$Owner.Font; $Dialog.StartPosition='CenterParent'
    $Layout=New-Object Windows.Forms.FlowLayoutPanel; $Layout.Dock='Fill'; $Layout.FlowDirection='TopDown'; $Layout.Padding=New-Object Windows.Forms.Padding(18); $Layout.AutoScroll=$true; $Dialog.Controls.Add($Layout)
    $Help=New-UiLabel 'Export a saved version, not unsaved Word edits. Current, history, and website files are not changed.'; $Help.MaximumSize=New-Object Drawing.Size(570,0); $Layout.Controls.Add($Help)
    $Versions=New-Object Windows.Forms.ComboBox; $Versions.Width=540; $Versions.DropDownStyle='DropDownList'; $Versions.AccessibleName='Saved version to export'
    $Archives=@(Get-ArchiveDirectories $UpdatesDir); foreach($Archive in $Archives){[void]$Versions.Items.Add((Format-ArchiveDate $Archive.Name))}; if($Archives.Count){$Versions.SelectedIndex=0}; $Layout.Controls.Add($Versions)
    $Kind=New-Object Windows.Forms.ComboBox; $Kind.Width=200; $Kind.DropDownStyle='DropDownList'; $Kind.AccessibleName='Public or private export'; [void]$Kind.Items.AddRange(@('Public','Private')); $Kind.SelectedIndex=0; $Layout.Controls.Add((New-UiLabel 'Contact variant')); $Layout.Controls.Add($Kind)
    $Format=New-Object Windows.Forms.ComboBox; $Format.Width=200; $Format.DropDownStyle='DropDownList'; $Format.AccessibleName='Export format'; [void]$Format.Items.AddRange(@('PDF','DOCX')); $Format.SelectedIndex=0; $Layout.Controls.Add((New-UiLabel 'File type')); $Layout.Controls.Add($Format)
    $Name=New-Object Windows.Forms.TextBox; $Name.Width=540; $Name.Text='Resume'; $Name.MaxLength=120; $Name.AccessibleName='Company or role filename'; $Layout.Controls.Add((New-UiLabel 'Filename (optional company or role, no extension)')); $Layout.Controls.Add($Name)
    $Reviewed=New-Object Windows.Forms.CheckBox; $Reviewed.Text='I will review the public document before sharing'; $Reviewed.AutoSize=$true; $Layout.Controls.Add($Reviewed)
    $Message=New-UiLabel 'Public export runs configured privacy checks. Images and unknown private details still need review.'; $Message.MaximumSize=New-Object Drawing.Size(570,0); $Layout.Controls.Add($Message)
    $Save=New-UiButton 'Export Copy...'; $Save.Enabled=$Archives.Count -gt 0; $Layout.Controls.Add($Save)
    $Save.Add_Click({
        try{
            if($Versions.SelectedIndex -lt 0){throw 'Select a saved version.'}
            if($Kind.SelectedItem -eq 'Public' -and -not $Reviewed.Checked){throw 'Confirm that you will review the public output before sharing.'}
            $Picker=New-Object Windows.Forms.SaveFileDialog; $Picker.Filter=$Format.SelectedItem+' file|*.'+$Format.SelectedItem.ToLower(); $Picker.FileName=$Name.Text+'-'+$Kind.SelectedItem.ToLower()+'.'+$Format.SelectedItem.ToLower(); $Picker.OverwritePrompt=$true
            try{
                if($Picker.ShowDialog($Dialog) -ne 'OK'){return}
                $Review=Export-ResumeCopy $Archives[$Versions.SelectedIndex].FullName $Kind.SelectedItem $Format.SelectedItem $Picker.FileName -Overwrite
                $Message.Text='Export saved. '+$Review.Checks+"`n"+($Review.Warnings -join '; ')
            }finally{$Picker.Dispose()}
        }catch{$Message.Text='Export failed: '+$_.Exception.Message}
    })
    $script:OperationBusy=$true
    try{[void]$Dialog.ShowDialog($Owner)}finally{$script:OperationBusy=$false;$Dialog.Dispose()}
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
        Assert-ProfileRoot $Loaded.Configuration.DataRoot $Path
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
    $Names=if($Issue.Kind -eq 'Setup'){@('Locate existing settings','Create new setup','Cancel')}else{@('Retry','Locate existing settings','Edit settings','Cancel')}
    foreach($Name in $Names){$Button=New-UiButton $Name; $Button.Tag=$Name; $Button.Add_Click({$Choice.Value=$this.Tag; $Dialog.Close()}); $Actions.Controls.Add($Button)}
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

function Save-LocalSettings([string]$Path, $Value, [switch]$BackupSettings) {
    $Code = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot)).TrimEnd('\') + '\'
    $Path = [IO.Path]::GetFullPath($Path)
    if ($Path.StartsWith($Code,[StringComparison]::OrdinalIgnoreCase)) { throw 'Settings must be outside the program repository.' }
    $Directory = Split-Path -Parent $Path
    [void][IO.Directory]::CreateDirectory($Directory)
    $Temporary = Join-Path $Directory ([guid]::NewGuid().ToString('N') + '.tmp')
    try {
        if($BackupSettings -and (Test-Path -LiteralPath $Path)){
            # Only validated prior settings are advertised as a recoverable configuration.
            $Valid=$false
            try{
                $Previous=& { . (Join-Path $PSScriptRoot 'resume_settings.ps1') -ConfigPath $Path; $PythonExe }
                [void](Invoke-ResumePython $Previous @((Join-Path $PSScriptRoot 'update_resume_projects.py'),'--validate-settings','--config',$Path)); $Valid=$true
            }catch{}
            if($Valid){
                $BackupDir=Join-Path $Directory 'Settings Backups'; [void][IO.Directory]::CreateDirectory($BackupDir)
                [IO.File]::Copy($Path,(Join-Path $BackupDir ('settings-'+(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')+'-'+[guid]::NewGuid().ToString('N')+'.json')),$false)
            }
        }
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
    $Command = "`$ProgressPreference='SilentlyContinue'; [Console]::OutputEncoding=New-Object Text.UTF8Encoding(`$false); `$env:PYTHONIOENCODING='utf-8'; try { & '$Workflow' -SourcePath '$($Source.Replace("'","''"))' -ConfigPath '$($Configuration.Replace("'","''"))' -ProgressPath '$($ProgressFile.Replace("'","''"))'; exit 0 } catch { [Console]::Error.WriteLine(`$_); exit 1 }"
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
        if ($Process.ExitCode -ne 0) { throw ((Get-Content -LiteralPath (Join-Path $JobDir 'error.txt') -Raw -Encoding UTF8) + "`nRecovery details: $JobDir") }
    }
    finally { $Owner.Text = $OriginalTitle; $Owner.Enabled = $true; $Process.Dispose() }
}

function Show-SettingsDialog([string]$Path, $Owner, $SeedSettings = $null, [string]$StartingSource = '', [string]$DefaultDataRoot = '') {
    $Existing=$null; $ReadError=''
    if(Test-Path -LiteralPath $Path){try{$Existing=Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json}catch{$ReadError='Saved settings could not be read. Review these values before replacing them.'}}
    if(-not $Existing -and $SeedSettings){$Existing=$SeedSettings | ConvertTo-Json -Depth 100 | ConvertFrom-Json}
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
    $SourceBox.Text=$StartingSource
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
        @('PythonExe','Python executable',$(if($script:PythonExe){$script:PythonExe}else{'python.exe'})),@('WebsitePublicPath','Website public DOCX (optional)',''),
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
    if($DefaultDataRoot){$Fields.DataRoot.Text=$DefaultDataRoot}
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
                if($Key -eq 'PrivateOnly'){$Value=@($Value)} # Keep a one-item privacy policy as a JSON array.
                $Candidate | Add-Member -NotePropertyName $Key -NotePropertyValue $Value -Force
            }
            $FocusControl=$Fields.PythonExe; $Candidate.PythonExe=Test-Prerequisites $Candidate.PythonExe
            $CandidatePath=Join-Path (Split-Path -Parent $Path) ('candidate-'+[guid]::NewGuid().ToString('N')+'.json')
            Save-LocalSettings $CandidatePath $Candidate
            $FocusControl=$Fields.PublicContact
            [void](Invoke-ResumePython $Candidate.PythonExe @((Join-Path $PSScriptRoot 'update_resume_projects.py'),'--validate-settings','--config',$CandidatePath))
            $FocusControl=$Fields.DataRoot
            $Validated=& { . (Join-Path $PSScriptRoot 'resume_settings.ps1') -ConfigPath $CandidatePath; [pscustomobject]@{Private=$CurrentPrivate;Public=$CurrentPublic;Root=$Root} }
            Assert-ProfileRoot $Validated.Root $Path
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
            Save-LocalSettings $Path $Candidate -BackupSettings
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

function Save-SessionState {
    if(-not $script:WorkingDir){return}
    Save-LocalSettings (Join-Path $script:WorkingDir 'session.local.json') ([pscustomobject]@{
        SettingsPath=$ConfigPath;DataRoot=$Root;Kind=$script:SessionKind;WorkingPath=$script:WorkingPath
        LastArchivedHash=$script:LastArchivedHash;Updated=[datetime]::UtcNow.ToString('o')
    })
}

function Get-RecoverableCopies {
    foreach($Directory in @(Get-ChildItem -LiteralPath $SessionRoot -Directory -ErrorAction SilentlyContinue)){
        $Metadata=Join-Path $Directory.FullName 'session.local.json'
        if(-not (Test-Path -LiteralPath $Metadata)){continue}
        try{
            $State=Get-Content -LiteralPath $Metadata -Raw -Encoding UTF8 | ConvertFrom-Json
            if([IO.Path]::GetFullPath($State.SettingsPath) -ne [IO.Path]::GetFullPath($ConfigPath) -or [IO.Path]::GetFullPath($State.DataRoot) -ne [IO.Path]::GetFullPath($Root)){continue}
            $File=Get-Item -LiteralPath $State.WorkingPath -ErrorAction Stop
            if($File.Name.StartsWith('~$') -or $File.Extension -ne '.docx'){continue}
            $Hash=Get-ResumeHash $File.FullName
            if($Hash -eq $State.LastArchivedHash){continue}
            [pscustomobject]@{Path=$File.FullName;Kind=$State.Kind;Saved=$File.LastWriteTime;Hash=$Hash}
        }catch{} # An incomplete/interrupted package is retained for manual inspection, never deleted.
    }
}

function Recover-ResumeCopy([string]$Source,$Owner) {
    if(Test-ManagerBusy){throw 'Close the editing session and finish saving before recovery.'}
    [void](Invoke-ResumePython $PythonExe @($VariantScript,'--check-source',$Source,'--config',$ConfigPath))
    $script:OperationBusy=$true
    try{[void](Backup-CurrentPair $Root); Invoke-WorkflowInteractive $Source $ConfigPath $Owner}
    finally{$script:OperationBusy=$false}
    $LatestValue.Text=Format-ArchiveDate (Get-LatestArchive); Update-PublicReview
    $ErrorPanel.Visible=$false; Set-Status ('Both versions saved at '+(Get-Date -Format 'h:mm:ss tt'))
}

function Show-RecoveryDialog($Owner) {
    if(Test-ManagerBusy){throw 'Finish the editing session before recovery.'}
    $Dialog=New-Object Windows.Forms.Form; $Dialog.Text='Resume Manager - Recover Saved Working Copy'; $Dialog.Size=New-Object Drawing.Size(800,440); $Dialog.MinimumSize=$Dialog.Size; $Dialog.Font=$Owner.Font; $Dialog.StartPosition='CenterParent'
    $Layout=New-Object Windows.Forms.TableLayoutPanel; $Layout.Dock='Fill'; $Layout.Padding=New-Object Windows.Forms.Padding(14); $Layout.ColumnCount=1; $Layout.RowCount=3
    [void]$Layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('AutoSize'))); [void]$Layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent',100))); [void]$Layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('AutoSize'))); $Dialog.Controls.Add($Layout)
    $Message=New-UiLabel 'Only saved DOCX content can be recovered, not unsaved Word text. Recovery creates a new version with the active contact settings; the source is retained.'; $Message.MaximumSize=New-Object Drawing.Size(740,0); $Layout.Controls.Add($Message,0,0)
    $List=New-Object Windows.Forms.ListView; $List.View='Details'; $List.FullRowSelect=$true; $List.MultiSelect=$false; $List.Dock='Fill'; $List.AccessibleName='Saved working copies for the active profile'; [void]$List.Columns.Add('Saved locally',180); [void]$List.Columns.Add('Variant',90); [void]$List.Columns.Add('Source',450)
    foreach($Copy in @(Get-RecoverableCopies | Sort-Object Saved -Descending)){$Item=New-Object Windows.Forms.ListViewItem($Copy.Saved.ToString('g')); [void]$Item.SubItems.Add($Copy.Kind); [void]$Item.SubItems.Add($Copy.Path); $Item.Tag=$Copy; [void]$List.Items.Add($Item)}; $Layout.Controls.Add($List,0,1)
    $Actions=New-Object Windows.Forms.FlowLayoutPanel; $Actions.AutoSize=$true; $Actions.Dock='Fill'; $Layout.Controls.Add($Actions,0,2)
    $Recover=New-UiButton 'Recover Selected Copy'; $Browse=New-UiButton 'Choose Saved DOCX...'; $Actions.Controls.AddRange(@($Recover,$Browse))
    $Run={param([string]$Source)
        if(-not (Confirm-ResumeReplacement $Dialog ("Use saved content from:`n$Source`n`nCreate both variants using the active privacy settings? Current will be backed up first. Unsaved Word content cannot be recovered.") 'Confirm recovery')){return}
        Recover-ResumeCopy $Source $Dialog; $Message.Text='Both versions saved. Recovery source retained.'
    }
    $Recover.Add_Click({try{if($List.SelectedItems.Count -ne 1){throw 'Select a saved working copy first.'}; & $Run $List.SelectedItems[0].Tag.Path}catch{$Message.Text='Recovery failed: '+$_.Exception.Message}})
    $Browse.Add_Click({$Picker=New-Object Windows.Forms.OpenFileDialog; $Picker.Filter='Saved Word document (*.docx)|*.docx'; try{if($Picker.ShowDialog($Dialog) -eq 'OK'){& $Run $Picker.FileName}}catch{$Message.Text='Recovery failed: '+$_.Exception.Message}finally{$Picker.Dispose()}})
    try{[void]$Dialog.ShowDialog($Owner)}finally{$Dialog.Dispose()}
}

function Set-ActiveProfile([string]$Path) {
    if(Test-ManagerBusy){throw 'Finish editing, saving, export, restore, or recovery before switching profiles.'}
    $Issue=Get-StartupIssue $Path
    if($Issue.Kind -ne 'Ready'){throw ($Issue.Message+"`n"+$Issue.Details)}
    $Loaded=& { . (Join-Path $PSScriptRoot 'resume_settings.ps1') -ConfigPath $Path; [pscustomobject]@{Root=$Root} }
    Assert-ProfileRoot $Loaded.Root $Path
    $script:ConfigPath=[IO.Path]::GetFullPath($Path)
    . (Join-Path $PSScriptRoot 'resume_settings.ps1') -ConfigPath $script:ConfigPath
    foreach($Name in @('Root','CurrentDir','UpdatesDir','PrivateName','PublicName','CurrentPrivate','CurrentPublic','CurrentPrivateName','CurrentPublicName','PythonExe','VariantScript','Settings','WebsitePublic')){Set-Variable -Name $Name -Scope Script -Value (Get-Variable -Name $Name -ValueOnly)}
    $script:PythonExe=$Issue.Python; $script:RecoveryPath=$Root
    $script:LastSavedAt=''
    $ErrorPanel.Visible=$false; $ErrorDetail.Text=''; $script:LastErrorDetail=''
    $LatestValue.Text=Format-ArchiveDate (Get-LatestArchive); Update-PublicReview; Update-ProfileCaption; Set-Status 'Ready'
}

function Save-ProfileRegistry {
    # ponytail: a small local list of settings paths, not a database or a second document workflow.
    if(Test-Path -LiteralPath $script:RegistryPath){
        [IO.File]::Copy($script:RegistryPath,($script:RegistryPath+'.backup-'+[guid]::NewGuid().ToString('N')),$false)
    }
    Save-LocalSettings $script:RegistryPath ([pscustomobject]@{ActivePath=$ConfigPath;Entries=@($script:ProfileEntries)})
}

function Show-ProfilesDialog($Owner) {
    if(Test-ManagerBusy){throw 'Finish the current editing or saving operation before managing profiles.'}
    $Dialog=New-Object Windows.Forms.Form; $Dialog.Text='Resume Manager - Optional Profiles'; $Dialog.Size=New-Object Drawing.Size(700,410); $Dialog.MinimumSize=$Dialog.Size; $Dialog.Font=$Owner.Font; $Dialog.StartPosition='CenterParent'
    $Layout=New-Object Windows.Forms.FlowLayoutPanel; $Layout.Dock='Fill'; $Layout.FlowDirection='TopDown'; $Layout.Padding=New-Object Windows.Forms.Padding(18); $Layout.AutoScroll=$true; $Dialog.Controls.Add($Layout)
    $Message=New-UiLabel 'Each profile has its own contacts, Current pair, history, recovery, and notes. Your existing resume stays in its original folder.'; $Message.MaximumSize=New-Object Drawing.Size(610,0); $Layout.Controls.Add($Message)
    $List=New-Object Windows.Forms.ListBox; $List.Size=New-Object Drawing.Size(610,150); $List.AccessibleName='Resume profiles'; foreach($Entry in @($script:ProfileEntries)){[void]$List.Items.Add($Entry.Name)}; $Layout.Controls.Add($List)
    $Actions=New-Object Windows.Forms.FlowLayoutPanel; $Actions.AutoSize=$true; $Actions.Width=610; $Layout.Controls.Add($Actions)
    $Switch=New-UiButton 'Use Selected Profile'; $New=New-UiButton 'New from Word Resume...'; $Copy=New-UiButton 'Copy Active Resume...'; $Actions.Controls.AddRange(@($Switch,$New,$Copy))
    $Switch.Add_Click({try{if($List.SelectedIndex -lt 0){throw 'Select a profile first.'}; Set-ActiveProfile $script:ProfileEntries[$List.SelectedIndex].SettingsPath; Save-ProfileRegistry; $Dialog.Close()}catch{$Message.Text=$_.Exception.Message}})
    $Create={param([bool]$CopyActive)
        $Name=Show-TextEntry $Dialog 'New Profile' 'Profile name, such as Developer or General.'
        if($null -eq $Name){return}; if(-not $Name -or @($script:ProfileEntries | Where-Object {$_.Name -eq $Name}).Count){throw 'Choose a nonempty, unique profile name.'}
        $Folder=Join-Path (Split-Path -Parent $script:RegistryPath) ('Profiles\'+[guid]::NewGuid().ToString('N'))
        $Path=Join-Path $Folder 'settings.local.json'; $Seed=$null; $Source=''
        if($CopyActive){
            $Seed=$Settings | ConvertTo-Json -Depth 100 | ConvertFrom-Json; $Seed.DataRoot=Join-Path $Folder 'Data'; $Seed.WebsitePublicPath=''; $Source=$CurrentPublic
        }
        if(Show-SettingsDialog $Path $Dialog $Seed $Source (Join-Path $Folder 'Data')){
            $script:ProfileEntries+=[pscustomobject]@{Name=$Name;SettingsPath=$Path}
            Save-ProfileRegistry; [void]$List.Items.Add($Name); $Message.Text='Profile created. Select it to switch. Website copying is disabled unless explicitly configured.'
        }
    }
    $New.Add_Click({try{& $Create $false}catch{$Message.Text=$_.Exception.Message}}); $Copy.Add_Click({try{& $Create $true}catch{$Message.Text=$_.Exception.Message}})
    try{[void]$Dialog.ShowDialog($Owner)}finally{$Dialog.Dispose()}
}

function Update-DesktopShortcut([string]$ShortcutPath = '') {
    if(-not $ShortcutPath){$ShortcutPath=Join-Path ([Environment]::GetFolderPath('Desktop')) 'Resume Manager.lnk'}
    $ShortcutPath=Assert-PlainPath $ShortcutPath
    if([IO.Path]::GetExtension($ShortcutPath) -ne '.lnk'){throw 'Desktop shortcut must be a .lnk file.'}
    $Temporary=Join-Path (Split-Path -Parent $ShortcutPath) ([guid]::NewGuid().ToString('N')+'.lnk')
    $Shell=New-Object -ComObject WScript.Shell
    try{
        $Shortcut=$Shell.CreateShortcut($Temporary)
        $Shortcut.TargetPath="$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
        $Shortcut.Arguments='-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "'+(Join-Path $PSScriptRoot 'Resume Manager.ps1')+'" -ConfigPath "'+$ConfigPath+'"'
        $Shortcut.WorkingDirectory=Split-Path -Parent $PSScriptRoot; $Shortcut.Description='Resume Manager '+$script:ProgramVersion+' - local settings'; $Shortcut.Save()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Shortcut)
        if(Test-Path -LiteralPath $ShortcutPath){
            $BackupDir=Join-Path (Split-Path -Parent $ConfigPath) 'Shortcut Backups'; [void][IO.Directory]::CreateDirectory($BackupDir)
            $Backup=Join-Path $BackupDir ('Resume Manager-'+[guid]::NewGuid().ToString('N')+'.lnk'); [IO.File]::Copy($ShortcutPath,$Backup,$false)
            [IO.File]::Replace($Temporary,$ShortcutPath,[NullString]::Value)
        }else{[IO.File]::Move($Temporary,$ShortcutPath)}
    }finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Shell); if(Test-Path -LiteralPath $Temporary){[IO.File]::Delete($Temporary)}}
}

function Show-AboutDialog($Owner) {
    $Dialog=New-Object Windows.Forms.Form; $Dialog.Text='Resume Manager - About and Help'; $Dialog.Size=New-Object Drawing.Size(660,460); $Dialog.MinimumSize=$Dialog.Size; $Dialog.Font=$Owner.Font; $Dialog.StartPosition='CenterParent'
    $Layout=New-Object Windows.Forms.FlowLayoutPanel; $Layout.Dock='Fill'; $Layout.Padding=New-Object Windows.Forms.Padding(18); $Layout.FlowDirection='TopDown'; $Layout.AutoScroll=$true; $Dialog.Controls.Add($Layout)
    $Message=New-UiLabel ('Resume Manager '+$script:ProgramVersion+"`nMIT licensed. Runs locally; no uploads or telemetry.`n`n1. Click Edit Private/Public Resume; Word opens automatically.`n2. Edit and save in Word. Keep the manager open.`n3. Wait for Both versions saved, then export a copy.`n`nThe program folder contains code only. Settings, profiles, documents, notes, backups, and history stay outside it.")
    $Message.MaximumSize=New-Object Drawing.Size(590,0); $Layout.Controls.Add($Message)
    $Shortcut=New-UiButton 'Create / Update Desktop Shortcut'; $Help=New-UiButton 'Open Instructions'; $Demo=New-UiButton 'Create Example Resume...'; $Layout.Controls.AddRange(@($Shortcut,$Help,$Demo))
    $Shortcut.Add_Click({try{if(-not (Confirm-ResumeReplacement $Dialog 'Create or update your desktop shortcut to this program and the active profile? Any old shortcut will be backed up locally.' 'Desktop shortcut')){return}; Update-DesktopShortcut; $Message.Text='Desktop shortcut updated for the active settings. Any old shortcut is in Shortcut Backups next to those settings.'}catch{$Message.Text=$_.Exception.Message}})
    $Help.Add_Click({try{Open-LocalFile (Join-Path (Split-Path -Parent $PSScriptRoot) 'README.md')}catch{$Message.Text=$_.Exception.Message}})
    $Demo.Add_Click({$Picker=New-Object Windows.Forms.SaveFileDialog; $Picker.Filter='Word resume (*.docx)|*.docx'; $Picker.FileName='example-resume.docx'; $Picker.OverwritePrompt=$false; try{if($Picker.ShowDialog($Dialog) -eq 'OK'){New-ExampleResume $Picker.FileName; $Message.Text='Example created. Import it only into a separate test profile; your active resume has not changed.'}}catch{$Message.Text=$_.Exception.Message}finally{$Picker.Dispose()}})
    try{[void]$Dialog.ShowDialog($Owner)}finally{$Dialog.Dispose()}
}

function New-ExampleResume([string]$Destination) {
    $Destination=Assert-PlainPath $Destination
    if([IO.Path]::GetExtension($Destination) -ne '.docx' -or (Test-Path -LiteralPath $Destination)){throw 'Choose a new DOCX filename; existing files are never overwritten.'}
    foreach($Protected in @((Split-Path -Parent $PSScriptRoot),$Root,(Split-Path -Parent $ConfigPath))){if(Test-PathWithin $Destination $Protected){throw 'Create the example outside the program, settings, and active data folders.'}}
    [void](Invoke-ResumePython $PythonExe @($VariantScript,'--create-example',$Destination))
}
