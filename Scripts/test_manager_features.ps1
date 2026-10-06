param([string]$PythonExe='python.exe',[string]$OutputDirectory='')
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'manager_dialogs.ps1')
$PythonExe=Test-Prerequisites $PythonExe
if(-not $OutputDirectory){$OutputDirectory=$env:TEMP}
$TestRoot=Join-Path $OutputDirectory ('resume-features-'+[guid]::NewGuid().ToString('N')+' '+[char]0xE9)
[void][IO.Directory]::CreateDirectory($TestRoot)
$SettingsDir=Join-Path $TestRoot 'Settings'; [void][IO.Directory]::CreateDirectory($SettingsDir)
$Config=Join-Path $SettingsDir 'settings.local.json'
$Source=Join-Path $TestRoot 'example.docx'
[void](Invoke-ResumePython $PythonExe @((Join-Path $PSScriptRoot 'update_resume_projects.py'),'--create-example',$Source))
$FakePhone='555'+'-0100'
$Fixture=[pscustomobject]@{DataRoot=(Join-Path $TestRoot 'General');PrivateContact=('public@example.com | personal@example.com | '+$FakePhone+' | Example City');PublicContact='public@example.com | Example City';ContactAnchor='public@example.com';PrivateOnly=@('personal@example.com',$FakePhone);PythonExe=$PythonExe;WebsitePublicPath=(Join-Path $TestRoot 'Website\public.docx');Extra=('UTF8 '+[char]0xE9)}
[void][IO.Directory]::CreateDirectory((Join-Path $TestRoot 'Website'))
Save-LocalSettings $Config $Fixture
function Assert-Test([bool]$Condition,[string]$Name){if(-not $Condition){throw ('FAIL: '+$Name)}; Write-Output ('PASS: '+$Name)}
function Assert-Fails([scriptblock]$Action,[string]$Name){$Failed=$false;try{& $Action | Out-Null}catch{$Failed=$true};Assert-Test $Failed $Name}
function Count-Versions {return @(Get-ArchiveDirectories $UpdatesDir).Count}
function Pump-Until([scriptblock]$Condition,[int]$Seconds=60){
    $Deadline=[datetime]::UtcNow.AddSeconds($Seconds)
    do{[Windows.Forms.Application]::DoEvents();if(& $Condition){return};Start-Sleep -Milliseconds 100}while([datetime]::UtcNow -lt $Deadline)
    throw ('Timed out: '+$StatusValue.Text+' '+$ErrorDetail.Text)
}
function Pump-Seconds([int]$Seconds){$Deadline=[datetime]::UtcNow.AddSeconds($Seconds);while([datetime]::UtcNow -lt $Deadline){[Windows.Forms.Application]::DoEvents();Start-Sleep -Milliseconds 100}}
function Add-TestText([string]$Marker){$Range=$script:WordDocument.Content;$Range.Collapse(0);$Range.InsertBefore($Marker+"`r");[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Range)}
function Tree-Hashes([string]$Folder){return (@(Get-ChildItem -LiteralPath $Folder -File -Recurse | Sort-Object FullName | ForEach-Object { $_.FullName+'|'+(Get-FileHash -LiteralPath $_.FullName).Hash}) -join "`n")}
function Controls($Parent){foreach($Child in $Parent.Controls){$Child;Controls $Child}}
function Inspect-Dialog([scriptblock]$Open,[scriptblock]$Check){
    $Result=@{Done=$false;Error=''};$Probe=New-Object Windows.Forms.Timer;$Probe.Interval=200
    $Probe.Add_Tick({$Window=@([Windows.Forms.Application]::OpenForms | Where-Object {$_.Text -like 'Resume Manager -*'})|Select-Object -Last 1;if(-not $Window -or $Result.Done){return};$Result.Done=$true;try{& $Check $Window @(Controls $Window)}catch{$Result.Error=$_.Exception.Message}finally{$Window.Close()}})
    $Probe.Start();try{& $Open}finally{$Probe.Stop();$Probe.Dispose()};if($Result.Error){throw $Result.Error};Assert-Test $Result.Done 'native dialog opened and cancelled'
}
$Workflow=Join-Path $PSScriptRoot 'resume_update_workflow.ps1'
$Stamp=Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
& $Workflow -SourcePath $Source -ConfigPath $Config -Timestamp $Stamp | Out-Null
& $Workflow -SourcePath $Source -ConfigPath $Config -Timestamp $Stamp | Out-Null
. (Join-Path $PSScriptRoot 'Resume Manager.ps1') -LoadOnly -ConfigPath $Config
$script:SessionRoot=Join-Path $TestRoot 'Sessions'
try{
    $Timer.Stop(); $Form.Show(); $Form.PerformLayout()
    Assert-Test ((Count-Versions) -eq 2) 'same-second saves create distinct complete folders'
    Assert-Test ($EditPrivateButton.Enabled -and $EditPublicButton.Enabled) 'valid existing setup reaches enabled edit buttons'
    Inspect-Dialog {Show-StartupRecovery ([pscustomobject]@{Kind='Setup';Message='Choose setup';Details=''}) | Out-Null} {param($Window,$All) Assert-Test (@($All | Where-Object {$_.Text -eq 'Locate existing settings'}).Count -eq 1) 'missing settings has Locate existing settings'; Assert-Test (@($All | Where-Object {$_.Text -eq 'Create new setup'}).Count -eq 1) 'first-run setup is an explicit separate action'}
    $OnePolicy=Join-Path $SettingsDir 'single-value.local.json'
    Inspect-Dialog {Show-SettingsDialog $OnePolicy $Form | Out-Null} {
        param($Window,$All)
        @($All | Where-Object {$_.AccessibleName -eq 'Resume data folder'})[0].Text=Join-Path $TestRoot 'Single Policy'
        @($All | Where-Object {$_.AccessibleName -eq 'Python executable'})[0].Text=$PythonExe
        @($All | Where-Object {$_.AccessibleName -eq 'Starting DOCX (optional for existing data)'})[0].Text=$Source
        @($All | Where-Object {$_.Text -eq 'Read contact details'})[0].PerformClick()
        $Contacts=@($All | Where-Object {$_ -is [Windows.Forms.CheckedListBox]})[0]
        $Contacts.SetItemChecked(0,$true);$Contacts.SetItemChecked($Contacts.Items.Count-1,$true)
        @($All | Where-Object {$_.Text -eq 'I reviewed the public contact details'})[0].Checked=$true
        @($All | Where-Object {$_.Text -eq '&Save and finish'})[0].PerformClick()
        if(-not (Test-Path -LiteralPath $OnePolicy)){throw 'Single-value setup failed'}
        $Policy=Get-Content -LiteralPath $OnePolicy -Raw -Encoding UTF8 | ConvertFrom-Json
        Assert-Test ($Policy.PrivateOnly -is [array] -and $Policy.PrivateOnly.Count -eq 1) 'one private detail stays a JSON array in actual setup'
    }
    $BeforeConfig=Get-Content -LiteralPath $Config -Raw -Encoding UTF8
    Save-LocalSettings $Config $Fixture -BackupSettings
    Assert-Test (@(Get-ChildItem -LiteralPath (Join-Path $SettingsDir 'Settings Backups') -File).Count -eq 1) 'previous valid settings backed up'
    Assert-Test ((Get-Content -LiteralPath $Config -Raw -Encoding UTF8 | ConvertFrom-Json).Extra -eq $Fixture.Extra) 'unknown UTF8 settings preserved'
    $ArchivesBefore=Tree-Hashes $UpdatesDir; $CurrentBefore=Tree-Hashes $CurrentDir; $WebsiteBefore=(Get-FileHash $WebsitePublic).Hash
    $First=@(Get-ArchiveDirectories $UpdatesDir)|Select-Object -First 1
    $ExportDir=Join-Path $TestRoot 'Exports';[void][IO.Directory]::CreateDirectory($ExportDir)
    foreach($Kind in @('Private','Public')){foreach($Format in @('PDF','DOCX')){
        $Target=Join-Path $ExportDir ($Kind+'.'+$Format.ToLower())
        $Review=Export-ResumeCopy $First.FullName $Kind $Format $Target
        $Name=if($Kind -eq 'Public'){$PublicName}else{$PrivateName};$Original=Join-Path $First.FullName ([IO.Path]::ChangeExtension($Name,$Format.ToLower()))
        Assert-Test ((Get-FileHash $Target).Hash -eq (Get-FileHash $Original).Hash) ($Kind+' '+$Format+' export is byte-identical')
    }}
    Assert-Fails {Export-ResumeCopy $First.FullName Public DOCX (Join-Path $ExportDir 'Public.docx')} 'overwrite requires explicit permission'
    Assert-Fails {Export-ResumeCopy $First.FullName Public DOCX $CurrentPublic -Overwrite} 'export cannot replace Current'
    Assert-Fails {Export-ResumeCopy $First.FullName Public DOCX (Join-Path $ExportDir 'CON.docx')} 'reserved Windows filename rejected'
    Set-VersionNote $Root $First.Name ('Added test '+[char]0xE9)
    Assert-Test ((Get-VersionNotes $Root).($First.Name) -match 'Added test') 'UTF8 label round-trip outside archive'
    $Diff=((Invoke-ResumePython $PythonExe @($VariantScript,'--compare-text',(Join-Path $First.FullName $PublicName),(Join-Path $First.FullName $PublicName))) -join "`n")|ConvertFrom-Json
    Assert-Test ($Diff -eq 'No body text changes.') 'unchanged text comparison'
    Inspect-Dialog {Show-ExportDialog $Form} {param($Window,$All) $Choices=@($All | Where-Object {$_ -is [Windows.Forms.ComboBox]});Assert-Test ($Choices[1].SelectedItem -eq 'Public' -and $Choices[2].SelectedItem -eq 'PDF') 'export defaults to public PDF'}
    Inspect-Dialog {Show-HistoryDialog $Form} {param($Window,$All) $List=@($All | Where-Object {$_ -is [Windows.Forms.ListView]})[0];Assert-Test ($List.MultiSelect -and $List.Columns[1].Text -eq 'Label') 'history supports labels and two-version comparison'}
    Assert-Test ((Tree-Hashes $UpdatesDir) -eq $ArchivesBefore -and (Tree-Hashes $CurrentDir) -eq $CurrentBefore -and (Get-FileHash $WebsitePublic).Hash -eq $WebsiteBefore) 'export, notes, comparison and cancellation never mutate documents'
    $Link=Join-Path $TestRoot 'Resume Test.lnk';Update-DesktopShortcut $Link;Update-DesktopShortcut $Link
    $Shell=New-Object -ComObject WScript.Shell;$Shortcut=$Shell.CreateShortcut($Link)
    Assert-Test ($Shortcut.Arguments.Contains('"'+$Config+'"') -and $Shortcut.Arguments.Contains('-STA') -and $Shortcut.Description.Contains($script:ProgramVersion)) 'shortcut quotes Unicode/spaces and current version/settings'
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Shortcut);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Shell)
    Assert-Test (@(Get-ChildItem -LiteralPath (Join-Path $SettingsDir 'Shortcut Backups') -File).Count -eq 1) 'old desktop shortcut is backed up'
    $Timer.Start()
    foreach($Kind in @('Private','Public')){
        $Before=Count-Versions;Start-EditSession $Kind;$script:Word.Visible=$false
        Add-TestText ('Test '+$Kind+' edit');Pump-Until {$StatusValue.Text -eq 'Unsaved changes in Word'}
        Assert-Test (-not $ProfilesButton.Enabled -and -not $RecoveryButton.Enabled) 'profile switching and recovery blocked during editing'
        Assert-Fails {Set-ActiveProfile $Config} 'programmatic profile switching is also blocked'
        $script:WordDocument.Save();Pump-Until {$null -ne $script:ArchiveProcess}
        if($Kind -eq 'Public'){
            Add-TestText 'Newer unsaved text';Pump-Until {$StatusValue.Text -eq 'Unsaved changes in Word'}
            Pump-Until {$null -eq $script:ArchiveProcess}
            Assert-Test ($StatusValue.Text -eq 'Unsaved changes in Word') 'earlier snapshot completion does not hide newer unsaved text'
            $script:WordDocument.Save()
        }
        Pump-Until {$null -eq $script:ArchiveProcess -and $StatusValue.Text -like 'Both versions saved at*'}
        $Count=Count-Versions;$script:WordDocument.Save();$script:WordDocument.Save();$script:WordDocument.Save();Pump-Seconds 4
        Assert-Test ((Count-Versions) -eq $Count -and $Count -gt $Before) ($Kind+' edits saved while Word open; repeated saves deduplicated')
        $script:WordDocument.Close($false);Pump-Until {-not $script:SessionActive}
        Assert-Test ((Count-Versions) -eq $Count) 'close after successful save is deduplicated'
    }
    Start-EditSession Public;$script:Word.Visible=$false;$Count=Count-Versions;$script:WordDocument.Close($false);Pump-Until {-not $script:SessionActive};Assert-Test ((Count-Versions) -eq $Count) 'unchanged close creates no archive'
    $Timer.Stop()
    $Newest=@(Get-ArchiveDirectories $UpdatesDir)|Select-Object -First 1
    Assert-Test ((Get-FileHash $CurrentPublic).Hash -eq (Get-FileHash (Join-Path $Newest.FullName $PublicName)).Hash -and (Get-FileHash $CurrentPrivate).Hash -eq (Get-FileHash (Join-Path $Newest.FullName $PrivateName)).Hash) 'Current pair matches newest archive'
    Assert-Test ((Get-FileHash $WebsitePublic).Hash -eq (Get-FileHash $CurrentPublic).Hash) 'website matches newest public DOCX byte-for-byte'
    $Review=((Invoke-ResumePython $PythonExe @($VariantScript,'--review',$Newest.FullName,'--config',$ConfigPath)) -join "`n") | ConvertFrom-Json
    Assert-Test ($Review.PrivatePages -eq 1 -and $Review.PublicPages -eq 1) 'both synthetic resumes remain one page'
    $Changed=((Invoke-ResumePython $PythonExe @($VariantScript,'--compare-text',(Join-Path $First.FullName $PublicName),$CurrentPublic)) -join "`n")|ConvertFrom-Json
    Assert-Test ($Changed -match 'Test Private edit' -and $Changed -match 'Test Public edit') 'text diff shows both propagated edits'
    $RecoverDir=Join-Path $SessionRoot 'interrupted';[void][IO.Directory]::CreateDirectory($RecoverDir);$RecoverSource=Join-Path $RecoverDir 'saved.docx';[IO.File]::Copy($CurrentPublic,$RecoverSource,$false)
    Save-LocalSettings (Join-Path $RecoverDir 'session.local.json') ([pscustomobject]@{SettingsPath=$Config;DataRoot=$Root;Kind='public';WorkingPath=$RecoverSource;LastArchivedHash='old'})
    $Copies=@(Get-RecoverableCopies);Assert-Test ($Copies.Count -eq 1 -and $Copies[0].Kind -eq 'public') 'saved working copy can be located with profile and variant'
    $RecoveryHash=(Get-FileHash $RecoverSource).Hash;$Count=Count-Versions;Recover-ResumeCopy $RecoverSource $Form
    Assert-Test ((Count-Versions) -eq $Count+1 -and (Get-FileHash $RecoverSource).Hash -eq $RecoveryHash) 'recovery archives a new pair and retains source'
    $Profile2=Join-Path $SettingsDir 'developer.local.json';$Second=$Fixture|ConvertTo-Json|ConvertFrom-Json;$Second.DataRoot=Join-Path $TestRoot 'Developer';$Second.WebsitePublicPath='';Save-LocalSettings $Profile2 $Second
    & $Workflow -SourcePath $Source -ConfigPath $Profile2 | Out-Null
    $script:ProfileEntries+=[pscustomobject]@{Name='Developer';SettingsPath=$Profile2}
    Assert-Fails {Assert-ProfileRoot (Join-Path $Fixture.DataRoot 'Child') $Profile2} 'overlapping profile roots rejected'
    $GeneralBefore=Tree-Hashes $Fixture.DataRoot;$WebsiteBefore=(Get-FileHash $WebsitePublic).Hash
    Set-ActiveProfile $Profile2;Save-ProfileRegistry
    Assert-Test ($ProfileCaption.Text -eq 'Profile: Developer' -and -not $WebsitePublic) 'profile has independent files and website disabled'
    Assert-Test (@(Get-RecoverableCopies).Count -eq 0) 'recovery copies do not cross profiles'
    Set-VersionNote $Root (@(Get-ArchiveDirectories $UpdatesDir)[0].Name) 'Developer only'
    $Timer.Start();Start-EditSession Public;$script:Word.Visible=$false;Add-TestText 'Developer only edit';$script:WordDocument.Save();Pump-Until {$null -eq $script:ArchiveProcess -and $StatusValue.Text -like 'Both versions saved at*'};$script:WordDocument.Close($false);Pump-Until {-not $script:SessionActive};$Timer.Stop()
    Assert-Test ((Tree-Hashes $Fixture.DataRoot) -eq $GeneralBefore -and (Get-FileHash $Fixture.WebsitePublicPath).Hash -eq $WebsiteBefore) 'profile edit and note do not touch General or its website'
    Set-ActiveProfile $Config;Save-ProfileRegistry
    Assert-Test ($ProfileCaption.Text -eq 'Profile: General') 'existing resume switches back without import or migration'
    Inspect-Dialog {Show-AboutDialog $Form} {param($Window,$All) Assert-Test (@($All | Where-Object {$_.Text -match ('Resume Manager '+[regex]::Escape($script:ProgramVersion))}).Count -gt 0) 'About displays single-source version'}
    Assert-Test ($script:ProgramVersion -and -not $ErrorPanel.Visible) 'final manager is ready without errors'
    $Form.PerformLayout();$Bitmap=New-Object Drawing.Bitmap($Form.Width,$Form.Height);try{$Form.DrawToBitmap($Bitmap,(New-Object Drawing.Rectangle(0,0,$Form.Width,$Form.Height)));$Bitmap.Save((Join-Path $TestRoot 'manager.png'))}finally{$Bitmap.Dispose()}
    Write-Output ('ALL FEATURE CHECKS PASSED. Synthetic QA retained at: '+$TestRoot)
}finally{
    $Timer.Stop();if($script:WordDocument){try{$script:WordDocument.Close($false)}catch{}};Release-WordObjects
    $Form.Dispose();$Timer.Dispose();try{$Mutex.ReleaseMutex()}catch{};$Mutex.Dispose()
}
