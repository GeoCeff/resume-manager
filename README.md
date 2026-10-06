# Resume Manager

A small native Windows program for editing resumes in Microsoft Word and keeping dated private/public DOCX and PDF versions. No custom editor, server, database, uploads, telemetry, or automatic dependency installation.

## Download

Download the program-only ZIP from [Releases](https://github.com/GeoCeff/resume-manager/releases/latest), extract it into a local folder, then double-click **Resume Manager.cmd**. Keep all extracted files together. The source distribution requires the prerequisites below; it does not bundle Word or Python and is not an installer. Version 1.1.0 is MIT licensed.

Your own settings and resumes are created locally, outside the program folder. No real resume, filled settings, or personal history is included in the download. When updating, close the manager after saving finishes, extract the new version into a separate folder, and launch it with your existing settings. Use About / Help to update the desktop shortcut. Keep the previous program folder until the new version works.

## Start and setup

Requires Windows PowerShell 5.1, Microsoft Word desktop, and Python 3.10 or newer with `lxml` and `pypdf`. Install missing libraries manually: `python -m pip install -r requirements.txt`.

Double-click **Resume Manager.cmd**. Valid settings open the manager directly. Missing settings offer **Locate existing settings** or **Create new setup**. Problems show Retry, Locate existing settings, and Edit settings. Nothing is silently reset or reimported. Locate can select existing settings, a validated settings backup, or a retained recovery candidate, without copying or changing its contents.

In Simple setup, choose a Word resume, read its contact details, and check only the details allowed in public. Fresh imports start with nothing selected for public use. Unchecked details stay private; existing configured private values are marked always private and cannot be selected. Review the exact private/public contact previews and tick the confirmation before Save and finish.

Advanced settings expands inside the same window. It contains the data folder, raw contact lines, shared contact anchor, never-public values, Python executable, optional website destination, and filename overrides. Expanding/collapsing does not reset values. Unknown existing JSON properties are preserved when saving. Contact fields use ` | ` separators; the anchor must occur in both lines and the first six body paragraphs of the source. Contact detection supports an unambiguous email contact line near the name and rejects unsupported markup without changing the source. Ambiguous layouts require manual configuration.

Initialization uses the existing workflow without editing the imported original. Import or contact/privacy changes in an existing data folder require confirmation and back up Current in `DataRoot\Recovery\before-replace-*`, then immediately create matching variants and a new archive. An unchanged settings save creates no archive. Cancelling or declining confirmation before processing leaves active settings/documents unchanged. Once a new version is saved, cancelling cannot undo it. If final settings publication fails, the new archive and recovery candidate settings are retained and the dialog explicitly reports that state.

Settings are written atomically at `%LOCALAPPDATA%\ResumeManager\settings.local.json`, outside the checkout. The Settings button changes them when no editing session is active. Advanced JSON options remain supported: `CurrentPrivateName`, `CurrentPublicName`, `ArchivePrivateName`, `ArchivePublicName`, `PythonExe`, and `WebsitePublicPath`. For another external configuration, pass `-ConfigPath` or set `RESUME_MANAGER_CONFIG`.

Use **About / Help > Create / Update Desktop Shortcut** for a shortcut to this exact program and active settings. An old shortcut is backed up beside settings in `Shortcut Backups`. A shortcut pins its selected settings; the command launcher remembers the last profile unless `-ConfigPath` or an environment override is supplied.

## Editing and preview

Edit either variant in Word: saved content updates both while contacts remain separate. Only one manager and editing session may run. The manager opens a temporary working copy so Current can refresh while Word remains open. Repeated unchanged saves and unchanged closes create no duplicate archives. Saves during export are queued. Save As is followed for DOCX only.

The window distinguishes **Unsaved changes in Word**, **Word saved; generating versions**, **Both versions saved at [time]**, **Waiting for Word**, and **Save failed; working copy retained**. Last successful snapshot and queued saves stay visible. Typing while an older save exports keeps the unsaved warning visible; an earlier snapshot does not imply the newest Word text was saved. Errors persist until retry/success or an explicit new operation. Show details expands phases, elapsed time, warnings and exact errors. PDF previews identify their archive date, never unsaved edits. Multi-page warnings never change layout. Shortcuts: Ctrl+P private editing, Ctrl+U public editing, Ctrl+H history; controls support tab navigation and access keys.

```text
Edit Private/Public -> Word opens -> Edit and Ctrl+S
                                     |
                                     v
                         Validated dated DOCX + PDF pair
                                     |
                          Current updated, then Export Copy
```

## Export for an application

**Export Resume** defaults to Public PDF. Choose a dated version, Public/Private, PDF/DOCX, and a company/role filename. It copies validated bytes without changing Current, history, settings, or the website. Old versions that fail today's contact/privacy policy are rejected, not silently substituted. Confirm public review and overwrites. Protected program/data/settings/recovery destinations and unsafe Windows filenames are rejected. The temporary copy is checked before atomic publication. Review warnings, images and unknown private details still need human review; this is not an anonymity guarantee.

## History and restoration

Version History lists dated folders and marks incomplete versions. Open a PDF/folder or Restore as New Version. Restore is blocked during editing and requires confirmation. It backs up Current byte-for-byte in Recovery, then creates a new archive from the selected public DOCX using today's contact settings. The selected archive is never changed. An old contact anchor that no longer matches today's setting causes a safe, reported failure.

Select one version and **Label Version** to add a short local note. Labels are written atomically to `DataRoot\version-notes.local.json`, outside immutable archives. Hold Ctrl to select two versions, choose Public or Private, and **Compare Two Versions**. The read-only view compares body text including tables, not formatting, images, headers, or layout. Use PDF previews for visual comparison. Neither labels nor comparisons change documents.

## Optional profiles

**Profiles...** is optional. General registers existing settings in place without moving/importing documents. Create from a chosen Word resume or explicitly copy the active resume into an independent folder. Both routes reuse setup and require public-contact review. Website copying starts disabled; enable it deliberately in that profile's Advanced settings.

Each profile owns settings, Current, Updates, Recovery, and notes. `profiles.local.json` beside initial settings stores names, settings paths, and the last active path; prior registry bytes are backed up before changes. Data folders cannot overlap, use junctions, or be the program/Windows/home/drive root. Switching is blocked during editing, export, restore, or recovery; the global application/workflow locks still apply. No profiles or archives are automatically deleted. Repair missing settings through Locate existing settings or the local registry with the app closed.

Each complete `DataRoot\Updates\YYYY-MM-DD_HH-mm-ss` archive contains private/public DOCX and PDFs. Timestamps use local system time; collisions receive `-2`, `-3`, etc. Archives are never replaced or deleted.

## Privacy and fidelity

Real settings, documents, PDFs, logs, sessions, and archives stay outside the repository. The program rejects config/data/website destinations inside its checkout. Development and distribution use an explicit program-file allowlist.

Variants preserve package contents except the contact paragraph and relationships used exclusively by it. Retained fields preserve their run styles and hyperlinks. Missing email fields receive mail links when the source uses linked emails. Unrelated links are unchanged. Removed private links are deleted only if unreferenced elsewhere; a private target shared elsewhere fails validation. Complex fields or graphics in the contact line fail safely rather than being flattened.

Configured never-public values are checked in XML, text boxes, comments, metadata, relationships, and PDF text/metadata. Passing is **not an anonymity guarantee**. Review warnings flag comments, tracked changes, author metadata, embedded files, external targets, and images. The program does not OCR images, discover every unknown private detail, accept tracked changes, or silently scrub metadata. Review the public PDF and document before sharing.

Macro-enabled packages are rejected even when renamed to DOCX. Programmatic Word opening temporarily disables macros and link updates and restores prior settings afterward. Protected View and Trust Center policies are not bypassed.

## Errors and recovery

The persistent panel offers Retry save, Open recovery folder, and Copy redacted diagnostics. Setup uses field-specific messages, focus/error indicators, and optional technical details instead of raw setting-key errors. Entered values survive failures. Actual save races/sharing locks are temporary waits; package and dependency failures are reported. Redacted diagnostics remove configured contact values, emails, username, and local/network paths. Still review them before sharing.

Archives are validated before Current changes. Build failures retain `.failed-*`; publishing failures roll back replacements and retain the complete archive. Failed/interrupted working copies remain under `%TEMP%\ResumeManagerSessions`. Slow exports warn after two minutes and keep monitoring; Word is never killed automatically. Check Word for a dialog and wait; recovery paths remain local.

**Recover Saved Copy** lists changed saved copies for the active settings/data folder with source, variant and local save time. It can also explicitly select an older saved DOCX. Confirm recovery: Current is backed up, a new pair is generated with today's contact settings, and the source is retained. Unsaved Word text cannot be recovered. Incomplete packages are kept for manual inspection. Windows may clear TEMP; copy important recovery files to a permanent local folder promptly.

Confirmed settings changes back up the previous validated settings in `Settings Backups` beside the settings file before atomic replacement. Unknown JSON fields and UTF-8 text survive. Backups are not automatically deleted. Cancelling startup/setup/recovery/export does not reset settings or alter existing documents.

Website copying reports updated, disabled, or skipped when its folder is missing. A website replacement failure rolls publishing back. Close Word and allow saving to finish before closing the manager.

Manual recovery:

```powershell
.\Scripts\resume_update_workflow.ps1 -SourcePath "$env:TEMP\path-to-your-recovered.docx"
```

Without a source, the workflow uses Current public or the latest public archive. `-NoPdf` is troubleshooting only; history marks its output incomplete.

## Development and distribution

Keep all fixtures synthetic and outside the checkout. Enable hooks on each clone:

```powershell
git config core.hooksPath .githooks
python .\Scripts\test_resume_program.py
powershell.exe -NoProfile -STA -File .\Scripts\test_manager_features.ps1 -PythonExe python.exe
& '.\Scripts\Resume Manager.ps1' -SelfTest
.\Scripts\check_repository.ps1 -WorkingTree
.\Scripts\check_repository.ps1 -History
```

Pre-commit checks staged files; pre-push checks reachable history. Guards reject non-program files, non-example emails, possible phone numbers, and absolute machine paths. They cannot guarantee safety if bypassed or private prose/images are added manually.

Build a source ZIP outside the checkout:

```powershell
.\Scripts\package_program.ps1 -OutputDirectory "$env:LOCALAPPDATA\ResumeManager\Distribution"
```

The packager audits source and history, includes only the allowlist, and checks ZIP names and text contents. It excludes `.git`, filled settings, documents, archives, logs, caches, and recovery files. Nothing is uploaded.

**About / Help** shows the version from one source value and MIT licensing. **Create Example Resume...** generates a purely synthetic example outside program/data/settings folders, without a personal Word template or new dependency. Import it into a separate test profile, not over real Current files. Generation code is shipped; DOCX/PDF fixtures are excluded from the source ZIP. The instructions above contain no personal screenshots or machine-specific paths.

Release checklist:

- MIT license is included by the owner's choice; retain LICENSE in distributed copies.
- Test a fresh extraction, absent prerequisites, cancelled setup, and synthetic import.
- Test both edit paths, repeated saves, queued saves, failure/retry, unchanged close, restore, and public privacy.
- Inspect Word-rendered PDFs and UI at different display scaling levels.
- Review ZIP contents, source, staged files, commit metadata, and all reachable history.
- Inspect release notes, GitHub attachments, and commit/tag identities as well as file contents; use GitHub no-reply commit identities.

Only audited program/support files may be published. A locally generated ZIP is a release candidate until it is uploaded to a published GitHub Release. Do not attach personal resumes, filled settings, local diagnostics, or test reports to issues or releases.
