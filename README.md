# Resume Manager

A small native Windows program for editing resumes in Microsoft Word and keeping dated private/public DOCX and PDF versions. No custom editor, server, database, uploads, telemetry, or automatic dependency installation.

## Start and setup

Requires Windows PowerShell 5.1, Microsoft Word desktop, and Python 3.10 or newer with `lxml` and `pypdf`. Install missing libraries manually: `python -m pip install -r requirements.txt`.

Double-click **Resume Manager.cmd**. Valid existing settings open the manager directly. Missing settings open Simple setup; configuration problems, missing Current files, and unavailable Word/Python instead show startup help with Retry and Edit settings. Existing settings are not silently reset or resumes automatically reimported.

In Simple setup, choose a Word resume, read its contact details, and check only the details allowed in public. Fresh imports start with nothing selected for public use. Unchecked details stay private; existing configured private values are marked always private and cannot be selected. Review the exact private/public contact previews and tick the confirmation before Save and finish.

Advanced settings expands inside the same window. It contains the data folder, raw contact lines, shared contact anchor, never-public values, Python executable, optional website destination, and filename overrides. Expanding/collapsing does not reset values. Unknown existing JSON properties are preserved when saving. Contact fields use ` | ` separators; the anchor must occur in both lines and the first six body paragraphs of the source. Contact detection supports an unambiguous email contact line near the name and rejects unsupported markup without changing the source. Ambiguous layouts require manual configuration.

Initialization uses the existing workflow without editing the imported original. Import or contact/privacy changes in an existing data folder require confirmation and back up Current in `DataRoot\Recovery\before-replace-*`, then immediately create matching variants and a new archive. An unchanged settings save creates no archive. Cancelling or declining confirmation before processing leaves active settings/documents unchanged. Once a new version is saved, cancelling cannot undo it. If final settings publication fails, the new archive and recovery candidate settings are retained and the dialog explicitly reports that state.

Settings are written atomically at `%LOCALAPPDATA%\ResumeManager\settings.local.json`, outside the checkout. The Settings button changes them when no editing session is active. Advanced JSON options remain supported: `CurrentPrivateName`, `CurrentPublicName`, `ArchivePrivateName`, `ArchivePublicName`, `PythonExe`, and `WebsitePublicPath`. For another external configuration, pass `-ConfigPath` or set `RESUME_MANAGER_CONFIG`.

Create a desktop shortcut with Windows **Send to > Desktop (create shortcut)** on the launcher.

## Editing and preview

Edit either variant in Word: saved content updates both while contacts remain separate. Only one manager and editing session may run. The manager opens a temporary working copy so Current can refresh while Word remains open. Repeated unchanged saves and unchanged closes create no duplicate archives. Saves during export are queued. Save As is followed for DOCX only.

The main window shows readable archive dates, a concise save status and PDF page counts. Show details expands phases, elapsed time, queued saves, routine review warnings and exact errors. Privacy/review failures and multi-page notices remain immediately visible. PDF buttons use the system viewer and explicitly identify their saved archive date: they never preview unsaved Word edits. Missing PDFs are disabled. Multi-page warnings never change fonts, margins, or content. Shortcuts: Ctrl+P private editing, Ctrl+U public editing, Ctrl+H history; controls also support tab navigation and access keys.

## History and restoration

Version History lists dated folders and marks incomplete versions. Open a PDF/folder or Restore as New Version. Restore is blocked during editing and requires confirmation. It backs up Current byte-for-byte in Recovery, then creates a new archive from the selected public DOCX using today's contact settings. The selected archive is never changed. An old contact anchor that no longer matches today's setting causes a safe, reported failure.

Each complete `DataRoot\Updates\YYYY-MM-DD_HH-mm-ss` archive contains private/public DOCX and PDFs. Timestamps use local system time; collisions receive `-2`, `-3`, etc. Archives are never replaced or deleted.

## Privacy and fidelity

Real settings, documents, PDFs, logs, sessions, and archives stay outside the repository. The program rejects config/data/website destinations inside its checkout. Development and distribution use an explicit program-file allowlist.

Variants preserve package contents except the contact paragraph and relationships used exclusively by it. Retained fields preserve their run styles and hyperlinks. Missing email fields receive mail links when the source uses linked emails. Unrelated links are unchanged. Removed private links are deleted only if unreferenced elsewhere; a private target shared elsewhere fails validation. Complex fields or graphics in the contact line fail safely rather than being flattened.

Configured never-public values are checked in XML, text boxes, comments, metadata, relationships, and PDF text/metadata. Passing is **not an anonymity guarantee**. Review warnings flag comments, tracked changes, author metadata, embedded files, external targets, and images. The program does not OCR images, discover every unknown private detail, accept tracked changes, or silently scrub metadata. Review the public PDF and document before sharing.

Macro-enabled packages are rejected even when renamed to DOCX. Programmatic Word opening temporarily disables macros and link updates and restores prior settings afterward. Protected View and Trust Center policies are not bypassed.

## Errors and recovery

The persistent panel offers Retry save, Open recovery folder, and Copy redacted diagnostics. Setup uses field-specific messages, focus/error indicators, and optional technical details instead of raw setting-key errors. Entered values survive failures. Actual save races/sharing locks are temporary waits; package and dependency failures are reported. Redacted diagnostics remove configured contact values, emails, username, and local/network paths. Still review them before sharing.

Archives are validated before Current changes. Build failures retain `.failed-*`; publishing failures roll back replacements and retain the complete archive. Failed/interrupted working copies remain under `%TEMP%\ResumeManagerSessions`. Slow exports warn after two minutes and keep monitoring; Word is never killed automatically. Check Word for a dialog and wait; recovery paths remain local.

Website copying reports updated, disabled, or skipped when its folder is missing. A website replacement failure rolls publishing back. Close Word and allow saving to finish before closing the manager.

Manual recovery:

```powershell
.\Scripts\resume_update_workflow.ps1 -SourcePath "$env:TEMP\path-to-your-recovered.docx"
```

Without a source, the workflow uses Current public or the latest public archive. `-NoPdf` is troubleshooting only; history marks its output incomplete.

## Development and local distribution

Keep all fixtures synthetic and outside the checkout. Enable hooks on each clone:

```powershell
git config core.hooksPath .githooks
python .\Scripts\test_resume_program.py
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

Before public release:

- MIT license is included by the owner's choice; retain LICENSE in distributed copies.
- Test a fresh extraction, absent prerequisites, cancelled setup, and synthetic import.
- Test both edit paths, repeated saves, queued saves, failure/retry, unchanged close, restore, and public privacy.
- Inspect Word-rendered PDFs and UI at different display scaling levels.
- Review ZIP contents, source, staged files, commit metadata, and all reachable history.
- Obtain explicit approval before committing, pushing, publishing, or changing visibility.

The repository remains private. A local source ZIP is a release candidate, not a published release.
