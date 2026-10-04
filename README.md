# Resume Manager

A small native Windows resume manager. Edit a DOCX in Microsoft Word; changed saves create dated private/public DOCX and PDF versions. No custom editor, web service, database, or installer.

## Privacy boundary

This repository contains **program code only**. No real resumes, PDFs, contact details, archives, local settings, or website assets belong here. All document processing stays on your computer. The program has no upload or Git functionality.

Personal settings and resume data must be outside the checkout. The program rejects settings, data roots, and website document destinations inside it. `.gitignore` uses an explicit source allowlist; a Git privacy check rejects other tracked files and obvious contact details. Hooks are a local guard, not an absolute guarantee: forced staging, disabled hooks, arbitrary private prose, and image contents still require human review.

## Requirements

- Windows PowerShell 5.1, Microsoft Word desktop, Python 3.10 or newer.
- The existing Python libraries in `requirements.txt`. If needed: `python -m pip install -r requirements.txt`.

## Local setup

Copy `settings.example.json` to `%LOCALAPPDATA%\ResumeManager\settings.local.json`, **outside this repository**, and replace its fake contact lines and anchor. Set `PrivateOnly` to the email, phone, or other contact values that must never be public. Keep the contact paragraph within the first six body paragraphs, normally below your name. `ContactAnchor` must appear in both contact lines, such as the shared email. Explicit `PrivateOnly` values and pipe-separated fields found only in `PrivateContact` are checked in public XML, relationships, and PDF text. Private information elsewhere (including images) must be removed manually; this is not a general-purpose anonymizer.

`DataRoot` defaults in the example to `%LOCALAPPDATA%\ResumeManager\Data`. Create its `Current` folder and copy your starting DOCX there as `resume-private.docx`. Then initialize both variants and PDFs:

```powershell
.\Scripts\resume_update_workflow.ps1 -SourcePath "$env:LOCALAPPDATA\ResumeManager\Data\Current\resume-private.docx"
```

Optional settings: `PythonExe`, `WebsitePublicPath`, `CurrentPrivateName`, `CurrentPublicName`, `ArchivePrivateName`, and `ArchivePublicName`. Names must be plain, distinct DOCX filenames. The website copy is optional; its destination folder must already exist. To use another local settings file, pass `-ConfigPath` to either PowerShell script, or set `RESUME_MANAGER_CONFIG` before launching. Never commit the filled settings.

Double-click **Resume Manager.cmd**. For a desktop shortcut, use Windows **Send to > Desktop (create shortcut)** on that launcher.

## Using the manager

The buttons open the private/public resume, Current folder, or version history. Only one manager instance and one Word editing session run at once. Edit normally, save, and allow a few seconds for stable content and PDF export. A safe temporary working copy lets both Current files refresh while Word is open. Repeated saves without content changes and unchanged closes do not duplicate archives. Saves during export are queued.

Each completed version is stored under `DataRoot\Updates\YYYY-MM-DD_HH-mm-ss`, using local system time. Same-second saves receive `-2`, `-3`, etc. Every archive contains private/public DOCX and PDFs; archived folders are never replaced or deleted. Both variants preserve package contents outside the contact paragraph. Rebuilding that paragraph retains its paragraph/run style, but replaces contact-line hyperlinks with plain text.

## Errors and recovery

The program displays exact workflow errors and supports **Retry save**. It builds and validates the archive before replacing Current or an optional website copy. Publish failures roll back replacements and retain the complete archive. Build failures retain `Updates\.failed-*`. Failed/interrupted editing sessions preserve the working copy under `%TEMP%\ResumeManagerSessions`; its path is shown on failure. Recover it with `resume_update_workflow.ps1 -SourcePath` and your external settings. Close Word and wait for saving to finish before closing the manager. `-NoPdf` is for troubleshooting only.

## Development and publishing

Keep tests synthetic; never copy a user's data into this checkout. Enable the local guards on each clone:

```powershell
git config core.hooksPath .githooks
python .\Scripts\test_resume_program.py
.\Scripts\check_repository.ps1 -History
```

The pre-commit guard checks the staged tree and pre-push checks every reachable commit, so deleting a leaked file in a later commit does not hide the old version. Review staged files and history before changing repository visibility. No public release or license is implied by this initial private repository.
