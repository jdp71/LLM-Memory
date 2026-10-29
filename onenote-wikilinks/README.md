# OneNote wiki links

Converts `[[Page Title]]` markers in OneNote pages into real links to the page with that title.

- `[[Linux]]` links to the page titled "Linux" (case-insensitive).
- `[[Linux|the kernel]]` links to the same page but displays "the kernel".

## Why it works on a locked-down work/government laptop

It talks to **OneNote desktop through its local COM interface**. There is no Azure app registration,
no admin consent, no tokens, no network calls, and nothing to install. It uses whatever account
OneNote is already signed in with.

Requirements:
- Windows with **OneNote for Microsoft 365 (desktop)**. The "OneNote for Windows 10" app does not expose COM.
- PowerShell (built into Windows). If script execution is restricted, run it with
  `powershell -ExecutionPolicy Bypass -File .\Update-WikiLinks.ps1` only if your policy allows that.
  Otherwise ask IT, or paste the script into an interactive PowerShell session.

Check your agency's policy on running custom scripts before using it.

## Usage

Open OneNote and let it finish syncing, then:

```powershell
# 1. Dry run on one section (changes nothing, prints what would happen)
.\Update-WikiLinks.ps1 -Section "Scratch"

# 2. Apply to that section
.\Update-WikiLinks.ps1 -Section "Scratch" -Apply

# 3. Whole account
.\Update-WikiLinks.ps1 -Apply
```

Options: `-Notebook`, `-Section` (wildcards ok) limit which pages are *edited*. Link targets are
always searched across every notebook. `-BackupDir` sets where pre-change page XML is saved
(default `.\onenote-backup`); `-NoBackup` skips them. Delete backups when you're satisfied,
especially if the content is sensitive.

## Behavior

- Markers with no matching page, or with a title shared by multiple pages, are left untouched and reported.
- Existing hyperlinks and self-links are never modified.
- Only pages that contain a marker are rewritten.
- A marker split by formatting (e.g. `[[Li<b>nux</b>]]`) is not matched. Retype it plainly.
- OneNote desktop may itself turn `[[Title]]` into a link as you type; those are already links and are skipped.

## Status

Written without access to a Windows/OneNote machine, so it is **untested against real OneNote**.
Run the dry run on a throwaway section first.
