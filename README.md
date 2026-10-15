---
tags: [office, tool]
---
# office-install-tool

One project for installing Office and fixing failed Office installs: a universal installer that generates the ODT configuration, plus the diagnose and cleanup scripts. No product key is stored anywhere in it.

Plans are in [roadmap.md](roadmap.md).

## Installer (project root)

Universal Office installer for Windows. It builds the Office Deployment Tool (ODT) configuration for the edition you pick, downloads the ODT from Microsoft if it is missing, and runs the install.

No product key is stored anywhere in this project.

## Quick start

Open PowerShell (Start, type `powershell`, Enter) and paste:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
irm https://raw.githubusercontent.com/SheepasTech/office-install-tool/main/bootstrap.ps1 | iex
```

`-Scope Process` lifts the script policy for this window only. The bootstrap downloads the tool to `%TEMP%\office-install-tool`, asks for administrator rights and starts `Install-Office.ps1`. Pin a branch, tag or commit with `$env:OIT_REF`, or read `bootstrap.ps1` first: `irm | iex` runs whatever the URL serves.

## Use

Double-click `Start-OfficeInstaller.cmd` (it asks for administrator rights), or from an administrator PowerShell:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Install-Office.ps1
```

Interactive runs open a short menu first: install now, read-only readiness check, scan for leftovers, or remove leftovers (the diagnose and cleanup scripts). After an install, a one-line job record (no key) goes to `Cases\<date>_<PC>\install-jobs.log`, and a failed setup prints a plain-words likely cause.

The script asks for: edition, language(s), apps to exclude, MSI removal, installer window, and a product key.

**Product key:** typed with hidden input. Press Enter to skip, then activate in Word with *Enter product key*. A typed key goes only into a temporary config file, which is deleted afterwards, and into the setup log, which the script redacts afterwards.

## Parameters

| Parameter | Meaning |
|---|---|
| `-Product` | ODT product ID, e.g. `HomeStudent2024Retail`. Without it the script is interactive. |
| `-Language` | `lt-lt,en-us`, first is primary. Default `lt-lt`. |
| `-Arch` | `64` (default) or `32`. |
| `-ExcludeApp` | `Teams,OneDrive`, or `none`. Default `Teams`. |
| `-Display` | `Full` (default) or `None`. |
| `-RemoveMSI` | Remove older MSI-based Office. |
| `-AskKey` / `-NoKey` | Force or skip the key prompt. |
| `-GenerateOnly` | Write the XML to `output\` and stop. Never includes a key. |
| `-DownloadOnly` | Download and verify the ODT into `ODT\` and stop. |
| `-OdtPath` | Folder that already holds `setup.exe`. |

Examples:

```powershell
.\Install-Office.ps1 -Product HomeStudent2024Retail -Language lt-lt,en-us -GenerateOnly
.\Install-Office.ps1 -DownloadOnly
```

## Editions

Home & Student / Home & Business / Home / Professional 2024, Home & Student / Home & Business / Professional 2021, LTSC Professional Plus 2024 and 2021 (volume), Microsoft 365 Apps (enterprise, business, Family/Personal). Add more in the `$Catalog` list at the top of the script.

## ODT download

`-DownloadOnly`, or any install without a `setup.exe` present, fetches the ODT from the official Microsoft download page, checks it is signed by Microsoft Corporation, and extracts `setup.exe` to `ODT\`. `setup.exe` is searched in `-OdtPath`, `ODT\`, the script folder, then `C:\ODT`.

# Diagnose, cleanup and install kit

## Folders

| Folder | Contents |
|---|---|
| `1-Diagnose` | `Test-OfficeInstallReadiness.ps1`: read-only checks (disk, file system, write test, services, antivirus/filter drivers, event log, network, Office leftovers, setup log errors). |
| `2-Cleanup` | `Remove-OfficeLeftovers.ps1`: finds all Office installs and leftovers, lists them, removes them only after you type `YES`. Backs up the registry first. |
| `3-Install` | Older fixed-config installer (`Install-Office.cmd` and two XML files, no key). A simple fallback; `Install-Office.ps1` replaces it. |
| `_Reference` | ODT sample configuration and old notes. |
| `Cases` | Logs from past jobs, one folder per PC. Product keys in these logs are redacted. |

## Running the PowerShell scripts

Open **Windows PowerShell as administrator** (64-bit, not "x86"), then:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
cd <USB drive>:\office-install-tool
```

Both scripts save their output to a new folder on the **client's Desktop**
(`OfficeDiag_<date>` / `OfficeCleanup_<date>`). Copy those into `Cases\<date>_<PC name>\`
when you're done.

## Install procedure

1. **Clean up**, if Office or a Microsoft 365 trial was ever on the PC:
   `.\2-Cleanup\Remove-OfficeLeftovers.ps1 -ScanOnly` to look, then
   `.\2-Cleanup\Remove-OfficeLeftovers.ps1` to remove. Restart Windows.
2. **Install without the key**: right-click `3-Install\Install-Office.cmd` >
   *Run as administrator* and choose **1** (Lithuanian) or **2** (English).
3. **Activate**: open Word, choose *Enter product key* (not *Sign in*), and enter the key.
4. If it fails, run `.\1-Diagnose\Test-OfficeInstallReadiness.ps1` straight away
   (`-Quick` skips the two slow scans). It copies the newest setup logs into its report folder.

## Error notes

| Error | Meaning / what to do |
|---|---|
| `30175-45 (1289)` followed by `0-2031 (17004)` within seconds, with no download | Setup failed at the *apply configuration* step, where the product key is applied. The key was rejected: install with a `NOKEY` config and enter the key in Word. If Word rejects it too, the key is not a valid retail key for Home & Student 2024, or it can only be redeemed at setup.office.com (Microsoft account required). See `Cases\2026-10-08_LAPTOP-KV76T046`. |
| `30175-45` after a long download / at 70-80% | Real install failure. Run Cleanup, restart, try again. If it repeats, check the Diagnose report (write tests, disk events, third-party antivirus/filter drivers). |
| `0-2031 (17004)` alone | The setup program reporting that the install failed. Look for the error that came before it. |
| Office "Couldn't install" on a fresh laptop | Usually a preinstalled Microsoft 365 trial or Office Store app. Run Cleanup first. |

## Security

- Never put a product key in a config file here. To install with a key, use `Install-Office.ps1` in the project root, which asks for it at run time and wipes it afterwards.
- Download the Office Deployment Tool only from microsoft.com. Never use an installer supplied by a key seller.
