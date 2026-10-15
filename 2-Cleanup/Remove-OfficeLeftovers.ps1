#Requires -Version 5.1
<#
.SYNOPSIS
    Finds every trace of Microsoft Office / Microsoft 365 on this PC, lists it,
    and - only after you type YES - removes it.

.DESCRIPTION
    Intended to fix failed Office installs such as error 30175-45, which are
    usually caused by leftovers of a previous or preinstalled Office.

    PHASE 1 - SCAN (read-only, nothing is changed)
        * Installed Office products ("Installed apps" registry entries, 64-bit,
          32-bit and per-user), with their type: Click-to-Run, MSI or orphaned
        * Click-to-Run product registration
        * Installed Office license keys (via OSPP.VBS, if present)
        * Running Office processes
        * Office services (ClickToRunSvc, OfficeSvc, osppsvc)
        * Office scheduled tasks (\Microsoft\Office\)
        * Start menu shortcuts pointing into Office folders
        * Office program / data folders (with size and file count)
        * Office registry keys

    PHASE 2 - CONFIRM
        Everything found is listed. Nothing is removed unless you type YES.

    PHASE 3 - REMOVE (in this order)
        1. Back up every affected registry key to .reg files
        2. Close Office apps
        3. Remove installed Office license keys
        4. Uninstall Click-to-Run products with Microsoft's own uninstaller
        5. Uninstall MSI-based Office products with msiexec
        6. Stop and delete Office services
        7. Delete Office scheduled tasks
        8. Delete Office shortcuts
        9. Delete leftover folders (taking ownership if needed)
       10. Delete leftover registry keys and orphaned "Installed apps" entries
       11. Re-scan and report what (if anything) is still left

    Your documents are never touched. User Office settings
    (HKCU\Software\Microsoft\Office and the local Office cache) are only
    removed if you pass -IncludeUserSettings.

    A full log (transcript), the registry backups, a JSON list of findings and
    MSI uninstall logs are saved to a folder under %TEMP%\Debug\office-install-tool\<date>.

.PARAMETER ScanOnly
    Only scan and list. Never removes anything.

.PARAMETER IncludeUserSettings
    Also remove HKCU\Software\Microsoft\Office and
    %LOCALAPPDATA%\Microsoft\Office\16.0 (your Office preferences, recent file
    lists, caches) for the current user. Documents are not affected.

.PARAMETER BackupPath
    Where to store the log and backups.
    Default: %TEMP%\Debug\office-install-tool\<date>\OfficeCleanup_<date>_<time>

.EXAMPLE
    .\Remove-OfficeLeftovers.ps1 -ScanOnly
    Just shows what would be removed.

.EXAMPLE
    .\Remove-OfficeLeftovers.ps1
    Scans, lists, asks for confirmation, then removes.

.NOTES
    Run from an elevated (Run as administrator) 64-bit Windows PowerShell.
    If PowerShell blocks the script, first run:
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
    Restart Windows afterwards, before reinstalling Office.
#>

[CmdletBinding()]
param(
    [switch]$ScanOnly,
    [switch]$IncludeUserSettings,
    [string]$BackupPath
)

Set-StrictMode -Version 2.0

# Debug log (own file per window) + license discovery. Both are optional: a missing lib must not stop the cleanup.
try { . (Join-Path $PSScriptRoot '..\lib\Debug-Log.ps1'); . (Join-Path $PSScriptRoot '..\lib\Office-Licenses.ps1') }
catch {
    Write-Warning "Debug/license helpers unavailable: $($_.Exception.Message)"
    function Start-DebugLog { param($Name) }; function Stop-DebugLog {}; function Write-DebugCrash { param($ErrorRecord) }
}
trap { Write-DebugCrash $_; break }
$ErrorActionPreference = 'Continue'

if (-not $BackupPath) {
    $BackupPath = Join-Path (Join-Path $env:TEMP ('Debug\office-install-tool\' + (Get-Date -Format 'yyyy-MM-dd'))) ('OfficeCleanup_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
}

# --------------------------------------------------------------------------
#  Settings
# --------------------------------------------------------------------------

# "Installed apps" display names that count as Office
$ProductNamePattern = '^(Microsoft 365|Microsoft Office|Office 16 Click-to-Run|Microsoft (OneNote|Project|Visio|Access|Excel|Word|PowerPoint|Outlook|Publisher)\b|Skype for Business)'
# ...but never these (used by other software, or not Office itself)
$ProductExcludePattern = 'Access database engine|Office Deployment Tool|Visual C\+\+|Visual Studio'

# Office apps closed before uninstalling
$OfficeAppProcesses = @('WINWORD','EXCEL','POWERPNT','OUTLOOK','ONENOTE','ONENOTEM','MSACCESS',
                        'MSPUB','VISIO','WINPROJ','lync','SDXHelper','MSOSYNC','MSOUC','msoia','MSOADFSB')
# Click-to-Run engine processes, closed only AFTER the uninstaller has run
$OfficeEngineProcesses = @('OfficeC2RClient','OfficeClickToRun','IntegratedOffice','AppVShNotify','OSPPSVC')

$OfficeServices = @('ClickToRunSvc','OfficeSvc','osppsvc')

$C2RExe = Join-Path $env:CommonProgramFiles 'Microsoft Shared\ClickToRun\OfficeClickToRun.exe'

$CategoryOrder = @('Installed product','Click-to-Run configuration (info)','Office license key','Office license (info)',
                   'Running process','Service','Scheduled task','Shortcut','Folder','Registry key')

$script:Results = New-Object System.Collections.Generic.List[object]

# --------------------------------------------------------------------------
#  Helpers
# --------------------------------------------------------------------------

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO','DETAIL','FOUND','OK','WARN','ERROR','STEP')][string]$Level = 'INFO'
    )
    $colors = @{ INFO = 'Gray'; DETAIL = 'DarkGray'; FOUND = 'Yellow'; OK = 'Green'
                 WARN = 'DarkYellow'; ERROR = 'Red'; STEP = 'Cyan' }
    $line = '[{0}] [{1,-6}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    Write-Host $line -ForegroundColor $colors[$Level]
}

function Write-Section {
    param([string]$Title)
    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host ('=' * 78) -ForegroundColor Cyan
}

function Get-Prop {
    param($Object, [string]$Name)
    if ($null -ne $Object -and $Object.PSObject.Properties[$Name]) { return $Object.$Name }
    return $null
}

function Format-Size {
    param([double]$Bytes)
    if ($Bytes -ge 1GB) { return '{0:N2} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:N1} MB' -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return '{0:N0} KB' -f ($Bytes / 1KB) }
    return ('{0} bytes' -f $Bytes)
}

function New-Finding {
    param([string]$Category, [string]$Name, [string]$Target, [string]$Detail, $Data)
    Write-Log ("Found {0}: {1}" -f $Category.ToLower(), $Name) 'FOUND'
    [pscustomobject]@{ Category = $Category; Name = $Name; Target = $Target; Detail = $Detail; Data = $Data }
}

function Add-Result {
    param([string]$Action, [string]$Target, [ValidateSet('OK','FAILED','SKIPPED')][string]$Status, [string]$Note = '')
    $script:Results.Add([pscustomobject]@{ Action = $Action; Target = $Target; Status = $Status; Note = $Note })
}

function Remove-FolderForce {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log "Already gone: $Path" 'DETAIL'
        return $true
    }
    try {
        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
        Write-Log "Deleted: $Path" 'OK'
        return $true
    }
    catch {
        Write-Log ("First attempt failed: {0}" -f $_.Exception.Message) 'WARN'
        Write-Log 'Taking ownership (Administrators) and granting full control, then retrying...' 'DETAIL'
    }
    # SID-based so it works on any Windows display language
    & icacls.exe "$Path" /setowner '*S-1-5-32-544' /T /C /Q 2>&1 | Out-Null
    & icacls.exe "$Path" /grant '*S-1-5-32-544:(OI)(CI)F' /T /C /Q 2>&1 | Out-Null
    try {
        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
        Write-Log "Deleted after taking ownership: $Path" 'OK'
        return $true
    }
    catch {
        Write-Log ("Could not delete {0}: {1}" -f $Path, $_.Exception.Message) 'ERROR'
        Write-Log 'Usually a file is locked. Restart Windows and run this script again.' 'WARN'
        return $false
    }
}

# --------------------------------------------------------------------------
#  Scan
# --------------------------------------------------------------------------

function Get-OfficeFindings {
    $found = New-Object System.Collections.Generic.List[object]

    # 1. Installed products --------------------------------------------------
    Write-Log 'Scanning "Installed apps" registry entries (64-bit, 32-bit, per-user)...' 'STEP'
    $roots = @(
        'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    foreach ($root in $roots) {
        $rootPath = "Registry::$root"
        if (-not (Test-Path -LiteralPath $rootPath)) { Write-Log "Not present: $root" 'DETAIL'; continue }
        $keys = @(Get-ChildItem -LiteralPath $rootPath -ErrorAction SilentlyContinue)
        Write-Log ("Checking {0} entries in {1}" -f $keys.Count, $root) 'DETAIL'
        foreach ($key in $keys) {
            $p = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
            $name = Get-Prop $p 'DisplayName'
            if (-not $name) { continue }
            if ($name -notmatch $ProductNamePattern) { continue }
            if ($name -match $ProductExcludePattern) {
                Write-Log "Ignoring (not Office itself / used by other software): $name" 'DETAIL'
                continue
            }
            $uninstall = Get-Prop $p 'UninstallString'
            $isMsi     = ((Get-Prop $p 'WindowsInstaller') -eq 1)
            $isSysComp = ((Get-Prop $p 'SystemComponent') -eq 1)

            if ($uninstall -and $uninstall -match 'OfficeClickToRun\.exe') { $kind = 'ClickToRun' }
            elseif ($isMsi -and $key.PSChildName -match '^\{[0-9A-Fa-f-]{36}\}$') { $kind = 'MSI' }
            else { $kind = 'Orphaned entry' }

            if ($kind -eq 'MSI' -and $isSysComp) {
                Write-Log "Skipping MSI sub-component (removed together with its parent product): $name" 'DETAIL'
                continue
            }
            $detail = 'Type: {0} | Version: {1}' -f $kind, (Get-Prop $p 'DisplayVersion')
            if ($uninstall) { $detail += " | Uninstall command: $uninstall" }
            $found.Add((New-Finding 'Installed product' $name $key.Name $detail @{
                Kind = $kind; UninstallString = $uninstall; ProductCode = $key.PSChildName; RegKey = $key.Name }))
        }
    }

    # 2. Click-to-Run registration (info) -------------------------------------
    Write-Log 'Checking Click-to-Run product registration...' 'STEP'
    $cfgKey = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if (Test-Path -LiteralPath "Registry::$cfgKey") {
        $cfg = Get-ItemProperty -LiteralPath "Registry::$cfgKey" -ErrorAction SilentlyContinue
        $ids = Get-Prop $cfg 'ProductReleaseIds'
        $detail = 'Products: {0} | Version: {1} | Platform: {2}' -f $ids, (Get-Prop $cfg 'VersionToReport'), (Get-Prop $cfg 'Platform')
        $found.Add((New-Finding 'Click-to-Run configuration (info)' 'Click-to-Run registration' $cfgKey $detail @{ ProductReleaseIds = $ids }))
    }
    else { Write-Log 'No Click-to-Run registration found.' 'DETAIL' }
    if (Test-Path -LiteralPath $C2RExe) { Write-Log "Click-to-Run engine present: $C2RExe" 'DETAIL' }

    # 3. License keys ---------------------------------------------------------
    Write-Log 'Checking installed Office license keys (OSPP.VBS)...' 'STEP'
    $osppList = @()
    foreach ($base in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if (-not $base) { continue }
        $candidate = Join-Path $base 'Microsoft Office\Office16\OSPP.VBS'
        if (Test-Path -LiteralPath $candidate) { $osppList += $candidate }
    }
    if ($osppList.Count -eq 0) { Write-Log 'OSPP.VBS not found - no Office license keys to query.' 'DETAIL' }
    $seenKeys = @{}
    foreach ($ospp in $osppList) {
        Write-Log "Running: cscript //Nologo `"$ospp`" /dstatus" 'DETAIL'
        $out = (& cscript.exe //Nologo "$ospp" /dstatus 2>&1) | Out-String
        foreach ($block in ($out -split '-{10,}')) {
            if ($block -match 'Last 5 characters of installed product key:\s*([A-Za-z0-9]{5})') {
                $last5 = $Matches[1]
                if ($seenKeys.ContainsKey($last5)) { continue }
                $seenKeys[$last5] = $true
                $lic = ''; $status = ''
                if ($block -match 'LICENSE NAME:\s*(.+)')   { $lic = $Matches[1].Trim() }
                if ($block -match 'LICENSE STATUS:\s*(.+)') { $status = $Matches[1].Trim() }
                $found.Add((New-Finding 'Office license key' "Key ending in $last5" $ospp "License: $lic | Status: $status" @{
                    Last5 = $last5; Ospp = $ospp }))
            }
        }
    }

    # 3b. Full license picture: every key/subscription and the install it belongs to
    Write-Log 'Looking for all Office licenses (licensing service, registry, account) and the install each belongs to...' 'STEP'
    try {
        $lic = Get-OfficeLicenseInfo
        foreach ($i in $lic.Installs) {
            Write-Log ('Install: {0} | {1} | version {2} | {3} {4}' -f $i.Kind, $i.Name, $i.Version, $i.Platform, $i.Channel) 'DETAIL'
        }
        foreach ($l in $lic.Licenses) {
            $tail = if ($l.Last5) { " key ending $($l.Last5)" } else { '' }
            $detail = 'Source: {0} | Channel: {1} | Status: {2} | Belongs to: {3}' -f $l.Source, $l.Channel, $l.Status, $l.BelongsTo
            if ($l.Description) { $detail += " | $($l.Description)" }
            if ($l.Note) { $detail += " | $($l.Note)" }
            $ospp = @($found | Where-Object { $_.Category -eq 'Office license key' -and $l.Last5 -and $_.Data.Last5 -eq $l.Last5 })
            if ($ospp.Count -gt 0) { $ospp[0].Detail += " | Channel: $($l.Channel) | Belongs to: $($l.BelongsTo)"; continue }
            $found.Add((New-Finding 'Office license (info)' ("$($l.Name)$tail") $l.Source $detail @{ License = $l }))
        }
    }
    catch { Write-Log "License discovery failed: $($_.Exception.Message)" 'WARN' }

    # 4. Running processes ----------------------------------------------------
    Write-Log 'Checking for running Office processes...' 'STEP'
    $procs = @(Get-Process -Name ($OfficeAppProcesses + $OfficeEngineProcesses) -ErrorAction SilentlyContinue)
    if ($procs.Count -eq 0) { Write-Log 'No Office processes running.' 'DETAIL' }
    foreach ($pr in $procs) {
        $path = ''
        try { $path = $pr.Path } catch { }
        $found.Add((New-Finding 'Running process' ("{0} (PID {1})" -f $pr.ProcessName, $pr.Id) $path '' @{
            Id = $pr.Id; ProcessName = $pr.ProcessName }))
    }

    # 5. Services -------------------------------------------------------------
    Write-Log 'Checking for Office services...' 'STEP'
    foreach ($svcName in $OfficeServices) {
        $svc = Get-CimInstance -ClassName Win32_Service -Filter "Name='$svcName'" -ErrorAction SilentlyContinue
        if ($svc) {
            $found.Add((New-Finding 'Service' ("{0} ({1})" -f $svc.Name, $svc.DisplayName) $svc.PathName `
                ("State: {0} | Start mode: {1}" -f $svc.State, $svc.StartMode) @{ ServiceName = $svc.Name }))
        }
        else { Write-Log "Service not present: $svcName" 'DETAIL' }
    }

    # 6. Scheduled tasks ------------------------------------------------------
    Write-Log 'Checking scheduled tasks under \Microsoft\Office\ ...' 'STEP'
    $tasks = @(Get-ScheduledTask -TaskPath '\Microsoft\Office\*' -ErrorAction SilentlyContinue)
    if ($tasks.Count -eq 0) { Write-Log 'No Office scheduled tasks.' 'DETAIL' }
    foreach ($t in $tasks) {
        $found.Add((New-Finding 'Scheduled task' $t.TaskName ($t.TaskPath + $t.TaskName) ("State: {0}" -f $t.State) @{
            TaskName = $t.TaskName; TaskPath = $t.TaskPath }))
    }

    # 7. Shortcuts ------------------------------------------------------------
    Write-Log 'Checking Start menu and desktop shortcuts that point into Office folders...' 'STEP'
    $shell = New-Object -ComObject WScript.Shell
    $startMenus = @("$env:ProgramData\Microsoft\Windows\Start Menu\Programs",
                    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs")
    $desktops   = @("$env:PUBLIC\Desktop", [Environment]::GetFolderPath('Desktop'))
    $lnkFiles = @()
    foreach ($r in $startMenus) { if (Test-Path -LiteralPath $r) { $lnkFiles += @(Get-ChildItem -LiteralPath $r -Filter *.lnk -Recurse -Force -ErrorAction SilentlyContinue) } }
    foreach ($r in $desktops)   { if (Test-Path -LiteralPath $r) { $lnkFiles += @(Get-ChildItem -LiteralPath $r -Filter *.lnk -Force -ErrorAction SilentlyContinue) } }
    Write-Log ("Inspecting {0} shortcut files" -f $lnkFiles.Count) 'DETAIL'
    foreach ($lnk in $lnkFiles) {
        $target = ''
        try { $target = $shell.CreateShortcut($lnk.FullName).TargetPath } catch { }
        if ($target -and $target -match '\\Microsoft Office( 15)?\\|OfficeClickToRun') {
            $found.Add((New-Finding 'Shortcut' $lnk.Name $lnk.FullName "Points to: $target" @{ Path = $lnk.FullName }))
        }
    }
    foreach ($r in $startMenus) {
        $toolsDir = Join-Path $r 'Microsoft Office Tools'
        if (Test-Path -LiteralPath $toolsDir) {
            $found.Add((New-Finding 'Shortcut' 'Microsoft Office Tools (Start menu folder)' $toolsDir '' @{ Path = $toolsDir }))
        }
    }

    # 8. Folders --------------------------------------------------------------
    Write-Log 'Checking Office program and data folders...' 'STEP'
    $folders = @()
    foreach ($base in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if ($base) { $folders += (Join-Path $base 'Microsoft Office'); $folders += (Join-Path $base 'Microsoft Office 15') }
    }
    foreach ($base in @($env:CommonProgramFiles, ${env:CommonProgramFiles(x86)})) {
        if ($base) { $folders += (Join-Path $base 'Microsoft Shared\ClickToRun') }
    }
    $folders += "$env:ProgramData\Microsoft\ClickToRun"
    $folders += "$env:ProgramData\Microsoft\Office\ClickToRunPackageLocker"
    $folders += "$env:LOCALAPPDATA\Microsoft\Office\Licenses"
    if ($IncludeUserSettings) { $folders += "$env:LOCALAPPDATA\Microsoft\Office\16.0" }

    foreach ($f in ($folders | Select-Object -Unique)) {
        if (-not (Test-Path -LiteralPath $f)) { Write-Log "Not present: $f" 'DETAIL'; continue }
        Write-Log "Measuring $f ..." 'DETAIL'
        $files = @(Get-ChildItem -LiteralPath $f -Recurse -File -Force -ErrorAction SilentlyContinue)
        $size = 0
        foreach ($x in $files) { $size += $x.Length }
        $found.Add((New-Finding 'Folder' $f $f ("{0} files, {1}" -f $files.Count, (Format-Size $size)) @{ Path = $f }))
    }

    # 9. Registry keys --------------------------------------------------------
    Write-Log 'Checking Office registry keys...' 'STEP'
    $regKeys = @(
        'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun',
        'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun',
        'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\AppVISV',
        'HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Licensing'
    )
    if ($IncludeUserSettings) { $regKeys += 'HKEY_CURRENT_USER\Software\Microsoft\Office' }
    foreach ($k in $regKeys) {
        if (Test-Path -LiteralPath "Registry::$k") {
            $sub = @(Get-ChildItem -LiteralPath "Registry::$k" -Recurse -ErrorAction SilentlyContinue).Count
            $found.Add((New-Finding 'Registry key' $k $k "$sub subkeys" @{ Key = $k }))
        }
        else { Write-Log "Not present: $k" 'DETAIL' }
    }

    return ,$found
}

function Show-Findings {
    param($Findings)
    $i = 0
    foreach ($cat in $CategoryOrder) {
        $items = @($Findings | Where-Object { $_.Category -eq $cat })
        if ($items.Count -eq 0) { continue }
        Write-Host ''
        Write-Host ('  {0} ({1})' -f $cat, $items.Count) -ForegroundColor White
        Write-Host ('  ' + ('-' * 74)) -ForegroundColor DarkGray
        foreach ($f in $items) {
            $i++
            Write-Host ('  [{0,2}] {1}' -f $i, $f.Name) -ForegroundColor Yellow
            if ($f.Target -and $f.Target -ne $f.Name) { Write-Host ('       Location : {0}' -f $f.Target) -ForegroundColor Gray }
            if ($f.Detail) { Write-Host ('       Details  : {0}' -f $f.Detail) -ForegroundColor DarkGray }
        }
    }
    Write-Host ''
}

# --------------------------------------------------------------------------
#  Pre-flight
# --------------------------------------------------------------------------

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host 'This script must be run as Administrator.' -ForegroundColor Red
    Write-Host 'Right-click Windows PowerShell > Run as administrator, then run the script again.' -ForegroundColor Red
    Read-Host 'Press Enter to close'
    exit 1
}
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    Write-Host 'This is 32-bit PowerShell on 64-bit Windows; registry results would be incomplete.' -ForegroundColor Red
    Write-Host 'Open "Windows PowerShell" (not "Windows PowerShell (x86)") as administrator and run again.' -ForegroundColor Red
    Read-Host 'Press Enter to close'
    exit 1
}

New-Item -ItemType Directory -Path $BackupPath -Force | Out-Null
$DebugLogFile = Start-DebugLog 'cleanup'

try {
    Write-Section 'Office leftovers cleanup - SCAN (nothing is changed yet)'
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    if ($os) { Write-Log ("Windows : {0} (build {1}, {2})" -f $os.Caption, $os.BuildNumber, $os.OSArchitecture) }
    Write-Log ("User    : {0}" -f [Security.Principal.WindowsIdentity]::GetCurrent().Name)
    Write-Log ("PS      : {0}" -f $PSVersionTable.PSVersion)
    Write-Log ("Log dir : {0}" -f $BackupPath)
    Write-Log ("Mode    : {0}{1}" -f $(if ($ScanOnly) { 'Scan only' } else { 'Scan, confirm, remove' }),
                                    $(if ($IncludeUserSettings) { ' (including user Office settings)' } else { '' }))

    $findings = Get-OfficeFindings

    Write-Section 'FOUND'
    if ($findings.Count -eq 0) {
        Write-Log 'No Office installations or leftovers were found. Nothing to do.' 'OK'
        return
    }
    Show-Findings $findings

    $actionable = @($findings | Where-Object { $_.Category -notlike '*(info)' })
    $findings | Select-Object Category, Name, Target, Detail |
        ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $BackupPath 'findings.json') -Encoding UTF8
    Write-Log ("{0} item(s) can be removed. Full list saved to findings.json" -f $actionable.Count)

    if ($ScanOnly) {
        Write-Log 'Scan-only mode: nothing was changed.' 'OK'
        return
    }
    if ($actionable.Count -eq 0) {
        Write-Log 'Only informational entries were found. Nothing to remove.' 'OK'
        return
    }

    Write-Host ''
    Write-Host '  WARNING: this removes ALL Microsoft Office / Microsoft 365 software listed above.' -ForegroundColor Red
    Write-Host '  Save and close any open Office documents first. Your document files are not deleted.' -ForegroundColor Red
    Write-Host '  Installed Office product keys are removed too; you can enter your key again after reinstalling.' -ForegroundColor Red
    Write-Host ''
    $answer = Read-Host '  Type YES (capitals) to remove everything listed, or anything else to cancel'
    if ($answer -cne 'YES') {
        Write-Log 'Cancelled by user. Nothing was changed.' 'WARN'
        return
    }

    # ---- 1. Registry backup ------------------------------------------------
    Write-Section 'Step 1/11 - Backing up registry keys'
    $toBackup = @()
    $toBackup += @($findings | Where-Object { $_.Category -eq 'Registry key' }      | ForEach-Object { $_.Target })
    $toBackup += @($findings | Where-Object { $_.Category -eq 'Installed product' } | ForEach-Object { $_.Target })
    $n = 0
    foreach ($k in ($toBackup | Select-Object -Unique)) {
        $n++
        $safe = ($k -replace '^HKEY_LOCAL_MACHINE', 'HKLM' -replace '^HKEY_CURRENT_USER', 'HKCU') -replace '[\\/:*?"<>| ]', '_'
        if ($safe.Length -gt 120) { $safe = $safe.Substring($safe.Length - 120) }
        $file = Join-Path $BackupPath ('{0:D2}_{1}.reg' -f $n, $safe)
        & reg.exe export "$k" "$file" /y 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) { Write-Log "Backed up $k" 'OK'; Add-Result 'Backup registry' $k 'OK' $file }
        else { Write-Log "Backup failed for $k (exit $LASTEXITCODE)" 'WARN'; Add-Result 'Backup registry' $k 'FAILED' }
    }

    # ---- 2. Close Office apps ----------------------------------------------
    Write-Section 'Step 2/11 - Closing Office apps'
    $running = @(Get-Process -Name $OfficeAppProcesses -ErrorAction SilentlyContinue)
    if ($running.Count -eq 0) { Write-Log 'No Office apps running.' 'DETAIL' }
    foreach ($pr in $running) {
        try {
            Stop-Process -Id $pr.Id -Force -ErrorAction Stop
            Write-Log ("Closed {0} (PID {1})" -f $pr.ProcessName, $pr.Id) 'OK'
            Add-Result 'Close app' $pr.ProcessName 'OK'
        }
        catch {
            Write-Log ("Could not close {0}: {1}" -f $pr.ProcessName, $_.Exception.Message) 'WARN'
            Add-Result 'Close app' $pr.ProcessName 'FAILED' $_.Exception.Message
        }
    }

    # ---- 3. License keys ---------------------------------------------------
    Write-Section 'Step 3/11 - Removing installed Office license keys'
    $licenses = @($findings | Where-Object { $_.Category -eq 'Office license key' })
    if ($licenses.Count -eq 0) { Write-Log 'No license keys to remove.' 'DETAIL' }
    foreach ($l in $licenses) {
        $last5 = $l.Data.Last5
        Write-Log "Running: cscript //Nologo `"$($l.Data.Ospp)`" /unpkey:$last5" 'DETAIL'
        $out = (& cscript.exe //Nologo "$($l.Data.Ospp)" "/unpkey:$last5" 2>&1) | Out-String
        Write-Log ($out.Trim() -replace '\s+', ' ') 'DETAIL'
        if ($out -match 'successful') { Write-Log "Removed key ending in $last5" 'OK'; Add-Result 'Remove license key' $last5 'OK' }
        else { Write-Log "Could not confirm removal of key ending in $last5" 'WARN'; Add-Result 'Remove license key' $last5 'FAILED' }
    }

    # ---- 4. Click-to-Run uninstall -----------------------------------------
    Write-Section 'Step 4/11 - Uninstalling Click-to-Run Office products'
    $c2rProducts = @($findings | Where-Object { $_.Category -eq 'Installed product' -and $_.Data.Kind -eq 'ClickToRun' })
    $c2rCfg      = @($findings | Where-Object { $_.Category -eq 'Click-to-Run configuration (info)' })

    $jobs = @()
    foreach ($prod in $c2rProducts) {
        $us = [string]$prod.Data.UninstallString
        $exe = $null; $arguments = ''
        if ($us -match '^\s*"([^"]+)"\s*(.*)$')         { $exe = $Matches[1]; $arguments = $Matches[2] }
        elseif ($us -match '^\s*(\S+?\.exe)\s*(.*)$')  { $exe = $Matches[1]; $arguments = $Matches[2] }
        if (-not $exe -or -not (Test-Path -LiteralPath $exe)) { $exe = $C2RExe }
        $jobs += [pscustomobject]@{ Name = $prod.Name; Exe = $exe; Args = $arguments; WatchKey = $prod.Target }
    }
    if ($jobs.Count -eq 0 -and $c2rCfg.Count -gt 0 -and (Test-Path -LiteralPath $C2RExe)) {
        # Half-installed Office (typical after a failed install): registered with
        # Click-to-Run but no "Installed apps" entry. Ask the engine to remove all products.
        Write-Log 'No Click-to-Run "Installed apps" entry, but Click-to-Run is registered. Using best-effort remove-all.' 'WARN'
        $jobs += [pscustomobject]@{ Name = 'All Click-to-Run products'; Exe = $C2RExe
                                    Args = 'scenario=install scenariosubtype=ARP sourcetype=None productstoremove=AllProducts'; WatchKey = $null }
    }
    if ($jobs.Count -eq 0) { Write-Log 'No Click-to-Run products to uninstall.' 'DETAIL' }

    foreach ($j in $jobs) {
        if (-not (Test-Path -LiteralPath $j.Exe)) {
            Write-Log "Uninstaller not found ($($j.Exe)). Leftovers will be removed manually in later steps." 'WARN'
            Add-Result 'Uninstall Click-to-Run' $j.Name 'SKIPPED' 'Uninstaller missing'
            continue
        }
        $arguments = $j.Args
        if ($arguments -notmatch 'DisplayLevel=')      { $arguments += ' DisplayLevel=False' }
        if ($arguments -notmatch 'forceappshutdown=')  { $arguments += ' forceappshutdown=True' }
        Write-Log "Uninstalling: $($j.Name)" 'STEP'
        Write-Log "Command: `"$($j.Exe)`" $arguments" 'DETAIL'
        try {
            $proc = Start-Process -FilePath $j.Exe -ArgumentList $arguments -PassThru -Wait -ErrorAction Stop
            Write-Log "Click-to-Run uninstaller returned exit code $($proc.ExitCode)" 'DETAIL'
        }
        catch {
            Write-Log ("Failed to start uninstaller: {0}" -f $_.Exception.Message) 'ERROR'
            Add-Result 'Uninstall Click-to-Run' $j.Name 'FAILED' $_.Exception.Message
            continue
        }
        # The uninstaller keeps working in the background; wait for it to finish.
        Start-Sleep -Seconds 10
        $deadline = (Get-Date).AddMinutes(15)
        while ((Get-Date) -lt $deadline) {
            $stillRegistered = $j.WatchKey -and (Test-Path -LiteralPath "Registry::$($j.WatchKey)")
            $clientRunning   = @(Get-Process -Name 'OfficeC2RClient' -ErrorAction SilentlyContinue).Count -gt 0
            if (-not $stillRegistered -and -not $clientRunning) { break }
            Write-Log 'Uninstaller still working, waiting 15 seconds...' 'DETAIL'
            Start-Sleep -Seconds 15
        }
        if ($j.WatchKey -and (Test-Path -LiteralPath "Registry::$($j.WatchKey)")) {
            Write-Log "$($j.Name) is still registered; its leftovers will be removed manually in later steps." 'WARN'
            Add-Result 'Uninstall Click-to-Run' $j.Name 'FAILED' 'Still registered after uninstall; cleaned manually'
        }
        else {
            Write-Log "Uninstall finished: $($j.Name)" 'OK'
            Add-Result 'Uninstall Click-to-Run' $j.Name 'OK'
        }
    }

    # ---- 5. MSI uninstall --------------------------------------------------
    Write-Section 'Step 5/11 - Uninstalling MSI-based Office products'
    $msiProducts = @($findings | Where-Object { $_.Category -eq 'Installed product' -and $_.Data.Kind -eq 'MSI' })
    if ($msiProducts.Count -eq 0) { Write-Log 'No MSI-based Office products.' 'DETAIL' }
    foreach ($m in $msiProducts) {
        $code = $m.Data.ProductCode
        $log = Join-Path $BackupPath ("msi_uninstall_{0}.log" -f ($code -replace '[{}]', ''))
        $arguments = "/x $code /qn /norestart /L*v `"$log`""
        Write-Log "Uninstalling: $($m.Name)" 'STEP'
        Write-Log "Command: msiexec.exe $arguments" 'DETAIL'
        $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -PassThru -Wait
        switch ($proc.ExitCode) {
            0       { Write-Log 'Uninstalled.' 'OK'; Add-Result 'Uninstall MSI' $m.Name 'OK' }
            3010    { Write-Log 'Uninstalled (restart required).' 'OK'; Add-Result 'Uninstall MSI' $m.Name 'OK' 'Restart required' }
            1605    { Write-Log 'Windows Installer says it is not installed (orphaned entry).' 'WARN'; Add-Result 'Uninstall MSI' $m.Name 'SKIPPED' 'Not installed' }
            default { Write-Log "msiexec failed with exit code $($proc.ExitCode). See $log" 'ERROR'; Add-Result 'Uninstall MSI' $m.Name 'FAILED' "Exit $($proc.ExitCode)" }
        }
    }

    # ---- 6. Services -------------------------------------------------------
    Write-Section 'Step 6/11 - Stopping Click-to-Run engine and removing Office services'
    foreach ($pr in @(Get-Process -Name $OfficeEngineProcesses -ErrorAction SilentlyContinue)) {
        try { Stop-Process -Id $pr.Id -Force -ErrorAction Stop; Write-Log ("Stopped {0} (PID {1})" -f $pr.ProcessName, $pr.Id) 'OK' }
        catch { Write-Log ("Could not stop {0}: {1}" -f $pr.ProcessName, $_.Exception.Message) 'WARN' }
    }
    $services = @($findings | Where-Object { $_.Category -eq 'Service' })
    if ($services.Count -eq 0) { Write-Log 'No Office services.' 'DETAIL' }
    foreach ($s in $services) {
        $svcName = $s.Data.ServiceName
        if (-not (Get-Service -Name $svcName -ErrorAction SilentlyContinue)) {
            Write-Log "Service already removed by the uninstaller: $svcName" 'OK'
            Add-Result 'Remove service' $svcName 'OK' 'Removed by uninstaller'
            continue
        }
        Write-Log "Stopping service $svcName ..." 'DETAIL'
        Stop-Service -Name $svcName -Force -ErrorAction SilentlyContinue
        & sc.exe delete "$svcName" 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) { Write-Log "Deleted service $svcName" 'OK'; Add-Result 'Remove service' $svcName 'OK' }
        else { Write-Log "sc.exe delete $svcName failed (exit $LASTEXITCODE)" 'ERROR'; Add-Result 'Remove service' $svcName 'FAILED' "Exit $LASTEXITCODE" }
    }

    # ---- 7. Scheduled tasks ------------------------------------------------
    Write-Section 'Step 7/11 - Removing Office scheduled tasks'
    $taskItems = @($findings | Where-Object { $_.Category -eq 'Scheduled task' })
    if ($taskItems.Count -eq 0) { Write-Log 'No Office scheduled tasks.' 'DETAIL' }
    foreach ($t in $taskItems) {
        $existing = Get-ScheduledTask -TaskName $t.Data.TaskName -TaskPath $t.Data.TaskPath -ErrorAction SilentlyContinue
        if (-not $existing) { Write-Log "Already removed: $($t.Target)" 'OK'; Add-Result 'Remove task' $t.Target 'OK' 'Removed by uninstaller'; continue }
        try {
            Unregister-ScheduledTask -TaskName $t.Data.TaskName -TaskPath $t.Data.TaskPath -Confirm:$false -ErrorAction Stop
            Write-Log "Removed task $($t.Target)" 'OK'; Add-Result 'Remove task' $t.Target 'OK'
        }
        catch { Write-Log ("Could not remove task {0}: {1}" -f $t.Target, $_.Exception.Message) 'ERROR'; Add-Result 'Remove task' $t.Target 'FAILED' $_.Exception.Message }
    }

    # ---- 8. Shortcuts ------------------------------------------------------
    Write-Section 'Step 8/11 - Removing Office shortcuts'
    $shortcuts = @($findings | Where-Object { $_.Category -eq 'Shortcut' })
    if ($shortcuts.Count -eq 0) { Write-Log 'No Office shortcuts.' 'DETAIL' }
    foreach ($s in $shortcuts) {
        $p = $s.Data.Path
        if (-not (Test-Path -LiteralPath $p)) { Write-Log "Already gone: $p" 'DETAIL'; Add-Result 'Remove shortcut' $p 'OK' 'Removed by uninstaller'; continue }
        try { Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction Stop; Write-Log "Deleted $p" 'OK'; Add-Result 'Remove shortcut' $p 'OK' }
        catch { Write-Log ("Could not delete {0}: {1}" -f $p, $_.Exception.Message) 'ERROR'; Add-Result 'Remove shortcut' $p 'FAILED' $_.Exception.Message }
    }

    # ---- 9. Folders --------------------------------------------------------
    Write-Section 'Step 9/11 - Removing leftover Office folders'
    $folderItems = @($findings | Where-Object { $_.Category -eq 'Folder' })
    if ($folderItems.Count -eq 0) { Write-Log 'No Office folders.' 'DETAIL' }
    foreach ($f in $folderItems) {
        Write-Log "Removing $($f.Data.Path) ..." 'STEP'
        if (Remove-FolderForce -Path $f.Data.Path) { Add-Result 'Remove folder' $f.Data.Path 'OK' }
        else { Add-Result 'Remove folder' $f.Data.Path 'FAILED' 'Locked files - restart and run again' }
    }

    # ---- 10. Registry ------------------------------------------------------
    Write-Section 'Step 10/11 - Removing leftover registry keys'
    $regItems = @()
    $regItems += @($findings | Where-Object { $_.Category -eq 'Registry key' }      | ForEach-Object { $_.Target })
    $regItems += @($findings | Where-Object { $_.Category -eq 'Installed product' } | ForEach-Object { $_.Target })
    foreach ($k in ($regItems | Select-Object -Unique)) {
        if (-not (Test-Path -LiteralPath "Registry::$k")) { Write-Log "Already gone: $k" 'DETAIL'; Add-Result 'Remove registry key' $k 'OK' 'Already gone'; continue }
        try {
            Remove-Item -LiteralPath "Registry::$k" -Recurse -Force -ErrorAction Stop
            Write-Log "Deleted $k" 'OK'; Add-Result 'Remove registry key' $k 'OK'
        }
        catch { Write-Log ("Could not delete {0}: {1}" -f $k, $_.Exception.Message) 'ERROR'; Add-Result 'Remove registry key' $k 'FAILED' $_.Exception.Message }
    }

    # ---- 11. Verify --------------------------------------------------------
    Write-Section 'Step 11/11 - Re-scanning to verify'
    $remaining = @(Get-OfficeFindings | Where-Object { $_.Category -notlike '*(info)' })

    Write-Section 'SUMMARY'
    $ok      = @($script:Results | Where-Object { $_.Status -eq 'OK' }).Count
    $failed  = @($script:Results | Where-Object { $_.Status -eq 'FAILED' })
    $skipped = @($script:Results | Where-Object { $_.Status -eq 'SKIPPED' }).Count
    Write-Log ("Actions succeeded: {0} | failed: {1} | skipped: {2}" -f $ok, $failed.Count, $skipped)
    foreach ($f in $failed) { Write-Log ("FAILED  {0}: {1}  {2}" -f $f.Action, $f.Target, $f.Note) 'ERROR' }
    $script:Results | Export-Csv -LiteralPath (Join-Path $BackupPath 'results.csv') -NoTypeInformation -Encoding UTF8

    if ($remaining.Count -eq 0) {
        Write-Log 'Re-scan is clean: no Office leftovers found.' 'OK'
    }
    else {
        Write-Log ("{0} item(s) are still present (usually files locked until restart):" -f $remaining.Count) 'WARN'
        Show-Findings $remaining
        Write-Log 'Restart Windows, then run this script again to finish the cleanup.' 'WARN'
    }
    Write-Log "Log, registry backups and results saved in: $BackupPath" 'INFO'
    Write-Host ''
    Write-Host '  Restart Windows BEFORE installing Office again.' -ForegroundColor Cyan
    $restart = Read-Host '  Restart now? (Y/N)'
    if ($restart -match '^[Yy]') {
        Write-Log 'Restarting in 10 seconds...' 'INFO'
        Stop-DebugLog
        Start-Sleep -Seconds 10
        Restart-Computer -Force
    }
}
catch {
    Write-DebugCrash $_
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Stop-DebugLog
}
Read-Host 'Press Enter to close'
