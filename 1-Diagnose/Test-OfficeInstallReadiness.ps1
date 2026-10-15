#Requires -Version 5.1
<#
.SYNOPSIS
    Read-only diagnostics for failed Office Click-to-Run installs
    (e.g. error 30175-45 (1289) = ERROR_BEYOND_VDL, a file-system write failure).

.DESCRIPTION
    Checks everything that commonly breaks an Office install and prints a
    PASS / WARN / FAIL / INFO line for each, with a suggested fix.
    NOTHING on the system is changed. The only writes are temporary test files
    (deleted straight away) and the report folder on your Desktop.

    Checks:
      System     - Windows version/build, admin, 64-bit, pending reboot
      Storage    - free space, file system type, dirty bit, NTFS compression,
                   BitLocker conversion, online file-system scan (= chkdsk /scan),
                   physical disk health and SMART-style counters
      Write test - writes, flushes, reads back and hashes a test file in
                   %TEMP%, ProgramData and Program Files (the 1289 error
                   happens while writing files)
      Windows    - component store health (= DISM /ScanHealth), Windows
                   Installer / Cryptographic / Click-to-Run services
      Security   - installed antivirus products, Defender Controlled Folder
                   Access, third-party file-system filter drivers (fltmc)
      Events     - disk / NTFS / storage-controller / WHEA hardware errors in
                   the last 14 days, last Windows Memory Diagnostic result
      Paths      - TEMP folder location, non-ASCII characters in TEMP path
      Network    - reachability of the Office CDN, proxy settings
      Office     - leftover Office installs, Office Deployment Tool folder and
                   signature, newest Office setup logs (errors extracted)

.PARAMETER Quick
    Skip the slow checks (component store scan, online volume scan).

.PARAMETER TestFileMB
    Size of each write-test file in MB. Default 256.

.PARAMETER OdtPath
    Folder with the Office Deployment Tool. Default: the kit's 3-Install folder
    (..\3-Install next to this script) if it contains setup.exe, otherwise C:\ODT

.EXAMPLE
    .\Test-OfficeInstallReadiness.ps1
.EXAMPLE
    .\Test-OfficeInstallReadiness.ps1 -Quick

.NOTES
    Run in an elevated 64-bit Windows PowerShell. If scripts are blocked:
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
    A full run takes about 5-15 minutes (the component store scan is slowest).
#>

[CmdletBinding()]
param(
    [switch]$Quick,
    [ValidateRange(16, 4096)][int]$TestFileMB = 256,
    [string]$OdtPath
)

if (-not $OdtPath) {
    $kitInstall = Join-Path (Split-Path -Parent $PSScriptRoot) '3-Install'
    if (Test-Path -LiteralPath (Join-Path $kitInstall 'setup.exe')) { $OdtPath = $kitInstall } else { $OdtPath = 'C:\ODT' }
}

$ErrorActionPreference = 'Continue'
try { . (Join-Path $PSScriptRoot '..\lib\Debug-Log.ps1'); . (Join-Path $PSScriptRoot '..\lib\Office-Licenses.ps1') }
catch {
    Write-Warning "Debug/license helpers unavailable: $($_.Exception.Message)"
    function Start-DebugLog { param($Name) }; function Stop-DebugLog {}; function Write-DebugCrash { param($ErrorRecord) }
}
trap { Write-DebugCrash $_; break }
$ReportDir = Join-Path ([Environment]::GetFolderPath('Desktop')) ('OfficeDiag_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
$script:Checks = New-Object System.Collections.Generic.List[object]
$SysDrive = $env:SystemDrive          # e.g. C:
$SysLetter = $SysDrive.TrimEnd(':')   # e.g. C

# --------------------------------------------------------------------------
#  Helpers
# --------------------------------------------------------------------------

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $colors = @{ INFO = 'Gray'; DETAIL = 'DarkGray'; STEP = 'Cyan'; PASS = 'Green'; WARN = 'Yellow'; FAIL = 'Red' }
    $c = $colors[$Level]; if (-not $c) { $c = 'Gray' }
    Write-Host ('[{0}] [{1,-6}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message) -ForegroundColor $c
}

function Write-Section {
    param([string]$Title)
    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host ('=' * 78) -ForegroundColor Cyan
}

function Add-Check {
    param(
        [string]$Area, [string]$Name,
        [ValidateSet('PASS','WARN','FAIL','INFO')][string]$Status,
        [string]$Detail = '', [string]$Fix = ''
    )
    $script:Checks.Add([pscustomobject]@{ Status = $Status; Area = $Area; Check = $Name; Detail = $Detail; Fix = $Fix })
    $lvl = $Status; if ($Status -eq 'INFO') { $lvl = 'INFO' }
    Write-Log ("{0}: {1}" -f $Name, $Detail) $lvl
    if ($Fix -and $Status -in 'WARN','FAIL') { Write-Log ("  -> {0}" -f $Fix) 'DETAIL' }
}

function Invoke-Check {
    # Runs one check; an unexpected error is reported instead of stopping the script.
    param([string]$Area, [string]$Name, [scriptblock]$Body)
    try { & $Body }
    catch { Add-Check $Area $Name 'WARN' ("Check could not run: {0}" -f $_.Exception.Message) 'Run the script as administrator in 64-bit Windows PowerShell.' }
}

function Format-Size {
    param([double]$Bytes)
    if ($Bytes -ge 1TB) { return '{0:N2} TB' -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return '{0:N1} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:N0} MB' -f ($Bytes / 1MB) }
    return '{0:N0} KB' -f ($Bytes / 1KB)
}

function Resolve-DriverPath {
    param([string]$Path)
    if (-not $Path) { return $null }
    $p = $Path.Trim('"')
    $p = $p -replace '^\\\?\?\\', ''
    $p = $p -replace '^\\SystemRoot', $env:windir
    $p = $p -replace '^(?i)system32', (Join-Path $env:windir 'System32')
    return $p
}

# --------------------------------------------------------------------------
#  Pre-flight
# --------------------------------------------------------------------------

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$IsAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Host 'Run this script as Administrator (right-click Windows PowerShell > Run as administrator).' -ForegroundColor Red
    Read-Host 'Press Enter to close'
    exit 1
}
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    Write-Host 'Use 64-bit "Windows PowerShell", not "Windows PowerShell (x86)".' -ForegroundColor Red
    Read-Host 'Press Enter to close'
    exit 1
}

New-Item -ItemType Directory -Path $ReportDir -Force | Out-Null
$DebugLogFile = Start-DebugLog 'diagnose'

try {

Write-Section 'Office install diagnostics (read-only)'
Write-Log "Report folder: $ReportDir"
Write-Log ("Mode: {0}, write test {1} MB per location" -f $(if ($Quick) { 'Quick' } else { 'Full' }), $TestFileMB)

# ==========================================================================
Write-Section '1. System'
# ==========================================================================

Invoke-Check 'System' 'Windows version' {
    $os = Get-CimInstance Win32_OperatingSystem
    $build = [int]$os.BuildNumber
    $ubr = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue).UBR
    $ver = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue).DisplayVersion
    $detail = '{0} {1} (build {2}.{3}, {4})' -f $os.Caption, $ver, $build, $ubr, $os.OSArchitecture
    if ($build -lt 10240) { Add-Check 'System' 'Windows version' 'FAIL' $detail 'Office 2024 needs Windows 10 or 11.' }
    elseif ($build -lt 19045) { Add-Check 'System' 'Windows version' 'WARN' $detail 'This Windows build is out of support. Update Windows before installing Office 2024.' }
    else { Add-Check 'System' 'Windows version' 'PASS' $detail }
    $script:OsLang = $os.MUILanguages -join ', '
    Add-Check 'System' 'Windows language' 'INFO' $script:OsLang
}

Invoke-Check 'System' 'Pending restart' {
    $reasons = @()
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $reasons += 'Windows component servicing' }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $reasons += 'Windows Update' }
    $pfro = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue).PendingFileRenameOperations
    if ($pfro) { $reasons += ('{0} pending file rename(s)' -f (@($pfro | Where-Object { $_ }).Count)) }
    if ($reasons.Count -gt 0) { Add-Check 'System' 'Pending restart' 'WARN' ($reasons -join '; ') 'Restart Windows before installing Office.' }
    else { Add-Check 'System' 'Pending restart' 'PASS' 'No restart pending.' }
}

Invoke-Check 'System' 'Uptime' {
    $boot = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
    $days = [math]::Round(((Get-Date) - $boot).TotalDays, 1)
    if ($days -gt 7) { Add-Check 'System' 'Uptime' 'WARN' "$days days since last restart (Fast Startup does not count as a restart)." 'Use Start > Power > Restart (not Shut down) before retrying.' }
    else { Add-Check 'System' 'Uptime' 'PASS' "$days days since last restart." }
}

# ==========================================================================
Write-Section '2. Storage and file system'
# ==========================================================================

Invoke-Check 'Storage' 'System drive' {
    $ld = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$SysDrive'"
    $free = [double]$ld.FreeSpace
    $detail = '{0} {1}, {2} free of {3}' -f $SysDrive, $ld.FileSystem, (Format-Size $free), (Format-Size ([double]$ld.Size))
    if ($ld.FileSystem -ne 'NTFS') { Add-Check 'Storage' 'File system type' 'FAIL' $detail 'Office must be installed on an NTFS system drive.' }
    else { Add-Check 'Storage' 'File system type' 'PASS' "$SysDrive is NTFS." }

    if ($free -lt 10GB) { Add-Check 'Storage' 'Free space' 'FAIL' $detail 'Free at least 10 GB (Office needs ~4 GB plus room for download and temp files).' }
    elseif ($free -lt 20GB) { Add-Check 'Storage' 'Free space' 'WARN' $detail 'Free space is tight; 20 GB+ is recommended.' }
    else { Add-Check 'Storage' 'Free space' 'PASS' $detail }

    if ($ld.VolumeDirty) { Add-Check 'Storage' 'Dirty bit' 'FAIL' "$SysDrive is flagged dirty (file system needs repair)." "Run: chkdsk $SysDrive /f  then restart; repair runs during boot." }
    else { Add-Check 'Storage' 'Dirty bit' 'PASS' "$SysDrive is not flagged dirty." }

    if ($ld.Compressed) { Add-Check 'Storage' 'Drive compression' 'WARN' "NTFS compression is enabled on all of $SysDrive." "Disable it in $SysDrive Properties > General > uncheck 'Compress this drive'." }
}

Invoke-Check 'Storage' 'Folder compression' {
    $paths = @($env:ProgramFiles, $env:CommonProgramFiles, $env:ProgramData, $env:TEMP, (Join-Path $env:windir 'Installer'))
    $compressed = @()
    foreach ($p in $paths) {
        if ($p -and (Test-Path -LiteralPath $p)) {
            $attr = (Get-Item -LiteralPath $p -Force).Attributes
            if ($attr -band [IO.FileAttributes]::Compressed) { $compressed += $p }
            if ($attr -band [IO.FileAttributes]::Encrypted)  { $compressed += "$p (EFS-encrypted)" }
        }
    }
    if ($compressed.Count) { Add-Check 'Storage' 'Folder compression/encryption' 'WARN' ($compressed -join '; ') 'Turn off compression/encryption on these folders (Properties > Advanced).' }
    else { Add-Check 'Storage' 'Folder compression/encryption' 'PASS' 'Program Files, ProgramData, TEMP and Installer are not compressed or encrypted.' }
}

Invoke-Check 'Storage' 'BitLocker' {
    if (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue) {
        $bl = Get-BitLockerVolume -MountPoint $SysDrive -ErrorAction Stop
        $detail = 'Status: {0}, {1}% encrypted, protection {2}' -f $bl.VolumeStatus, $bl.EncryptionPercentage, $bl.ProtectionStatus
        if ("$($bl.VolumeStatus)" -match 'InProgress|Paused') { Add-Check 'Storage' 'BitLocker' 'WARN' $detail 'Wait until BitLocker encryption/decryption finishes before installing.' }
        else { Add-Check 'Storage' 'BitLocker' 'PASS' $detail }
    }
    else { Add-Check 'Storage' 'BitLocker' 'INFO' 'BitLocker cmdlets not available on this edition (normal for Windows Home).' }
}

Invoke-Check 'Storage' 'Physical disk health' {
    $disks = @(Get-PhysicalDisk -ErrorAction Stop)
    foreach ($d in $disks) {
        $detail = '{0} ({1}, {2}, {3}) Health: {4}, Status: {5}' -f $d.FriendlyName, $d.MediaType, $d.BusType, (Format-Size ([double]$d.Size)), $d.HealthStatus, ($d.OperationalStatus -join ',')
        if ("$($d.HealthStatus)" -ne 'Healthy') { Add-Check 'Storage' "Disk $($d.DeviceId)" 'FAIL' $detail 'Back up your data now and replace or check this drive.' }
        else { Add-Check 'Storage' "Disk $($d.DeviceId)" 'PASS' $detail }

        $rc = $d | Get-StorageReliabilityCounter -ErrorAction SilentlyContinue
        if ($rc) {
            $parts = @()
            foreach ($prop in 'Wear','Temperature','ReadErrorsTotal','ReadErrorsUncorrected','WriteErrorsTotal','WriteErrorsUncorrected','PowerOnHours') {
                $v = $rc.$prop
                if ($null -ne $v) { $parts += "$prop=$v" }
            }
            $rcDetail = if ($parts.Count) { $parts -join ', ' } else { 'Drive does not report counters (use CrystalDiskInfo for SMART data).' }
            $bad = ($rc.ReadErrorsUncorrected -gt 0) -or ($rc.WriteErrorsUncorrected -gt 0) -or ($rc.Wear -ge 90)
            if ($bad) { Add-Check 'Storage' "Disk $($d.DeviceId) counters" 'FAIL' $rcDetail 'Uncorrected errors or high wear: the drive is likely failing. Back up and replace.' }
            else { Add-Check 'Storage' "Disk $($d.DeviceId) counters" 'INFO' $rcDetail }
        }
    }
}

if (-not $Quick) {
    Invoke-Check 'Storage' 'Online file-system scan' {
        Write-Log "Scanning $SysDrive file system online (same as chkdsk /scan, read-only). This can take a few minutes..." 'STEP'
        $res = Repair-Volume -DriveLetter $SysLetter -Scan -ErrorAction Stop
        $detail = "Result: $res"
        if ("$res" -eq 'NoErrorsFound') { Add-Check 'Storage' 'Online file-system scan' 'PASS' $detail }
        else { Add-Check 'Storage' 'Online file-system scan' 'FAIL' $detail "Run in an admin prompt: chkdsk $SysDrive /f  then restart (repair happens at boot)." }
    }
}
else { Add-Check 'Storage' 'Online file-system scan' 'INFO' 'Skipped (-Quick).' }

# ==========================================================================
Write-Section '3. Write / read-back test'
# ==========================================================================

function Test-WriteReadBack {
    param([string]$Folder, [int]$SizeMB)
    $file = Join-Path $Folder ('officediag_{0}.tmp' -f [guid]::NewGuid().ToString('N'))
    $chunk = 4MB
    $buf = New-Object byte[] $chunk
    $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
    $sha = [Security.Cryptography.SHA256]::Create()
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $fs = $null
    try {
        $fs = New-Object IO.FileStream($file, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None, $chunk, [IO.FileOptions]::WriteThrough)
        $total = [int64]$SizeMB * 1MB
        $written = [int64]0
        while ($written -lt $total) {
            $rng.GetBytes($buf)
            $fs.Write($buf, 0, $chunk)
            [void]$sha.TransformBlock($buf, 0, $chunk, $null, 0)
            $written += $chunk
        }
        [void]$sha.TransformFinalBlock((New-Object byte[] 0), 0, 0)
        $fs.Flush($true)
        $fs.Close(); $fs = $null
        $writeSec = $sw.Elapsed.TotalSeconds
        $expected = ([BitConverter]::ToString($sha.Hash)) -replace '-', ''
        $actual = (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash
        $mbps = [math]::Round($SizeMB / [math]::Max($writeSec, 0.001), 0)
        if ($expected -eq $actual) { return [pscustomobject]@{ Ok = $true;  Detail = "$SizeMB MB written and verified ($mbps MB/s write)" } }
        else { return [pscustomobject]@{ Ok = $false; Detail = "Data read back does NOT match what was written ($SizeMB MB)" } }
    }
    catch {
        $ex = $_.Exception
        while ($ex.InnerException) { $ex = $ex.InnerException }
        $hr = ''
        if ($ex.HResult) {
            $hr = ' (HRESULT 0x{0:X8}' -f $ex.HResult
            # 0x8007xxxx = wrapped Win32 error; 1289 here would be ERROR_BEYOND_VDL
            if (($ex.HResult -band 0xFFFF0000) -eq 0x80070000) { $hr += ', Win32 error {0}' -f ($ex.HResult -band 0xFFFF) }
            $hr += ')'
        }
        return [pscustomobject]@{ Ok = $false; Detail = ("Write/read failed: {0}{1}" -f $ex.Message, $hr) }
    }
    finally {
        if ($fs) { $fs.Dispose() }
        Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue
    }
}

foreach ($loc in @(@{ Name = 'TEMP folder'; Path = $env:TEMP },
                   @{ Name = 'ProgramData'; Path = $env:ProgramData },
                   @{ Name = 'Program Files'; Path = $env:ProgramFiles })) {
    Invoke-Check 'Write test' $loc.Name {
        Write-Log ("Writing {0} MB test file to {1} ..." -f $TestFileMB, $loc.Path) 'STEP'
        $r = Test-WriteReadBack -Folder $loc.Path -SizeMB $TestFileMB
        if ($r.Ok) { Add-Check 'Write test' $loc.Name 'PASS' ("{0}: {1}" -f $loc.Path, $r.Detail) }
        else { Add-Check 'Write test' $loc.Name 'FAIL' ("{0}: {1}" -f $loc.Path, $r.Detail) "Disk, file-system or filter-driver problem. Run chkdsk $SysDrive /f, check disk health, remove third-party antivirus/backup/encryption tools." }
    }
}

# ==========================================================================
Write-Section '4. Windows components and services'
# ==========================================================================

if (-not $Quick) {
    Invoke-Check 'Windows' 'Component store' {
        Write-Log 'Scanning Windows component store (same as DISM /ScanHealth, read-only). This can take 5-10 minutes...' 'STEP'
        $r = Repair-WindowsImage -Online -ScanHealth -ErrorAction Stop
        $state = "$($r.ImageHealthState)"
        if ($state -eq 'Healthy') { Add-Check 'Windows' 'Component store' 'PASS' 'Healthy' }
        else { Add-Check 'Windows' 'Component store' 'FAIL' "State: $state" 'Run: DISM /Online /Cleanup-Image /RestoreHealth  then  sfc /scannow  then restart.' }
    }
}
else { Add-Check 'Windows' 'Component store' 'INFO' 'Skipped (-Quick).' }

Invoke-Check 'Windows' 'Services' {
    $svcList = @(
        @{ Name = 'msiserver';     Label = 'Windows Installer';      Need = 'NotDisabled' },
        @{ Name = 'CryptSvc';      Label = 'Cryptographic Services'; Need = 'Running' },
        @{ Name = 'ClickToRunSvc'; Label = 'Office Click-to-Run';    Need = 'Info' },
        @{ Name = 'BITS';          Label = 'Background Intelligent Transfer'; Need = 'NotDisabled' }
    )
    foreach ($s in $svcList) {
        $svc = Get-CimInstance Win32_Service -Filter ("Name='{0}'" -f $s.Name) -ErrorAction SilentlyContinue
        if (-not $svc) {
            if ($s.Need -eq 'Info') { Add-Check 'Windows' $s.Label 'INFO' 'Not installed (normal before Office is installed).' }
            else { Add-Check 'Windows' $s.Label 'FAIL' 'Service is missing.' 'Repair Windows: DISM /Online /Cleanup-Image /RestoreHealth, then sfc /scannow.' }
            continue
        }
        $detail = 'State: {0}, Start mode: {1}' -f $svc.State, $svc.StartMode
        switch ($s.Need) {
            'Running'     { if ($svc.State -eq 'Running') { Add-Check 'Windows' $s.Label 'PASS' $detail } else { Add-Check 'Windows' $s.Label 'FAIL' $detail "Start it: Start-Service $($s.Name)" } }
            'NotDisabled' { if ($svc.StartMode -eq 'Disabled') { Add-Check 'Windows' $s.Label 'FAIL' $detail "Set it to Manual: Set-Service $($s.Name) -StartupType Manual" } else { Add-Check 'Windows' $s.Label 'PASS' $detail } }
            default       { Add-Check 'Windows' $s.Label 'INFO' $detail }
        }
    }
}

# ==========================================================================
Write-Section '5. Security software and file-system filters'
# ==========================================================================

Invoke-Check 'Security' 'Antivirus products' {
    $av = @(Get-CimInstance -Namespace 'root/SecurityCenter2' -ClassName AntiVirusProduct -ErrorAction Stop)
    if ($av.Count -eq 0) { Add-Check 'Security' 'Antivirus products' 'INFO' 'None registered.' ; return }
    foreach ($a in $av) {
        $hex = '{0:X6}' -f [int]$a.productState
        $enabled = $hex.Substring(2, 2) -in '10','11'
        $detail = '{0} ({1})' -f $a.displayName, $(if ($enabled) { 'active' } else { 'inactive' })
        if ($a.displayName -match 'Windows Defender|Microsoft Defender') { Add-Check 'Security' 'Antivirus' 'PASS' $detail }
        elseif ($enabled) { Add-Check 'Security' 'Antivirus' 'WARN' $detail 'Third-party antivirus can block Office file writes. Uninstall it temporarily (disabling is often not enough), restart, install Office, then reinstall it.' }
        else { Add-Check 'Security' 'Antivirus' 'INFO' $detail }
    }
}

Invoke-Check 'Security' 'Controlled Folder Access' {
    $mp = Get-MpPreference -ErrorAction Stop
    $v = [int]$mp.EnableControlledFolderAccess
    switch ($v) {
        1 { Add-Check 'Security' 'Controlled Folder Access' 'WARN' 'Enabled (blocking).' 'Windows Security > Virus & threat protection > Ransomware protection > turn off temporarily.' }
        2 { Add-Check 'Security' 'Controlled Folder Access' 'INFO' 'Audit mode.' }
        default { Add-Check 'Security' 'Controlled Folder Access' 'PASS' 'Off.' }
    }
}

Invoke-Check 'Security' 'File-system filter drivers' {
    $lines = & fltmc.exe filters 2>&1
    $filters = @()
    foreach ($l in $lines) {
        if ("$l" -match '^\s*(\S+)\s+(\d+)\s+([\d.]+)\s*(\d*)\s*$') { $filters += [pscustomobject]@{ Name = $Matches[1]; Altitude = $Matches[3] } }
    }
    if ($filters.Count -eq 0) { Add-Check 'Security' 'Filter drivers' 'WARN' 'Could not read filter list (fltmc).' ''; return }
    $thirdParty = @()
    $all = @()
    foreach ($f in $filters) {
        $drv = Get-CimInstance Win32_SystemDriver -Filter ("Name='{0}'" -f $f.Name) -ErrorAction SilentlyContinue
        $path = if ($drv) { Resolve-DriverPath $drv.PathName } else { Join-Path $env:windir ("System32\drivers\{0}.sys" -f $f.Name) }
        $company = ''
        if ($path -and (Test-Path -LiteralPath $path)) { $company = (Get-Item -LiteralPath $path).VersionInfo.CompanyName }
        $desc = if ($drv) { $drv.DisplayName } else { '' }
        $all += ('{0} [{1}] {2} - {3}' -f $f.Name, $f.Altitude, $desc, $company)
        if ($company -notmatch 'Microsoft') { $thirdParty += ('{0} ({1}{2})' -f $f.Name, $(if ($company) { $company } else { 'unknown vendor' }), $(if ($desc) { ", $desc" } else { '' })) }
    }
    $all | Set-Content -LiteralPath (Join-Path $ReportDir 'filter_drivers.txt') -Encoding UTF8
    if ($thirdParty.Count) { Add-Check 'Security' 'Third-party filter drivers' 'WARN' ($thirdParty -join '; ') 'These intercept every file write (antivirus, backup, encryption, cloud sync). Uninstall the owning software temporarily if the install keeps failing.' }
    else { Add-Check 'Security' 'Third-party filter drivers' 'PASS' ("{0} filters loaded, all Microsoft." -f $filters.Count) }
}

# ==========================================================================
Write-Section '6. Event logs (last 14 days)'
# ==========================================================================

Invoke-Check 'Events' 'Disk and file-system errors' {
    $since = (Get-Date).AddDays(-14)
    $providers = @('disk','Ntfs','Microsoft-Windows-Ntfs','volmgr','stornvme','storahci','iaStorA','iaStorAC','iaStorAVC','nvme','Microsoft-Windows-StorPort','Microsoft-Windows-WHEA-Logger')
    $events = @()
    foreach ($p in $providers) {
        try { $events += @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; ProviderName = $p; Level = 1, 2, 3; StartTime = $since } -ErrorAction Stop) } catch { }
    }
    if ($events.Count -eq 0) { Add-Check 'Events' 'Disk/NTFS/hardware errors' 'PASS' 'None in the last 14 days.'; return }
    $events = $events | Sort-Object TimeCreated -Descending
    $events | Select-Object TimeCreated, ProviderName, Id, LevelDisplayName, Message |
        Format-List | Out-String -Width 300 | Set-Content -LiteralPath (Join-Path $ReportDir 'disk_events.txt') -Encoding UTF8
    $groups = $events | Group-Object ProviderName, Id | Sort-Object Count -Descending | Select-Object -First 6
    $summary = ($groups | ForEach-Object { '{0} x{1}' -f $_.Name, $_.Count }) -join '; '
    $serious = @($events | Where-Object { $_.ProviderName -match 'disk|Ntfs|WHEA|stor|nvme|iaStor' -and $_.Level -le 2 })
    if ($serious.Count) { Add-Check 'Events' 'Disk/NTFS/hardware errors' 'FAIL' ("{0} error(s): {1}" -f $events.Count, $summary) "See disk_events.txt. Repeated disk/NTFS/WHEA errors mean a failing drive, cable or controller; run chkdsk $SysDrive /f and check the drive." }
    else { Add-Check 'Events' 'Disk/NTFS/hardware errors' 'WARN' ("{0} warning(s): {1}" -f $events.Count, $summary) 'See disk_events.txt.' }
}

Invoke-Check 'Events' 'Memory diagnostic' {
    $m = Get-WinEvent -FilterHashtable @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-MemoryDiagnostics-Results' } -MaxEvents 1 -ErrorAction SilentlyContinue
    if ($m) {
        if ($m.Id -eq 1201) { Add-Check 'Events' 'Memory diagnostic' 'FAIL' ("{0}: hardware problems detected" -f $m.TimeCreated) 'Faulty RAM corrupts written data. Test/replace RAM modules.' }
        else { Add-Check 'Events' 'Memory diagnostic' 'PASS' ("Last run {0}: no errors" -f $m.TimeCreated) }
    }
    else { Add-Check 'Events' 'Memory diagnostic' 'INFO' 'Never run. If write tests fail randomly, run mdsched.exe.' }
}

# ==========================================================================
Write-Section '7. Paths and environment'
# ==========================================================================

Invoke-Check 'Paths' 'TEMP folder' {
    $t = $env:TEMP
    $detail = $t
    $issues = @()
    if ($t -match '[^\x20-\x7E]') { $issues += 'contains non-ASCII characters (e.g. Lithuanian letters in the user name)' }
    if (-not (Test-Path -LiteralPath $t)) { $issues += 'does not exist' }
    if ($t -and -not $t.StartsWith($SysDrive, [StringComparison]::OrdinalIgnoreCase)) { $issues += "is not on $SysDrive" }
    if ($issues.Count) { Add-Check 'Paths' 'TEMP folder' 'WARN' ("$detail - " + ($issues -join '; ')) 'Try installing from a new local admin account with a plain ASCII name, or set TEMP/TMP to C:\Temp for the install session.' }
    else { Add-Check 'Paths' 'TEMP folder' 'PASS' $detail }
}

Invoke-Check 'Paths' 'User profile path' {
    if ($env:USERPROFILE -match '[^\x20-\x7E]') { Add-Check 'Paths' 'User profile path' 'WARN' $env:USERPROFILE 'Non-ASCII profile path. If the install keeps failing, try from a new local admin account with an ASCII name.' }
    else { Add-Check 'Paths' 'User profile path' 'PASS' $env:USERPROFILE }
}

# ==========================================================================
Write-Section '8. Network'
# ==========================================================================

Invoke-Check 'Network' 'Office CDN' {
    foreach ($h in 'officecdn.microsoft.com', 'officecdn.microsoft.com.edgesuite.net') {
        $ok = Test-NetConnection -ComputerName $h -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue
        if ($ok) { Add-Check 'Network' "Reach $h" 'PASS' 'Port 443 reachable.' }
        else { Add-Check 'Network' "Reach $h" 'FAIL' 'Cannot connect on port 443.' 'Check internet, VPN, firewall, DNS or proxy. Or download first with setup.exe /download on a working network.' }
    }
}

Invoke-Check 'Network' 'Proxy' {
    $is = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -ErrorAction SilentlyContinue
    if ($is -and $is.ProxyEnable -eq 1) { Add-Check 'Network' 'User proxy' 'WARN' ("Proxy enabled: {0}" -f $is.ProxyServer) 'Disable the proxy or VPN while installing if downloads fail.' }
    elseif ($is -and $is.AutoConfigURL) { Add-Check 'Network' 'User proxy' 'INFO' ("PAC script: {0}" -f $is.AutoConfigURL) }
    else { Add-Check 'Network' 'User proxy' 'PASS' 'No proxy configured.' }
}

# ==========================================================================
Write-Section '9. Office state'
# ==========================================================================

Invoke-Check 'Office' 'Existing Office installs' {
    $pattern = '^(Microsoft 365|Microsoft Office|Office 16 Click-to-Run|Microsoft (OneNote|Project|Visio|Access|Excel|Word|PowerPoint|Outlook|Publisher)\b)'
    $exclude = 'Access database engine|Visual'
    $found = @()
    foreach ($root in 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall') {
        if (-not (Test-Path $root)) { continue }
        foreach ($k in Get-ChildItem $root -ErrorAction SilentlyContinue) {
            $p = Get-ItemProperty $k.PSPath -ErrorAction SilentlyContinue
            if ($p -and $p.DisplayName -and $p.DisplayName -match $pattern -and $p.DisplayName -notmatch $exclude -and $p.SystemComponent -ne 1) {
                $found += ('{0} {1}' -f $p.DisplayName, $p.DisplayVersion)
            }
        }
    }
    $c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
    if ($c2r) { $found += ('Click-to-Run registration: {0} {1}' -f $c2r.ProductReleaseIds, $c2r.VersionToReport) }
    if (Test-Path (Join-Path $env:ProgramFiles 'Microsoft Office\root')) { $found += "$env:ProgramFiles\Microsoft Office\root exists" }
    if (${env:ProgramFiles(x86)} -and (Test-Path (Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\root'))) { $found += "${env:ProgramFiles(x86)}\Microsoft Office\root exists (32-bit Office)" }
    if ($found.Count) { Add-Check 'Office' 'Existing Office installs/leftovers' 'WARN' ($found -join '; ') 'Run Remove-OfficeLeftovers.ps1, restart, then install again.' }
    else { Add-Check 'Office' 'Existing Office installs/leftovers' 'PASS' 'No Office installs or leftovers.' }
}

Invoke-Check 'Office' 'Office licenses and the install they belong to' {
    $lic = Get-OfficeLicenseInfo
    foreach ($i in $lic.Installs) { Add-Check 'Office' ('Install: ' + $i.Kind) 'INFO' ('{0} | version {1} | {2} {3}' -f $i.Name, $i.Version, $i.Platform, $i.Channel) }
    if (-not $lic.Installs) { Add-Check 'Office' 'Installed Office' 'INFO' 'No Click-to-Run, MSI or Store Office found.' }
    foreach ($l in $lic.Licenses) {
        $what = if ($l.Last5) { 'Key ending ' + $l.Last5 } else { $l.Source }
        $detail = '{0} | {1} | channel: {2} | status: {3} | belongs to: {4}' -f $l.Name, $l.Description, $l.Channel, $l.Status, $l.BelongsTo
        $status = if ($l.BelongsTo -like 'no matching install*') { 'WARN' } else { 'INFO' }
        Add-Check 'Office' ('License: ' + $what) $status $detail $(if ($status -eq 'WARN') { 'Orphan license of a removed install: Remove-OfficeLeftovers.ps1 removes it.' } else { '' })
    }
    if (-not $lic.Licenses) { Add-Check 'Office' 'Office licenses' 'INFO' 'No Office license keys or subscriptions found.' }
}

Invoke-Check 'Office' 'Office Deployment Tool' {
    $setup = Join-Path $OdtPath 'setup.exe'
    if (-not (Test-Path -LiteralPath $setup)) { Add-Check 'Office' 'Office Deployment Tool' 'INFO' "Not found at $setup (use -OdtPath if it is elsewhere)."; return }
    $sig = Get-AuthenticodeSignature -LiteralPath $setup
    $signer = if ($sig.SignerCertificate) { $sig.SignerCertificate.Subject } else { 'none' }
    if ($sig.Status -eq 'Valid' -and $signer -match 'Microsoft') { Add-Check 'Office' 'ODT setup.exe signature' 'PASS' ("Valid, {0}, version {1}" -f ($signer -replace ',.*$', ''), (Get-Item $setup).VersionInfo.FileVersion) }
    else { Add-Check 'Office' 'ODT setup.exe signature' 'FAIL' ("Status: {0}, signer: {1}" -f $sig.Status, $signer) 'Download the Office Deployment Tool again from microsoft.com only.' }

    $xmls = @(Get-ChildItem -LiteralPath $OdtPath -Filter *.xml -ErrorAction SilentlyContinue)
    foreach ($x in $xmls) {
        try {
            [xml]$doc = Get-Content -LiteralPath $x.FullName -Raw
            $add = $doc.Configuration.Add
            $prods = @($doc.Configuration.Add.Product | ForEach-Object { $_.ID }) -join ', '
            $src = if ($add.SourcePath) { $add.SourcePath } else { '(online)' }
            Add-Check 'Office' ("Config {0}" -f $x.Name) 'INFO' ("Products: {0}; {1}-bit; channel {2}; source {3}" -f $prods, $add.OfficeClientEdition, $add.Channel, $src)
        }
        catch { Add-Check 'Office' ("Config {0}" -f $x.Name) 'FAIL' ("XML is invalid: {0}" -f $_.Exception.Message) 'Fix the XML syntax (unclosed tag, quotes).' }
    }

    $dl = Join-Path $OdtPath 'Office\Data'
    if (Test-Path -LiteralPath $dl) {
        $files = @(Get-ChildItem -LiteralPath $dl -Recurse -File -ErrorAction SilentlyContinue)
        $size = 0; foreach ($f in $files) { $size += $f.Length }
        $zero = @($files | Where-Object { $_.Length -eq 0 })
        $detail = '{0} files, {1}' -f $files.Count, (Format-Size $size)
        if ($zero.Count) { Add-Check 'Office' 'Downloaded Office files' 'WARN' ("$detail; {0} empty file(s)" -f $zero.Count) "Delete $OdtPath\Office and run setup.exe /download again." }
        elseif ($size -lt 1GB) { Add-Check 'Office' 'Downloaded Office files' 'WARN' "$detail (looks incomplete; a full download is ~2-4 GB)" "Delete $OdtPath\Office and run setup.exe /download again." }
        else { Add-Check 'Office' 'Downloaded Office files' 'PASS' $detail }
    }
}

Invoke-Check 'Office' 'Setup logs' {
    $dirs = @($env:TEMP, (Join-Path $env:windir 'Temp')) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique
    $logs = @()
    foreach ($d in $dirs) {
        $logs += @(Get-ChildItem -LiteralPath $d -Filter '*.log' -File -ErrorAction SilentlyContinue |
                   Where-Object { $_.Name -match ('^{0}-\d{{8}}-\d{{4}}' -f [regex]::Escape($env:COMPUTERNAME)) })
    }
    if ($logs.Count -eq 0) { Add-Check 'Office' 'Setup logs' 'INFO' 'No Office setup logs found in TEMP folders.'; return }
    $logs = $logs | Sort-Object LastWriteTime -Descending | Select-Object -First 3
    $out = Join-Path $ReportDir 'office_setup_log_errors.txt'
    $hits = @()
    foreach ($l in $logs) {
        Copy-Item -LiteralPath $l.FullName -Destination $ReportDir -Force -ErrorAction SilentlyContinue
        $m = @(Select-String -LiteralPath $l.FullName -Pattern '30175|1289|ErrorCode|ErrorMessage|ErrorDetails|HRESULT|0x[0-9A-F]{8}|failed|Failure' -ErrorAction SilentlyContinue)
        Add-Content -LiteralPath $out -Value ("===== {0} ({1}) - {2} matching lines =====" -f $l.Name, $l.LastWriteTime, $m.Count) -Encoding UTF8
        $m | Select-Object -Last 200 | ForEach-Object { Add-Content -LiteralPath $out -Value ('{0,6}: {1}' -f $_.LineNumber, $_.Line.Trim()) -Encoding UTF8 }
        $hits += @($m | Where-Object { $_.Line -match '30175|1289' })
    }
    $detail = 'Newest: {0} ({1}); copied with error extract to report folder' -f $logs[0].Name, $logs[0].LastWriteTime
    if ($hits.Count) {
        Add-Check 'Office' 'Setup logs' 'WARN' ("$detail. {0} line(s) mention 30175/1289." -f $hits.Count) 'Send office_setup_log_errors.txt for analysis.'
        Write-Log 'Last log lines mentioning 30175 / 1289:' 'DETAIL'
        $hits | Select-Object -Last 8 | ForEach-Object { Write-Log ('  ' + ($_.Line.Trim() -replace '\s+', ' ')) 'DETAIL' }
    }
    else { Add-Check 'Office' 'Setup logs' 'INFO' $detail }
}

# ==========================================================================
Write-Section 'SUMMARY'
# ==========================================================================

$order = @{ FAIL = 0; WARN = 1; PASS = 2; INFO = 3 }
$sorted = $script:Checks | Sort-Object { $order[$_.Status] }, Area, Check
$counts = $script:Checks | Group-Object Status | ForEach-Object { '{0}: {1}' -f $_.Name, $_.Count }
Write-Log ($counts -join ' | ')
Write-Host ''
foreach ($c in ($sorted | Where-Object { $_.Status -in 'FAIL','WARN' })) {
    $color = if ($c.Status -eq 'FAIL') { 'Red' } else { 'Yellow' }
    Write-Host ('  [{0}] {1} / {2}' -f $c.Status, $c.Area, $c.Check) -ForegroundColor $color
    Write-Host ('         {0}' -f $c.Detail) -ForegroundColor Gray
    if ($c.Fix) { Write-Host ('         Fix: {0}' -f $c.Fix) -ForegroundColor White }
}
if (-not ($sorted | Where-Object { $_.Status -in 'FAIL','WARN' })) {
    Write-Host '  No problems found. If Office still fails, send office_setup_log_errors.txt.' -ForegroundColor Green
}

$sorted | Export-Csv -LiteralPath (Join-Path $ReportDir 'report.csv') -NoTypeInformation -Encoding UTF8
$sorted | Format-Table Status, Area, Check, Detail -AutoSize -Wrap | Out-String -Width 400 |
    Set-Content -LiteralPath (Join-Path $ReportDir 'report.txt') -Encoding UTF8
Write-Host ''
Write-Log "Report saved to: $ReportDir (report.txt, report.csv, transcript.log and extracts)"

}
catch {
    Write-DebugCrash $_
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Stop-DebugLog
    try { if ($DebugLogFile -and (Test-Path -LiteralPath $DebugLogFile)) { Copy-Item -LiteralPath $DebugLogFile -Destination (Join-Path $ReportDir 'transcript.log') -Force } } catch { }
}
Read-Host 'Press Enter to close'
