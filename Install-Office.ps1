<#
.SYNOPSIS
  Universal Office installer: generates an Office Deployment Tool (ODT) configuration XML
  for the chosen product, downloads the ODT from Microsoft if needed, and runs the install.

.DESCRIPTION
  Interactive by default (menus). Every choice can also be passed as a parameter.
  The product key is never stored: it is asked for at run time (hidden input, optional),
  written only to a temporary config file that is deleted afterwards, and scrubbed from
  the setup logs. Press Enter at the key prompt to install without a key and activate in the app.

.EXAMPLE
  .\Install-Office.ps1
.EXAMPLE
  .\Install-Office.ps1 -Product HomeStudent2024Retail -Language lt-lt,en-us -GenerateOnly
.EXAMPLE
  .\Install-Office.ps1 -DownloadOnly
#>
[CmdletBinding()]
param(
    [string]$Product,                          # ODT product ID, e.g. HomeStudent2024Retail
    [string[]]$Language,                       # e.g. lt-lt,en-us (first one is primary)
    [ValidateSet('64', '32')][string]$Arch = '64',
    [string[]]$ExcludeApp,                     # e.g. Teams,OneDrive; 'none' for nothing
    [ValidateSet('Full', 'None')][string]$Display = 'Full',
    [switch]$RemoveMSI,                        # remove older MSI-based Office
    [switch]$AskKey,                           # always ask for a key (default: ask in interactive mode)
    [switch]$NoKey,                            # never ask for a key
    [switch]$GenerateOnly,                     # write the XML to .\output and stop (never contains a key)
    [switch]$DownloadOnly,                     # only fetch the ODT into .\ODT and stop
    [string]$OdtPath                           # folder holding setup.exe (default: .\ODT, script folder, C:\ODT)
    ,[switch]$Elevated                         # internal: set on the elevated relaunch to prevent a relaunch loop
)

$ErrorActionPreference = 'Stop'
# -File passes 'a,b' as one string: split it
if ($Language)   { $Language   = @($Language   | ForEach-Object { $_ -split ',' } | Where-Object { $_ }) }
if ($ExcludeApp) { $ExcludeApp = @($ExcludeApp | ForEach-Object { $_ -split ',' } | Where-Object { $_ }) }
$Root = $PSScriptRoot

# Log folder: %TEMP%\Debug\office-install-tool\<date>\ (one transcript per window: launcher, main, each step; plus Office's own setup logs).
$script:LogRoot = Join-Path $env:TEMP 'Debug\office-install-tool'
$script:LogDir = Join-Path $script:LogRoot (Get-Date -Format 'yyyy-MM-dd')

# Relaunch elevated before any prompt, so the answers are not lost. -GenerateOnly needs no admin rights.
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole('Administrator')
if (-not $isAdmin -and -not $GenerateOnly -and -not $Elevated) {
    try { New-Item -ItemType Directory -Force $script:LogDir | Out-Null; Start-Transcript -Path (Join-Path $script:LogDir ('debug-launcher-{0}.log' -f (Get-Date -Format 'HHmmss'))) | Out-Null } catch {}
    Write-Host 'Administrator rights are needed; relaunching elevated...' -ForegroundColor Yellow
    $relaunch = @('-NoExit', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"", '-Elevated')
    foreach ($kv in $PSBoundParameters.GetEnumerator()) {
        if ($kv.Value -is [switch]) { if ($kv.Value.IsPresent) { $relaunch += "-$($kv.Key)" } }
        else { $relaunch += "-$($kv.Key)"; $relaunch += "`"$(@($kv.Value) -join ',')`"" }
    }
    Write-Host "Relaunch arguments: $($relaunch -join ' ')"
    Start-Process powershell -Verb RunAs -ArgumentList $relaunch
    try { Stop-Transcript | Out-Null } catch {}
    return
}

# Debug log: always on. A full transcript goes to %TEMP%\Debug\office-install-tool\<date>\, and a shortcut in Desktop\Debug points at it.
$script:Transcribing = $false
try {
    New-Item -ItemType Directory -Force $script:LogDir | Out-Null
    $script:LogFile = Join-Path $script:LogDir ('debug-install-{0}.log' -f (Get-Date -Format 'HHmmss'))
    Start-Transcript -Path $script:LogFile | Out-Null
    $script:Transcribing = $true
    Write-Host "Debug log: $script:LogDir" -ForegroundColor DarkGray
} catch { Write-Warning "Could not start the debug log in $($script:LogDir): $($_.Exception.Message)" }
try {
    $lnkDir = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Debug'
    $lnk = Join-Path $lnkDir 'office-install-tool.lnk'
    if (-not (Test-Path -LiteralPath $lnk)) {
        $make = -not [Environment]::UserInteractive -or ((Read-Host '  Put a shortcut to the logs in Desktop\Debug? (Y/n)') -notmatch '^[nN]')
        if ($make) {
            New-Item -ItemType Directory -Force $lnkDir | Out-Null
            $sc = (New-Object -ComObject WScript.Shell).CreateShortcut($lnk); $sc.TargetPath = $script:LogRoot; $sc.Save()
            Write-Host "Shortcut created: $lnk" -ForegroundColor Green
        }
    }
} catch { Write-Warning "Could not create the desktop shortcut ($($_.Exception.Message)). Logs are in $($script:LogDir)" }

try {
$OdtDir = Join-Path $Root 'ODT'
$OutDir = Join-Path $Root 'output'
$OdtPage = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'

# ---------------------------------------------------------------- catalog
# Name, ODT product ID, update channel. Add rows here to support more products.
$Catalog = @(
    @{ Name = 'Office Home & Student 2024';        Id = 'HomeStudent2024Retail';   Channel = 'Current' }
    @{ Name = 'Office Home & Business 2024';       Id = 'HomeBusiness2024Retail';  Channel = 'Current' }
    @{ Name = 'Office Home 2024';                  Id = 'Home2024Retail';          Channel = 'Current' }
    @{ Name = 'Office Professional 2024';          Id = 'Professional2024Retail';  Channel = 'Current' }
    @{ Name = 'Office LTSC Professional Plus 2024 (volume)'; Id = 'ProPlus2024Volume'; Channel = 'PerpetualVL2024' }
    @{ Name = 'Office Home & Student 2021';        Id = 'HomeStudent2021Retail';   Channel = 'Current' }
    @{ Name = 'Office Home & Business 2021';       Id = 'HomeBusiness2021Retail';  Channel = 'Current' }
    @{ Name = 'Office Professional 2021';          Id = 'Professional2021Retail';  Channel = 'Current' }
    @{ Name = 'Office LTSC Professional Plus 2021 (volume)'; Id = 'ProPlus2021Volume'; Channel = 'PerpetualVL2021' }
    @{ Name = 'Microsoft 365 Apps for enterprise'; Id = 'O365ProPlusRetail';       Channel = 'Current' }
    @{ Name = 'Microsoft 365 Apps for business';   Id = 'O365BusinessRetail';      Channel = 'Current' }
    @{ Name = 'Microsoft 365 Family/Personal';     Id = 'O365HomePremRetail';      Channel = 'Current' }
)
$AllApps = 'Access', 'Excel', 'Groove', 'Lync', 'OneDrive', 'OneNote', 'Outlook', 'PowerPoint', 'Publisher', 'Teams', 'Word'
$KeyPattern = '^[A-Z0-9]{5}(-[A-Z0-9]{5}){4}$'

# ---------------------------------------------------------------- helpers
function Read-Choice([string]$Title, [string[]]$Items) {
    Write-Host "`n$Title" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Items.Count; $i++) { Write-Host ("  {0,2}) {1}" -f ($i + 1), $Items[$i]) }
    while ($true) {
        $a = Read-Host 'Choose a number'
        if ($a -match '^\d+$' -and [int]$a -ge 1 -and [int]$a -le $Items.Count) { return [int]$a - 1 }
        Write-Host 'Invalid choice.' -ForegroundColor Yellow
    }
}

function New-OfficeConfig {
    param($ProductId, $Channel, $Arch, [string[]]$Languages, [string[]]$Exclude,
          [string]$Display, [bool]$RemoveMSI, [string]$Key)
    $xml = New-Object System.Xml.XmlDocument
    $cfg = $xml.AppendChild($xml.CreateElement('Configuration'))
    $add = $cfg.AppendChild($xml.CreateElement('Add'))
    $add.SetAttribute('OfficeClientEdition', $Arch)
    $add.SetAttribute('Channel', $Channel)
    $prod = $add.AppendChild($xml.CreateElement('Product'))
    $prod.SetAttribute('ID', $ProductId)
    if ($Key) { $prod.SetAttribute('PIDKEY', $Key) }
    foreach ($l in $Languages) { $e = $prod.AppendChild($xml.CreateElement('Language')); $e.SetAttribute('ID', $l) }
    foreach ($a in $Exclude)   { $e = $prod.AppendChild($xml.CreateElement('ExcludeApp')); $e.SetAttribute('ID', $a) }
    if ($RemoveMSI) { [void]$cfg.AppendChild($xml.CreateElement('RemoveMSI')) }
    $d = $cfg.AppendChild($xml.CreateElement('Display'))
    $d.SetAttribute('Level', $Display); $d.SetAttribute('AcceptEULA', 'TRUE')
    if ($Key) {
        $p = $cfg.AppendChild($xml.CreateElement('Property'))
        $p.SetAttribute('Name', 'AUTOACTIVATE'); $p.SetAttribute('Value', '1')
    }
    $sw = New-Object System.IO.StringWriter
    $set = New-Object System.Xml.XmlWriterSettings; $set.Indent = $true; $set.OmitXmlDeclaration = $true
    $w = [System.Xml.XmlWriter]::Create($sw, $set); $xml.Save($w); $w.Close()
    $sw.ToString()
}

function Find-Setup {
    $dirs = @($OdtPath, $OdtDir, $Root, 'C:\ODT') | Where-Object { $_ }
    foreach ($d in $dirs) { $p = Join-Path $d 'setup.exe'; if (Test-Path $p) { return $p } }
}

function Get-Odt {
    Write-Host 'Downloading the Office Deployment Tool from microsoft.com...' -ForegroundColor Cyan
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $page = Invoke-WebRequest -Uri $OdtPage -UseBasicParsing
    $url = [regex]::Match($page.Content, 'https://download\.microsoft\.com/download/[^"\s]+?officedeploymenttool_[\d\-]+\.exe').Value
    if (-not $url) { throw "Could not find the download link on $OdtPage. Download the ODT manually and put setup.exe in $OdtDir" }
    New-Item -ItemType Directory -Force $OdtDir | Out-Null
    $exe = Join-Path $env:TEMP 'officedeploymenttool.exe'
    Invoke-WebRequest -Uri $url -OutFile $exe -UseBasicParsing
    $sig = Get-AuthenticodeSignature $exe
    if ($sig.Status -ne 'Valid' -or $sig.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
        Remove-Item $exe -Force; throw "Downloaded file is not signed by Microsoft ($($sig.Status)). Aborting."
    }
    Start-Process $exe -ArgumentList "/quiet /extract:`"$OdtDir`"" -Wait
    Remove-Item $exe -Force
    $setup = Join-Path $OdtDir 'setup.exe'
    if (-not (Test-Path $setup)) { throw 'Extraction did not produce setup.exe.' }
    Write-Host "ODT ready: $setup" -ForegroundColor Green
    $setup
}

# Plain-words explanations of common setup errors (see README error table).
$ErrorHints = @(
    @{ Pattern = '30175-45.*1289|1289.*30175-45'; Text = 'Setup failed within seconds while applying the configuration: the product key was rejected. Reinstall without a key and enter it in Word (Enter product key).' }
    @{ Pattern = '30175-45';                      Text = 'Real install failure. Run the cleanup step, restart Windows, try again. If it repeats, run the readiness check and look at antivirus/filter drivers and disk errors.' }
    @{ Pattern = '17004';                         Text = 'Setup reports that the install failed. Look for the error that came before it in the log.' }
    @{ Pattern = '30015|30088|30094';             Text = 'Another install is running, or a restart is pending. Restart Windows and retry.' }
    @{ Pattern = '0x80070070|disk full|30182';    Text = 'Not enough disk space. Free at least 10 GB on the system drive.' }
)

function Get-ErrorHint([string]$Text) {
    foreach ($h in $ErrorHints) { if ($Text -match $h.Pattern) { return $h.Text } }
}

function Invoke-Step([string]$Script, [string[]]$StepArgs) {
    $p = Join-Path $Root $Script
    if (-not (Test-Path $p)) { Write-Warning "Not found: $p"; return }
    $stepLog = Join-Path $script:LogDir ('debug-{0}-{1}.log' -f [IO.Path]::GetFileNameWithoutExtension($Script), (Get-Date -Format 'HHmmss'))
    $inner = "Start-Transcript -Path '$stepLog' | Out-Null; try { & '$p' $($StepArgs -join ' ') } finally { Stop-Transcript | Out-Null }"
    Write-Host "Step log: $stepLog" -ForegroundColor DarkGray
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -Command $inner
}

# ---------------------------------------------------------------- download only
if ($DownloadOnly) { [void](Get-Odt); return }

# ---------------------------------------------------------------- guided preparation
# Interactive runs start with a readiness check and an optional cleanup, then go on to the install.
if (-not $Product -and -not $GenerateOnly) {
    $guided = $true
    while ($guided) {
        $step = Read-Choice 'What do you want to do?' @(
            'Install Office now'
            'Check this PC first (read-only readiness check)'
            'Scan for leftover Office (read-only)'
            'Remove leftover Office (asks YES before removing)')
        switch ($step) {
            0 { $guided = $false }
            1 { Invoke-Step '1-Diagnose\Test-OfficeInstallReadiness.ps1' @('-Quick') }
            2 { Invoke-Step '2-Cleanup\Remove-OfficeLeftovers.ps1' @('-ScanOnly') }
            3 { Invoke-Step '2-Cleanup\Remove-OfficeLeftovers.ps1' @(); Write-Host 'Restart Windows before installing.' -ForegroundColor Yellow }
        }
    }
}

# ---------------------------------------------------------------- choices
$interactive = -not $Product
if ($Product) {
    $item = $Catalog | Where-Object { $_.Id -eq $Product }
    if (-not $item) { $item = @{ Name = $Product; Id = $Product; Channel = 'Current' }; Write-Warning "Product '$Product' is not in the catalog; using channel Current." }
} else {
    $item = $Catalog[(Read-Choice 'Which Office product?' ($Catalog | ForEach-Object { $_.Name }))]
}

if (-not $Language) {
    if ($interactive) {
        $in = Read-Host "`nLanguage(s), comma-separated, first is primary [lt-lt]"
        $Language = if ($in) { $in -split '\s*,\s*' } else { @('lt-lt') }
    } else { $Language = @('lt-lt') }
}

if (-not $ExcludeApp) {
    if ($interactive) {
        Write-Host "`nApps to exclude: $($AllApps -join ', ')"
        $in = Read-Host 'Comma-separated, "none" for nothing [Teams]'
        $ExcludeApp = if (-not $in) { @('Teams') } elseif ($in -eq 'none') { @() } else { $in -split '\s*,\s*' }
    } else { $ExcludeApp = @('Teams') }
} elseif ($ExcludeApp -contains 'none') { $ExcludeApp = @() }

if ($interactive) {
    if ((Read-Host "`nRemove older MSI-based Office first? (y/N)") -match '^[yY]') { $RemoveMSI = $true }
    if ((Read-Host 'Show the installer window? (Y/n)') -match '^[nN]') { $Display = 'None' }
}

# ---------------------------------------------------------------- generate-only
if ($GenerateOnly) {
    $xml = New-OfficeConfig $item.Id $item.Channel $Arch $Language $ExcludeApp $Display $RemoveMSI.IsPresent $null
    New-Item -ItemType Directory -Force $OutDir | Out-Null
    $file = Join-Path $OutDir "$($item.Id)-$($Language -join '+')-x$Arch.xml"
    Set-Content -Path $file -Value $xml -Encoding UTF8
    Write-Host "Written: $file (no key included)" -ForegroundColor Green
    return
}

# ---------------------------------------------------------------- admin + setup.exe
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole('Administrator')
if (-not $isAdmin) { throw 'Run this script as administrator (use Start-OfficeInstaller.cmd).' }

$setup = Find-Setup
if (-not $setup) {
    if ($interactive -and (Read-Host 'setup.exe (Office Deployment Tool) not found. Download it from Microsoft now? (Y/n)') -match '^[nN]') { throw 'setup.exe is required.' }
    $setup = Get-Odt
}

# ---------------------------------------------------------------- key (never stored)
$key = $null
if (-not $NoKey -and ($AskKey -or $interactive)) {
    $sec = Read-Host "`nProduct key (hidden input; Enter to skip and activate later in the app)" -AsSecureString
    $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)).Trim().ToUpper()
    if ($plain) {
        if ($plain -notmatch $KeyPattern) { throw 'That is not a valid key format (XXXXX-XXXXX-XXXXX-XXXXX-XXXXX).' }
        $key = $plain
    }
}

# ---------------------------------------------------------------- run
$cfgPath = Join-Path $env:TEMP ("office-" + [guid]::NewGuid().ToString('N') + '.xml')
$started = Get-Date
try {
    $xml = New-OfficeConfig $item.Id $item.Channel $Arch $Language $ExcludeApp $Display $RemoveMSI.IsPresent $key
    Set-Content -Path $cfgPath -Value $xml -Encoding UTF8
    Write-Host "`nInstalling: $($item.Name) | x$Arch | $($Language -join ', ') | key: $(if ($key) { 'provided' } else { 'none' })" -ForegroundColor Cyan
    $p = Start-Process $setup -ArgumentList "/configure `"$cfgPath`"" -Wait -PassThru
    if ($p.ExitCode -eq 0) { Write-Host 'Finished (exit code 0).' -ForegroundColor Green }
    else {
        Write-Host "Setup exited with code $($p.ExitCode)." -ForegroundColor Red
        # explain the newest setup log errors in plain words
        $recent = Get-ChildItem $env:TEMP -Filter "$env:COMPUTERNAME-*.log" -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -ge $started } | Select-Object -ExpandProperty FullName
        $logText = ($recent | ForEach-Object { Get-Content $_ -ErrorAction SilentlyContinue } | Select-String -Pattern '30175|17004|30015|30088|30094|30182|0x80070070' | Out-String)
        $hint = Get-ErrorHint "$logText 0-2031 ($($p.ExitCode))"
        if ($hint) { Write-Host "Likely cause: $hint" -ForegroundColor Yellow }
    }
    # short job log in the Cases folder; never contains a key
    $caseDir = Join-Path $Root ("Cases\{0}_{1}" -f (Get-Date -Format 'yyyy-MM-dd'), $env:COMPUTERNAME)
    New-Item -ItemType Directory -Force $caseDir | Out-Null
    Add-Content (Join-Path $caseDir 'install-jobs.log') ("{0:s} | {1} | x{2} | {3} | key: {4} | exit {5}" -f (Get-Date), $item.Id, $Arch, ($Language -join '+'), $(if ($key) { 'provided' } else { 'none' }), $p.ExitCode)
    if ($key -and $p.ExitCode -ne 0) { Write-Host 'If it failed within seconds with 30175-45 / 17004, the key was rejected: rerun and press Enter at the key prompt, then enter the key in Word.' -ForegroundColor Yellow }
}
finally {
    Remove-Item $cfgPath -Force -ErrorAction SilentlyContinue
    if ($key) {
        # setup logs keep the PIDKEY in plain text: redact it
        Get-ChildItem $env:TEMP -Filter '*.log' -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -ge $started -and $_.Name -like "$env:COMPUTERNAME-*" } |
            ForEach-Object { try { (Get-Content $_.FullName -Raw).Replace($key, 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX') | Set-Content $_.FullName } catch {} }
    }
    # Office's own setup logs (the separate setup window): copy them next to the transcripts
    Get-ChildItem $env:TEMP -Filter "$env:COMPUTERNAME-*.log" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $started } |
        ForEach-Object { try { Copy-Item $_.FullName (Join-Path $script:LogDir ('officesetup-' + $_.Name)) -Force } catch {} }
    $key = $null; $plain = $null
}
} finally {
    if ($script:Transcribing) { try { Stop-Transcript | Out-Null } catch {}; $script:Transcribing = $false }
    # a typed key is never echoed, but scrub the transcript anyway
    Get-ChildItem $script:LogDir -Filter '*.log' -ErrorAction SilentlyContinue | ForEach-Object {
        try { (Get-Content $_.FullName -Raw) -replace '[A-Z0-9]{5}(-[A-Z0-9]{5}){4}', 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX' | Set-Content $_.FullName } catch {}
    }
}
