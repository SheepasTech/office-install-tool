# Office license discovery: which license keys/subscriptions exist and which install they belong to.
# Only the LAST 5 characters of a key are ever read or shown. Every source is wrapped in try/catch.

$script:OfficeAppId = '0ff1ce15-a989-479d-af46-f275c6370663'
$script:LicStatusText = @{ 0 = 'Unlicensed'; 1 = 'Licensed'; 2 = 'Out-of-box grace'; 3 = 'Out-of-tolerance grace'; 4 = 'Non-genuine grace'; 5 = 'Notification'; 6 = 'Extended grace' }
$script:C2RChannels = @{
    '492350f6-3a01-4f97-b9c0-c7c6ddf67d60' = 'Current'
    '64256afe-f5d9-4f86-8936-8840a6a4f5be' = 'Current (Preview)'
    '55336b82-a18d-4dd6-b5f6-9e5095c314a6' = 'Monthly Enterprise'
    '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114' = 'Semi-Annual Enterprise'
    'b8f9b850-328d-4355-9145-c59439a0c4cf' = 'Semi-Annual Enterprise (Preview)'
    'f2e724c1-748f-4b47-8fb8-8e0d210e9208' = 'Perpetual VL 2019 (LTSC)'
    '5030841d-c919-4594-8d2d-84ae4f96e58e' = 'Perpetual VL 2021 (LTSC)'
}

function Get-SafeProp {
    param($Object, [string]$Name)
    if ($null -ne $Object -and $Object.PSObject.Properties[$Name]) { return $Object.$Name }
    return $null
}

function Hide-Email {
    param([string]$Text)
    if (-not $Text) { return $Text }
    return [regex]::Replace($Text, '([A-Za-z0-9._%+-]{2})[A-Za-z0-9._%+-]*@', '$1***@')
}

# What Office is installed, and by which mechanism (Click-to-Run / MSI / Microsoft Store).
function Get-OfficeInstallSummary {
    $list = New-Object System.Collections.Generic.List[object]
    try {
        $cfgKey = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
        if (Test-Path -LiteralPath $cfgKey) {
            $cfg = Get-ItemProperty -LiteralPath $cfgKey -ErrorAction Stop
            $cdn = [string](Get-SafeProp $cfg 'CDNBaseUrl')
            $guid = if ($cdn -match '([0-9a-fA-F-]{36})$') { $Matches[1].ToLower() } else { '' }
            $chan = if ($guid -and $script:C2RChannels.ContainsKey($guid)) { $script:C2RChannels[$guid] } elseif ($guid) { "unknown channel $guid" } else { '' }
            $ids = @(([string](Get-SafeProp $cfg 'ProductReleaseIds')) -split ',' | Where-Object { $_ })
            $list.Add([pscustomobject]@{ Kind = 'Click-to-Run'; Name = ($ids -join ', '); ReleaseIds = $ids
                Version = [string](Get-SafeProp $cfg 'VersionToReport'); Platform = [string](Get-SafeProp $cfg 'Platform'); Channel = $chan })
        }
    } catch { Write-Host "  (license scan) Click-to-Run configuration unreadable: $($_.Exception.Message)" -ForegroundColor DarkYellow }
    try {
        foreach ($root in 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall') {
            if (-not (Test-Path -LiteralPath $root)) { continue }
            foreach ($k in Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue) {
                $p = Get-ItemProperty -LiteralPath $k.PSPath -ErrorAction SilentlyContinue
                $n = [string](Get-SafeProp $p 'DisplayName')
                if ($n -match '^Microsoft Office (?!.*(Click-to-Run|Language Pack|Proofing|Shared|Web Components))' -and (Get-SafeProp $p 'WindowsInstaller') -eq 1 -and (Get-SafeProp $p 'SystemComponent') -ne 1) {
                    $list.Add([pscustomobject]@{ Kind = 'MSI'; Name = $n; ReleaseIds = @($n); Version = [string](Get-SafeProp $p 'DisplayVersion'); Platform = ''; Channel = '' })
                }
            }
        }
    } catch { Write-Host "  (license scan) MSI entries unreadable: $($_.Exception.Message)" -ForegroundColor DarkYellow }
    try {
        foreach ($a in @('Microsoft.Office.Desktop*', 'Microsoft.MicrosoftOfficeHub' | ForEach-Object { Get-AppxPackage -Name $_ -ErrorAction Stop })) {
            $list.Add([pscustomobject]@{ Kind = 'Microsoft Store'; Name = $a.Name; ReleaseIds = @($a.Name); Version = [string]$a.Version; Platform = [string]$a.Architecture; Channel = '' })
        }
    } catch { Write-Host "  (license scan) Store packages unreadable: $($_.Exception.Message)" -ForegroundColor DarkYellow }
    return ,$list.ToArray()
}

# Match a license name such as "Office 16, Office21ProPlus2021R_Grace edition" to an install.
function Find-LicenseOwner {
    param([string]$LicenseName, $Installs)
    $ln = ([string]$LicenseName).ToLower()
    foreach ($i in $Installs) {
        if ($i.Kind -eq 'Click-to-Run') {
            foreach ($id in @($i.ReleaseIds)) {
                if ($id -notmatch '^(O365\w*?|ProPlus|Professional|HomeStudent|HomeBusiness|Home|Standard|Mondo)(20\d\d)?(Retail|Volume)?$') { continue }
                $fam = $Matches[1].ToLower(); $yr = $Matches[2]
                if ($ln -notmatch [regex]::Escape($fam)) { continue }
                if ($yr -and $ln -notmatch "$yr|office$($yr.Substring(2))") { continue }
                return "Click-to-Run: $id"
            }
        }
        elseif ($i.Kind -eq 'MSI') {
            $fam = (($i.Name -replace '^Microsoft Office ', '' -replace '[^A-Za-z0-9]', '')).ToLower()
            if ($fam.Length -ge 6 -and $ln -replace '[^a-z0-9]', '' -match [regex]::Escape($fam.Substring(0, 6))) { return "MSI: $($i.Name)" }
        }
    }
    return $null
}

function Get-OfficeLicenseInfo {
    $installs = [object[]](Get-OfficeInstallSummary)
    $out = New-Object System.Collections.Generic.List[object]

    # 1. Licensing service (WMI): the authoritative list of installed product keys
    foreach ($cls in 'SoftwareLicensingProduct', 'OfficeSoftwareProtectionProduct') {
        try {
            $rows = @(Get-CimInstance -ClassName $cls -Filter "ApplicationId='$script:OfficeAppId' AND PartialProductKey IS NOT NULL" -ErrorAction Stop)
            foreach ($r in $rows) {
                $desc = [string](Get-SafeProp $r 'Description')
                $chan = [string](Get-SafeProp $r 'ProductKeyChannel')
                if (-not $chan) {
                    $chan = if ($desc -match 'KMSCLIENT') { 'Volume: KMS client' } elseif ($desc -match 'MAK') { 'Volume: MAK' } elseif ($desc -match 'OEM') { 'OEM' } elseif ($desc -match 'RETAIL') { 'Retail' } elseif ($desc -match 'O365|SUBSCRIPTION|TIMEBASED') { 'Subscription' } else { '' }
                }
                $st = Get-SafeProp $r 'LicenseStatus'
                $kms = [string](Get-SafeProp $r 'KeyManagementServiceMachine')
                $owner = Find-LicenseOwner ([string]$r.Name) $installs
                $out.Add([pscustomobject]@{
                    Source = "Licensing service ($cls)"; Name = [string]$r.Name; Description = $desc; Channel = $chan
                    Status = $(if ($null -ne $st -and $script:LicStatusText.ContainsKey([int]$st)) { $script:LicStatusText[[int]$st] } else { "$st" })
                    Last5 = [string](Get-SafeProp $r 'PartialProductKey'); Family = [string](Get-SafeProp $r 'LicenseFamily')
                    Note = $(if ($kms) { "KMS host: $kms" } else { '' })
                    BelongsTo = $(if ($owner) { $owner } else { 'no matching install (orphan license)' }) })
            }
        } catch { Write-Host "  (license scan) $cls not available: $($_.Exception.Message)" -ForegroundColor DarkGray }
    }

    # 2. Registration entries (product name / edition the install was registered with)
    try {
        foreach ($base in 'HKLM:\SOFTWARE\Microsoft\Office\16.0\Registration', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Registration', 'HKLM:\SOFTWARE\Microsoft\Office\15.0\Registration', 'HKLM:\SOFTWARE\Microsoft\Office\14.0\Registration') {
            if (-not (Test-Path -LiteralPath $base)) { continue }
            foreach ($k in Get-ChildItem -LiteralPath $base -ErrorAction SilentlyContinue) {
                $p = Get-ItemProperty -LiteralPath $k.PSPath -ErrorAction SilentlyContinue
                $pn = [string](Get-SafeProp $p 'ProductName')
                if (-not $pn) { continue }
                $out.Add([pscustomobject]@{ Source = 'Registry registration'; Name = $pn; Description = "Product ID: $(Get-SafeProp $p 'ProductID')"; Channel = ''
                    Status = ''; Last5 = ''; Family = [string](Get-SafeProp $p 'ConvertToEdition'); Note = $k.Name; BelongsTo = (Find-LicenseOwner $pn $installs) })
            }
        }
    } catch { Write-Host "  (license scan) registration keys unreadable: $($_.Exception.Message)" -ForegroundColor DarkYellow }

    # 3. Microsoft 365 / account-based licensing (per user, no product key exists)
    try {
        $ln = 'HKCU:\Software\Microsoft\Office\16.0\Common\Licensing\LicensingNext'
        if (Test-Path -LiteralPath $ln) {
            $p = Get-ItemProperty -LiteralPath $ln -ErrorAction Stop
            foreach ($prop in $p.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' }) {
                $out.Add([pscustomobject]@{ Source = 'Account licensing (HKCU)'; Name = $prop.Name; Description = (Hide-Email ([string]$prop.Value)); Channel = 'Subscription / account'
                    Status = ''; Last5 = ''; Family = ''; Note = 'tied to a Microsoft account, not a product key'; BelongsTo = (Find-LicenseOwner $prop.Name $installs) })
            }
        }
        $id = 'HKCU:\Software\Microsoft\Office\16.0\Common\Identity\Identities'
        if (Test-Path -LiteralPath $id) {
            foreach ($k in Get-ChildItem -LiteralPath $id -ErrorAction SilentlyContinue) {
                $p = Get-ItemProperty -LiteralPath $k.PSPath -ErrorAction SilentlyContinue
                $em = Hide-Email ([string](Get-SafeProp $p 'EmailAddress'))
                if ($em) { $out.Add([pscustomobject]@{ Source = 'Signed-in account'; Name = $em; Description = [string](Get-SafeProp $p 'FriendlyName'); Channel = 'Account'; Status = ''; Last5 = ''; Family = ''; Note = ''; BelongsTo = '' }) }
            }
        }
    } catch { Write-Host "  (license scan) account licensing unreadable: $($_.Exception.Message)" -ForegroundColor DarkYellow }
    try {
        $dir = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\Licenses'
        if (Test-Path -LiteralPath $dir) {
            $n = @(Get-ChildItem -LiteralPath $dir -Recurse -File -ErrorAction SilentlyContinue).Count
            if ($n) { $out.Add([pscustomobject]@{ Source = 'License token cache'; Name = "$n token file(s)"; Description = $dir; Channel = 'Subscription / account'; Status = ''; Last5 = ''; Family = ''; Note = 'cached sign-in licenses; cleared by cleanup with -IncludeUserSettings'; BelongsTo = '' }) }
        }
    } catch { Write-Host "  (license scan) token cache unreadable: $($_.Exception.Message)" -ForegroundColor DarkYellow }

    return [pscustomobject]@{ Installs = $installs; Licenses = $out.ToArray() }
}
