# Shared debug logging. Dot-source it, then:
#   $null = Start-DebugLog 'cleanup'      # one transcript per window/process
#   try { ... } catch { Write-DebugCrash $_; throw } finally { Stop-DebugLog }
# Logs: %TEMP%\Debug\office-install-tool\<yyyy-MM-dd>\<name>-<yyyyMMdd-HHmmss>.log
# Desktop\Debug.lnk points at %TEMP%\Debug.

$script:DebugRoot    = Join-Path $env:TEMP 'Debug'
$script:DebugProject = Join-Path $script:DebugRoot 'office-install-tool'
$script:DebugDir     = Join-Path $script:DebugProject (Get-Date -Format 'yyyy-MM-dd')
$script:DebugFile    = $null
$script:DebugActive  = $false

function Start-DebugLog {
    param([Parameter(Mandatory = $true)][string]$Name)
    try {
        New-Item -ItemType Directory -Force $script:DebugDir | Out-Null
        $script:DebugFile = Join-Path $script:DebugDir ('{0}-{1}.log' -f $Name, (Get-Date -Format 'yyyyMMdd-HHmmss'))
        Start-Transcript -Path $script:DebugFile -Force | Out-Null
        $script:DebugActive = $true
        Write-Host "Debug log: $script:DebugFile" -ForegroundColor DarkGray
        return $script:DebugFile
    } catch {
        Write-Warning "Could not start the debug log in $($script:DebugDir): $($_.Exception.Message)"
        return $null
    }
}

# Written to a separate file, so it survives even when the transcript is lost or locked.
function Write-DebugCrash {
    param($ErrorRecord)
    try {
        New-Item -ItemType Directory -Force $script:DebugDir | Out-Null
        $f = Join-Path $script:DebugDir ('crash-{0}.txt' -f (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
        $text = @(
            "Time    : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            "Process : PID $PID, $([Environment]::CommandLine)"
            "Log     : $script:DebugFile"
            "Error   : $($ErrorRecord.Exception.Message)"
            "Where   : $($ErrorRecord.InvocationInfo.PositionMessage)"
            "Stack   : $($ErrorRecord.ScriptStackTrace)"
        ) -join "`r`n"
        Set-Content -LiteralPath $f -Value $text -Encoding UTF8
        Write-Host "Crash details saved: $f" -ForegroundColor Red
    } catch { }
}

# A typed key is never echoed, but scrub key-shaped text from this window's log anyway.
function Stop-DebugLog {
    if ($script:DebugActive) { try { Stop-Transcript | Out-Null } catch { }; $script:DebugActive = $false }
    if ($script:DebugFile -and (Test-Path -LiteralPath $script:DebugFile)) {
        try { (Get-Content -LiteralPath $script:DebugFile -Raw) -replace '[A-Z0-9]{5}(-[A-Z0-9]{5}){4}', 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX' | Set-Content -LiteralPath $script:DebugFile } catch { }
    }
}

# Desktop\Debug.lnk -> %TEMP%\Debug (asks once; -Silent creates it without asking).
function Initialize-DebugShortcut {
    param([switch]$Silent)
    try {
        $lnk = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Debug.lnk'
        if (Test-Path -LiteralPath $lnk) { return }
        $make = $Silent -or -not [Environment]::UserInteractive -or ((Read-Host '  Put a "Debug" shortcut (to the logs) on the desktop? (Y/n)') -notmatch '^[nN]')
        if ($make) {
            New-Item -ItemType Directory -Force $script:DebugRoot | Out-Null
            $sc = (New-Object -ComObject WScript.Shell).CreateShortcut($lnk); $sc.TargetPath = $script:DebugRoot; $sc.Save()
            Write-Host "Shortcut created: $lnk" -ForegroundColor Green
        }
    } catch { Write-Warning "Could not create the Debug shortcut ($($_.Exception.Message)). Logs are in $($script:DebugRoot)" }
}
