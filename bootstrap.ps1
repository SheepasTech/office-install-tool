# One-liner entry point: downloads the repo zip to %TEMP% and starts Install-Office.ps1 as administrator.
# Pin a branch, tag or commit with $env:OIT_REF (default: main).
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ref  = if ($env:OIT_REF) { $env:OIT_REF } else { 'main' }
$dest = Join-Path $env:TEMP 'office-install-tool'
$zip  = "$dest.zip"
if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
Invoke-WebRequest "https://github.com/SheepasTech/office-install-tool/archive/$ref.zip" -OutFile $zip -UseBasicParsing
Expand-Archive $zip -DestinationPath $dest -Force
Remove-Item $zip -Force
$root = (Get-ChildItem $dest -Directory | Select-Object -First 1).FullName
Start-Process powershell -Verb RunAs -ArgumentList "-NoExit -ExecutionPolicy Bypass -Command `"Set-Location '$root'; .\Install-Office.ps1`""
