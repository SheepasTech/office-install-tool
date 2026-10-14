@echo off
:: Launches Install-Office.ps1 elevated, with the execution policy bypassed for this run only.
net session >nul 2>&1
if errorlevel 1 (
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -ArgumentList '%*' -Verb RunAs"
  exit /b
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-Office.ps1" %*
pause
