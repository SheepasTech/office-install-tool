@echo off
setlocal
title Office install (Office Deployment Tool)
cd /d "%~dp0"
:: Simple fixed-config installer. For other editions/languages use ..\Install-Office.ps1.

net session >nul 2>&1
if errorlevel 1 goto :notadmin

set "SETUP=%~dp0setup.exe"
if not exist "%SETUP%" if exist "C:\ODT\setup.exe" set "SETUP=C:\ODT\setup.exe"
if not exist "%SETUP%" goto :nosetup

set "CFG=%~1"
if not "%CFG%"=="" goto :run

echo.
echo  Office Home and Student 2024 - choose a configuration:
echo.
echo    1 = Lithuanian - recommended, enter the key in Word afterwards
echo    2 = English
echo    Q = quit
echo.
choice /c 12Q /n /m " Select 1, 2 or Q: "
if errorlevel 3 exit /b 0
if errorlevel 2 set "CFG=%~dp0HomeStudent2024-en-NOKEY.xml"& goto :run
set "CFG=%~dp0HomeStudent2024-lt-NOKEY.xml"

:run
if not exist "%CFG%" goto :nocfg
echo.
echo  Setup    : %SETUP%
echo  Config   : %CFG%
echo  Started  : %date% %time%
echo.
"%SETUP%" /configure "%CFG%"
set "RC=%errorlevel%"
echo.
echo  Finished with exit code %RC%  (0 = success)
if not "%RC%"=="0" echo  Failed? Run 1-Diagnose\Test-OfficeInstallReadiness.ps1 now - it collects the setup logs.
pause
exit /b %RC%

:notadmin
echo  Right-click Install-Office.cmd and choose "Run as administrator".
pause
exit /b 1

:nosetup
echo  setup.exe not found in this folder or in C:\ODT.
echo  Download the Office Deployment Tool from microsoft.com, run it,
echo  and extract setup.exe into this 3-Install folder.
pause
exit /b 1

:nocfg
echo  Configuration file not found: %CFG%
pause
exit /b 1
