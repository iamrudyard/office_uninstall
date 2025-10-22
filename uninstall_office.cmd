@echo off
title Microsoft Office Full Uninstaller
color 0C
echo ==============================================
echo     Microsoft Office Uninstall Utility
echo ==============================================
echo.

:: Download Microsoft Support and Recovery Assistant (SaRA)
set "SaRAExe=%TEMP%\SaRAsetup.exe"
echo [*] Downloading Microsoft SaRA tool...
powershell -Command "Invoke-WebRequest -Uri 'https://aka.ms/SaRA-OfficeUninstall' -OutFile '%SaRAExe%'"

echo [*] Installing SaRA...
"%SaRAExe%" /quiet /norestart

:: Run SaRA Office uninstall silently
echo [*] Running SaRA Office uninstall (includes OneNote)...
"%ProgramFiles%\SaRA\SaRA.exe" /SOfficeUninstall

:: For 64-bit Windows, also try Program Files (x86)
if exist "%ProgramFiles(x86)%\SaRA\SaRA.exe" "%ProgramFiles(x86)%\SaRA\SaRA.exe" /SOfficeUninstall

echo.
echo [+] Office (all versions) uninstall completed.
pause
exit
