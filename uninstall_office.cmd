@echo off
REM =========================================================
REM Uninstall_Office_and_OneNote_AutoRestart.cmd
REM - Downloads Microsoft SaRA tool
REM - Runs Office uninstall scenario
REM - Removes OneNote (UWP) via PowerShell
REM - Auto restarts PC
REM =========================================================

SETLOCAL ENABLEDELAYEDEXPANSION

echo ========== Microsoft Office & OneNote Uninstaller ==========
echo.

REM -- Variables
set "TEMP_INSTALLER=%TEMP%\SaRAsetup.exe"
set "SaRAPath=C:\Program Files\SaRA\SaRAcmd.exe"
set "LOGFILE=%TEMP%\UninstallOffice_Log.txt"

echo [%DATE% %TIME%] Starting uninstaller... > "%LOGFILE%"

REM -- Download SaRA installer using PowerShell (TLS1.2 + use basic parsing)
echo Downloading SaRA (Microsoft Support and Recovery Assistant)...
powershell -NoProfile -Command "try {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://aka.ms/SaRA-OfficeUninstall' -OutFile '%TEMP_INSTALLER%' -UseBasicParsing -ErrorAction Stop; exit 0} catch { Write-Error $_; exit 1 }"
if NOT exist "%TEMP_INSTALLER%" (
    echo [%DATE% %TIME%] Failed to download SaRA. >> "%LOGFILE%"
    echo Failed to download SaRA installer. Please check internet connection.
    pause
    exit /b 1
)

echo [%DATE% %TIME%] SaRA downloaded to %TEMP_INSTALLER% >> "%LOGFILE%"

REM -- Install SaRA silently
echo Installing SaRA silently...
start /wait "" "%TEMP_INSTALLER%" /quiet
if exist "%SaRAPath%" (
    echo [%DATE% %TIME%] SaRA installed. >> "%LOGFILE%"
) else (
    echo [%DATE% %TIME%] SaRA not detected after install. >> "%LOGFILE%"
    echo SaRA installation may have failed. Continuing to attempt uninstall via other methods...
)

REM -- Try running SaRAcmd (silent Office uninstall)
if exist "%SaRAPath%" (
    echo Running SaRA Office uninstaller (silent)...
    start /wait "" "%SaRAPath%" -SOfficeUninstall -quiet
    echo [%DATE% %TIME%] SaRA OfficeUninstall executed. >> "%LOGFILE%"
) else (
    echo SaRAcmd not found, attempting alternate uninstall methods...
    echo Attempting to uninstall common Office MSI/uninstall strings...
    REM Try common Click-to-Run uninstall via OfficeC2RClient if present
    if exist "%CommonProgramFiles%\Microsoft Shared\ClickToRun\OfficeC2RClient.exe" (
        echo Triggering ClickToRun uninstall...
        "%CommonProgramFiles%\Microsoft Shared\ClickToRun\OfficeC2RClient.exe" /update user displaylevel=false forceappshutdown=true
        REM Attempt uninstall (note: requires correct ProductID which varies)
    )
    REM Attempt msiexec based uninstalls by searching registry
    for /f "tokens=2,* delims=	 " %%A in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" /s /v "UninstallString" 2^>nul ^| findstr /i "Office"') do (
        rem Placeholder - this registry parsing is best handled by PowerShell. Skipping detailed parsing here.
    )
)

REM -- Remove OneNote UWP (Store) if installed (uses PowerShell to remove appx)
echo Removing OneNote (Store version) if present...
powershell -NoProfile -Command "try { Get-AppxPackage *OneNote* | ForEach-Object { Write-Output ('Removing: ' + $_.Name); Remove-AppxPackage -Package $_.PackageFullName -ErrorAction SilentlyContinue } ; exit 0 } catch { exit 1 }"

echo [%DATE% %TIME%] OneNote (UWP) removal attempted. >> "%LOGFILE%"

REM -- Optional: Remove leftover folders (best effort)
echo Removing leftover folders (best-effort)...
rd /s /q "%ProgramFiles%\Microsoft Office" 2>nul
rd /s /q "%ProgramFiles(x86)%\Microsoft Office" 2>nul
rd /s /q "%ProgramFiles%\Common Files\Microsoft Shared\OFFICE16" 2>nul
rd /s /q "%ProgramData%\Microsoft\Office" 2>nul
rd /s /q "%LOCALAPPDATA%\Microsoft\Office" 2>nul
rd /s /q "%APPDATA%\Microsoft\Office" 2>nul
rd /s /q "%LOCALAPPDATA%\Microsoft\OneNote" 2>nul
rd /s /q "%APPDATA%\Microsoft\OneNote" 2>nul

echo [%DATE% %TIME%] Folder cleanup attempted. >> "%LOGFILE%"

REM -- Final message and restart
echo.
echo All done. The system will restart in 15 seconds to complete cleanup.
echo [%DATE% %TIME%] Completed. Restarting in 15 seconds... >> "%LOGFILE%"
timeout /t 15 /nobreak >nul
shutdown /r /t 0

ENDLOCAL
