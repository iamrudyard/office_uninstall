@echo off
setlocal
title Microsoft Office Silent Uninstaller 

echo [*] Detecting and uninstalling Microsoft Office...

:: --- Uninstall Click-to-Run installations (Office 2013–2021 / 365) ---
for %%A in (
    "%ProgramFiles%\Microsoft Office\root\Office16\setup.exe"
    "%ProgramFiles(x86)%\Microsoft Office\root\Office16\setup.exe"
) do (
    if exist "%%~A" (
        echo [*] Found Click-to-Run Office installation.
        echo [*] Uninstalling silently...
        "%%~A" /configure "%~dp0config.xml"
        goto :CLEAN
    )
)

:: --- Uninstall MSI-based installations (older Office versions) ---
echo [*] Checking for MSI-based Office installations...
for /f "tokens=2 delims== " %%i in ('wmic product where "name like 'Microsoft Office%%'" get IdentifyingNumber /value ^| find "IdentifyingNumber"') do (
    echo [*] Removing MSI product %%i ...
    msiexec /x %%i /quiet /norestart
)

:CLEAN
echo [*] Cleaning up leftover Office files and registry keys...
rmdir /s /q "%ProgramFiles%\Microsoft Office"
rmdir /s /q "%ProgramFiles(x86)%\Microsoft Office"
rmdir /s /q "%ProgramData%\Microsoft\Office"
rmdir /s /q "%AppData%\Microsoft\Office"
rmdir /s /q "%LocalAppData%\Microsoft\Office"

reg delete "HKCU\Software\Microsoft\Office" /f >nul 2>&1
reg delete "HKLM\Software\Microsoft\Office" /f >nul 2>&1
reg delete "HKLM\Software\WOW6432Node\Microsoft\Office" /f >nul 2>&1

echo [✓] Office uninstallation complete.
echo [*] System will reboot in 15 seconds...
shutdown /r /t 15 /c "Microsoft Office was removed successfully. Restarting your system to finalize cleanup."
exit /b 0
