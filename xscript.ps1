# =====================================================
#  Silent Office + Office & OneNote Uninstaller (Auto Restart)
#  Compatible with: irm <url> | iex
#  Requires: Admin rights
# =====================================================

# --- Ensure running as Administrator ---
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "🔒 Elevating privileges..."
    Start-Process powershell "-ExecutionPolicy Bypass -NoProfile -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Write-Host "🧹 Starting full silent uninstall of Microsoft Office and OneNote..." -ForegroundColor Cyan

# --- Gather uninstall commands for Office and OneNote ---
$officeProducts = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall" , "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" |
    Where-Object { ($_.GetValue("DisplayName") -like "*Microsoft Office*" -or $_.GetValue("DisplayName") -like "*OneNote*") -and ($_.GetValue("UninstallString")) }

foreach ($app in $officeProducts) {
    $displayName = $app.GetValue("DisplayName")
    $uninstallCmd = $app.GetValue("UninstallString")

    if ($uninstallCmd) {
        Write-Host "🗑️ Uninstalling: $displayName ..."
        try {
            Start-Process cmd.exe "/c $uninstallCmd /quiet /norestart" -Wait -WindowStyle Hidden
        } catch {
            Write-Warning "Failed to uninstall $displayName"
        }
    }
}

# --- Remove leftover folders ---
$folders = @(
    "$env:ProgramFiles\Microsoft Office",
    "$env:ProgramFiles (x86)\Microsoft Office",
    "$env:ProgramFiles\Common Files\Microsoft Shared\OFFICE*",
    "$env:ProgramData\Microsoft\Office",
    "$env:LOCALAPPDATA\Microsoft\Office",
    "$env:APPDATA\Microsoft\Office",
    "$env:LOCALAPPDATA\Microsoft\OneNote",
    "$env:APPDATA\Microsoft\OneNote"
)
foreach ($f in $folders) {
    if (Test-Path $f) {
        Write-Host "🧽 Removing folder: $f"
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue $f
    }
}

# --- Clean registry remnants ---
$regPaths = @(
    "HKCU:\Software\Microsoft\Office",
    "HKLM:\Software\Microsoft\Office",
    "HKLM:\Software\WOW6432Node\Microsoft\Office"
)
foreach ($reg in $regPaths) {
    if (Test-Path $reg) {
        Write-Host "🧹 Cleaning registry key: $reg"
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue $reg
    }
}

Write-Host "✅ Microsoft Office and OneNote have been completely removed." -ForegroundColor Green
Write-Host "💻 System will restart automatically in 10 seconds..." -ForegroundColor Yellow

Start-Sleep -Seconds 10
Restart-Computer -Force
