# =====================================================
#  Visible Office + OneNote Uninstaller (Auto Restart)
#  Works with irm <url> | iex
#  Displays progress while running
# =====================================================

# --- Force window visible ---
$Host.UI.RawUI.WindowTitle = "Microsoft Office & OneNote Uninstaller"
Write-Host ""
Write-Host "This script is running created by RICTU"
Write-Host "🧹 Starting full uninstall of Microsoft Office and OneNote..." -ForegroundColor Cyan
Write-Host "------------------------------------------------------------"
Write-Host ""

# --- Ensure running as Administrator ---
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "🔒 Elevating privileges..." -ForegroundColor Yellow
    Start-Process powershell "-ExecutionPolicy Bypass -NoProfile -WindowStyle Normal -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# --- Uninstall Office and OneNote products ---
$officeProducts = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall" , "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" |
    Where-Object { ($_.GetValue("DisplayName") -like "*Microsoft Office*" -or $_.GetValue("DisplayName") -like "*OneNote*") -and ($_.GetValue("UninstallString")) }

if ($officeProducts.Count -eq 0) {
    Write-Host "⚠️ No Office or OneNote products detected." -ForegroundColor Yellow
} else {
    foreach ($app in $officeProducts) {
        $displayName = $app.GetValue("DisplayName")
        $uninstallCmd = $app.GetValue("UninstallString")
        Write-Host "🗑️ Uninstalling: $displayName ..." -ForegroundColor White
        try {
            Start-Process cmd.exe "/c $uninstallCmd /quiet /norestart" -Wait -WindowStyle Normal
            Write-Host "✅ Removed: $displayName" -ForegroundColor Green
        } catch {
            Write-Warning "❌ Failed to uninstall $displayName"
        }
    }
}

# --- Terminate Click-to-Run processes (Office 365) ---
Write-Host ""
Write-Host "🧩 Stopping Office Click-to-Run services..." -ForegroundColor Cyan
Stop-Service -Name "ClickToRunSvc" -ErrorAction SilentlyContinue
Get-Process | Where-Object { $_.ProcessName -match "OfficeClickToRun|OfficeC2RClient|OneNote|WinWord|Excel|PowerPoint|Outlook" } | Stop-Process -Force -ErrorAction SilentlyContinue

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

Write-Host ""
Write-Host "✅ Microsoft Office and OneNote have been completely removed." -ForegroundColor Green
Write-Host "💻 System will restart automatically in 15 seconds..." -ForegroundColor Yellow

Start-Sleep -Seconds 15
Restart-Computer -Force
