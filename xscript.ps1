# ============================
# PowerShell Office Uninstall Bootstrap (No SaRA)
# ============================
# Usage:
# irm https://raw.githubusercontent.com/iamrudyard/office_uninstall/main/xscript.ps1 | iex

try {
    $url = 'https://raw.githubusercontent.com/iamrudyard/office_uninstall/refs/heads/main/uninstall_office.cmd'
    Write-Host "[*] Downloading Office uninstaller..." -ForegroundColor Cyan

    $response = Invoke-RestMethod -Uri $url -UseBasicParsing
    if ($response) {
        $tempFile = "$env:TEMP\office_uninstall_nosara.cmd"
        Set-Content -Path $tempFile -Value $response -Encoding ASCII
        Write-Host "[*] Running Office uninstaller silently..." -ForegroundColor Yellow
        Start-Process cmd.exe -ArgumentList "/c", $tempFile -Verb RunAs -WindowStyle Hidden
    } else {
        Write-Host "[!] Failed to download the uninstaller script." -ForegroundColor Red
    }
}
catch {
    Write-Host "[!] Error: $_" -ForegroundColor Red
}
