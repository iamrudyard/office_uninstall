# ============================
# PowerShell Office Uninstall Bootstrap
# ============================
# Download and execute .cmd uninstaller
# Usage: irm https://raw.githubusercontent.com/iamrudyard/office_uninstall/main/xscript.ps1 | iex

try {
    $url = 'https://raw.githubusercontent.com/iamrudyard/office_uninstall/main/uninstall_office.cmd'

    Write-Host "[*] Downloading Office uninstaller script..." -ForegroundColor Cyan
    $response = Invoke-RestMethod -Uri $url -UseBasicParsing

    if ($response) {
        $tempFile = "$env:TEMP\office_uninstall.cmd"
        Set-Content -Path $tempFile -Value $response -Encoding ASCII

        Write-Host "[*] Running uninstaller as Administrator..." -ForegroundColor Yellow
        Start-Process cmd.exe -ArgumentList "/c", $tempFile -Verb RunAs

        Write-Host "[+] Script executed successfully." -ForegroundColor Green
    } else {
        Write-Host "[!] Failed to download script content." -ForegroundColor Red
    }
}
catch {
    Write-Host "[!] Error occurred: $_" -ForegroundColor Red
}
