# ========================================================
# run_remote_uninstaller.ps1
# - Downloads a remote .cmd/.bat from multiple mirrors (raw URLs)
# - Optional SHA256 verification (set $releaseHash)
# - Writes temp file, runs via cmd.exe elevated (Start-Process -Verb RunAs)
# - Waits for completion and removes temp files
# - Designed to be used with: irm <raw-ps1-url> | iex
# ========================================================

param(
    [string[]] $ExtraArgs  # any args you want passed to the .cmd
)

& {
    $psv = (Get-Host).Version.Major
    $troubleshoot = 'https://example.com/troubleshoot'  # change as needed

    function Fail([string]$msg) {
        Write-Host $msg -ForegroundColor Red
        return
    }

    # URLs (replace with your actual raw URLs - GitHub raw links or other hosts)
    $URLs = @(
        'https://raw.githubusercontent.com/iamrudyard/office_uninstall/main/uninstall_office.cmd',
        # add mirrors if you have them:
        # 'https://raw.fastgit.org/<yourusername>/office-uninstall-tool/main/Uninstall_Office_and_OneNote_AutoRestart.cmd',
        # 'https://your.other.host/path/Uninstall_Office_and_OneNote_AutoRestart.cmd'
    )

    # Optional: expected SHA256 (UPPERCASE, no dashes). Set to $null to skip verification.
    $releaseHash = $null
    # Example: 'D60752A27BDED6887C5CEC88503F0F975ACB5BC849673693CA7BA7C95BCB3EF4'

    # Basic environment checks
    if ($ExecutionContext.SessionState.LanguageMode.value__ -ne 0) {
        Fail "PowerShell is not running in Full Language Mode. Aborting."
        return
    }

    try {
        [void][System.AppDomain]::CurrentDomain.GetAssemblies()
    } catch {
        Fail "PowerShell failed to load required .NET assemblies. Aborting."
        return
    }

    # Download loop
    Write-Progress -Activity "Downloading uninstaller" -Status "Starting..."
    $response = $null
    $errors = @()
    foreach ($url in $URLs | Sort-Object { Get-Random }) {
        try {
            Write-Host "Trying: $url"
            if ($psv -ge 3) {
                $response = Invoke-RestMethod -Uri $url -UseBasicParsing -ErrorAction Stop
            } else {
                $wc = New-Object System.Net.WebClient
                $response = $wc.DownloadString($url)
            }
            if ($response) { break }
        } catch {
            $errors += $_
            Write-Host "Failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    Write-Progress -Activity "Downloading uninstaller" -Status "Done" -Completed

    if (-not $response) {
        foreach ($e in $errors) { Write-Host "Error: $($e.Exception.Message)" -ForegroundColor Red }
        Fail "Failed to retrieve uninstaller from all mirrors. Aborting."
        return
    }

    # Optional hash verification
    if ($releaseHash) {
        $ms = New-Object IO.MemoryStream
        $sw = New-Object IO.StreamWriter($ms)
        $sw.Write($response)
        $sw.Flush()
        $ms.Position = 0
        $computed = [BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash($ms)) -replace '-'
        if ($computed -ne $releaseHash) {
            Fail "SHA256 mismatch. Computed: $computed Expected: $releaseHash. Aborting."
            return
        } else {
            Write-Host "Hash OK." -ForegroundColor Green
        }
    }

    # Prepare unique temp file location (use system temp when elevated, otherwise user temp)
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($isAdmin) {
        $tempFolder = Join-Path $env:SystemRoot "Temp"
    } else {
        $tempFolder = Join-Path $env:LOCALAPPDATA "Temp"
    }
    if (-not (Test-Path $tempFolder)) { New-Item -Path $tempFolder -ItemType Directory -Force | Out-Null }
    $rand = [Guid]::NewGuid().Guid
    $targetName = "UNINSTALL_$rand.cmd"
    $FilePath = Join-Path $tempFolder $targetName

    # Write content (ensure proper CRLF and ASCII)
    $finalContent = $response -replace "`n", "`r`n"
    Set-Content -Path $FilePath -Value $finalContent -Encoding ASCII -Force

    if (-not (Test-Path $FilePath)) {
        Fail "Failed to create temp file at $FilePath"
        return
    }

    Write-Host "Saved uninstaller to $FilePath"

    # Test cmd.exe availability and point to 64-bit cmd
    $env:ComSpec = Join-Path $env:SystemRoot "system32\cmd.exe"
    try {
        $chk = & $env:ComSpec /c "echo CMD is working" 2>$null
        if ($chk -notlike '*CMD is working*') {
            Write-Warning "cmd.exe did not respond as expected."
        }
    } catch {
        Write-Warning "cmd.exe test failed: $($_.Exception.Message)"
    }

    # Build argument list for Start-Process
    $argList = @("/c", "`"$FilePath`"")
    if ($ExtraArgs) { $argList += $ExtraArgs }

    Write-Host "Launching uninstaller elevated..."
    try {
        Start-Process -FilePath $env:ComSpec -ArgumentList $argList -Verb RunAs -Wait
    } catch {
        Write-Warning "Failed to start elevated process: $($_.Exception.Message)"
        Write-Host "Attempting non-elevated in-memory execution as fallback..."
        try {
            $response | cmd.exe
        } catch {
            Fail "Fallback execution failed: $($_.Exception.Message)"
        }
    }

    # Cleanup temp files created by this pattern
    try {
        Get-ChildItem -Path (Join-Path $tempFolder "UNINSTALL_*.cmd") -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "UNINSTALL_*" } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    } catch {}

    Write-Host "Bootstrapper finished." -ForegroundColor Green
}
