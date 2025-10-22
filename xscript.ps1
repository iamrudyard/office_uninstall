# ========================================================
#  Remote .cmd/.bat bootstrapper (massgrave-style)
#  - Downloads the target CMD/BAT from multiple mirrors
#  - Verifies SHA256 if provided
#  - Writes to a temp file
#  - Runs via cmd.exe elevated (Start-Process -Verb RunAs)
#  - Cleans up temp files afterwards
# ========================================================

param(
    [Parameter(Mandatory=$false)][string[]] $ArgsToPass
)

& {
    $psv = (Get-Host).Version.Major
    $troubleshoot = 'https://example.com/troubleshoot'  # replace as needed

    # Quick prechecks
    if ($ExecutionContext.SessionState.LanguageMode.value__ -ne 0) {
        Write-Host "PowerShell is not running in Full Language Mode." -ForegroundColor Yellow
        Write-Host "If you need help, check your PowerShell execution policy or run in full language mode." -ForegroundColor Cyan
        return
    }

    try {
        # Basic .NET sanity check
        [void][System.AppDomain]::CurrentDomain.GetAssemblies()
        [void][System.Math]::Sqrt(144)
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Powershell failed to load .NET commands. See $troubleshoot" -ForegroundColor Cyan
        return
    }

    function Check3rdAV {
        try {
            $cmd = if ($psv -ge 3) { 'Get-CimInstance' } else { 'Get-WmiObject' }
            $avList = & $cmd -Namespace root\SecurityCenter2 -Class AntiVirusProduct -ErrorAction SilentlyContinue |
                        Where-Object { $_.displayName -and ($_.displayName -notlike '*windows*') } |
                        Select-Object -ExpandProperty displayName -ErrorAction SilentlyContinue
            if ($avList) {
                Write-Host '3rd party Antivirus might be blocking the script - ' -NoNewline -ForegroundColor White -BackgroundColor Blue
                Write-Host " $($avList -join ', ')" -ForegroundColor DarkRed -BackgroundColor White
            }
        } catch {}
    }

    function CheckFile {
        param ([string]$FilePath)
        if (-not (Test-Path $FilePath)) {
            Check3rdAV
            Write-Host "Failed to create target file in temp folder, aborting!" -ForegroundColor Red
            Write-Host "Help - $troubleshoot" -ForegroundColor Cyan
            throw "File creation failed"
        }
    }

    # OPTIONAL: Put your target raw URLs here (in order of preference)
    $URLs = @(
        # Example GitHub raw (replace with your raw URL)
        'https://raw.githubusercontent.com/<yourusername>/office-uninstall-tool/main/Uninstall_Office_and_OneNote_AutoRestart.cmd',

        # Alternative mirror(s) - replace or remove as needed
        'https://raw.fastgit.org/<yourusername>/office-uninstall-tool/main/Uninstall_Office_and_OneNote_AutoRestart.cmd',
        'https://your.cdn.example/path/Uninstall_Office_and_OneNote_AutoRestart.cmd'
    )

    # OPTIONAL: Put the expected SHA256 hash here (UPPERCASE, no dashes) to verify integrity.
    # If you don't want hash checking, set $releaseHash = $null
    $releaseHash = $null
    # Example: 'D60752A27BDED6887C5CEC88503F0F975ACB5BC849673693CA7BA7C95BCB3EF4'

    Write-Progress -Activity "Downloading uninstaller" -Status "Please wait..." -PercentComplete 0
    $response = $null
    $errors = @()
    foreach ($URL in $URLs | Sort-Object { Get-Random }) {
        try {
            if ($psv -ge 3) {
                # Use Invoke-RestMethod to fetch text
                $response = Invoke-RestMethod -Uri $URL -UseBasicParsing -ErrorAction Stop
            }
            else {
                $wc = New-Object System.Net.WebClient
                $response = $wc.DownloadString($URL)
            }
            Write-Progress -Activity "Downloading uninstaller" -Status "Fetched from $URL" -PercentComplete 50
            break
        } catch {
            $errors += $_
        }
    }
    Write-Progress -Activity "Downloading uninstaller" -Status "Done" -PercentComplete 100
    if (-not $response) {
        Check3rdAV
        foreach ($err in $errors) {
            Write-Host "Error: $($err.Exception.Message)" -ForegroundColor Red
        }
        Write-Host "Failed to retrieve the uninstaller from any mirror. Aborting." -ForegroundColor Red
        Write-Host "Help - $troubleshoot" -ForegroundColor Cyan
        return
    }

    # If a hash is provided, verify SHA256
    if ($releaseHash) {
        try {
            $ms = New-Object IO.MemoryStream
            $sw = New-Object IO.StreamWriter($ms)
            $sw.Write($response)
            $sw.Flush()
            $ms.Position = 0
            $computed = [BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash($ms)) -replace '-'
            if ($computed -ne $releaseHash) {
                Write-Warning "Hash ($computed) does not match expected ($releaseHash). Aborting."
                return
            } else {
                Write-Host "Hash verified." -ForegroundColor Green
            }
        } catch {
            Write-Warning "Hash verification failed: $($_.Exception.Message)"
            return
        }
    }

    # Check for Autorun which could break CMD behavior
    $paths = "HKCU:\SOFTWARE\Microsoft\Command Processor", "HKLM:\SOFTWARE\Microsoft\Command Processor"
    foreach ($path in $paths) {
        if (Get-ItemProperty -Path $path -Name "Autorun" -ErrorAction SilentlyContinue) {
            Write-Warning "Autorun registry found at $path. This may make CMD behave unexpectedly. Consider removing the Autorun value."
        }
    }

    # Prepare a unique temp filename (admin vs user temp)
    $rand = [Guid]::NewGuid().Guid
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $FilePath = if ($isAdmin) { Join-Path $env:SystemRoot "Temp" "UNINSTALL_$rand.cmd" } else { Join-Path $env:LOCALAPPDATA "Temp" "UNINSTALL_$rand.cmd" }

    # Prepend a small marker (optional) to ensure file starts correctly
    $finalContent = "@echo off`r`n" + $response
    Set-Content -Path $FilePath -Value $finalContent -Encoding ASCII -Force
    CheckFile $FilePath

    # Ensure cmd works and point ComSpec to 64-bit cmd if available
    $env:ComSpec = Join-Path $env:SystemRoot "system32\cmd.exe"
    try {
        $chkcmd = & $env:ComSpec /c "echo CMD is working" 2>$null
        if ($chkcmd -notlike "*CMD is working*") {
            Write-Warning "cmd.exe did not respond correctly."
        }
    } catch {
        Write-Warning "cmd.exe test failed: $($_.Exception.Message)"
    }

    # Build argument list to pass through any args the user provided to the bootstrapper
    $argList = @("/c", "`"$FilePath`"")
    if ($ArgsToPass) { $argList += $ArgsToPass }

    Write-Host "Launching uninstaller via cmd.exe (elevated)..." -ForegroundColor Cyan

    try {
        # Use Start-Process to elevate and wait for exit (works like saps pattern)
        $psi = @{
            FilePath = $env:ComSpec
            ArgumentList = $argList
            Verb = "RunAs"
            Wait = $true
        }
        Start-Process @psi
    } catch {
        Write-Host "Failed to start cmd.exe elevated: $($_.Exception.Message)" -ForegroundColor Red
        # Try a non-elevated fallback (pipe)
        Write-Host "Attempting non-elevated in-memory execution as fallback..." -ForegroundColor Yellow
        try {
            # Pipe content to cmd.exe (non-elevated)
            $response | cmd.exe
        } catch {
            Write-Host "Fallback execution failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Clean up temp files matching our pattern
    try {
        $FilePatterns = @(Join-Path $env:SystemRoot "Temp\UNINSTALL_*.cmd", Join-Path $env:LOCALAPPDATA "Temp\UNINSTALL_*.cmd")
        foreach ($pattern in $FilePatterns) {
            Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        }
    } catch {}

    Write-Host "Done." -ForegroundColor Green
}
