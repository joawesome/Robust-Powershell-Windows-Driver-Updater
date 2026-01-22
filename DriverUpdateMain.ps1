# Description: This script will install the PSWindowsUpdate module and check for driver updates.
# It will run in a separate window and check for updates. If there are no updates, it will exit.
# You can run this script by itself:
#   irm https://raw.githubusercontent.com/joawesome/Robust-Powershell-Windows-Driver-Updater/main/DriverUpdateMain.ps1 | iex

# ---------------------------
# Relaunch elevated (single UAC) if needed
# ---------------------------
if (-not (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    # Determine script path (works when saved to disk; with iex fallback to MyInvocation)
    $scriptPath = $PSCommandPath
    if ([string]::IsNullOrEmpty($scriptPath)) { $scriptPath = $MyInvocation.MyCommand.Definition }

    # Build argument list to re-run the same script file
    $argList = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    try {
        # Start elevated copy and exit current (this triggers a single UAC prompt)
        Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -Verb RunAs
        exit
    } catch {
        Write-Error "Failed to relaunch elevated: $($_.Exception.Message)"
        exit 1
    }
}

# ---------------------------
# Minimize interactive prompts
# ---------------------------
# Temporarily suppress confirmation prompts (restore later if needed)
$oldConfirmPreference = $ConfirmPreference
$ConfirmPreference = 'None'

# ---------------------------
# Settings
# ---------------------------
$MaxAttempts = 5

# Phrase to detect in module output (case-insensitive). Adjust if your module prints a different phrase.
$RebootPhraseRegex = '(?i)\breboot\b.*\brequired\b'

# ---------------------------
# Helper: capture restart from module output by transcript
# ---------------------------
function Invoke-InstallWindowsUpdate-WithTranscript {
    param(
        [string]$TranscriptPath,
        [int]$DelayBeforeRestartSeconds = 5
    )

    # Start transcript to capture Write-Host / host output
    Start-Transcript -Path $TranscriptPath -Force | Out-Null

    try {
        # Run driver updates; keep -IgnoreReboot so we control restart from script
        Install-WindowsUpdate -AcceptAll -UpdateType Driver -IgnoreReboot -ErrorAction Stop
    } finally {
        # Always stop transcript
        Stop-Transcript | Out-Null
    }

    # Read transcript and search for the reboot phrase
    $foundReboot = $false
    try {
        $text = Get-Content -Path $TranscriptPath -Raw -ErrorAction SilentlyContinue
        if ($text -and ($text -match $RebootPhraseRegex)) {
            $foundReboot = $true
        }
    } catch {
        Write-Warning "Failed to read transcript: $($_.Exception.Message)"
    }

    if ($foundReboot) {
        Write-Host "Detected reboot message in module output."
        if ($DelayBeforeRestartSeconds -gt 0) {
            Write-Host "Restarting in $DelayBeforeRestartSeconds seconds..."
            Start-Sleep -Seconds $DelayBeforeRestartSeconds
        }

        try {
            Restart-Computer -Force -Confirm:$false -ErrorAction Stop
        } catch {
            Write-Error "Automatic reboot failed: $($_.Exception.Message)"
            # As a fallback, try to spawn an elevated one-liner restart (should not be necessary since we relaunched elevated)
            try {
                $restartArgs = '-NoProfile -WindowStyle Hidden -Command "Restart-Computer -Force -Confirm:$false"'
                Start-Process -FilePath 'powershell.exe' -ArgumentList $restartArgs -Verb RunAs -ErrorAction Stop
            } catch {
                Write-Error "Fallback elevated restart also failed: $($_.Exception.Message)"
            }
        }
    } else {
        Write-Host "No reboot message found in module output."
    }

    # Clean up transcript file
    try { Remove-Item -Path $TranscriptPath -ErrorAction SilentlyContinue } catch {}
}

# ---------------------------
# Main loop (module install + update attempts)
# ---------------------------
for ($i = 1; $i -le $MaxAttempts; $i++) {
    # Check if PSWindowsUpdate is installed/available
    if (-not (Get-Module -Name PSWindowsUpdate -ListAvailable)) {
        Write-Progress -Activity "Preparing PSWindowsUpdate Module" -Status "Attempt $i of $MaxAttempts"
        Write-Host 'Getting Package Provider'
        try { Get-PackageProvider -Name Nuget -ForceBootstrap -ErrorAction Stop | Out-Null } catch { Write-Warning "NuGet bootstrap failed: $($_.Exception.Message)" }

        Write-Host 'Setting Repository'
        try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop } catch { Write-Warning "Set-PSRepository warning: $($_.Exception.Message)" }

        Write-Host 'Installing PSWindowsUpdate Module...'
        try {
            Install-Module -Name PSWindowsUpdate -Force -AllowClobber -ErrorAction Stop
            Start-Sleep -Seconds 1

            Write-Host 'Importing module...'
            Import-Module -Name PSWindowsUpdate -Force -ErrorAction Stop
        } catch {
            Write-Warning "Module install/import failed on attempt $i: $($_.Exception.Message)"
            Start-Sleep -Seconds (5 * $i)
            continue
        } finally {
            Write-Progress -Activity "Preparing PSWindowsUpdate Module" -Completed
        }
    }

    # If module is available, proceed to updates
    if (Get-Module -Name PSWindowsUpdate -ListAvailable) {
        Write-Host "Checking for updates..."

        $WUmaxAttempts = 10
        $WUattempt = 0
        $WUsuccess = $false

        while (-not $WUsuccess -and $WUattempt -lt $WUmaxAttempts) {
            try {
                $WUattempt++
                Write-Host "Windows Update Attempt $WUattempt of $WUmaxAttempts..."

                # Use a transcript to capture host-bound output (Write-Host) from the module
                $transcriptPath = Join-Path $env:TEMP ("PSWU_Transcript_{0}.txt" -f ([guid]::NewGuid()))
                Invoke-InstallWindowsUpdate-WithTranscript -TranscriptPath $transcriptPath -DelayBeforeRestartSeconds 10

                Write-Host "Windows Update attempt completed (module run finished)."
                # If the system rebooted, the script will terminate on restart; if not, we set success to true here
                $WUsuccess = $true
            }
            catch {
                Write-Warning "Windows Update failed with error: $($_.Exception.Message)"

                if ($_.Exception.HResult -eq -2145124329) {  # 0x80248007
                    Write-Warning "Encountered 0x80248007 (Windows Update cache issue). Retrying..."
                } else {
                    Write-Warning "Unexpected error. Retrying..."
                }

                Start-Sleep -Seconds 15
            }
        }

        if (-not $WUsuccess) {
            Write-Error "Driver update failed after $WUmaxAttempts attempts."
        }

        break
    } else {
        Write-Warning "PSWindowsUpdate module still not available after attempt $i."
    }
}

# Restore confirm preference
$ConfirmPreference = $oldConfirmPreference

Write-Host "Script finished."
