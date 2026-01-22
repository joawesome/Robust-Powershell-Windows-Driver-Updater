# Description: Installs PSWindowsUpdate (if needed), runs driver updates, detects
# "Reboot is required, but do it manually." in the module output, and reboots automatically.
param(
    [switch]$DebugMode
)

# Relaunch elevated if needed (single UAC prompt)
if (-not (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    $scriptPath = $PSCommandPath
    if ([string]::IsNullOrEmpty($scriptPath)) { $scriptPath = $MyInvocation.MyCommand.Definition }
    if ([string]::IsNullOrEmpty($scriptPath)) {
        Write-Error "Cannot determine script path to relaunch elevated. Save the script to disk and run it as Administrator."
        exit 1
    }
    $argList = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    try {
        Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -Verb RunAs
        exit
    } catch {
        Write-Error "Failed to relaunch elevated: $($_.Exception.Message)"
        exit 1
    }
}

# Make confirmations non-interactive
$oldConfirmPreference = $ConfirmPreference
$ConfirmPreference = 'None'

# Settings
$MaxAttempts = 5
$WUmaxAttempts = 10

# Exact phrase (flexible) we expect from PSWindowsUpdate
$RebootPhraseRegex = '(?i)\breboot\s+is\s+required[,;:]?\s*but\s+do\s+it\s+manually\.?'

function Run-InstallWindowsUpdate-AndReturnRebootFlag {
    param(
        [Parameter(Mandatory=$true)][string]$TranscriptPath,
        [int]$DelaySeconds = 0
    )

    # Start transcript to capture Write-Host etc.
    Start-Transcript -Path $TranscriptPath -Force | Out-Null
    try {
        # Run update (we keep -IgnoreReboot to control reboot here)
        Install-WindowsUpdate -AcceptAll -UpdateType Driver -IgnoreReboot -ErrorAction Stop
    } finally {
        Stop-Transcript | Out-Null
    }

    # Read transcript and check for phrase
    $found = $false
    try {
        $text = Get-Content -Path $TranscriptPath -Raw -ErrorAction SilentlyContinue
        if ($null -ne $text -and $text -match $RebootPhraseRegex) {
            $found = $true
        }
    } catch {
        Write-Warning "Error reading transcript: $($_.Exception.Message)"
    }

    if ($DelaySeconds -gt 0) { Start-Sleep -Seconds $DelaySeconds }

    return $found
}

# Helper to attempt restart and log any error
function Try-RestartComputer {
    param([int]$DelaySeconds = 0)
    if ($DelaySeconds -gt 0) {
        Write-Host "Restarting in $DelaySeconds seconds..."
        Start-Sleep -Seconds $DelaySeconds
    }

    try {
        Restart-Computer -Force -Confirm:$false -ErrorAction Stop
        # If restart succeeds, script will be terminated by OS
    } catch {
        Write-Error "Restart-Computer failed: $($_.Exception.Message)"
        if ($_.Exception.HResult) {
            Write-Error ("HResult: 0x{0:X8}" -f $_.Exception.HResult)
        }
        return $false
    }
    return $true
}

# --- Main flow ---
for ($i = 1; $i -le $MaxAttempts; $i++) {

    if (-not (Get-Module -Name PSWindowsUpdate -ListAvailable)) {
        Write-Host "Preparing PSWindowsUpdate (attempt ${i} of $MaxAttempts)..."
        try { Get-PackageProvider -Name NuGet -ForceBootstrap -ErrorAction Stop | Out-Null } catch { Write-Warning "NuGet bootstrap failed: $($_.Exception.Message)" }
        try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop } catch { Write-Warning "Set-PSRepository: $($_.Exception.Message)" }

        try {
            Install-Module -Name PSWindowsUpdate -Force -AllowClobber -ErrorAction Stop
            Import-Module -Name PSWindowsUpdate -Force -ErrorAction Stop
        } catch {
            Write-Warning "Module install/import failed on attempt ${i}: $($_.Exception.Message)"
            Start-Sleep -Seconds (5 * $i)
            continue
        }
    }

    if (Get-Module -Name PSWindowsUpdate -ListAvailable) {
        Write-Host "Running driver updates..."
        $WUattempt = 0
        $WUsuccess = $false

        while (-not $WUsuccess -and $WUattempt -lt $WUmaxAttempts) {
            $WUattempt++
            Write-Host "Windows Update Attempt $WUattempt of $WUmaxAttempts..."

            # Create transcript path
            $transcriptPath = Join-Path $env:TEMP ("PSWU_Transcript_{0}.txt" -f ([guid]::NewGuid()))

            try {
                $needsReboot = Run-InstallWindowsUpdate-AndReturnRebootFlag -TranscriptPath $transcriptPath -DelaySeconds 0

                if ($DebugMode) {
                    Write-Host "Transcript file preserved at: $transcriptPath"
                    # Print the last ~2000 chars to inspect output
                    try {
                        $full = Get-Content -Path $transcriptPath -Raw -ErrorAction SilentlyContinue
                        if ($full) {
                            $tail = if ($full.Length -gt 2000) { $full.Substring($full.Length - 2000) } else { $full }
                            Write-Host "---- Transcript tail ----"
                            Write-Host $tail
                            Write-Host "---- End transcript tail ----"
                        } else {
                            Write-Host "Transcript empty or unreadable."
                        }
                    } catch { Write-Warning "Could not read transcript for debug output: $($_.Exception.Message)" }
                }

                if ($needsReboot) {
                    Write-Host "Reboot was requested"
                    # Attempt restart — script is elevated because we relaunched earlier
                    $ok = Try-RestartComputer -DelaySeconds 10
                    if ($ok) { 
                        # If restart succeeded, the computer will reboot and script won't continue.
                        return
                    } else {
                        Write-Warning "Restart attempt failed. Leaving transcript for debugging: $transcriptPath"
                        # keep the transcript for inspection if debug mode off
                    }
                } else {
                    Write-Host "Updates were successful. No reboots were required."
                    
                    Remove-Item "$([Environment]::GetFolderPath('Startup'))\DriverUpdaterStartup.cmd" -Force -ErrorAction SilentlyContinue

                    if (-not $DebugMode) {
                        # remove transcript if not debugging
                        try { Remove-Item -Path $transcriptPath -ErrorAction SilentlyContinue } catch {}
                    }
                }

                $WUsuccess = $true
            } catch {
                Write-Warning "Windows Update attempt failed: $($_.Exception.Message)"
                Start-Sleep -Seconds 15
            }
        }

        if (-not $WUsuccess) {
            Write-Error "Driver update failed after $WUmaxAttempts attempts."
        }

        break
    } else {
        Write-Warning "PSWindowsUpdate not available after attempt ${i}."
    }
}

# Restore confirm preference
$ConfirmPreference = $oldConfirmPreference

Write-Host "Operation completed... exiting"
