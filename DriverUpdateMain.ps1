# Description: Installs PSWindowsUpdate (if needed), runs driver updates, detects
# "Reboot is required, but do it manually." in the module output, and reboots automatically.
Add-Type -AssemblyName System.Windows.Forms

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(
        IntPtr hWnd, IntPtr hWndInsertAfter,
        int X, int Y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

$HWND_TOPMOST = [IntPtr]::new(-1)
$SWP_SHOWWINDOW = 0x0040
$SW_RESTORE = 9   # un-maximize first, or SetWindowPos will be ignored

# Get the primary screen's working area (excludes taskbar)
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$halfWidth = [int]($screen.Width / 2)

# --- Position THIS console window on the LEFT half, always on top ---
do {
    Start-Sleep -Milliseconds 100
    $hwnd = (Get-Process -Id $PID).MainWindowHandle
} while ($hwnd -eq 0)

[Win32]::ShowWindow($hwnd, $SW_RESTORE) | Out-Null
[Win32]::SetWindowPos(
    $hwnd, $HWND_TOPMOST,
    $screen.X, $screen.Y, $halfWidth, $screen.Height,
    $SWP_SHOWWINDOW
)

# --- Launch Device Manager and position it on the RIGHT half ---
$devMgr = Start-Process -FilePath "devmgmt.msc" -PassThru

# devmgmt.msc runs inside mmc.exe — wait for its window handle
$mmcHandle = [IntPtr]::Zero
$deadline = (Get-Date).AddSeconds(15)
while ($mmcHandle -eq [IntPtr]::Zero -and (Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 200
    $mmc = Get-Process -Name mmc -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne 0 } |
        Select-Object -First 1
    if ($mmc) { $mmcHandle = $mmc.MainWindowHandle }
}

if ($mmcHandle -ne [IntPtr]::Zero) {
    [Win32]::ShowWindow($mmcHandle, $SW_RESTORE) | Out-Null
    [Win32]::SetWindowPos(
        $mmcHandle, [IntPtr]::Zero,
        $screen.X + $halfWidth, $screen.Y, $screen.Width - $halfWidth, $screen.Height,
        $SWP_SHOWWINDOW
    )
} else {
    Write-Warning "Could not find Device Manager window to position."
}


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
                   # Start-Sleep -Seconds 5
                    
                   # & "$env:WINDIR\System32\Sysprep\sysprep.exe" /oobe /generalize /shutdown
                   # Write-Host "Now shutting down into Out-Of-Box-Experience"
                   # Start-Sleep -Seconds 5

                    if (-not $DebugMode) {
                        # remove transcript if not debugging
                        try { Remove-Item -Path $transcriptPath -ErrorAction SilentlyContinue } catch {}
                    }
                }

                $WUsuccess = $true
            } catch {
                Write-Warning "Windows Update attempt failed!! Trying again in 10 Seconds. Failed with: $($_.Exception.Message)"
                Start-Sleep -Seconds 10
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

Start-Sleep -Seconds 5
