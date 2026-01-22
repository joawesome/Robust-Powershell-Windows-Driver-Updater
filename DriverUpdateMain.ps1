# Description: This script will install the PSWindowsUpdate module and check for driver updates.
# It will run in a separate window and check for updates. If there are no updates, it will exit.
# You can just run this script by itself by running the command:
# irm  https://raw.githubusercontent.com/joawesome/Robust-Powershell-Windows-Driver-Updater/main/DriverUpdateMain.ps1 | iex

# How many attempts we want to try and install the module.
$MaxAttempts = 5

# Function: Test if a reboot is pending (checks common indicators in both registry views)
function Test-PendingReboot {
    Add-Type -AssemblyName Microsoft.Win32.Registry

    # Helper to check key existence in a specific view
    $checkKey = {
        param($hive, $subKey, $view)
        try {
            $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
            $k = $base.OpenSubKey($subKey)
            if ($k) { return $true }
        } catch { }
        return $false
    }

    # Helper to read a value safely
    $readValue = {
        param($hive, $subKey, $valueName, $view)
        try {
            $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
            $k = $base.OpenSubKey($subKey)
            if ($k) {
                return $k.GetValue($valueName, $null)
            }
        } catch { }
        return $null
    }

    $views = @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)
    foreach ($view in $views) {
        # Component Based Servicing RebootPending
        if (&$checkKey([Microsoft.Win32.RegistryHive]::LocalMachine, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending', $view)) {
            return $true
        }

        # Windows Update RebootRequired (this is the key you mentioned)
        if (&$checkKey([Microsoft.Win32.RegistryHive]::LocalMachine, 'SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\RebootRequired', $view)) {
            return $true
        }

        # PendingFileRenameOperations
        $pfro = &$readValue([Microsoft.Win32.RegistryHive]::LocalMachine, 'SYSTEM\CurrentControlSet\Control\Session Manager', 'PendingFileRenameOperations', $view)
        if ($pfro -and ($pfro -is [string[]] -or $pfro.Length -gt 0)) {
            return $true
        }

        # Pending computer rename (ActiveComputerName vs ComputerName)
        try {
            $activeName = &$readValue([Microsoft.Win32.RegistryHive]::LocalMachine, 'SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName', 'ComputerName', $view)
            $pendingName = &$readValue([Microsoft.Win32.RegistryHive]::LocalMachine, 'SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName', 'ComputerName', $view)
            if ($activeName -and $pendingName -and ($activeName -ne $pendingName)) {
                return $true
            }
        } catch { }
    }

    return $false
}

for ($i = 1; $i -le $MaxAttempts; $i++) {
    if (-Not (Get-Module -Name PSWindowsUpdate)) {
        Write-Progress -Activity "Preparing PSWindowsUpdate Module" -Status "Attempt $i of $MaxAttempts"
        Write-Host 'Getting Package Provider'
        Get-PackageProvider -Name Nuget -ForceBootstrap | Out-Null

        Write-Host 'Setting Repository'
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

        Write-Host 'Installing PSWindowsUpdate Module...'
        Install-Module -Name PSWindowsUpdate -Force
        Start-Sleep -Seconds 1

        Write-Host 'Importing module...'
        Import-Module -Name PSWindowsUpdate -Force
        Write-Progress -Activity "Preparing PSWindowsUpdate Module" -Completed
    }
    else {
        Write-Host "Checking for updates..."

        # Retry loop for driver installation
        $WUmaxAttempts = 10
        $WUattempt = 0
        $WUsuccess = $false

        while (-not $WUsuccess -and $WUattempt -lt $WUmaxAttempts) {
            try {
                $WUattempt++
                Write-Host "Windows Update Attempt $WUattempt of $WUmaxAttempts..."

                Install-WindowsUpdate -AcceptAll -UpdateType Driver -IgnoreReboot -ErrorAction Stop

                Write-Host "Windows Update completed successfully."

                # Check for pending reboot and restart automatically if needed
                Write-Host "Checking for pending reboot..."
                if (Test-PendingReboot) {
                    Write-Host "Reboot is required. Attempting to restart the computer..."

                    # Detect elevation
                    $isElevated = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)

                    if ($isElevated) {
                        try {
                            Restart-Computer -Force -Confirm:$false -ErrorAction Stop
                        } catch {
                            Write-Error "Restart-Computer failed: $($_.Exception.Message)"
                        }
                    } else {
                        # Relaunch an elevated PowerShell to perform the restart (will prompt UAC)
                        try {
                            Start-Process -FilePath 'powershell.exe' -ArgumentList '-NoProfile', '-WindowStyle', 'Hidden', '-Command', 'Restart-Computer -Force -Confirm:$false' -Verb RunAs -ErrorAction Stop
                        } catch {
                            Write-Error "Failed to launch elevated restart: $($_.Exception.Message). Run the script as Administrator or remove -IgnoreReboot to let PSWindowsUpdate reboot."
                        }
                    }

                    # Give a moment for the restart process to start
                    Start-Sleep -Seconds 5
                } else {
                    Write-Host "No reboot required."
                }

                Write-Host "Windows Update completed successfully."
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
    }
}
