# Description: This script will install the PSWindowsUpdate module, check for driver updates, 
# and automatically reboot if required.
# It will run in a separate window. If there are no updates, it will exit.
# Run this script directly:
# irm https://raw.githubusercontent.com/joawesome/Robust-Powershell-Windows-Driver-Updater/main/DriverUpdateMain.ps1 | iex

# Maximum attempts for module installation
$MaxAttempts = 5

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

                # Install driver updates
                Install-WindowsUpdate -AcceptAll -UpdateType Driver -ErrorAction Stop

                Write-Host "Windows Update completed successfully."

                # Reboot automatically if required
                if (Get-WURebootStatus) {
                    Write-Host "Reboot is required. Rebooting now..."
                    shutdown /r /t 5 /f
                }

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
