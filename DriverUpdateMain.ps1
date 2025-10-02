# Description: This script will install the PSWindowsUpdate module and check for driver updates.


# How many attempts we want to try and install the module.
$MaxAttempts   = 10
$AutoReboot    = $true   # set to $true for automatic reboot, $false for user prompt

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
        $WUmaxAttempts    = 5
        $WUattempt        = 0
        $WUsuccess        = $false
        $ConsecutiveFails = 0

        while (-not $WUsuccess -and $WUattempt -lt $WUmaxAttempts) {
            try {
                $WUattempt++
                Write-Host "Windows Update Attempt $WUattempt of $WUmaxAttempts..."

                # Capture the result instead of discarding it
                $results = Install-WindowsUpdate -AcceptAll -UpdateType Driver -IgnoreReboot -ErrorAction Stop

                # Look at the statuses
                $failedCount    = ($results | Where-Object { $_.Status -eq "Failed" }).Count
                $installedCount = ($results | Where-Object { $_.Status -eq "Installed" }).Count

                if ($failedCount -gt 0) {
                    Write-Warning "$failedCount update(s) failed in this attempt."
                    $ConsecutiveFails++
                } else {
                    $ConsecutiveFails = 0
                }

                if ($installedCount -gt 0 -and $failedCount -eq 0) {
                    Write-Host "Windows Update completed successfully."
                    $WUsuccess = $true
                } elseif ($ConsecutiveFails -ge 3) {
                    Write-Warning "3 consecutive failures detected. Retrying after delay..."
                    $ConsecutiveFails = 0
                    Start-Sleep -Seconds 30
                }
            }
            catch {
                Write-Warning "Windows Update threw an error: $($_.Exception.Message)"

                if ($_.Exception.HResult -eq -2145124329) {  # 0x80248007
                    Write-Warning "Encountered 0x80248007 (cache issue). Retrying..."
                } elseif ($_.Exception.Message -like "*value does not fall within the expected range*") {
                    Write-Warning "Driver installation may still be in progress. Waiting 2 minutes..."
                    Start-Sleep -Seconds 120
                } else {
                    Write-Warning "Unexpected error. Retrying..."
                }

                Start-Sleep -Seconds 15
            }
        }

        if (-not $WUsuccess) {
            Write-Error "Driver update failed after $WUmaxAttempts attempts."
        }

        # --- Reboot handling ---
        if (Get-WURebootStatus) {
            if ($AutoReboot -eq $true) {
                Write-Host "Reboot is required. Restarting automatically..."
                Restart-Computer -Force
            } else {
                $response = Read-Host "Reboot is required. Do you want to restart now? (y/n)"
                if ($response -match '^[Yy]$') {
                    Write-Host "Rebooting system..."
                    Restart-Computer -Force
                } else {
                    Write-Host "Reboot skipped by user."
                }
            }
        }

        break
    }
}
