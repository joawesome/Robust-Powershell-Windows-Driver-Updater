# Get current user's Startup folder
$startupPath = [Environment]::GetFolderPath("Startup")

# Path to the CMD file
$cmdFilePath = Join-Path $startupPath "DriverUpdaterStartup.cmd"

# Command to run
$psCommand = 'start powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/joawesome/Robust-Powershell-Windows-Driver-Updater/main/DriverUpdateMain.ps1 | iex"'

# Create the CMD file
Set-Content -Path $cmdFilePath -Value $psCommand -Encoding ASCII

Write-Output "Starting Driver updator.. Startup file created at: $cmdFilePath"
Start-Sleep -Seconds 2
Write-Output "Startup file will be removed automatically once updates are successful"
Start-Sleep -Seconds 5

# ---- RUN IT NOW ----
Start-Process powershell.exe -ArgumentList '-WindowStyle Normal -ExecutionPolicy Bypass -Command "irm ''https://raw.githubusercontent.com/joawesome/Robust-Powershell-Windows-Driver-Updater/main/DriverUpdateMain.ps1'' | iex"'
