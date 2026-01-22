# Get current user's Startup folder
$startupPath = [Environment]::GetFolderPath("Startup")

# Path to the CMD file
$cmdFilePath = Join-Path $startupPath "DriverUpdaterStartup.cmd"

# Command to run
$psCommand = 'powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/joawesome/Robust-Powershell-Windows-Driver-Updater/main/DriverUpdateMain.ps1 | iex"'

# Create the CMD file
Set-Content -Path $cmdFilePath -Value $psCommand -Encoding ASCII

Write-Output "Startup CMD file created at: $cmdFilePath"

# ---- RUN IT NOW ----
Invoke-Expression 'irm https://raw.githubusercontent.com/joawesome/Robust-Powershell-Windows-Driver-Updater/main/DriverUpdateMain.ps1 | iex'
