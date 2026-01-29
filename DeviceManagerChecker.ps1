Write-Host "Running unknown device scan..."

$devices = Get-PnpDevice | Where-Object Problem -eq 28

if (-not $devices) {
    Write-Host "No unknown devices detected."
} else {
    $devices | Format-Table FriendlyName, Status, InstanceId -AutoSize
}
