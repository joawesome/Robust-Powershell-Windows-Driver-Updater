Write-Host "Running device health scan..."

$devices = Get-PnpDevice | Where-Object {
    $_.Status -ne "OK" -or          # Not working properly
    $_.Problem -ne 0 -or            # Any Device Manager error code
    $_.FriendlyName -eq "Unknown device"
}

if (-not $devices) {
    Write-Host "No problematic devices detected."
}
else {
    Write-Host "`nDevices requiring attention:`n"

    $devices |
        Sort-Object Problem, Class, FriendlyName |
        Select-Object `
            FriendlyName,
            Class,
            Status,
            Problem,
            @{Name="Meaning";Expression={
                switch ($_.Problem) {
                    0  { "OK" }
                    10 { "Device cannot start" }
                    14 { "Reboot required" }
                    18 { "Driver must be reinstalled" }
                    19 { "Registry issue" }
                    21 { "Device disabled" }
                    22 { "Disabled by user" }
                    24 { "Device not present" }
                    28 { "No driver installed" }
                    29 { "Disabled by firmware" }
                    31 { "Driver failed to load" }
                    default { "Other problem" }
                }
            }} |
        Format-Table -AutoSize
}
