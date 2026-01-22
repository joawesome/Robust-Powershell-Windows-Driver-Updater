<#
.SYNOPSIS
  Install PSWindowsUpdate (if needed) and attempt driver updates with retries.

.DESCRIPTION
  - Ensures TLS 1.2, installs/loads PSWindowsUpdate if missing, then runs Install-WindowsUpdate
    filtering for driver updates. Retries are implemented for both module install and update steps.

.PARAMETER MaxModuleAttempts
  Number of times to try installing/loading the module.

.PARAMETER MaxWUAttempts
  Number of times to try Install-WindowsUpdate.

.PARAMETER CurrentUserInstall
  If set, uses -Scope CurrentUser for Install-Module to avoid requiring elevation.

.EXAMPLE
  .\DriverUpdateMain.ps1 -Verbose
#>

[CmdletBinding()]
param(
    [int]$MaxModuleAttempts = 5,
    [int]$MaxWUAttempts     = 10,
    [switch]$CurrentUserInstall
)

function Ensure-Tls12 {
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Verbose "Set TLS to Tls12"
    } catch {
        Write-Warning "Failed to set TLS 1.2: $($_.Exception.Message)"
    }
}

function Ensure-PSWindowsUpdate {
    param(
        [int]$Attempts = 5,
        [switch]$CurrentUser
    )

    $scopeArg = if ($CurrentUser) { "-Scope CurrentUser" } else { "" }

    for ($i = 1; $i -le $Attempts; $i++) {
        Write-Verbose "Module check attempt $i of $Attempts"

        # Check installed modules (not only loaded)
        $installed = Get-InstalledModule -Name PSWindowsUpdate -ErrorAction SilentlyContinue
        if ($installed) {
            Write-Verbose "PSWindowsUpdate is installed (version $($installed.Version)). Importing..."
            try {
                Import-Module -Name PSWindowsUpdate -Force -ErrorAction Stop
                return $true
            } catch {
                Write-Warning "Failed importing PSWindowsUpdate: $($_.Exception.Message)"
                Start-Sleep -Seconds (5 * $i)
                continue
            }
        }

        # Not installed: attempt to install
        try {
            Write-Verbose "Bootstrapping NuGet provider (if needed)"
            Get-PackageProvider -Name NuGet -ForceBootstrap -ErrorAction Stop | Out-Null

            # Ensure PSGallery is available and trusted
            try {
                $repo = Get-PSRepository -Name PSGallery -ErrorAction Stop
                if ($repo.InstallationPolicy -ne 'Trusted') {
                    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
                }
            } catch {
                Write-Verbose "PSGallery repository not registered; registering..."
                Register-PSRepository -Name PSGallery -SourceLocation 'https://www.powershellgallery.com/api/v2' -InstallationPolicy Trusted -ErrorAction Stop
            }

            Write-Verbose "Installing PSWindowsUpdate (attempt $i)"
            # Build common parameters: use -Scope CurrentUser optionally to avoid admin requirement
            if ($CurrentUser) {
                Install-Module -Name PSWindowsUpdate -Force -AllowClobber -Scope CurrentUser -Repository PSGallery -ErrorAction Stop
            } else {
                Install-Module -Name PSWindowsUpdate -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
            }

            Import-Module -Name PSWindowsUpdate -Force -ErrorAction Stop
            Write-Verbose "PSWindowsUpdate installed and imported"
            return $true
        }
        catch {
            Write-Warning "Install attempt $i failed: $($_.Exception.Message)"
            Start-Sleep -Seconds (5 * $i)  # simple backoff
        }
    }

    return $false
}

function Run-DriverUpdates {
    param(
        [int]$Attempts = 10
    )

    $attempt = 0
    while ($attempt -lt $Attempts) {
        $attempt++
        Write-Verbose "Windows Update attempt $attempt of $Attempts"

        try {
            # Use -AcceptAll and -IgnoreReboot to avoid prompts; -ErrorAction Stop to catch failures
            Install-WindowsUpdate -AcceptAll -UpdateType Driver -IgnoreReboot -ErrorAction Stop

            Write-Host "Windows Update completed successfully."
            return $true
        } catch {
            $ex = $_.Exception
            Write-Warning "Windows Update failed: $($ex.Message)"

            # Log HResult/Win32Exception code if available
            if ($ex.HResult) {
                Write-Verbose ("Exception HResult: 0x{0:X8}" -f $ex.HResult)
            }

            # Specific remediation hint for cache error 0x80248007
            if ($ex.HResult -eq -2145124329) {
                Write-Warning "Encountered 0x80248007 (Windows Update cache). Consider resetting SoftwareDistribution and Background Intelligent Transfer Service (BITS)."
            }

            Start-Sleep -Seconds (15 * [Math]::Min($attempt, 6))
        }
    }

    return $false
}

# --- Main ---
Ensure-Tls12

# Optional: detect if running elevated (informational only)
$elevated = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if (-not $elevated) {
    Write-Verbose "Not running elevated. Install-Module without -Scope CurrentUser will fail. Use -CurrentUserInstall to avoid elevation or run as Administrator."
}

$installed = Ensure-PSWindowsUpdate -Attempts $MaxModuleAttempts -CurrentUser:$CurrentUserInstall
if (-not $installed) {
    Write-Error "Failed to install/import PSWindowsUpdate after $MaxModuleAttempts attempts."
    exit 2
}

$success = Run-DriverUpdates -Attempts $MaxWUAttempts
if (-not $success) {
    Write-Error "Driver update failed after $MaxWUAttempts attempts."
    exit 3
}

Write-Host "Driver update flow completed."
exit 0
