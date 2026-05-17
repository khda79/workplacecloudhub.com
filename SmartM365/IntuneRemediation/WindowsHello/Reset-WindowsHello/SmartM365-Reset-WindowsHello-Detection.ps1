
<#
.SYNOPSIS
    Version: 1.0
    Intune detection script for Windows Hello for Business with multilingual AD connectivity check.
.DESCRIPTION
    This script:
    - Detects if Windows Hello is configured (NGC folder or registry keys).
    - Checks Hybrid Join status using dsregcmd.
    - Dynamically checks AD connectivity (supports English and French outputs).
    Exit codes:
      0 = Windows Hello configured and AD reachable.
      1 = Windows Hello not configured.
      2 = Windows Hello configured but AD unreachable.
.NOTES
    Run as SYSTEM via Intune.
#>

$LogRoot = Join-Path -Path $env:ProgramData -ChildPath 'SmartM365\IntuneRemediation\Logs\Reset-WindowsHello'
$LogPath = Join-Path -Path $LogRoot -ChildPath 'Reset-WindowsHello-Detection.log'
if (-not (Test-Path -LiteralPath $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
if (!(Test-Path (Split-Path $LogPath))) {
    New-Item -ItemType Directory -Path (Split-Path $LogPath) -Force
}
Start-Transcript -Path $LogPath -Append

Write-Host "[INFO] Starting Windows Hello detection..."

# Check NGC folder
$ngcPath = "C:\\Windows\\ServiceProfiles\\LocalService\\AppData\\Local\\Microsoft\\NGC"
$ngcConfigured = (Test-Path $ngcPath -and (Get-ChildItem $ngcPath -Recurse -ErrorAction SilentlyContinue).Count -gt 0)

# Check registry for PIN presence
$pinConfigured = $false
try {
    $regPath = "HKLM:\\SOFTWARE\\Microsoft\\PassportForWork\\"  # Windows Hello for Business key
    $keys = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue
    if ($keys) {
        $pinConfigured = $true
    }
} catch {
    Write-Host "[WARN] Registry path not found."
}

Write-Host "[INFO] Windows Hello configured: $($ngcConfigured -or $pinConfigured)"

# Check Hybrid Join status
$dsreg = dsregcmd /status | Out-String
$domainJoined = $dsreg -match "DomainJoined.*YES"
$azureJoined = $dsreg -match "AzureAdJoined.*YES"
Write-Host "[INFO] DomainJoined: $domainJoined, AzureAdJoined: $azureJoined"

# Dynamically check AD connectivity (supports FR and EN output)
$adReachable = $false
try {
    $nltestResult = nltest /dsgetdc: | Out-String
    if ($nltestResult -match "DC:" -or $nltestResult -match "Domain Name" -or 
        $nltestResult -match "Nom du domaine" -or $nltestResult -match "Contrôleur de domaine") {
        $adReachable = $true
    }
} catch {
    $adReachable = $false
}
Write-Host "[INFO] AD reachable: $adReachable"

Stop-Transcript

# Determine exit code
if (($ngcConfigured -or $pinConfigured) -and $adReachable) {
    Write-Host "[RESULT] Windows Hello configured and AD reachable."
    exit 0
} elseif (($ngcConfigured -or $pinConfigured) -and (-not $adReachable)) {
    Write-Host "[RESULT] Windows Hello configured but AD unreachable (Key Trust issue)."
    exit 2
} else {
    Write-Host "[RESULT] Windows Hello not configured."
    exit 1
}
