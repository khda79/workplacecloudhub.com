
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
      0 = Healthy or not applicable.
      1 = Windows Hello configured but AD unreachable, or technical error.
.NOTES
    Run as SYSTEM via Intune.
#>

$ErrorActionPreference = "Stop"
$Scenario = "Reset-WindowsHello"

function Write-DetectionResult {
    param([string]$Message)
    Write-Output "$Scenario $Message"
}

# Check NGC folder
$ngcPath = "C:\\Windows\\ServiceProfiles\\LocalService\\AppData\\Local\\Microsoft\\NGC"
$ngcConfigured = ((Test-Path -LiteralPath $ngcPath) -and ((Get-ChildItem -LiteralPath $ngcPath -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1).Count -gt 0))

# Check registry for PIN presence
$pinConfigured = $false
try {
    $regPath = "HKLM:\\SOFTWARE\\Microsoft\\PassportForWork\\"  # Windows Hello for Business key
    $keys = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue
    if ($keys) {
        $pinConfigured = $true
    }
} catch {
    $pinConfigured = $false
}

$helloConfigured = $ngcConfigured -or $pinConfigured

# Check Hybrid Join status
$dsreg = dsregcmd /status | Out-String
$domainJoined = $dsreg -match "DomainJoined.*YES"
$azureJoined = $dsreg -match "AzureAdJoined.*YES"

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
# Determine exit code
if (-not $helloConfigured) {
    Write-DetectionResult "Status=NotApplicable Reason=WindowsHelloNotConfigured"
    exit 0
}

if (-not $domainJoined) {
    Write-DetectionResult "Status=Healthy Reason=WindowsHelloConfiguredWithoutHybridJoin DomainJoined=$domainJoined AzureAdJoined=$azureJoined"
    exit 0
}

if ($adReachable) {
    Write-DetectionResult "Status=Healthy Reason=WindowsHelloConfiguredAndADReachable DomainJoined=$domainJoined AzureAdJoined=$azureJoined"
    exit 0
}

Write-DetectionResult "Status=RemediationRequired Reason=WindowsHelloConfiguredButADUnreachable DomainJoined=$domainJoined AzureAdJoined=$azureJoined"
exit 1
