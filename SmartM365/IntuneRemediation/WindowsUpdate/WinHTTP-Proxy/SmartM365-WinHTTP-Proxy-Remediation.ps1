<#
.SYNOPSIS
    Version: 1.0
    Repairs inconsistent WinHTTP proxy state for WinHTTP-Proxy.
.DESCRIPTION
    Reads the current WinHTTP proxy state and resets it to direct access when the configuration is unreadable or inconsistent. This does not change user-level browser proxy settings.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$Scenario = 'WinHTTP-Proxy'
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"
function Write-Log { param([string]$Message) $line = '{0} [{1}] {2}' -f (Get-Date -Format 's'), $Scenario, $Message; Write-Output $line; Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8 }

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-Log 'Remediation started.'
    $proxyOutput = netsh winhttp show proxy 2>&1
    $proxyText = ($proxyOutput -join ' ')
    Write-Log "Current WinHTTP proxy output: $proxyText"
    if ($LASTEXITCODE -ne 0 -or ($proxyText -notmatch 'Direct access' -and $proxyText -notmatch 'Proxy Server')) {
        netsh winhttp reset proxy | Out-Null
        Write-Log 'WinHTTP proxy reset to direct access.'
    }
    else {
        Write-Log 'WinHTTP proxy configuration is already valid; no reset required.'
    }
    exit 0
}
catch {
    Write-Log "Remediation failed: $($_.Exception.Message)"
    exit 1
}

