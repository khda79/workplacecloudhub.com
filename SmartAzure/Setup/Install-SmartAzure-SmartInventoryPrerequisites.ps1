<#
.SYNOPSIS
Installs PowerShell prerequisites for SmartAzure SmartInventory scripts.

.DESCRIPTION
Installs the PowerShell modules required by SmartAzure SmartInventory scripts from PSGallery,
then validates that each module can be imported.

The current SmartAzure SmartInventory scripts use Az modules for Azure Resource Manager data.
Microsoft.Graph.Authentication is included for shared SmartAzure notification features that
can send Graph app-only mail on failure.

.PARAMETER Scope
PowerShellGet installation scope. Defaults to CurrentUser.

.PARAMETER Repository
PowerShell repository name. Defaults to PSGallery.

.PARAMETER TrustRepository
Sets the repository installation policy to Trusted before installing modules.

.PARAMETER Force
Passes -Force to Install-Module.

.PARAMETER AllowClobber
Passes -AllowClobber to Install-Module.

.PARAMETER SkipPublisherCheck
Passes -SkipPublisherCheck to Install-Module.

.PARAMETER WhatIf
Shows what would be installed without installing modules.

.PARAMETER SkipImportValidation
Skips Import-Module validation after installation.

.NOTES
Version: 1.0
Author: https://github.com/khda79/workplacecloudhub.com
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]$Scope = 'CurrentUser',
    [string]$Repository = 'PSGallery',
    [switch]$TrustRepository,
    [switch]$Force,
    [switch]$AllowClobber,
    [switch]$SkipPublisherCheck,
    [switch]$SkipImportValidation
)

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    exit 1
}

$ErrorActionPreference = 'Stop'

$requiredModules = @(
    'Az.Accounts',
    'Az.Resources',
    'Az.Compute',
    'Az.Network',
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Applications'
)

function Write-SmartAzureSetupLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    Write-Host ("{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message)
}

function Ensure-PackageProvider {
    [CmdletBinding()]
    param()

    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        if ($PSCmdlet.ShouldProcess('NuGet package provider', 'Install')) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope $Scope -Force | Out-Null
        }
    }
}

function Ensure-RepositoryTrust {
    [CmdletBinding()]
    param([string]$Name)

    $repository = Get-PSRepository -Name $Name -ErrorAction Stop
    if ($TrustRepository -and $repository.InstallationPolicy -ne 'Trusted') {
        if ($PSCmdlet.ShouldProcess($Name, 'Set PSGallery installation policy to Trusted')) {
            Set-PSRepository -Name $Name -InstallationPolicy Trusted
        }
    }
}

function Install-RequiredModule {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    $installed = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    if ($installed -and -not $Force) {
        Write-SmartAzureSetupLog -Level SUCCESS -Message ("Module already installed: {0} {1}" -f $Name, $installed.Version)
        return
    }

    $installParams = @{
        Name       = $Name
        Scope      = $Scope
        Repository = $Repository
        ErrorAction = 'Stop'
    }
    if ($Force) { $installParams['Force'] = $true }
    if ($AllowClobber) { $installParams['AllowClobber'] = $true }
    if ($SkipPublisherCheck) { $installParams['SkipPublisherCheck'] = $true }

    if ($PSCmdlet.ShouldProcess($Name, "Install-Module -Scope $Scope")) {
        Write-SmartAzureSetupLog -Message ("Installing module: {0}" -f $Name)
        Install-Module @installParams
    }
}

Write-SmartAzureSetupLog -Message 'Installing SmartAzure SmartInventory prerequisites.'
Ensure-PackageProvider
Ensure-RepositoryTrust -Name $Repository

foreach ($moduleName in $requiredModules) {
    Install-RequiredModule -Name $moduleName
}

if (-not $SkipImportValidation) {
    foreach ($moduleName in $requiredModules) {
        Write-SmartAzureSetupLog -Message ("Validating module import: {0}" -f $moduleName)
        Import-Module -Name $moduleName -ErrorAction Stop
    }
}

if ($WhatIfPreference) {
    Write-SmartAzureSetupLog -Level SUCCESS -Message 'SmartAzure SmartInventory prerequisite audit completed. No installation was performed because -WhatIf is enabled.'
}
else {
    Write-SmartAzureSetupLog -Level SUCCESS -Message 'SmartAzure SmartInventory prerequisites are ready.'
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBOMK0+6G53GsL4
# 4Sy9wOkOOrN7QZ8h2wzThhFamPnwBqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCm1TNikJwss1rNt3eG
# ecBIBKGEEcdarIVDRiknAepzxDANBgkqhkiG9w0BAQEFAASCAYCfvdPdKB9JnMuD
# xjQ0LCfjK6zxQw/twuq6hYpP/yz2XfA+Rwy9YoDPSBV2UCcX3ZRhLBuvSHtM626/
# ZPL7dWMT73vaBAQu7txmFXKMSr+rVIDMmNxK+6gTBZ3SxuY3RAAe23I2KdZkKVkQ
# EnzexBzY/wVceO3K0m9/mY3+GutxUadO/XziKPUtOBil9Gi1FqVRoKMWNM+/ypWr
# 1IRVb9HiTHxODOwUioCxePk+HEjybfcdTz5odKbRrhIp3LRJPRJesy+MOD6w/Fvp
# d8oowdtqrlzYfqPYWTiLyIVMfmmsblbF0bbzILRh3ZFbMga0kDUnZmMHYy7zemY8
# EpMcKbpvSQuMf8D1eNjLJtcFPgpTkT4anNrHdl1Tg0zZLuocgbJ33hbQb8cZC9q3
# KpEdflBatzcR2qIn9Py8fSfSnt55R73C+/nqpX15v+yYfSsINAD9rNr2PqyDy7i8
# ALYiwGXcH51kwuNN83tP3WXo79YxhRFG0CyT4i0gsNQnbVRtRLI=
# SIG # End signature block
