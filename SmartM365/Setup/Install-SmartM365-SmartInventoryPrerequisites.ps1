<#
.SYNOPSIS
Installs PowerShell prerequisites for SmartM365 SmartInventory scripts.

.DESCRIPTION
Installs the PowerShell modules required by SmartM365 SmartInventory scripts from PSGallery,
then validates that cloud modules can be imported.

ActiveDirectory is a Windows RSAT / Windows Server feature, not a PSGallery module. This script
checks for it and prints the appropriate install hint when it is missing.

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

.PARAMETER SkipMicrosoftGraphMetaModule
Skips the Microsoft.Graph rollup module. Use only when Graph submodules are managed separately.

.PARAMETER WhatIf
Shows what would be installed without installing modules.

.PARAMETER SkipImportValidation
Skips Import-Module validation after installation.

.NOTES
Version: 1.2
Author: https://github.com/khda79/workplacecloudhub.com
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]$Scope = 'CurrentUser',
    [string]$Repository = 'PSGallery',
    [switch]$TrustRepository,
    [switch]$Force = $true,
    [switch]$AllowClobber = $true,
    [switch]$SkipPublisherCheck,
    [switch]$SkipMicrosoftGraphMetaModule,
    [switch]$SkipImportValidation
)

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    exit 1
}

$ErrorActionPreference = 'Stop'

$requiredModules = @()
if (-not $SkipMicrosoftGraphMetaModule) {
    $requiredModules += 'Microsoft.Graph'
}

$requiredModules += @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Applications',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'Microsoft.Graph.Sites',
    'Microsoft.Graph.Teams',
    'Microsoft.Graph.DeviceManagement',
    'Microsoft.Graph.DeviceManagement.Administration',
    'Microsoft.Graph.DeviceManagement.Enrollment',
    'Microsoft.Graph.Beta.DeviceManagement',
    'Microsoft.Graph.Beta.Reports',
    'Microsoft.Graph.Policies',
    'Microsoft.Graph.Reports',
    'ExchangeOnlineManagement',
    'PnP.PowerShell',
    'MSAL.PS'
)

$requiredModules = @($requiredModules | Select-Object -Unique)

function Write-SmartM365SetupLog {
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
        Write-SmartM365SetupLog -Level SUCCESS -Message ("Module already installed: {0} {1}" -f $Name, $installed.Version)
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
        Write-SmartM365SetupLog -Message ("Installing module: {0}" -f $Name)
        Install-Module @installParams
    }
}

function Test-ActiveDirectoryPrerequisite {
    [CmdletBinding()]
    param()

    if (Get-Module -ListAvailable -Name ActiveDirectory) {
        Write-SmartM365SetupLog -Level SUCCESS -Message 'ActiveDirectory module is available.'
        return
    }

    Write-SmartM365SetupLog -Level WARN -Message 'ActiveDirectory module is missing. It is installed through RSAT or Windows Server features, not PSGallery.'
    Write-Host ''
    Write-Host 'Windows 10/11 client hint:'
    Write-Host '  Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'
    Write-Host ''
    Write-Host 'Windows Server hint:'
    Write-Host '  Install-WindowsFeature RSAT-AD-PowerShell'
    Write-Host ''
}

Write-SmartM365SetupLog -Message 'Installing SmartM365 SmartInventory prerequisites.'
Ensure-PackageProvider
Ensure-RepositoryTrust -Name $Repository

foreach ($moduleName in $requiredModules) {
    Install-RequiredModule -Name $moduleName
}

if (-not $SkipImportValidation) {
    foreach ($moduleName in $requiredModules) {
        Write-SmartM365SetupLog -Message ("Validating module import: {0}" -f $moduleName)
        Import-Module -Name $moduleName -ErrorAction Stop
        $loadedModule = Get-Module -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $loadedModule) {
            $loadedModule = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
        }
        if ($loadedModule) {
            Write-SmartM365SetupLog -Level SUCCESS -Message ("Module ready: {0} {1}; Path={2}" -f $loadedModule.Name, $loadedModule.Version, $loadedModule.Path)
        }
        else {
            throw ("Module import reported success but module version could not be resolved: {0}" -f $moduleName)
        }
    }
}

Test-ActiveDirectoryPrerequisite

if ($WhatIfPreference) {
    Write-SmartM365SetupLog -Level SUCCESS -Message 'SmartM365 SmartInventory prerequisite audit completed. No installation was performed because -WhatIf is enabled.'
}
else {
    Write-SmartM365SetupLog -Level SUCCESS -Message 'SmartM365 SmartInventory prerequisites are ready.'
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA8BWGahD86hsjn
# b9jVvchMJljTIysj8CoBSE0vT7v2zqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDGoIXa2CO6ER7/oriF
# rjvsOYg8YbZEDEIJchqsBgLX0TANBgkqhkiG9w0BAQEFAASCAYAHRbad0agNlvXR
# E+z9N9IMYBP6ZmoGtowRc9J8X1+oGMLnuA3uGZqRsDiBot4gQk6VA4nvM7VvpS3H
# T1VQCUjskfNvhudwgArLzcgPQZpc2zFk6ue8HDc9ziYuxi4DA8FhlDNEu9XgeK+1
# 31DzI62KM4bkRXW/jDdCCBldABr/jDSv+m7xBuF0mHLkQkUqXdIFfv1TM6AdaBXz
# xobvBDXSso9z/BHM0pIowUzrQuhWWSC8dVPYc21wGrUCDh6GjJEvMBdghhQLrnH0
# TXz6pd6jg1/feLo27nmeZtBOCNrCeJefvVuk2DqHPIlM0/dLgp7WFy7dNAYkDwCR
# JRfMphz1VtO71k3/xg1RC9e7DuxI5S/RlIu80tEbd4rVFDrxpGQwNIOwBWpqn5eR
# UUa25gODti2nzA4uMPrBsQCtjvWU1NGYNGAy+fRDt6x3kPMZn55CO7aTfGxFl2um
# bpLaZUDdqJVNDn9B5g2ESN4DoowYYMICcpTnuEvHp65wnKEz1Ws=
# SIG # End signature block
