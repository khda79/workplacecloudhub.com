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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA8BWGahD86hsjn
# b9jVvchMJljTIysj8CoBSE0vT7v2zqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCDGoIXa2CO6ER7/oriFrjvsOYg8YbZEDEIJchqsBgLX0TANBgkqhkiG9w0B
# AQEFAASCAYCK6Q+38jNJ1h7upK+X3KqP4B2lCNseEXdzTUzcCUvP/WiZBbK01iaK
# MT9lJjKW8IMOICjOMXBaZA8QMDbGQxOlFPHZqrxpHCt5gI8lHKpkmkD3z370R6pA
# BYQcjZBoOBFTR1Rp3FCdVTUSWJg7aYbjW2NTmXypgb6/kA7/0arfSYy76dDJV6nQ
# eBX3KHYe86odxmku+VTvAvGqloWidOEJ70b78y1tyt13hanxgusDOdSAM8qVXvbX
# 02sHAw52c2qb+2e/sYmFgWVNSYQGma7i4L8pcQbUFsaxmQBcit3GbbMdG2TBThL3
# /gV9kT+2Uwrz/YV5Q9XJ0v/Xun17aNbnS5rxxNMlx64SKtmYQoiefZkYw69EDPDs
# sKXXZrQSZoSrIQ0qoZkBE6ZH/o/czmCFg8w8uIJdO4HMInPupti4t6O0bymMc0P7
# +ZEPyxYzVUC9sTLxmwlXx2gQ57JzidhULOmUrijE4Wng04+VKnFRfjs0Dll1HMVE
# 0+WqEUTz+uY=
# SIG # End signature block
