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
Version: 1.1
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
    'Microsoft.Graph.DeviceManagement',
    'Microsoft.Graph.DeviceManagement.Administration',
    'Microsoft.Graph.DeviceManagement.Enrollment',
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
    }
}

Test-ActiveDirectoryPrerequisite

if ($WhatIfPreference) {
    Write-SmartM365SetupLog -Level SUCCESS -Message 'SmartM365 SmartInventory prerequisite audit completed. No installation was performed because -WhatIf is enabled.'
}
else {
    Write-SmartM365SetupLog -Level SUCCESS -Message 'SmartM365 SmartInventory prerequisites are ready.'
}
