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
