<#
.SYNOPSIS
    Smart Endpoint Migration Toolkit shared engine module.

.DESCRIPTION
    Loads the engine function files in a fixed order. The module is technician-side only:
    endpoint scripts never depend on it and remain deployable as single files.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

Set-StrictMode -Version 2.0

# Refuse unsupported engines before any function can write state.
$semtEdition = [string]$PSVersionTable['PSEdition']
$semtVersion = $PSVersionTable['PSVersion']
if (($semtEdition -eq 'Core' -and $semtVersion -lt [version]'7.2') -or ($semtEdition -ne 'Core' -and $semtVersion -lt [version]'5.1')) {
    throw ('SEMT_UNSUPPORTED_POWERSHELL: {0} {1} is not supported. Use Windows PowerShell 5.1 or PowerShell 7.2 and later.' -f $semtEdition,$semtVersion)
}

$script:SemtModuleVersion = '1.0.0'
$script:SemtModuleRoot = $PSScriptRoot

foreach ($functionFile in @(
    'Semt.Common.ps1',
    'Semt.Options.ps1',
    'Semt.ComputerList.ps1',
    'Semt.Concurrency.ps1',
    'Semt.Inventory.ps1',
    'Semt.Network.ps1',
    'Semt.Run.ps1',
    'Semt.Remote.ps1',
    'Semt.Reports.ps1'
)) {
    $functionPath = Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Functions') -ChildPath $functionFile
    if (-not (Test-Path -LiteralPath $functionPath -PathType Leaf)) {
        throw "SmartEndpointMigration.Core component is missing: $functionPath"
    }
    . $functionPath
}
