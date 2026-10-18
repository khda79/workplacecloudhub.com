#Requires -Version 7.0
<#
.SYNOPSIS
    Runs the Smart IntunePackager Pester 5 tests.

.DESCRIPTION
    Unit tests are offline. The Integration tag runs the real Win32 Content Prep Tool and
    requires IP_TEST_CONTENT_PREP_TOOL to point to a verified IntuneWinAppUtil.exe.
    The runner relaunches itself in an STA process so that the WPF tests run.

.VERSION
    1.1.0

.EXAMPLE
    pwsh -NoProfile -ExecutionPolicy Bypass -File .\Tests\SmartM365-IntunePackager-RunTests.ps1
#>
[CmdletBinding()]
param(
    [switch]$IncludeIntegration,
    [ValidateSet('None', 'Normal', 'Detailed', 'Diagnostic')][string]$Output = 'Detailed'
)

$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA) {
    $arguments = @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-Output', $Output)
    if ($IncludeIntegration) { $arguments += '-IncludeIntegration' }
    & (Get-Process -Id $PID).Path @arguments
    exit $LASTEXITCODE
}
$pester =Get-Module -ListAvailable -Name Pester | Where-Object { $_.Version -ge [version]'5.0.0' } | Sort-Object Version -Descending | Select-Object -First 1
if (-not $pester) {
    throw 'Pester 5 or later is required. Install it with: Install-Module Pester -MinimumVersion 5.0.0 -Scope CurrentUser -SkipPublisherCheck'
}
Import-Module $pester.Path -Force

$configuration = New-PesterConfiguration
$configuration.Run.Path = $PSScriptRoot
$configuration.Run.Exit = $true
$configuration.Output.Verbosity = $Output
if (-not $IncludeIntegration) { $configuration.Filter.ExcludeTag = @('Integration') }
Invoke-Pester -Configuration $configuration
