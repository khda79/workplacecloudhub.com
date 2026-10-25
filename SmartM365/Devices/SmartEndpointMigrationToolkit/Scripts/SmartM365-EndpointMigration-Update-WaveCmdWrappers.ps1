<#
.SYNOPSIS
Regenerates the Loop/Once launch wrappers of Smart Endpoint Migration Toolkit Waves.

.DESCRIPTION
Calls Write-SemtWaveWrappers for every Wave folder of the selected Campaign (both
Campaigns when -Campaign is omitted), or for the single Wave given with -WavePath.
Obsolete Start-Wave-*.cmd files are removed by the engine. Prints one timestamped line
per Wave and a final summary. Exit code 0 when every Wave was updated, 1 otherwise.

.PARAMETER Campaign
HybridJoin or Windows11Upgrade. Both Campaigns are processed when omitted.

.PARAMETER WavePath
Optional single Wave folder (<Package>\Waves\<Campaign>\<Wave>).

.PARAMETER PackageRoot
Package root. Defaults to the package that contains this script.

.VERSION
1.0.0

.NOTES
Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [ValidateSet('HybridJoin', 'Windows11Upgrade')]
    [string]$Campaign,
    [string]$WavePath,
    [string]$PackageRoot
)

$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.0.0'

function Write-Host {
    param(
        [Parameter(Position = 0)][object]$Object = '',
        [switch]$NoNewline,
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor
    )

    foreach ($line in @(([string]$Object) -split "`r?`n")) {
        $parameters = @{ Object = ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $line) }
        if ($PSBoundParameters.ContainsKey('ForegroundColor')) { $parameters['ForegroundColor'] = $ForegroundColor }
        if ($PSBoundParameters.ContainsKey('BackgroundColor')) { $parameters['BackgroundColor'] = $BackgroundColor }
        if ($PSBoundParameters.ContainsKey('NoNewline')) { $parameters['NoNewline'] = $NoNewline }
        Microsoft.PowerShell.Utility\Write-Host @parameters
    }
}

$exitCode = 0
$updated = 0
$failed = 0
try {
    $scriptsRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $coreModulePath = [System.IO.Path]::GetFullPath((Join-Path $scriptsRoot '..\Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1'))
    Import-Module $coreModulePath -MinimumVersion 1.0.0 -Force
    $effectivePackageRoot = if ([string]::IsNullOrWhiteSpace($PackageRoot)) { Get-SemtPackageRoot -StartPath $scriptsRoot } else { [System.IO.Path]::GetFullPath($PackageRoot) }

    Write-Host ("Smart Endpoint Migration Toolkit - Update-WaveCmdWrappers version {0}" -f $ScriptVersion) -ForegroundColor Cyan
    Write-Host ("Package     : {0}" -f $effectivePackageRoot)

    $targets = New-Object System.Collections.Generic.List[object]
    if (-not [string]::IsNullOrWhiteSpace($WavePath)) {
        $waveFull = [System.IO.Path]::GetFullPath($WavePath).TrimEnd('\')
        if (-not (Test-Path -LiteralPath $waveFull -PathType Container)) { throw "Wave folder not found: $waveFull" }
        $waveCampaign = Split-Path -Leaf (Split-Path -Parent $waveFull)
        if (@('HybridJoin', 'Windows11Upgrade') -notcontains $waveCampaign) {
            throw "The Wave must be under Waves\HybridJoin or Waves\Windows11Upgrade: $waveFull"
        }
        if (-not [string]::IsNullOrWhiteSpace($Campaign) -and $Campaign -ne $waveCampaign) {
            throw ("The Wave belongs to Campaign {0}, not {1}: {2}" -f $waveCampaign, $Campaign, $waveFull)
        }
        $expectedRoot = [System.IO.Path]::GetFullPath((Get-SemtWaveRoot -PackageRoot $effectivePackageRoot -Campaign $waveCampaign)).TrimEnd('\')
        if ((Split-Path -Parent $waveFull) -ine $expectedRoot) {
            throw ("The Wave is not inside the package Waves folder {0}: {1}" -f $expectedRoot, $waveFull)
        }
        $targets.Add([pscustomobject]@{ Campaign = $waveCampaign; Name = (Split-Path -Leaf $waveFull); Path = $waveFull })
    }
    else {
        $campaigns = if ([string]::IsNullOrWhiteSpace($Campaign)) { @('HybridJoin', 'Windows11Upgrade') } else { @($Campaign) }
        foreach ($campaignName in $campaigns) {
            foreach ($wave in @(Get-SemtWaves -PackageRoot $effectivePackageRoot -Campaign $campaignName)) {
                $targets.Add([pscustomobject]@{ Campaign = $campaignName; Name = [string]$wave.Name; Path = [string]$wave.Path })
            }
        }
    }

    if ($targets.Count -eq 0) {
        Write-Host 'No Wave folder found.' -ForegroundColor Yellow
    }
    foreach ($target in $targets) {
        try {
            if (-not (Test-SemtWaveName -Name $target.Name)) {
                throw ("SEMT_WAVE_NAME_INVALID: '{0}' is not a valid Wave name." -f $target.Name)
            }
            $written = @(Write-SemtWaveWrappers -WavePath $target.Path)
            $updated++
            Write-Host ("OK     {0}\{1}: {2} wrapper(s) written" -f $target.Campaign, $target.Name, $written.Count) -ForegroundColor Green
        }
        catch {
            $failed++
            Write-Host ("FAILED {0}\{1}: {2}" -f $target.Campaign, $target.Name, $_.Exception.Message) -ForegroundColor Red
        }
    }
    Write-Host ("Summary: Waves={0}; Updated={1}; Failed={2}" -f $targets.Count, $updated, $failed) -ForegroundColor $(if ($failed -gt 0) { 'Yellow' } else { 'Green' })
    if ($failed -gt 0) { $exitCode = 1 }
}
catch {
    Write-Host ("ERROR: {0}" -f $_.Exception.Message) -ForegroundColor Red
    $exitCode = 1
}
exit $exitCode
