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
            $written = @(Write-SemtWaveWrappers -WavePath $target.Path -PackageRoot $effectivePackageRoot)
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

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDggY/BC/XymT6h
# sb2KeD5liSgdjGhhakTmKWrAZ9Xpa6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDU6AaUJrRAWkmD/8O8VC0g
# MePLdPWSVOfUz4tg/1lHyjANBgkqhkiG9w0BAQEFAASCAYAv6YwNdRlyNWJhjVYf
# 1GRtj/DBCQePptdw8/aEVv8fxJUKSQM+Qi9TCCXMwHW+WAprhxJymF5rvvChgWJv
# uA3OqH0jV9FBb4jq7CzJbSSSU+Vtzo9Kd2gE10WutM+4hnDs1PiD/vddgfG1uluQ
# z+aN6ywj+KboN8dpLjlAWLoOUf4Qok4PufRBQV1z9XX2YU4xwuiyFzhClo2/eldV
# 1crU131niRgbWo+RX4Ciy1vVrSpGEIwEVNwkblvRfzJW2uCs4XPwsS8e3VVu5HYs
# oX//DXDW9NfVyL6eY0g3EPDfUq9kIgdUp+wbFbv3cZLKkrqt2nxKH/HZzdC/qJVl
# gIkcxT6rvgf0uD1U6aRDV9DPHKDk1/LDXt8LEzmr7M5qAiHSW7EAwduP+wnMr1xK
# NavE/F1NbnYCTkPex9ptW/GWt8tOvOi1hQsPEMezMfBIVANkRmY2j81t/SGN+sCE
# r6xH4hw4sSBDJJh195doxQMfkiJwk61kJ22kgFhZyXyLHRk=
# SIG # End signature block
