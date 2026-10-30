<#
.SYNOPSIS
    Regenerates the configuration templates from the option catalog.

.DESCRIPTION
    Writes Config\SmartEndpointMigrationToolkit.local.json.template (every option, inherit marker)
    and Campaigns\<Campaign>\Templates\Wave\Wave.local.json.template (options applicable to the
    campaign). Tests fail when the committed templates differ from this generator.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param([switch]$Check)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$packageRoot = Split-Path -Path $PSScriptRoot -Parent
Import-Module (Join-Path -Path $packageRoot -ChildPath 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
$catalog = Import-SemtOptionCatalog -PackageRoot $packageRoot

function ConvertTo-TemplateText {
    param([Parameter(Mandatory = $true)]$Object)
    return (($Object | ConvertTo-Json -Depth 6) -replace "`r`n","`n") + "`n"
}

$targets = New-Object System.Collections.Generic.List[object]
$globalOptions = [ordered]@{}
foreach ($definition in @($catalog.Options)) { $globalOptions[[string]$definition.Name] = '__INHERIT__' }
$targets.Add([pscustomobject]@{
    Path = Join-Path $packageRoot 'Config\SmartEndpointMigrationToolkit.local.json.template'
    Text = ConvertTo-TemplateText -Object ([ordered]@{
        SchemaVersion = 1
        Product = 'SmartEndpointMigrationToolkit'
        Storage = [ordered]@{ WavesPath = ''; RunsPath = '' }
        Description = 'Global options of this technician installation. __INHERIT__ keeps the campaign default of Config\SmartEndpointMigrationToolkit.options.json.txt. A Wave, a SEMT_* environment variable, a GUI choice or an explicit CLI option overrides it.'
        Options = $globalOptions
    })
})
foreach ($campaign in @('HybridJoin','Windows11Upgrade')) {
    $waveOptions = [ordered]@{}
    foreach ($definition in @(Get-SemtCampaignOptionDefinitions -Catalog $catalog -Campaign $campaign)) { $waveOptions[[string]$definition.Name] = '__INHERIT__' }
    $targets.Add([pscustomobject]@{
        Path = Join-Path $packageRoot ("Campaigns\{0}\Templates\Wave\Wave.local.json.template" -f $campaign)
        Text = ConvertTo-TemplateText -Object ([ordered]@{
            SchemaVersion = 1
            Product = 'SmartEndpointMigrationToolkit'
            Campaign = $campaign
            Description = ''
            Source = 'Manual'
            CreatedUtc = ''
            Options = $waveOptions
        })
    })
}

$different = @()
foreach ($target in $targets) {
    $current = if (Test-Path -LiteralPath $target.Path -PathType Leaf) { ([System.IO.File]::ReadAllText($target.Path) -replace "`r`n","`n") } else { '' }
    # Compared as JSON content: Windows PowerShell 5.1 and PowerShell 7 indent ConvertTo-Json
    # differently, and a template must not be reported out of date for that reason alone.
    $same = $false
    if ($current) {
        try { $same = (($current | ConvertFrom-Json | ConvertTo-Json -Depth 10 -Compress) -ceq ($target.Text | ConvertFrom-Json | ConvertTo-Json -Depth 10 -Compress)) } catch { $same = $false }
    }
    if (-not $same) {
        $different += $target.Path
        if (-not $Check) {
            New-Item -ItemType Directory -Path (Split-Path -Path $target.Path -Parent) -Force | Out-Null
            [System.IO.File]::WriteAllText($target.Path, ($target.Text -replace "`n","`r`n"), (New-Object System.Text.UTF8Encoding($false)))
        }
    }
}
if ($Check -and $different.Count -gt 0) { throw ("Configuration templates are out of date: {0}" -f ($different -join ', ')) }
[pscustomobject]@{ Checked = $targets.Count; Updated = $(if ($Check) { 0 } else { $different.Count }); Different = $different }

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAstGMqWF2zNnF4
# 8a+eA+mJQSX6HdazMRnxNot/dWP5VKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCUWeLzPv+8btEhYJZ7yWit
# 8K9kmMUKjwOneZnacFHLZzANBgkqhkiG9w0BAQEFAASCAYAVTCBldJJsvAQlmVl8
# L/U4lWu94y0HTfd0viHx5wAbkNYEQgxkmSF7QgDLpuq3eIM3T6heX4HWX9XMw9Mi
# OrTitUJLYXxicO3H1dFXYDs8NxZEJ+A0vimwlqQBsiNk4BCb+CfehyLN+DhWBQoW
# NtgJ+TIAQVCBFxuwTzDRUQOFRODZrUVOkXgcjMZdf1uv0cg9Dx0SKiXX5Kv++g7x
# hMdlt22kKDRfsHMpnba4SOWCdnq5/EGFo9yuHrfN8UUFg/xNkctipBEFCC94Xelv
# WftAiD92YPHRnNCtRTifRtZGgzL1jsUccMnBcwSRlmoM5U84WdgHOchCyabX783k
# jfJ3xEkVf+lgGbtJLQvU+dgQuvCH41XI4hVfHEpiQa5+SCdIqN7iz2jXRQzyyvwv
# yNSgTx/rxN2SILxgTQ98dRIZwnaiUDOrkcDCY5WtkI4cDXoaKeZF9rZqT7v0ugIV
# tOhgfTCPO8qS5jHSjTJPrrsOabCJBq33Q2Vp+nPH8Y45Ae8=
# SIG # End signature block
