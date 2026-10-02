<#
.SYNOPSIS
    Launches a reviewed SharePoint Online site operation for one migration.
.DESCRIPTION
    Interactive use offers a preview by default and requires typed confirmation
    before applying a change. Scheduled use previews unless -Execute is given.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$MigrationRoot,

    [Parameter(Mandatory = $true)]
    [ValidateSet('DisablePageComments', 'SetSiteLockState')]
    [string]$Operation,

    [ValidateSet('ReadOnly', 'Unlock', 'NoAccess')]
    [string]$LockState,

    [switch]$Execute,
    [switch]$NonInteractive
)

$ErrorActionPreference = 'Stop'

function Write-LauncherMessage {
    param([string]$Message)
    Write-Host ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
}

function Read-LauncherInput {
    param([string]$Message)
    Read-Host ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
}

$migrationDirectory = (Resolve-Path -LiteralPath $MigrationRoot -ErrorAction Stop).Path
$configPath = Join-Path $migrationDirectory 'migration.config.psd1'
if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
    throw "Migration configuration not found: $configPath"
}
Import-LocalizedData -BindingVariable config -BaseDirectory $migrationDirectory -FileName 'migration.config.psd1' -ErrorAction Stop
if ((Split-Path -Leaf $migrationDirectory) -eq '_Template' -or [string]$config.Name -eq 'NewMigration') {
    throw 'The migration template cannot be executed directly. Copy Migrations\_Template to a named migration and update migration.config.psd1 first.'
}
if ([string]$config.Target.Type -ine 'SPO') {
    throw 'These operations require a SharePoint Online target.'
}
if ([string]::IsNullOrWhiteSpace([string]$config.Target.SiteUrl) -or
    [string]::IsNullOrWhiteSpace([string]$config.Target.TenantAdminUrl)) {
    throw 'Target.SiteUrl and Target.TenantAdminUrl are required in migration.config.psd1.'
}

if ($Operation -eq 'SetSiteLockState' -and [string]::IsNullOrWhiteSpace($LockState)) {
    if ($NonInteractive) {
        throw 'SetSiteLockState requires -LockState ReadOnly, Unlock, or NoAccess in noninteractive mode.'
    }
    $LockState = Read-LauncherInput 'Lock state (ReadOnly, Unlock, NoAccess; blank cancels)'
    if ([string]::IsNullOrWhiteSpace($LockState)) {
        Write-LauncherMessage 'Cancelled; no site change was made.'
        exit 2
    }
    if ($LockState -notin @('ReadOnly', 'Unlock', 'NoAccess')) {
        throw "Invalid lock state '$LockState'. Use ReadOnly, Unlock, or NoAccess."
    }
}

$action = if ($Operation -eq 'DisablePageComments') {
    'disable page comments'
}
else {
    "set the site lock state to '$LockState'"
}
Write-LauncherMessage ("Target site: {0}" -f $config.Target.SiteUrl)
Write-LauncherMessage ("Requested action: {0}" -f $action)

if (-not $NonInteractive) {
    if (-not $Execute) {
        $choice = Read-LauncherInput 'Press Enter for preview, or type EXECUTE to apply the change'
        if ($choice -ceq 'EXECUTE') {
            $Execute = $true
        }
        elseif (-not [string]::IsNullOrWhiteSpace($choice)) {
            Write-LauncherMessage 'Cancelled; no site change was made.'
            exit 2
        }
    }
    if ($Execute) {
        $confirmation = Read-LauncherInput ("Type YES to {0} on this site" -f $action)
        if ($confirmation -cne 'YES') {
            Write-LauncherMessage 'Cancelled; no site change was made.'
            exit 2
        }
    }
}

$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..') -ErrorAction Stop).Path
$scriptName = if ($Operation -eq 'DisablePageComments') {
    'SmartM365-SharePointTarget-PageCommentsDisable.ps1'
}
else {
    'SmartM365-SharePointTarget-SiteLockStateSet.ps1'
}
$operationScript = Join-Path $projectRoot (Join-Path 'Scripts\Operations' $scriptName)
if (-not (Test-Path -LiteralPath $operationScript -PathType Leaf)) {
    throw "Operation script not found: $operationScript"
}

$operationArgs = @{
    SiteUrl = [string]$config.Target.SiteUrl
    TenantAdminUrl = [string]$config.Target.TenantAdminUrl
}
if ($Operation -eq 'DisablePageComments') {
    $operationArgs.OutputName = if ([string]::IsNullOrWhiteSpace([string]$config.Name)) {
        (Split-Path -Leaf $migrationDirectory)
    }
    else {
        [string]$config.Name
    }
}
else {
    $operationArgs.LockState = $LockState
}
if (-not $Execute) {
    $operationArgs.WhatIfMode = $true
    Write-LauncherMessage 'Preview mode: no site change will be applied.'
}

& $operationScript @operationArgs

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBYKvOBCe23KEbv
# ApperKExwZ45TCadxwVD84NLJa6qLKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCKn9auFmLZvehajsOamD5j
# bPQEh7GqwnlZ2LrZ7X70bTANBgkqhkiG9w0BAQEFAASCAYBahIIsgmr0EZtjxh7g
# Z/ivsy1PuFrR+EVUWsN7RadOM6fThkQVyX9220hDJOfkExTeMHDfdVkA0LRfhN4K
# hhL5d4kO5eyRG8roboEiH7SmKwSUNjx45gvFekd0rnPu4KPXWmjYoYBuOUd91LK3
# EugtnF9WBjYHEXnHY8pgXT07LoxUQoTmRNLa4e9Dt30NNMwvBswEBnhwVAAYrcbq
# iCR4V7fK8PZBy6ulBg2Svkt3aflTi6fCaVFmO/qHYirFQtS7i7Ow6T1ogkAvmj03
# znhVwuRmLvhGmO/qZ38327pPH5yT+kgyF4EFlE1XeYG9EXHoB5koYObHoK7Jnurl
# JE7sgGnLvDdLEN6rsML0yBHZMuhOJu6vcRPyKsRi3OhxHJHNZu/gBkuauRyY/NMR
# fpYKdf27ZjhReYd9tEDHpqxTC9YBZKB3DVcgQy8UhL1SwtDKnN3F0TrWVWObnMf8
# 6A1etXVgE4W7/ZkVsGFCnoX+43D9MqnXh9b3nsglKoFjyLY=
# SIG # End signature block
