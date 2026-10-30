<#
.SYNOPSIS
Read-only targeted user license lookup for an identified Single PC user.
.DESCRIPTION
Uses a unique on-premises SID and the confirmed device tenant ID. The operator must
obtain these from a current endpoint audit. No user, license or endpoint is modified.
.VERSION
1.0.0
#>
#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidatePattern('^S-1-(?:\d+-)+\d+$')][string]$UserSid,
    [Parameter(Mandatory=$true)][guid]$TenantId,
    [Parameter(Mandatory=$true)][string]$OutputPath,
    [switch]$NoConnect
)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $root 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
$module=Find-SemtGraphAuthenticationModule
if (-not $module) { throw 'Microsoft.Graph.Authentication is not installed.' }
Import-Module $module.Path -ErrorAction Stop
if (-not $NoConnect) {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Connect-MgGraph -TenantId $TenantId.ToString() -Scopes 'User.Read.All' -ContextScope Process -NoWelcome -ErrorAction Stop | Out-Null
}
try {
    $context=Get-MgContext
    if (-not $context -or $context.AuthType -ne 'Delegated' -or $context.TenantId -ine $TenantId.ToString()) { throw 'Graph context or tenant mismatch.' }
    $filter=[uri]::EscapeDataString("onPremisesSecurityIdentifier eq '$UserSid'")
    $uri='https://graph.microsoft.com/v1.0/users?$filter='+$filter+'&$count=true&$select=id,userPrincipalName,onPremisesSecurityIdentifier,accountEnabled,usageLocation'
    $response=Invoke-MgGraphRequest -Method GET -Uri $uri -Headers @{ConsistencyLevel='eventual'} -ErrorAction Stop
    if (-not $response.Contains('value') -or $null -eq $response['value'] -or $response['@odata.nextLink']) { throw 'User lookup incomplete or paginated.' }
    $users=@($response['value'])
    if ($users.Count -ne 1 -or $users[0]['onPremisesSecurityIdentifier'] -ine $UserSid) { throw "User SID is missing or ambiguous: $($users.Count) matches." }
    $user=$users[0]
    $details=Invoke-MgGraphRequest -Method GET -Uri ('https://graph.microsoft.com/v1.0/users/'+$user['id']+'/licenseDetails') -ErrorAction Stop
    if (-not $details.Contains('value') -or $null -eq $details['value'] -or $details['@odata.nextLink']) { throw 'License lookup incomplete or paginated.' }
    $report=[ordered]@{ CollectedUtc=[datetime]::UtcNow.ToString('o'); TenantId=$TenantId.ToString(); User=$user; LicenseDetails=$details['value']; MdmUserScope='Unknown: verify in Entra'; Provenance='TargetedGraphRead' }
    $full=[IO.Path]::GetFullPath($OutputPath)
    $shared=[IO.Path]::GetFullPath((Join-Path $root 'Inventory\UsersLicenses.csv'))
    if ($full -ieq $shared -or [IO.File]::Exists($full)) { throw 'Output already exists or targets the shared cache.' }
    $parent=Split-Path -Parent $full
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) { throw 'Output directory does not exist.' }
    $report | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $full -Encoding UTF8 -ErrorAction Stop
    Write-Host "Targeted license evidence: $full"
} finally { if (-not $NoConnect) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } }

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCZmWsm1fakiGfW
# UfgiPs48QdTK4tLZ7zlNXbml8m9X9qCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCgtVzNlzTl8yDTqXxQE9S3
# isibvvW+XTs905QP0Qj5FzANBgkqhkiG9w0BAQEFAASCAYBcuIwiNR943qY1WIah
# ERuRj1RvLhweV48Q63LXSeqXFAN77ipjz0Ybf4OeZNjPLvEmTuG5OowCpaCogg5J
# 2UXzcuswEpgM4crp68QaPp3RnJHlRiFt9NjI1RROP4WC8/rWbQhiaHudvW+lBGrk
# Ko/S3jB2VRzfFgZOjEOssOuW475g/CoYB3vuRpWbSZjBSw981bXOMsAHxAnZgrpZ
# tlSqlreFRKDbREPo/x/mkPMHOexjNKRcsSYwY3PvXNKqZ3R93xg9Vgb2zufZeT6B
# nYNFqZdU6lWY4uht9o5lddWD16mtYapfFbgsEoP/NgxnvESCuW6JDDDIDCLFGMTI
# 3GlQ5wM5z02Y21OBwOSX0qlqx939dROxtI4ayzEqcbRbndJy4upogMQZxmaDLnY7
# O232C149QA+5z0rSQq/dwhYrbU9pKf2MhcOxqt/smOK1CxuwdW8oW/jrhjevVHSc
# 3fLffIZN+Znbl4yI0qIe4eWH3niYvpnYhUuSHgPqnRSkclE=
# SIG # End signature block
