<#
.VERSION
1.0.0
#>
<#
.SYNOPSIS
Exports user identity and assigned service plans for SEMT diagnostics.
.DESCRIPTION
Read-only delegated Graph export. A complete tenant-wide response is required before
replacing the cache. The file contains personal data and must be protected accordingly.
#>
#requires -Version 5.1
[CmdletBinding()]
param([string]$OutputPath,[string]$TenantId,[switch]$NoConnect,[switch]$ForceRefresh)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $root 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
if (-not $OutputPath) { $OutputPath = Join-Path $root 'Inventory\UsersLicenses.csv' }

# Cache provenance is carried in each row; an empty, incomplete or stale export is never trusted.
function Test-UserLicenseCache {
    param([string]$Path,[string]$ExpectedTenant)
    if (-not $ExpectedTenant -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    try {
        $age = ([datetime]::UtcNow - (Get-Item -LiteralPath $Path -ErrorAction Stop).LastWriteTimeUtc).TotalHours
        if ($age -lt 0 -or $age -ge 24) { return $false }
        $first = Import-Csv -LiteralPath $Path -ErrorAction Stop | Select-Object -First 1
        if (-not $first -or $first.InventoryTenantId -ine $ExpectedTenant -or $first.InventoryScope -ne 'AllUsers' -or $first.InventoryComplete -ne 'True') { return $false }
        $stamp = ([string]$first.CollectedUtc).Trim()
        if ($stamp -notmatch '(?i)(Z|[+-]\d{2}:\d{2})$') { return $false }
        $collected = ConvertTo-SemtUtcDateTime -Value $stamp
        if ($null -eq $collected) { return $false }
        $recordAge = ([datetime]::UtcNow - $collected).TotalHours
        return ($recordAge -ge 0 -and $recordAge -lt 24)
    } catch { return $false }
}
if (-not $ForceRefresh -and (Test-UserLicenseCache -Path $OutputPath -ExpectedTenant $TenantId)) {
    Write-Host "Using verified user inventory: $OutputPath"
    return
}
$module = Find-SemtGraphAuthenticationModule
if (-not $module) { throw 'Microsoft.Graph.Authentication is not installed.' }
Import-Module $module.Path -ErrorAction Stop
if (-not $NoConnect) {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    $connection = @{ Scopes = @('User.Read.All'); ContextScope = 'Process'; NoWelcome = $true; ErrorAction = 'Stop' }
    if ($TenantId) { $connection.TenantId = $TenantId }
    Connect-MgGraph @connection | Out-Null
}
try {
    $context = Get-MgContext
    if (-not $context -or $context.AuthType -ne 'Delegated' -or ($TenantId -and $context.TenantId -ne $TenantId)) { throw 'Delegated Graph context or tenant mismatch.' }
    $tenant = [string]$context.TenantId
    if (-not $tenant) { throw 'Graph tenant unavailable.' }
    if (-not $ForceRefresh -and (Test-UserLicenseCache -Path $OutputPath -ExpectedTenant $tenant)) {
        Write-Host "Using verified user inventory: $OutputPath"
        return
    }
    $next = 'https://graph.microsoft.com/v1.0/users?$select=id,userPrincipalName,onPremisesSamAccountName,onPremisesDomainName,onPremisesSecurityIdentifier,accountEnabled,usageLocation,assignedPlans,assignedLicenses&$top=999'
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $seenPages = @{}
    while ($next) {
        if ($seenPages.ContainsKey($next) -or $next -notmatch '^https://graph\.microsoft\.com/') { throw 'Invalid or repeated Graph continuation URL.' }
        $seenPages[$next] = $true
        $page = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
        if (-not $page.Contains('value') -or $null -eq $page['value']) { throw 'Graph response is missing its value collection.' }
        foreach ($u in $page['value']) {
            if (-not $u['id'] -or -not $u.Contains('assignedPlans') -or $null -eq $u['assignedPlans'] -or -not $u.Contains('assignedLicenses') -or $null -eq $u['assignedLicenses']) { throw 'Graph returned a user without required license fields.' }
            $plans = @($u['assignedPlans'] | ForEach-Object { '{0}:{1}' -f $_['servicePlanId'],$_['capabilityStatus'] }) -join ';'
            $licenses = @($u['assignedLicenses'] | ForEach-Object { [string]$_['skuId'] }) -join ';'
            $rows.Add([pscustomobject]@{
                UserId=[string]$u['id']; UserPrincipalName=[string]$u['userPrincipalName']
                OnPremisesSamAccountName=[string]$u['onPremisesSamAccountName']; OnPremisesDomainName=[string]$u['onPremisesDomainName']
                OnPremisesSecurityIdentifier=[string]$u['onPremisesSecurityIdentifier']; AccountEnabled=[string]$u['accountEnabled']
                UsageLocation=[string]$u['usageLocation']; AssignedPlans=$plans; AssignedSkuIds=$licenses
                InventoryTenantId=$tenant; InventoryScope='AllUsers'; InventoryComplete='True'; CollectedUtc=[datetime]::UtcNow.ToString('o')
            })
        }
        Write-Host ("Users and licenses: {0} user(s) read across {1} page(s)." -f $rows.Count,$seenPages.Count)
        $next = [string]$page['@odata.nextLink']
    }
    if ($rows.Count -eq 0) { throw 'Empty user inventory; existing cache was retained.' }
    $parent = Split-Path -Parent ([IO.Path]::GetFullPath($OutputPath))
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
    $temp = Join-Path $parent ('.semt-users-' + [guid]::NewGuid().ToString('N') + '.tmp')
    try {
        $rows | Export-Csv -LiteralPath $temp -NoTypeInformation -Encoding UTF8
        $target = [IO.Path]::GetFullPath($OutputPath)
        if ([IO.File]::Exists($target)) { [IO.File]::Replace($temp, $target, $null) }
        else { [IO.File]::Move($temp, $target) }
    } finally { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
    Write-Host "User inventory exported: $($rows.Count) users; $OutputPath"
} finally {
    if (-not $NoConnect) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBP4Kg9u7+4P+Fm
# A/zdpC18vszudYIu/Uf61rcwURcojaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDt2FbYpif3TY9VUPtusgoe
# m8Of0s/nxKgW88O9yzk70jANBgkqhkiG9w0BAQEFAASCAYBOnu/BZZtCB9Jl6OZP
# DqaibB/Clzii4hjwmgXxwR+OBQmGglAxnmiPjQMA1nNPiQ6YGPEhF5s73wzvSWuT
# AlG79wnr73+UejW4KgA760EyU9xJS43CT+0OSfKUNXj6Hen4P8An9ASRImqgPLQp
# Ag4fRi85/C8Uk156l9+kGYXSkGsTPAOTwi5Pq11UykofsVeZiItnA3cU1nFB4mU+
# WX/R7LIp4xddhpA5P//lR/JcxX2PfO+VYYnzvewhxbSUJXWRw0ylkV6L/SKACZ2Q
# UAqRm/MwsXVv6mgX0ILQb2EECL0qgNyDV143KDfZdlHCWyOFDd9yvDpUNeq26xPy
# Km2exc50pR6VSR7K5PmwVxsH673IRbZ7S2NAG7KypRhQGhqsvnogs4m7PJotPrCE
# EBWis2XsFGaqzEkLhPOHX2fD5VfhUaF8gwmQlMMc4DnSIRIU9dHpYGEQZaQW9ThN
# xmUs2I3iy+0xORibh3vG9KkuuyAkst9AvUIVYZ+8dqFySqQ=
# SIG # End signature block
