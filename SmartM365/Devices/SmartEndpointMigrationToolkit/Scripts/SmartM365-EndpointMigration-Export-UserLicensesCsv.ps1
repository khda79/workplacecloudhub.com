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

# Cache provenance is carried in each row; an empty or partial export is never trusted.
if (-not $ForceRefresh -and $TenantId -and (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
    $age = ((Get-Date).ToUniversalTime() - (Get-Item -LiteralPath $OutputPath).LastWriteTimeUtc).TotalHours
    if ($age -ge 0 -and $age -lt 24) {
        $first = Import-Csv -LiteralPath $OutputPath | Select-Object -First 1
        if ($first -and $first.InventoryTenantId -eq $TenantId -and $first.InventoryScope -eq 'AllUsers' -and $first.InventoryComplete -eq 'True') {
            Write-Host "Using verified user inventory: $OutputPath"
            return
        }
    }
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCf9ZByf9N8m1t3
# q1XndZsfajfVGZjyeWcbgon041+N3qCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCjq+mjVhdSCuNNxjts5scd
# Y3UDCUKZawyTfl8N+hZi/zANBgkqhkiG9w0BAQEFAASCAYBkHDLUXl8yyDwEVGSm
# 35+JrkdSvxfhv7PZRPKiBM3OL8FbPbTEitiq17bI1CEDVtKDcxDY5a7wkLiJ2zOj
# f7xKDhuTBsecbLoNS1F3fEImpMzVFtUXYcGgkQCXI8Miyu6KVwS6qy/EHJFgGf4W
# Mmmjv0dBwa3BW+BG1lzITMs0njJRMItsuH8fmKTOZmW7tDqzGxFharm9NYrToNZM
# DoXDoy9rOLQ4pfubBh7rOl8HxaUGwe5Nk+boxJDpXPuqUtPaSddHhm0I20mzefNK
# Asju2YmnR5RjeVbB8at6MU2Fn0JYUJbwrwq/oa+oWUCYxkBVoslCHBh2uJtvwWaV
# GcVE6LWe45LOuLVr1RYMBzwlmh1ipfAxDvutqxE5rCoRBlW8qrUP3TfJhwHM/OmP
# 4b8sKmp5zF5PSwvPfLG4vfQ/SRYaFOwj3G6ALwmaxlmK5Mm4WVsh2m07YN8LUmWx
# eXQQuAV1OAIpDX6sPqj4XOpr3Oi4wMKkq/LImKgy5K4K3o8=
# SIG # End signature block
