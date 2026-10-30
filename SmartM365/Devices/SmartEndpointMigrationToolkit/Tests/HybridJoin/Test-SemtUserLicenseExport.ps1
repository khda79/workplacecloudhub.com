<#
.VERSION
1.0.0
#>
#requires -Version 5.1
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SEMT-users-test-' + [guid]::NewGuid().ToString('N'))
$oldModulePath = $env:PSModulePath
try {
    $moduleDir = Join-Path $testRoot 'Modules\Microsoft.Graph.Authentication'
    New-Item -ItemType Directory -Path $moduleDir -Force | Out-Null
    $moduleSource = @'
function Get-MgContext { @{ AuthType='Delegated'; TenantId=$(if($env:SEMT_TEST_TENANT_ID){$env:SEMT_TEST_TENANT_ID}else{'22222222-2222-2222-2222-222222222222'}) } }
function Invoke-MgGraphRequest {
    param($Method,$Uri,$Headers)
    if ($Uri -like '*/licenseDetails') { return @{ value=@(@{skuId='sku1';servicePlans=@()}) } }
    if ($Uri -like '*$filter=*onPremisesSecurityIdentifier*') { return @{ value=@(@{ id='u1'; userPrincipalName='u1@example.test'; onPremisesSecurityIdentifier='S-1-5-21-42' }) } }
    if ($Uri -like '*page2*' -and $env:SEMT_TEST_FAIL_PAGE -eq '1') { throw 'simulated second page failure' }
    $id = if ($Uri -like '*page2*') { 'u2' } else { 'u1' }
    $user = @{ id=$id; userPrincipalName="$id@example.test"; onPremisesSecurityIdentifier="S-1-5-21-$id"; assignedPlans=@(); assignedLicenses=@() }
    $result = @{ value=@($user) }
    if ($id -eq 'u1') { $result['@odata.nextLink']='https://graph.microsoft.com/v1.0/users?page2' }
    return $result
}
Export-ModuleMember -Function Get-MgContext,Invoke-MgGraphRequest
'@
    Set-Content -LiteralPath (Join-Path $moduleDir 'Microsoft.Graph.Authentication.psm1') -Value $moduleSource -Encoding UTF8
    New-ModuleManifest -Path (Join-Path $moduleDir 'Microsoft.Graph.Authentication.psd1') -RootModule 'Microsoft.Graph.Authentication.psm1' -ModuleVersion '999.0.0' -FunctionsToExport @('Get-MgContext','Invoke-MgGraphRequest')
    $env:PSModulePath = (Join-Path $testRoot 'Modules') + ';' + $oldModulePath
    $output = Join-Path $testRoot 'Users.csv'
    $exporter = Join-Path $root 'Scripts\SmartM365-EndpointMigration-Export-UserLicensesCsv.ps1'
    & $exporter -NoConnect -TenantId '22222222-2222-2222-2222-222222222222' -OutputPath $output -ForceRefresh
    $rows = @(Import-Csv -LiteralPath $output)
    if ($rows.Count -ne 2 -or $rows[0].InventoryComplete -ne 'True') { throw 'Paged export incomplete.' }
    Write-Host 'PASS: all user pages exported with provenance'
    $targeted = Join-Path $testRoot 'Targeted.json'
    & (Join-Path $root 'Scripts\SmartM365-EndpointMigration-Get-UserLicense.ps1') -UserSid 'S-1-5-21-42' -TenantId '22222222-2222-2222-2222-222222222222' -OutputPath $targeted -NoConnect
    $targetData = Get-Content -LiteralPath $targeted -Raw | ConvertFrom-Json
    if ($targetData.User.id -ne 'u1' -or @($targetData.LicenseDetails).Count -ne 1) { throw 'Targeted lookup returned incorrect identity or licenses.' }
    Write-Host 'PASS: targeted Graph lookup matches SID and reads license details'
    $before = (Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash
    $env:SEMT_TEST_FAIL_PAGE = '1'
    & $exporter -NoConnect -TenantId '22222222-2222-2222-2222-222222222222' -OutputPath $output
    if ((Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash -ne $before) { throw 'Fresh cache was changed.' }
    Write-Host 'PASS: fresh user cache reused'
    & $exporter -NoConnect -OutputPath $output
    if ((Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash -ne $before) { throw 'Same-tenant context did not reuse cache.' }
    Write-Host 'PASS: a fresh cache is reused after confirming the signed-in tenant'
    $env:SEMT_TEST_TENANT_ID='33333333-3333-3333-3333-333333333333'
    $crossTenantFailed=$false
    try { & $exporter -NoConnect -OutputPath $output } catch { $crossTenantFailed=$true }
    Remove-Item Env:\SEMT_TEST_TENANT_ID -ErrorAction SilentlyContinue
    if (-not $crossTenantFailed -or (Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash -ne $before) { throw 'Cache was reused across tenants or replaced after a failed refresh.' }
    Write-Host 'PASS: a different tenant does not reuse or damage the existing cache'
    $staleEvidence=Join-Path $testRoot 'StaleEvidence.csv'
    $staleRows=@(Import-Csv -LiteralPath $output)
    foreach ($entry in $staleRows) { $entry.CollectedUtc=[datetime]::UtcNow.AddHours(-25).ToString('o') }
    $staleRows | Export-Csv -LiteralPath $staleEvidence -NoTypeInformation -Encoding UTF8
    $staleHash=(Get-FileHash -LiteralPath $staleEvidence -Algorithm SHA256).Hash
    $staleFailed=$false
    try { & $exporter -NoConnect -OutputPath $staleEvidence } catch { $staleFailed=$true }
    if (-not $staleFailed -or (Get-FileHash -LiteralPath $staleEvidence -Algorithm SHA256).Hash -ne $staleHash) { throw 'Stale collection metadata was accepted or overwritten after a failed refresh.' }
    Write-Host 'PASS: recently touched files do not make stale user evidence fresh'
    (Get-Item -LiteralPath $output).LastWriteTimeUtc = [datetime]::UtcNow.AddHours(-25)
    $failed = $false
    try { & $exporter -NoConnect -TenantId '22222222-2222-2222-2222-222222222222' -OutputPath $output } catch { $failed = $true }
    if (-not $failed -or (Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash -ne $before) { throw 'Failed refresh did not preserve cache.' }
    Write-Host 'PASS: stale cache refresh fails closed and preserves previous file'
} finally {
    $env:PSModulePath = $oldModulePath
    Remove-Item Env:\SEMT_TEST_FAIL_PAGE -ErrorAction SilentlyContinue
    Remove-Item Env:\SEMT_TEST_TENANT_ID -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB+BCuyueHNDi9Y
# iiI2pN9KgXVMxn2um2qr0cUK7UGdwKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAESOuy2Vq9uFJzpPe0IXX0
# C1M8X4KvrmOSf2MKkTFUQTANBgkqhkiG9w0BAQEFAASCAYArhJnyJdhOLPvoRgOo
# TnXBXYW5rPTw6S2kBQRZDVud18aziPuRlex3/FogRHtpDS479aMQi/NZ/UbADsh3
# zOAbsLwztpKpmNDrDv1kK9i/CddZtv4pkTvr5+ByzJ+zxjqnASwXA0yYLfkuDvga
# KkYh6OldBwN9isafrx12iNziQBoKURXtra6LQ1eSegioyYhXPr5k8AmCO1g34FMy
# XW6p+3TNEaobbfZtQUSciwQhWuPD9wSzny9aTqRkEWilgwCNdsUAQAVOW1cOWSnd
# jeqCaXtTau+XvSlDzoxFQHT5XfsUJOURsN1VAtZPznsU+MUAuk0vghXCKEYMgdK+
# XoemtiY4mKTTq2jDWUjVg9jG1j5QrjPysqa5e2uZ1y3H0fHXt2qjstyeeIpB+IN4
# LGKN8y6RkMecUmTSv/4z1Nbb2bhmdAJT7ma/dd7MORozrdovRJxVAcms9icrp5Ep
# 0EaeaUW0eDeqZsULJ5aIws79Yn5uXZhi4hOIn1g0bzrR/fg=
# SIG # End signature block
