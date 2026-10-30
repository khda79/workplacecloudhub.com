<#
.SYNOPSIS
Offline test of the user-license CMD launcher using a synthetic exporter.
.VERSION
1.0.0
#>
#requires -Version 5.1
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SEMT launcher ' + [guid]::NewGuid().ToString('N'))
$priorPause = $env:SEMT_NO_PAUSE
$priorRecord = $env:SEMT_TEST_LAUNCHER_RECORD
$priorExit = $env:SEMT_TEST_LAUNCHER_EXIT
try {
    New-Item -ItemType Directory -Path (Join-Path $testRoot 'Scripts') -Force | Out-Null
    $launcher = Join-Path $testRoot 'Export-UserLicensesCsv.cmd'
    Copy-Item -LiteralPath (Join-Path $root 'Export-UserLicensesCsv.cmd') -Destination $launcher
    $stub = Join-Path $testRoot 'Scripts\SmartM365-EndpointMigration-Export-UserLicensesCsv.ps1'
    @'
param([string]$TenantId,[switch]$ForceRefresh)
[pscustomobject]@{ Tenant=$TenantId; Force=[bool]$ForceRefresh; WorkingDirectory=(Get-Location).Path } |
    ConvertTo-Json | Set-Content -LiteralPath $env:SEMT_TEST_LAUNCHER_RECORD -Encoding UTF8
exit ([int]$env:SEMT_TEST_LAUNCHER_EXIT)
'@ | Set-Content -LiteralPath $stub -Encoding UTF8
    $env:SEMT_NO_PAUSE='1'
    $env:SEMT_TEST_LAUNCHER_RECORD=Join-Path $testRoot 'result.json.txt'
    $env:SEMT_TEST_LAUNCHER_EXIT='0'
    $command='"' + $launcher + '" -TenantId 22222222-2222-2222-2222-222222222222 -ForceRefresh'
    & $env:ComSpec /d /c $command
    if ($LASTEXITCODE -ne 0) { throw "Launcher failed: $LASTEXITCODE" }
    $result=Get-Content -LiteralPath $env:SEMT_TEST_LAUNCHER_RECORD -Raw | ConvertFrom-Json
    if ($result.Tenant -ne '22222222-2222-2222-2222-222222222222' -or -not $result.Force -or $result.WorkingDirectory -ine $testRoot) { throw 'Launcher arguments or working directory were not preserved.' }
    Write-Host 'PASS: launcher forwards arguments and works from a path with spaces'
    $env:SEMT_TEST_LAUNCHER_EXIT='7'
    & $env:ComSpec /d /c $command
    if ($LASTEXITCODE -ne 7) { throw "Launcher did not preserve exporter exit code: $LASTEXITCODE" }
    Write-Host 'PASS: launcher preserves the exporter exit code'
    Remove-Item -LiteralPath $stub -Force
    & $env:ComSpec /d /c $command
    if ($LASTEXITCODE -eq 0) { throw 'Missing exporter was incorrectly reported as success.' }
    Write-Host 'PASS: launcher reports a missing exporter'
} finally {
    $env:SEMT_NO_PAUSE=$priorPause
    $env:SEMT_TEST_LAUNCHER_RECORD=$priorRecord
    $env:SEMT_TEST_LAUNCHER_EXIT=$priorExit
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAy/3b9Z5IcBbNP
# U5iFJz2YbTEiQOmnEbt7mWMcOr5LDqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAH2rdrnD0XoxVnXPvb/4zK
# leWQxkOx84d2FYNASCBLODANBgkqhkiG9w0BAQEFAASCAYAaQXw5El/EpHkY1GCP
# 8+PGyEGzb32ol2/OkrIoCWiZ8xV5PtZsQE5639JSZ+wAXOxeK9KqMoICJFHroSNJ
# FxYfeyn6aPRdx91ZX5xb+NCz3vK1qQbIfnNd6cyfpXoZngn6/KAjHpB6pgPg3SYy
# dqN7unjcGXhObHj6Wv4h3cIe0PPJBkNwJAIFNTeA4B0QGKXVJu/vyfEnOD0ZZjfv
# MJjrGt1mYNyACQV72DApYLiFdthVxKYjtBNkOG5Qu8r9gwBre/5XszZ6Iys5UtUY
# sovtSUIC0MiNcptCSd4qSOWTksD6N+u3jrGGbTqwiqCRwGubT6ZTw7E3EuNV3aXn
# 7VpN4MyIZHhhSt4MdhleHIHQZHbnGFb9vJO3nssDqqf1YLRYZTCxinQhXlzV8Afn
# FP719oWP97+jRktsiNzIfiMP9ZIkCrE6/Sra4M0QSglvF2qFFAj+iPpNXxPMixbQ
# nTqy43QrwB3yoHRMHdQ8MJJtFLOgfj4GBUbaHNMmSMOOLu8=
# SIG # End signature block
