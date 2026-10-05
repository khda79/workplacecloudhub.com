<#
.SYNOPSIS
    Exercise real XLSX conversion and snapshot provenance with synthetic reports.
.VERSION
    1.0.0
#>
#Requires -Version 7.4
$ErrorActionPreference='Stop'
if (-not (Get-Module -ListAvailable ImportExcel)) { throw 'This offline test requires ImportExcel already installed; it does not install modules.' }
Import-Module ImportExcel
$source=Join-Path $PSScriptRoot '..'
$tempRoot=[IO.Path]::GetFullPath((Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-WorkbookTest-'+[guid]::NewGuid().ToString('N'))))
if (-not $tempRoot.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe temporary test root.' }
try {
    foreach($relative in @(
        'Scripts/Diagnostics/SmartM365-SharePointMigration-Diagnostics.ps1',
        'Scripts/Diagnostics/SmartM365-SharePointMigration-ImportExcel.ps1',
        'Scripts/Diagnostics/analyze_sharegate_reports.py',
        'Scripts/Launchers/SmartM365-SharePointMigration-ConsoleLifecycle.ps1',
        'Scripts/console_lifecycle.py','Scripts/Compare/report_html.py',
        'Config/sharegate-diagnostics.columns.json.template','Config/sharegate-diagnostics.rules.json.template'
    )) {
        $target=Join-Path $tempRoot $relative
        [void](New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force)
        Copy-Item -LiteralPath (Join-Path $source $relative) -Destination $target
    }
    $project=Join-Path $tempRoot 'Project'
    [void](New-Item -ItemType Directory -Path $project)
    $book=Join-Path $project 'report.xlsx'
    [pscustomobject]@{
        'Session ID'='test-session'; ID='1'; Date='2026-10-05'; Status='Success'; Type='File'; Title='fixture.txt'
        'Source site address'='https://source.example.test/sites/a'; 'Source list title'='Docs'; 'Source ID'='1'; Messages='Copied'
        'Destination site address'='https://target.example.test/sites/a'; 'Destination list title'='Docs'
    } | Export-Excel -Path $book -WorksheetName Data
    $diagnostics=Join-Path $tempRoot 'Scripts/Diagnostics/SmartM365-SharePointMigration-Diagnostics.ps1'
    $output=Join-Path $project 'output'
    $messages=@(& (Get-Command pwsh).Source -NoProfile -File $diagnostics -ProjectRoot $project -InputPath $book -OutputDirectory $output 2>&1)
    if ($LASTEXITCODE -ne 0) { throw ("XLSX analysis failed: " + ($messages -join "`n")) }
    $summary=Get-Content (Join-Path $output 'Summary.json.txt') -Raw | ConvertFrom-Json
    if ($summary.Lines -ne 1 -or $summary.InputEvidence[0].Path -ne $book -or
        $summary.InputEvidence[0].Sha256 -ne (Get-FileHash $book).Hash.ToLowerInvariant() -or
        -not (Test-Path -LiteralPath $summary.ReportPath) -or
        @(Get-ChildItem -LiteralPath $output -Filter '.converted-*').Count) { throw 'Snapshot conversion lost rows, original provenance or temporary cleanup.' }
    $messages=@(& (Get-Command pwsh).Source -NoProfile -File $diagnostics -ProjectRoot $project -InputPath $book -DryRun 2>&1)
    if ($LASTEXITCODE -ne 0 -or -not ($messages -match 'used CSV=0, XLSX=1')) { throw ("XLSX DryRun failed: " + ($messages -join "`n")) }
    'PASS: real ImportExcel conversion, original SHA256/path, HTML output, temporary cleanup and XLSX DryRun.'
}
finally {
    if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDgnUi4MN4hR8gO
# +l+9kVo9YVirQV1T17XasZShOjelnKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBYdH4y4ge/Yo+FBYHGvF5O
# vNKIWr3OmuDv4RJhZdvbdTANBgkqhkiG9w0BAQEFAASCAYAYk4rWjlsxGo3NhzMq
# ifb3vmL3HU8r6Z+ZJdcF8mZirkPncwT2MFUeJhS0KgaAilr9Wg9tSlpK+VyyyhEA
# SjpvfnEbB2VO9ycci/lPsWAhieZJlY/fM7Jcteb2mp0DThT+tYpv4K/6o1a8NgLR
# n1rty3W0CagRzNaNF5tRYq1ucx5I45dlt7+Nf6fAHuLeUOTVNK3aX1Iv7smZlDMB
# JqqQj7pCNvBJndJFt5ycjwFyUzTa7tCZz3KwU6R3aNnZUvF/8t5Nd6thZ/HH3Bxb
# pMB01lcGBScFao+WRK66GnzodRYPibkSbVRCZTNxrAEymkM3tRftCpw04kv7qxcR
# GPfzxR2S7uvGkK5miBAH50szgB7bP1uWKDZIPkvSWeMUAeJRhN8L807JJ3yLz027
# P4R5KCKLU0K5z14uRklUvtPgPRPwYFJJOx5aWQZbYobrozb+VhIkvPGn6Ihk7MAS
# 0py1zqL7fiFL0zWH0s0Kn3CMH5gCcFU6c0tK7P9096J9pHw=
# SIG # End signature block
