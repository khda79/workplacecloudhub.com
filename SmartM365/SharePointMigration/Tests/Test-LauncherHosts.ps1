<#
.SYNOPSIS
    Tests real child-process host routing with mock inventory files only.
.VERSION
    1.0.0
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $PSScriptRoot
$root = Join-Path ([IO.Path]::GetTempPath()) ('SharePointMigration-hosts-' + [guid]::NewGuid().ToString('N'))
$name = "Folder O'Brien"
try {
    foreach ($dir in @('Scripts/Launchers/Generic', 'Scripts/Inventory', 'Scripts/Compare', "Migrations/$name")) {
        New-Item -ItemType Directory -Path (Join-Path $root $dir) -Force | Out-Null
    }
    $launcher = Join-Path $root 'Scripts/Launchers/Generic/SmartM365-SharePointMigration-Launcher.ps1'
    Copy-Item (Join-Path $project 'Scripts/Launchers/Generic/SmartM365-SharePointMigration-Launcher.ps1') $launcher
    Copy-Item (Join-Path $project 'Scripts/Launchers/SmartM365-SharePointMigration-LauncherCommon.ps1') (Join-Path $root 'Scripts/Launchers')
    Copy-Item (Join-Path $project 'Scripts/Compare/scan_evidence.py') (Join-Path $root 'Scripts/Compare')
    $fixture = @'
[CmdletBinding()]
param($OutputPath, $LogPath, $SiteUrl, $WebUrlsFile, $DocumentLibrariesOnly, $IncludeItemPermissions, $ItemProgressInterval, $Interactive, $ForceAuthentication)
if ($env:SPMIG_TEST_FAIL -eq '1') { throw 'Synthetic failure' }
[IO.File]::WriteAllText($OutputPath, "Host`n$($PSVersionTable.PSVersion.Major)`n")
'@
    foreach ($side in @('Source', 'Target')) {
        foreach ($kind in @('File', 'Permission')) {
            [IO.File]::WriteAllText((Join-Path $root "Scripts/Inventory/SmartM365-SharePoint$side-$($kind)Inventory.ps1"), $fixture)
        }
    }
    $count = 0
    foreach ($type in @('SP2016', 'SP2019', 'SPO')) {
        $config = @"
@{
 Name = 'ReportLabel'
 Source = @{ Type = '$type'; SiteUrl = 'https://workplacecloudhub.sharepoint.com/source'; UrlsFile = 'urls.txt' }
 Target = @{ Type = '$type'; SiteUrl = 'https://workplacecloudhub.sharepoint.com/target'; UrlsFile = 'urls.txt' }
 Output = @{ SourceFileScans = 'sf'; TargetFileScans = 'tf'; SourcePermissionScans = 'sp'; TargetPermissionScans = 'tp'; Logs = 'logs' }
 Permissions = @{ IncludeItemPermissions = `$true; ItemProgressInterval = 500 }
}
"@
        [IO.File]::WriteAllText((Join-Path $root "Migrations/$name/migration.config.psd1"), $config)
        [IO.File]::WriteAllText((Join-Path $root "Migrations/$name/urls.txt"), 'https://workplacecloudhub.sharepoint.com/example')
        foreach ($action in @('ScanSourceFiles', 'ScanTargetFiles', 'ScanSourcePermissions', 'ScanTargetPermissions')) {
            $result = & pwsh -NoProfile -ExecutionPolicy Bypass -File $launcher -MigrationName $name -Action $action -NonInteractive -ForceAuthentication:$false 2>&1
            if ($LASTEXITCODE -ne 0) { throw ($result | Out-String) }
            $consoleText = $result | Out-String
            if ($consoleText -notmatch 'SmartM365 by WorkplaceCloudHub' -or $consoleText -notmatch 'Status\s+: SUCCESS' -or $consoleText -notmatch ('Action\s+: {0}' -f $action)) {
                throw "Launcher intro or success summary missing for $type $action"
            }
            $latest = Get-ChildItem (Join-Path $root "Migrations/$name") -Recurse -Filter '*.csv' | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
            $expected = if ($type -eq 'SPO') { '7' } else { '5' }
            if (-not $latest -or [IO.File]::ReadAllText($latest.FullName) -ne "Host`n$expected`n") { throw "Wrong host for $type $action" }
            if (-not (Test-Path -LiteralPath ($latest.FullName + '.manifest.json.txt') -PathType Leaf)) { throw "Scan manifest missing for $type $action" }
            $count++
        }
    }
    # A failed comparison must end with a visible status and persist that status in its run log.
    $compareConfig = @"
@{
 Name = 'ReportLabel'
 Source = @{ Type = 'SPO'; SiteUrl = 'https://workplacecloudhub.sharepoint.com/source'; UrlsFile = 'urls.txt' }
 Target = @{ Type = 'SPO'; SiteUrl = 'https://workplacecloudhub.sharepoint.com/target'; UrlsFile = 'urls.txt' }
 Output = @{ SourceFileScans = 'sf'; TargetFileScans = 'tf'; FileComparisons = 'cf'; Logs = 'logs' }
 Comparison = @{ MaxScanAgeDifferenceHours = 12; ModifiedDateToleranceMinutes = 0; SizeToleranceBytes = 0; ShareGateReplacementCharacter = '_' }
}
"@
    [IO.File]::WriteAllText((Join-Path $root "Migrations/$name/migration.config.psd1"), $compareConfig)
    $missingCsv = Join-Path $root 'missing.csv'
    $result = & pwsh -NoProfile -ExecutionPolicy Bypass -File $launcher -MigrationName $name -Action CompareFiles -SourceCsv $missingCsv -TargetCsv $missingCsv -NonInteractive 2>&1
    if ($LASTEXITCODE -eq 0 -or ($result | Out-String) -notmatch 'Status\s+: FAILED') { throw 'Failed comparison did not return a failure summary.' }
    $compareLog = Get-ChildItem (Join-Path $root "Migrations/$name/logs") -Filter 'ReportLabel-CompareFiles-*.log' | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
    if (-not $compareLog -or (Get-Content -LiteralPath $compareLog.FullName -Raw) -notmatch 'Status\s+: FAILED') { throw 'Failed comparison summary was not appended to its run log.' }

    # Verify failure propagation across the PowerShell 7 -> 5.1 boundary.
    $config = $config.Replace("Type = 'SPO'", "Type = 'SP2019'")
    [IO.File]::WriteAllText((Join-Path $root "Migrations/$name/migration.config.psd1"), $config)
    $env:SPMIG_TEST_FAIL = '1'
    $result = & pwsh -NoProfile -ExecutionPolicy Bypass -File $launcher -MigrationName $name -Action ScanSourceFiles -NonInteractive 2>&1
    if ($LASTEXITCODE -eq 0) { throw 'Child failure was reported as success.' }
    $consoleText = $result | Out-String
    if ($consoleText -notmatch 'Status\s+: FAILED' -or $consoleText -notmatch 'Reason\s+: Synthetic failure') {
        throw 'Launcher failure summary missing or misleading.'
    }
    Write-Output "PASS: $count host-routing cases and child failure propagation; only mock inventories executed."
}
finally {
    Remove-Item Env:SPMIG_TEST_FAIL -ErrorAction SilentlyContinue
    $resolved = [IO.Path]::GetFullPath($root)
    $allowed = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if ($resolved.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'SharePointMigration-hosts-*') {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB4cUP4n6pY2psF
# Cbvr0cOAcj2Kx3IULoPHdontl5vWxqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAzzXGCZtLFGaGrbdKRIE5I
# WVbfkppEXk88JN82J4pLszANBgkqhkiG9w0BAQEFAASCAYBfqZT3d58WJ9KEZydD
# +D4IprysKyZWmmMsWBBVzPidJ4ZjjHdkqR7W0wYMVDEol9xHnQhinVVS2kv7sl2y
# 5PFgW0lzJE1uBVbTT6wsR6VIYkBCErOtSAWKRPX/t1K4qTOn6D9kx086UxrjQ+ev
# fCsCjpgcfyuXoqZuBCU1FgNHZYIkAT9JoHTna15waqlkajwUDBwgb/s7+CdqcuOf
# Ny7iPSwjtozyCFtxTUJOw9Gf/16ujNzxshr6xl7MV0BfDDnX/JT2aiEJMVo8nXnn
# tzYcbw5ipTesK3xTyWR/zQ7syDqCPRllgv2r6pYpWqHz78YXw9sXpNNl10df0PLv
# YJzt3cXQDzCmhJNQJdZheB7BAlGDNCQUio4kiPkZtDRamqEwCrqHzC0qcgzPNC75
# w8QA1q6ud8vt7boRes5ARTXNscQmcDTioqTwdaEqu9ooVouRCn/vxC5nOPSnKnik
# Lt3BoAY1P0SdgZYTU1DKC30SOgCHJaw9B1OSlin9fRhP7u0=
# SIG # End signature block
