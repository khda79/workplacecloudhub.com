<#
.SYNOPSIS
    Tests real child-process host routing with mock inventory files only.
.VERSION
    1.0.4
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
    Copy-Item (Join-Path $project 'Scripts/Launchers/SmartM365-SharePointMigration-Storage.ps1') (Join-Path $root 'Scripts/Launchers')
    Copy-Item (Join-Path $project 'Scripts/console_lifecycle.py') (Join-Path $root 'Scripts')
    Copy-Item (Join-Path $project 'Scripts/Compare/scan_evidence.py') (Join-Path $root 'Scripts/Compare')
    $fixture = @'
[CmdletBinding()]
param($OutputPath, $LogPath, $SiteUrl, $WebUrlsFile, $DocumentLibrariesOnly, $IncludeItemPermissions, $ItemProgressInterval, $Interactive, $DeviceLogin, $Thumbprint, $ForceAuthentication, $AssessmentRulesDirectory, $ClientId, $Tenant, $TenantId, $IncludeVersionCounts)
if ($env:SPMIG_TEST_FAIL -eq '1') { throw 'Synthetic failure' }
[IO.File]::WriteAllText($OutputPath, "Host`n$($PSVersionTable.PSVersion.Major)`n")
if ($ClientId) { [IO.File]::WriteAllText(($OutputPath + '.auth-test'), "$ClientId|$Tenant|$TenantId|$([bool]$IncludeVersionCounts)|$([bool]$DeviceLogin)|$([bool]$Interactive)|$Thumbprint") }
'@
    foreach ($side in @('Source', 'Target')) {
        foreach ($kind in @('File', 'Permission')) {
            [IO.File]::WriteAllText((Join-Path $root "Scripts/Inventory/SmartM365-SharePoint$side-$($kind)Inventory.ps1"), $fixture)
        }
    }
    $count = 0
    foreach ($type in @('SP2016', 'SP2019', 'SPO')) {
        $sourceAuth = if ($type -eq 'SPO') { "; AuthConfigPath = 'source.local.psd1'; IncludeVersionCounts = `$true; AuthMode = 'DeviceLogin'" } else { '' }
        $targetAuth = if ($type -eq 'SPO') { "; AuthConfigPath = 'target.local.psd1'; AuthMode = 'Certificate'" } else { '' }
        $config = @"
@{
 Name = 'ReportLabel'
 Source = @{ Type = '$type'; SiteUrl = 'https://workplacecloudhub.sharepoint.com/source'; UrlsFile = 'urls.txt' $sourceAuth }
 Target = @{ Type = '$type'; SiteUrl = 'https://workplacecloudhub.sharepoint.com/target'; UrlsFile = 'urls.txt' $targetAuth }
 Output = @{ SourceFileScans = 'sf'; TargetFileScans = 'tf'; SourcePermissionScans = 'sp'; TargetPermissionScans = 'tp'; Logs = 'logs' }
 Permissions = @{ IncludeItemPermissions = `$true; ItemProgressInterval = 500 }
}
"@
        [IO.File]::WriteAllText((Join-Path $root "Migrations/$name/migration.config.psd1"), $config)
        [IO.File]::WriteAllText((Join-Path $root "Migrations/$name/urls.txt"), 'https://workplacecloudhub.sharepoint.com/example')
        [IO.File]::WriteAllText((Join-Path $root "Migrations/$name/source.local.psd1"), "@{ ClientId='source-app'; Tenant='source.example'; TenantId='source-id' }")
        [IO.File]::WriteAllText((Join-Path $root "Migrations/$name/target.local.psd1"), "@{ ClientId='target-app'; Tenant='target.example'; TenantId='target-id'; Thumbprint='target-thumb' }")
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
            if ($type -eq 'SPO') {
                $authSide = if ($action -like '*Source*') { 'source' } else { 'target' }
                $expectedVersions = $action -eq 'ScanSourceFiles'
                $expectedDevice = $authSide -eq 'source'
                $expectedInteractive = $false
                $expectedThumb = if ($authSide -eq 'target') { 'target-thumb' } else { '' }
                $expectedAuth = "$authSide-app|$authSide.example|$authSide-id|$expectedVersions|$expectedDevice|$expectedInteractive|$expectedThumb"
                if ([IO.File]::ReadAllText($latest.FullName + '.auth-test') -ne $expectedAuth) { throw "Wrong authentication profile for $action" }
            }
            if (-not (Test-Path -LiteralPath ($latest.FullName + '.manifest.json.txt') -PathType Leaf)) { throw "Scan manifest missing for $type $action" }
            $count++
        }
    }
    # An explicitly selected profile must never silently fall back to another tenant.
    $missingProfileConfig = $config.Replace("source.local.psd1", "missing.local.psd1")
    [IO.File]::WriteAllText((Join-Path $root "Migrations/$name/migration.config.psd1"), $missingProfileConfig)
    $result = & pwsh -NoProfile -ExecutionPolicy Bypass -File $launcher -MigrationName $name -Action ScanSourceFiles -NonInteractive -ForceAuthentication:$false 2>&1
    if ($LASTEXITCODE -eq 0 -or ($result | Out-String) -notmatch 'Source AuthConfigPath does not exist') {
        throw 'A missing explicit source authentication profile did not fail closed.'
    }
    [IO.File]::WriteAllText((Join-Path $root "Migrations/$name/migration.config.psd1"), $config)
    $result = & pwsh -NoProfile -ExecutionPolicy Bypass -File $launcher -MigrationName $name -Action ComparePermissions -NonInteractive 2>&1
    if ($LASTEXITCODE -eq 0 -or ($result | Out-String) -notmatch 'Cross-tenant permission comparison requires') {
        throw 'A cross-tenant permission comparison was not rejected.'
    }
    $otherHostConfig = $config.Replace('https://workplacecloudhub.sharepoint.com/target', 'https://other.sharepoint.com/target').Replace("; AuthConfigPath = 'source.local.psd1'", '').Replace("; AuthConfigPath = 'target.local.psd1'", '')
    [IO.File]::WriteAllText((Join-Path $root "Migrations/$name/migration.config.psd1"), $otherHostConfig)
    $result = & pwsh -NoProfile -ExecutionPolicy Bypass -File $launcher -MigrationName $name -Action ComparePermissions -NonInteractive 2>&1
    if ($LASTEXITCODE -eq 0 -or ($result | Out-String) -notmatch 'Cross-tenant permission comparison requires') { throw 'Different SPO hosts were not rejected for the shared authentication profile.' }
    [IO.File]::WriteAllText((Join-Path $root "Migrations/$name/migration.config.psd1"), $config)
    [IO.File]::WriteAllText((Join-Path $root "Migrations/$name/target.local.psd1"), "@{ ClientId='target-app'; Tenant='source.example'; TenantId='source-id' }")
    $result = & pwsh -NoProfile -ExecutionPolicy Bypass -File $launcher -MigrationName $name -Action ComparePermissions -NonInteractive 2>&1
    if (($result | Out-String) -match 'Cross-tenant permission comparison requires') {
        throw 'Same-tenant profiles were incorrectly rejected as cross-tenant.'
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBGETi+VGfWZn6q
# 9vTfMHK3t9j50HKG4hA9UgmsNtfCaaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIPvj1uer3uovHaBqHHwsSyX73YsLEKWYZ9bh6x9GeZaKMA0GCSqG
# SIb3DQEBAQUABIIBgFDaLZU4iDrTp45th7mFmn38EZ6xY6i2EnHwLqNoiC+f3vYI
# uRXiIJ0M3oI+8IVi5g9oIbF6UJEcsOM1MmsNoS3t9upFKUuamkxhOMaAE3SyrmhK
# TFb2pVfE2ybG4UOa2+uLJjlU+UcVAF+B1quu8Tb5p1BV1D6HXHZnmCK2zaY3OvA9
# Nx7h0a/OTiqnbvQrC/DB9Zg4IjQ32JneZYRMMHqeuHn/yjqrJdOFsT89toIxcUYy
# 4+zl2KLNGnzF19wL5DaOHe7N6Vd/5GHyLJQp6vnaaqscoFmufT0kn5eX7DzhYCEZ
# 2GdmHsWZgnU/LFOwTfDSxECSGF3qi2knEwSTb9yVmeYGbO5ClFr1pyYEExvULUTe
# HhNfCnygw9HTw4IVEI5JjRI4/S2Do1DYFa31Z73D8yiJXwlPjW/uRXjMq8Bb58iR
# ihJ7iKZ6uzYWbImx74cw5xYzN3zc5Xl1KQBaWwC1VrXhc10i6oLT+5XYzUKOWBAT
# lq8oFI9Zj3KueN52gaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDQyMjM4
# MjlaMC8GCSqGSIb3DQEJBDEiBCBhhIfmGv2ZQjTv72DSEnzGXJPse1V741QLbZZ/
# NJK3KjANBgkqhkiG9w0BAQEFAASCAgCh9A8b+2H1yOQgoXL2wTo5I2UMXaaQ78Y4
# MwrxB+cRqK9q0tQbyRaKlEc2VzSC6i/9YMU7rnvPgPFbbW/hx+f3GDD0fbWOhKHY
# 6ND+wLoo/PF5ovME8bpt7DWq7BDzJORuimvYgyO3dOj9m13g0B51vSBESIhjIZAX
# 5qx4ZKk4B4x506llJc/mjbhqwmLi+e2KeWvrISJmlse2vQjAVVAl+ENJXPFGny4+
# HSRgTZmOwlIdIRiUYHNbNnnCdANxZktmdKoG9tii6UW7EbScRayA3/INCwMUwuZC
# HCWx9XYmVv0TDf2i1N5iU9APkC7hDkwQDD8R1XdjilrpdRQf20wkYL2M0IjKTQDm
# xOzq0kxKAIxA6+kXJCx5/BYP5fklKDHx6VmCVI8h8NUX4uAzbydp6LvgxylLY5Ud
# VOws//V32C4w0kioWS/6w+u0dIMSJWeZInCXlR5Y9/QHZWp09rTDKIsdeVN+Qzeo
# aWZdTOVXEcDBcIV7EuGClk5jCIKQBYw+uFCE7A4MLwwgsQ5YjMvHiAznil+k75FS
# O2ghyhjvna9lwvWaMbZITaYWsQBc35mnXaRENMEsehPN4PFTp5Tz5R0Odpv3vAd3
# /GSwo5OZ8sXxUv9XlywIPV6MQhv1kx7YASBxez5BK1yI/M0zfNKRX1DnUKvexliV
# iC6uOhmUQA==
# SIG # End signature block
