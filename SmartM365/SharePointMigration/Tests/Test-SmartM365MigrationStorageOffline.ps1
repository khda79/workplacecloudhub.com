<#
.SYNOPSIS
    Offline checks for the configurable SharePointMigration storage roots (Config\storage.json.txt).
.DESCRIPTION
    Uses synthetic toolkit and migration folders under TEMP only. Covers default resolution,
    template self-heal, shared roots, launcher config lookup, GUI activity and batch log roots,
    cross-volume archive move and the toolkit.root.txt pointer read by the template launchers.
.VERSION
    1.0.0
#>
[CmdletBinding()]
# CrossVolumeRoot: optional writable folder on another volume than TEMP for the archive copy path.
param([string]$CrossVolumeRoot = '')
$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $PSScriptRoot
$launchers = Join-Path $project 'Scripts\Launchers'
. (Join-Path $launchers 'SmartM365-SharePointMigration-Storage.ps1')
. (Join-Path $launchers 'SmartM365-SharePointMigration-LauncherCommon.ps1')
. (Join-Path $launchers 'Generic\SmartM365-SharePointMigration-GuiActivity.ps1')
. (Join-Path $launchers 'Generic\SmartM365-SharePointMigration-BatchGui.ps1')

$script:Checks = 0
function Assert-Storage {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}
function Write-StorageJson {
    param([string]$Toolkit, [hashtable]$Values)
    $json = [pscustomobject]$Values | ConvertTo-Json
    [IO.File]::WriteAllText((Join-Path $Toolkit 'Config\storage.json.txt'), $json, [Text.UTF8Encoding]::new($false))
}

$tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')
$root = Join-Path $tempRoot ('SharePointMigration-storage-test-' + [guid]::NewGuid().ToString('N'))
try {
    $toolkit = Join-Path $root 'Toolkit'
    $shared = Join-Path $root 'Shared'
    [void](New-Item -ItemType Directory -Path (Join-Path $toolkit 'Config') -Force)
    Copy-Item -LiteralPath (Join-Path $project 'Config\storage.json.template') -Destination (Join-Path $toolkit 'Config')

    # 1. Default layout and runtime creation from the template.
    $storage = Get-SmartM365MigrationStorage -ProjectRoot $toolkit
    Assert-Storage ($storage.MigrationsRoot -eq (Join-Path $toolkit 'Migrations')) 'default Migrations root'
    Assert-Storage ($storage.ArchivedMigrationsRoot -eq (Join-Path $toolkit 'ArchivedMigrations')) 'default ArchivedMigrations root'
    Assert-Storage ($storage.LogsRoot -eq (Join-Path $toolkit 'Migrations\logs')) 'default logs root'
    Assert-Storage $storage.IsDefault 'default flag'
    Assert-Storage (Test-Path -LiteralPath (Join-Path $toolkit 'Config\storage.json.txt')) 'runtime created from template'

    # 2. Self-heal adds missing template keys without touching existing values.
    [IO.File]::WriteAllText((Join-Path $toolkit 'Config\storage.json.txt'), '{ "MigrationsRoot": "" }')
    [void](Get-SmartM365MigrationStorage -ProjectRoot $toolkit)
    $healed = Get-Content -LiteralPath (Join-Path $toolkit 'Config\storage.json.txt') -Raw | ConvertFrom-Json
    Assert-Storage ($healed.PSObject.Properties['ArchivedMigrationsRoot'] -and $healed.PSObject.Properties['LogsRoot']) 'self-heal added keys'

    # 3. Shared roots; logs follow Migrations when LogsRoot is empty.
    Start-Sleep -Milliseconds 20
    Write-StorageJson $toolkit @{ MigrationsRoot = (Join-Path $shared 'Migrations'); ArchivedMigrationsRoot = (Join-Path $shared 'Archive'); LogsRoot = '' }
    $storage = Get-SmartM365MigrationStorage -ProjectRoot $toolkit
    Assert-Storage ($storage.MigrationsRoot -eq (Join-Path $shared 'Migrations')) 'shared Migrations root'
    Assert-Storage ($storage.ArchivedMigrationsRoot -eq (Join-Path $shared 'Archive')) 'shared archive root'
    Assert-Storage ($storage.LogsRoot -eq (Join-Path $shared 'Migrations\logs')) 'logs follow shared Migrations'
    Assert-Storage (-not $storage.IsDefault) 'configured flag'

    $env:SPMIG_STORAGE_TEST = $shared
    Start-Sleep -Milliseconds 20
    Write-StorageJson $toolkit @{ MigrationsRoot = '%SPMIG_STORAGE_TEST%\Migrations\'; ArchivedMigrationsRoot = 'Old'; LogsRoot = (Join-Path $shared 'Logs') }
    $storage = Get-SmartM365MigrationStorage -ProjectRoot $toolkit
    Assert-Storage ($storage.MigrationsRoot -eq (Join-Path $shared 'Migrations')) 'environment variable and trailing slash'
    Assert-Storage ($storage.ArchivedMigrationsRoot -eq (Join-Path $toolkit 'Old')) 'relative path resolved from the toolkit'
    Assert-Storage ($storage.LogsRoot -eq (Join-Path $shared 'Logs')) 'dedicated logs root'
    Assert-Storage ((Get-SmartM365MigrationLogsRoot -ProjectRoot $toolkit -ChildPath 'gui-activity') -eq (Join-Path $shared 'Logs\gui-activity')) 'logs child path'

    # 4. Consumers: launcher config, GUI activity, batch results.
    $migration = Join-Path $shared 'Migrations\Alpha'
    [void](New-Item -ItemType Directory -Path $migration -Force)
    [IO.File]::WriteAllText((Join-Path $migration 'migration.config.psd1'), "@{ Name='Alpha'; Source=@{Type='SP2019'}; Target=@{Type='SPO'} }")
    $loaded = Import-LauncherMigrationConfig -ProjectRoot $toolkit -MigrationName 'Alpha'
    Assert-Storage ($loaded.Root -eq $migration -and $loaded.Data.Name -eq 'Alpha') 'launcher reads the shared migration'
    $activity = New-SmartM365GuiActivity -ProjectRoot $toolkit -Migration 'Alpha' -Action 'Test'
    Assert-Storage ($activity.StartsWith((Join-Path $shared 'Logs\gui-activity\'), [StringComparison]::OrdinalIgnoreCase)) 'GUI activity under the shared logs root'
    $batchId = '20260101-010101-abcdef12'
    $batchDir = New-Item -ItemType Directory -Path (Join-Path $shared "Logs\comparison-batches\$batchId") -Force
    [IO.File]::WriteAllText((Join-Path $batchDir.FullName 'batch.log'), 'Batch finished: 1 successful')
    $result = Get-SmartM365BatchResult -Root $toolkit -Kind Comparison
    Assert-Storage ($result.Directory -eq $batchDir.FullName) 'batch results under the shared logs root'

    # 5. Invalid configuration fails explicitly instead of silently using the toolkit folders.
    Start-Sleep -Milliseconds 20
    [IO.File]::WriteAllText((Join-Path $toolkit 'Config\storage.json.txt'), '{ invalid')
    $failed = $false
    try { [void](Get-SmartM365MigrationStorage -ProjectRoot $toolkit) } catch { $failed = $_.Exception.Message -like '*Invalid SharePointMigration storage configuration*' }
    Assert-Storage $failed 'invalid storage configuration rejected'

    # 6. Archive move: same volume and forced copy path (verification, source removal).
    $src = New-Item -ItemType Directory -Path (Join-Path $root 'MoveA\Beta\scans') -Force
    [IO.File]::WriteAllText((Join-Path $src.FullName 'a.csv'), 'x,y')
    Move-SmartM365MigrationDirectory -Source (Join-Path $root 'MoveA\Beta') -Destination (Join-Path $root 'MoveB\Beta' | ForEach-Object { [void](New-Item -ItemType Directory -Path (Split-Path $_) -Force); $_ })
    Assert-Storage ((Test-Path -LiteralPath (Join-Path $root 'MoveB\Beta\scans\a.csv')) -and -not (Test-Path -LiteralPath (Join-Path $root 'MoveA\Beta'))) 'same-volume move'
    $crossBase = if ($CrossVolumeRoot) { $CrossVolumeRoot } else { $PSScriptRoot }
    if (-not [string]::Equals([IO.Path]::GetPathRoot($crossBase), [IO.Path]::GetPathRoot($root), [StringComparison]::OrdinalIgnoreCase)) {
        $crossRoot = Join-Path $crossBase ('.storage-move-test-' + [guid]::NewGuid().ToString('N'))
        try {
            [void](New-Item -ItemType Directory -Path $crossRoot)
            Move-SmartM365MigrationDirectory -Source (Join-Path $root 'MoveB\Beta') -Destination (Join-Path $crossRoot 'Beta')
            Assert-Storage ((Test-Path -LiteralPath (Join-Path $crossRoot 'Beta\scans\a.csv')) -and -not (Test-Path -LiteralPath (Join-Path $root 'MoveB\Beta'))) 'cross-volume move'
            Assert-Storage (@(Get-ChildItem -LiteralPath $crossRoot -Force -Filter '.moving-*').Count -eq 0) 'no staging folder left'
        }
        finally { if (Test-Path -LiteralPath $crossRoot) { Remove-Item -LiteralPath $crossRoot -Recurse -Force } }
    }
    else { Write-Host 'SKIP cross-volume move: no folder on another volume (use -CrossVolumeRoot).' }
    $refused = $false
    try { Move-SmartM365MigrationDirectory -Source $migration -Destination (Join-Path $shared 'Migrations') } catch { $refused = $true }
    Assert-Storage $refused 'move refuses an existing destination'

    # 7. Template launcher reads toolkit.root.txt when Migrations lives outside the toolkit.
    $stub = Join-Path $toolkit 'Scripts\Launchers\Generic'
    [void](New-Item -ItemType Directory -Path $stub -Force)
    [IO.File]::WriteAllText((Join-Path $stub 'SmartM365-SharePointMigration-Launcher.ps1'),
        "param([string]`$MigrationName, [string]`$Action) [IO.File]::WriteAllText('$root\launcher-called.txt', `"`$MigrationName|`$Action|`$PSCommandPath`")")
    $pointer = Set-SmartM365MigrationToolkitPointer -MigrationsRoot (Join-Path $shared 'Migrations') -ToolkitRoot $toolkit
    Assert-Storage ((Get-Content -LiteralPath $pointer -Raw).Trim() -eq $toolkit) 'pointer content'
    $launcherDir = New-Item -ItemType Directory -Path (Join-Path $migration 'launchers\scheduled-tasks\files') -Force
    Copy-Item -LiteralPath (Join-Path $project 'Migrations_Template\launchers\scheduled-tasks\files\03-Compare-Files-Scheduled.cmd') -Destination $launcherDir.FullName
    & cmd.exe /d /c ('"{0}"' -f (Join-Path $launcherDir.FullName '03-Compare-Files-Scheduled.cmd')) | Out-Null
    $called = if (Test-Path -LiteralPath "$root\launcher-called.txt") { (Get-Content -LiteralPath "$root\launcher-called.txt" -Raw).Split('|') } else { @() }
    Assert-Storage ($called.Count -eq 3 -and $called[0] -eq 'Alpha' -and $called[1] -eq 'CompareFiles' -and
        $called[2].StartsWith($toolkit, [StringComparison]::OrdinalIgnoreCase)) 'template launcher resolved the toolkit through toolkit.root.txt'

    Write-Host ("PASS Test-SmartM365MigrationStorageOffline: {0} checks" -f $script:Checks)
}
finally {
    Remove-Item Env:SPMIG_STORAGE_TEST -ErrorAction SilentlyContinue
    if ($root.StartsWith($tempRoot + '\', [StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $root)) {
        Remove-Item -LiteralPath $root -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDjXKwthAn18h5P
# SPn71AfB6f3LRXOeSN/qXN8DpNnueqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIO6OgmS/hE9Hmb6JsgByThkCJqGzQPT0QLVyN5Jto96DMA0GCSqG
# SIb3DQEBAQUABIIBgJEqjySlKgjj5o2aore808PpbC3lQJTevnKMefsCquUkro7c
# 2QhscG5zPZ5DKS3HB7L/qXVCldCeKlFdvTw5RrSYqWXQ9ARavDZDvNGJpVEyuqWW
# oLMhYVGEqnsZKM6zizGdaX6nHIL9xNkn/lP/m/8smwpRMACTjT1luMVpI1SmyMrR
# gUdz+mOg3HFmWObhSb/GqEUyWzSfONevNmupWt57SdnyPooEVwhlBKp6WjSMOReF
# 2aL91x0w6Ui8XanXwT50Q+Fk4BFhOi87qv3h8RnxNZXVOzuMj5luy29xXiqd5j3/
# 1crYDMyeeImJHSLqYtCRJoQff7rMLvQlvSQYtz/SvN8MFmQ4brK9lzqXPYI1dLqf
# A8063fBCHvP7tcGQ8Hqnni1Uu/oIGP4nGcgxKLdqJO12OBWJu+2sme00Qn+q/Xzt
# U8svelszKZfrxUVu72vHydAyufmW62f6c1ElEpk7QTcrYcJgqldSmaXxAFrNofRc
# qm9zlJPQAgc2yPhi26GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxNzQy
# MzhaMC8GCSqGSIb3DQEJBDEiBCBrOxAMSZsKFtpzm9o0xLNQWcD6XBD8MLL2NQan
# BALO7zANBgkqhkiG9w0BAQEFAASCAgAbU11NGOX8VoaJ5Z7YlmKen4vSSRwA42zZ
# jyq5nkeqqBrWDsiqB3U9YEJll9HVwTiscCqOBjZYu4C3B3JKhxAqv/l1Ijui3K4e
# Wesghjl/UB32+5kTq1tB+3xsBEzZ2Gkqk419hBbXpWIMZ86rLfxJb9eSedM/GhD/
# GDjTsOE9EcfhTpRAvpqeoaDHFuf58d1FLSPXEjoAueT9m5+3yaP2bUTsTi6Q+uR6
# Fpk5bYOdNh8P2R2ove36CyjJgjEawEfuFhmOmAtAoFnubp7G+1p/0To2lsq9N9mf
# Tgbjzj4QIOlD6CqCgbt4SKF5UHQwUzZQGl/OVyp8hLzw7vnKmE+6w3/iG7Ad03Sd
# VF4cIag0dXZ5NTL1NYGCZxgWr4XYMzsrVUm5VyDHEKWX5mdcrd5WBXc4f/mNf2Od
# h9wCcCdq60uNKmmunyWp7CVg+2J2v+c5kDI++eCl0xwXyRoC2KL+9305ND/umVLE
# CYl4x4trah8IqwLAGJUFcwydRMIxG+mxJJ8kZiSzax/v/Y9ETBd8GjsgkMls9V7m
# g7FvSBmKi4mfxho8+D6dkg7i/bb2kNHSeLcJ/B8xg6h/0TgXBbDI2sLxv+i4yIxQ
# 6j/NrDqKi3gB6ovPfMivFtnarY/KoEeLfjs/+89Dg3tPF6OPdPQw+BXiidgLMPA+
# 9REkVemCUw==
# SIG # End signature block
