# Stable V1 fixture cases, executed inside Test-SmartWorkplaceCMDB-CIRegistry.ps1.
function Write-HardwareFixture {
    param([AllowEmptyCollection()][object[]]$Rows)
    $schema=Get-Content -Raw (Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.hardware.tables.json') | ConvertFrom-Json
    if ($Rows.Count) { $Rows | Select-Object $schema.tables[0].columns | Export-Csv -LiteralPath $hardwarePath -NoTypeInformation -Encoding UTF8 }
    else { ($schema.tables[0].columns -join ',') | Set-Content -LiteralPath $hardwarePath -Encoding UTF8 }
    $script:hardwareState=[ordered]@{
        Version=1;SourceName='Intune_DeviceHardware.csv';RunId='synthetic-run';Status='Completed'
        Mode='Fixture';Coverage='Fixture';MaxItems=0;RowCount=$Rows.Count
        SHA256=(Get-FileHash $hardwarePath).Hash
        TenantKey='example-test';OrganizationKey='example';EnvironmentKey='test';TenantId=''
        StartedUtc='2026-01-02T00:00:00Z';CompletedUtc='2026-01-02T00:01:00Z'
    }
    Write-HardwareState
}
function Write-HardwareState { $hardwareState | ConvertTo-Json | Set-Content -LiteralPath ($hardwarePath+'.status.json') -Encoding UTF8 }
function New-HardwareFixture {
    New-RawFixture
    $script:hardwarePath=Join-Path (Split-Path $outputPath) 'Intune_DeviceHardware.csv'
    $script:hardwareRow=[pscustomobject][ordered]@{
        TenantKey='example-test';OrganizationKey='example';EnvironmentKey='test';TenantId=''
        SourceSystem='MicrosoftIntune';ManagedDeviceId='intune-managed-1';AzureAdDeviceId='source-device-1'
        SerialNumber='SYNTHETIC-SERIAL';SerialNumberStatus='Reported';Manufacturer='Example manufacturer';ManufacturerStatus='Reported'
        Model='Example model';ModelStatus='Reported';TotalStorageSpaceInBytes='512110190592';StorageStatus='Reported'
        SourceCollectedDateTime='2026-01-02T00:00:30Z'
    }
    Write-HardwareFixture @($hardwareRow)
    $script:argsCI.RawRootPath=$rawPath; $script:argsCI.HardwareInputPath=$hardwarePath
}
Test-Case 'Hardware absent preserves existing outputs and explicit NotProvided status' {
    $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI
    Assert-True ($r.Hardware.Status -ceq 'NotProvided' -and -not (Test-Path (Join-Path $outputPath 'CMDB_CIDeviceHardware.csv'))) 'Unexpected hardware output'
}
Test-Case 'Hardware requires original raw mapping and ValidateOnly is read-only' {
    New-HardwareFixture
    $argsCI.Remove('RawRootPath')
    Assert-Throws { Export-SmartWorkplaceCMDBCIRegistry @argsCI -ValidateOnly } 'requires RawRootPath'
    $argsCI.RawRootPath=$rawPath
    $before=(Get-FileHash $hardwarePath).Hash
    $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -ValidateOnly
    Assert-True ($r.Hardware.RowCount -eq 1 -and $r.Hardware.CIsWithHardware -eq 1 -and -not (Test-Path $outputPath) -and (Get-FileHash $hardwarePath).Hash -ceq $before) 'Validation changed inputs or outputs'
}
Test-Case 'Hardware output preserves CI identity, source dates, contract and capacity precision' {
    New-HardwareFixture
    $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -IncludeContext
    $row=Import-Csv (Join-Path $outputPath 'CMDB_CIDeviceHardware.csv')
    $schema=Get-Content -Raw (Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.ci.hardware.json') | ConvertFrom-Json
    Assert-True (($row.PSObject.Properties.Name -join ',') -ceq ($schema.columns -join ',')) 'Hardware header drift'
    Assert-True ($row.CI_ID -ceq 'example-test|device|internal-1' -and $row.ManagedDeviceId -ceq 'intune-managed-1' -and $row.SerialNumber -ceq 'SYNTHETIC-SERIAL') 'Native identity or value changed'
    Assert-True ($row.HardwareCollectedDateTime -ceq '2026-01-02T00:00:30Z' -and $row.InventoryCollectedDateTime -ceq '2026-01-01T00:00:00Z' -and $row.TotalStorageSpaceInBytes -ceq '512110190592') 'Dates or capacity changed'
    Assert-True ($r.Hardware.AttributeFreshness -ceq 'NotAssessed' -and $r.Context.Status -ceq 'Validated') 'False freshness or context regression'
    $ci=Import-Csv (Join-Path $outputPath 'CMDB_ConfigurationItems.csv') | Where-Object CI_Type -eq Device
    Assert-True ($ci.OwnershipStatus -ceq 'NotCollected' -and $ci.LifecycleStatus -ceq 'Unknown') 'Hardware inferred governance'
    $manifest=Get-Content -Raw (Join-Path $outputPath 'CIRegistry.manifest.json') | ConvertFrom-Json
    Assert-True ($manifest.Hardware.RowCount -eq 1 -and @($manifest.Hardware.InputHashes.PSObject.Properties).Count -eq 4) 'Hardware evidence missing from manifest'
}
Test-Case 'Hardware partial coverage and missing versus zero stay explicit' {
    New-HardwareFixture
    $native=Import-Csv (Get-RawPath 'Intune_ManagedDevices.csv'); $other=$native.PSObject.Copy(); $other.ManagedDeviceId='another-native'
    Write-Raw 'Intune_ManagedDevices.csv' @($native,$other)
    $hardwareRow.SerialNumber=''; $hardwareRow.SerialNumberStatus='Missing'; $hardwareRow.TotalStorageSpaceInBytes='0'; $hardwareRow.StorageStatus='ZeroReported'
    Write-HardwareFixture @($hardwareRow); $hardwareState.Coverage='Bounded'; $hardwareState.MaxItems=500; Write-HardwareState
    $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI
    Assert-True ($r.Hardware.InventoryNativeDevicesWithoutHardware -eq 1 -and $r.Hardware.Coverage -ceq 'Bounded' -and $r.Hardware.MissingAttributes.SerialNumber -eq 1 -and $r.Hardware.ZeroReportedStorageCount -eq 1) 'Partial coverage or zero hidden'
}
Test-Case 'Multiple Intune candidates and repeated serials remain separate on the same CI' {
    New-HardwareFixture
    $native=Import-Csv (Get-RawPath 'Intune_ManagedDevices.csv'); $other=$native.PSObject.Copy(); $other.ManagedDeviceId='another-native'
    Write-Raw 'Intune_ManagedDevices.csv' @($native,$other)
    $otherHardware=$hardwareRow.PSObject.Copy(); $otherHardware.ManagedDeviceId='another-native'
    Write-HardwareFixture @($hardwareRow,$otherHardware)
    $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI
    Assert-True ($r.Hardware.RowCount -eq 2 -and $r.Hardware.CIsWithHardware -eq 1 -and $r.Hardware.DuplicateSerialValueCount -eq 1) 'Candidate or serial incorrectly merged'
}
Test-Case 'Hardware empty snapshot exports header and reports uncovered inventory' {
    New-HardwareFixture; Write-HardwareFixture @()
    $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI
    Assert-True ($r.Hardware.RowCount -eq 0 -and $r.Hardware.InventoryNativeDevicesWithoutHardware -eq 1 -and @(Import-Csv (Join-Path $outputPath 'CMDB_CIDeviceHardware.csv')).Count -eq 0) 'Empty snapshot inferred hardware'
}
foreach ($case in @('missingState','Failed','InProgress','hash','rowCount','tenant','coverage','mode','future','reversed','unqualified')) {
    Test-Case "Reject hardware state $case" {
        New-HardwareFixture
        switch ($case) {
            missingState { Remove-Item -LiteralPath ($hardwarePath+'.status.json') }
            Failed { $hardwareState.Status='Failed' }
            InProgress { $hardwareState.Status='InProgress' }
            hash { $hardwareState.SHA256='A'*64 }
            rowCount { $hardwareState.RowCount=2 }
            tenant { $hardwareState.TenantKey='other-test' }
            coverage { $hardwareState.Coverage='Bounded' }
            mode { $hardwareState.Mode='Unknown' }
            future { $hardwareState.CompletedUtc=[datetime]::UtcNow.AddDays(1).ToString('o') }
            reversed { $hardwareState.StartedUtc='2026-01-03T00:00:00Z' }
            unqualified { $hardwareState.CompletedUtc='01/02/2026 12:00:00' }
        }
        if ($case -ne 'missingState') { Write-HardwareState }
        Assert-Throws { Export-SmartWorkplaceCMDBCIRegistry @argsCI } 'hardware|Hardware'
        Assert-True (-not (Test-Path $outputPath) -and @(Get-ChildItem (Split-Path $outputPath) -Force -Directory | Where-Object Name -like '.cmdb-ci-*').Count -eq 0) 'Failed build left partial hardware output'
    }
}
foreach ($case in @('tenant','source','duplicate','unmapped','correlation','status','storage','date','header')) {
    Test-Case "Reject hardware row $case" {
        New-HardwareFixture
        switch ($case) {
            tenant { $hardwareRow.TenantId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' }
            source { $hardwareRow.SourceSystem='Other' }
            duplicate { }
            unmapped { $hardwareRow.ManagedDeviceId='missing-native' }
            correlation { $hardwareRow.AzureAdDeviceId='different-correlation' }
            status { $hardwareRow.SerialNumberStatus='Missing' }
            storage { $hardwareRow.TotalStorageSpaceInBytes='9223372036854775808' }
            date { $hardwareRow.SourceCollectedDateTime='01/02/2026 12:00:00' }
            header { }
        }
        if ($case -eq 'duplicate') { Write-HardwareFixture @($hardwareRow,$hardwareRow) } else { Write-HardwareFixture @($hardwareRow) }
        if ($case -eq 'header') { 'WrongHeader' | Set-Content $hardwarePath; $hardwareState.SHA256=(Get-FileHash $hardwarePath).Hash; Write-HardwareState }
        Assert-Throws { Export-SmartWorkplaceCMDBCIRegistry @argsCI } 'hardware|Hardware|header'
        Assert-True (-not (Test-Path $outputPath) -and @(Get-ChildItem (Split-Path $outputPath) -Force -Directory | Where-Object Name -like '.cmdb-ci-*').Count -eq 0) 'Failed build left stage'
    }
}
Test-Case 'Corrected hardware can be retried without changing registry keys' {
    New-HardwareFixture; $hardwareState.RowCount=9; Write-HardwareState
    Assert-Throws { Export-SmartWorkplaceCMDBCIRegistry @argsCI } 'row count'
    Write-HardwareFixture @($hardwareRow)
    $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI
    Assert-True ($r.Status -ceq 'Exported' -and $r.CICount -eq 5 -and $r.RelationshipCount -eq 1) 'Retry changed registry'
}
Test-Case 'Hardware collector fixture output imports through the CI adapter end to end' {
    New-HardwareFixture
    $tenantId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
    foreach ($folder in @($inputPath,$rawPath)) {
        Get-ChildItem -LiteralPath $folder -Filter '*.csv' -Recurse -File | ForEach-Object {
            $file=$_.FullName; $rows=@(Import-Csv -LiteralPath $file)
            foreach ($row in $rows) { $row.TenantId=$tenantId }
            if ($rows.Count) { $rows | Export-Csv -LiteralPath $file -NoTypeInformation -Encoding UTF8 }
        }
    }
    $argsCI.TenantId=$tenantId
    $fixture=Join-Path (Split-Path $outputPath) 'hardware-fixture.json'
    '{"value":[{"id":"intune-managed-1","azureADDeviceId":"source-device-1","serialNumber":"SYNTHETIC-E2E","manufacturer":"Example","model":"Example"}]}' | Set-Content -LiteralPath $fixture -Encoding UTF8
    $collector=Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneHardware-Collect.ps1'
    $collected=& $collector -Tenant audit -OrganizationKey example -EnvironmentKey test -TenantKey example-test -TenantId $tenantId -DataRootPath (Join-Path (Split-Path $outputPath) 'Collection') -InputJsonPath $fixture
    $argsCI.HardwareInputPath=$collected.RawLatestOutputPath
    $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI
    $result=Import-Csv (Join-Path $outputPath 'CMDB_CIDeviceHardware.csv')
    Assert-True ($r.Hardware.RowCount -eq 1 -and $result.SerialNumber -ceq 'SYNTHETIC-E2E' -and $result.StorageStatus -ceq 'Missing' -and $result.HardwareCollectedDateTime -ceq $collected.CollectedDateTime) 'Collector/CI contract mismatch'
}
Test-Case 'Missing Azure ID uses the existing Intune-only inventory identity' {
    New-HardwareFixture
    Write-Raw 'Entra_Devices.csv' @()
    $device=Import-Csv (Join-Path $inputPath 'CMDB_Devices.csv'); $device.SourceSystem='MicrosoftIntune'; $device.SourceDeviceId='intune:intune-managed-1'; Write-Row 'CMDB_Devices.csv' @($device)
    $native=Import-Csv (Get-RawPath 'Intune_ManagedDevices.csv'); $native.AzureAdDeviceId=''; Write-Raw 'Intune_ManagedDevices.csv' @($native)
    $hardwareRow.AzureAdDeviceId='00000000-0000-0000-0000-000000000000'; Write-HardwareFixture @($hardwareRow)
    $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -ValidateOnly
    Assert-True ($r.Hardware.CIsWithHardware -eq 1) 'Fallback mapping changed'
}
Test-Case 'Hardware rows from mixed retrieval times cannot masquerade as one run' {
    New-HardwareFixture
    $native=Import-Csv (Get-RawPath 'Intune_ManagedDevices.csv'); $other=$native.PSObject.Copy(); $other.ManagedDeviceId='another-native'; Write-Raw 'Intune_ManagedDevices.csv' @($native,$other)
    $otherHardware=$hardwareRow.PSObject.Copy(); $otherHardware.ManagedDeviceId='another-native'; $otherHardware.SourceCollectedDateTime='2026-01-02T00:00:40Z'
    Write-HardwareFixture @($hardwareRow,$otherHardware)
    Assert-Throws { Export-SmartWorkplaceCMDBCIRegistry @argsCI -ValidateOnly } 'mixed hardware collection'
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAcH7LyvRQolLSu
# V8bYPp96WUCpCAt/ZBM64n9PIpiY5KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFEuFbAGZkwRDHwRrkEBw6knIsR1zzzTkF5BPq1tiUnKMA0GCSqG
# SIb3DQEBAQUABIIBgACPA18JhsxhEMH0L6VGlbxHN0bkUy7K6oI0r968TvgM4tUM
# /1T8ZM3Pe41IlbIG3/GJxFrDA7AlGr7R/Tx0BLHfLeH4pb/rxQ2Us3yzrUlu4A62
# wAgpeOZsZiAz1eToi0naVNXmHa0gNtJMk+aoLN47X7pCCtFDYMY6WJqzUeVn5Xtt
# hfWsneH15MihiIg7ed3lT2i3GBu6ece+andmXljWFDIor1/uUd5JNKq2TvuT6fFs
# zvqyamhK2JoWLzFCGgh9p2I7kBLuSJVCyzKqBjzRAK0+V3Sc/eFi9J5E0/aZOH3h
# Gw4feKFgUunXHimT8Myt5HTz/JBNYYX9r1+ioCy6GLVVYdRn9gXsq0tb1tx6zX7G
# kdU8UxJsaWK0aNJbyBjSnVU+KSqEjg3Tl2C6P4YNvFbVo623D31Y5AJ3h7d5hPe6
# BFxx/HSVX3AvocNdlUKt94uXxdLlLPZUfIG2wdP0LH5qcVFetaN8fxQmMchb2W/a
# blczGQ547DICuBCuwKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxMzI1
# MjNaMC8GCSqGSIb3DQEJBDEiBCCVCttCHfjbl/HjUf628709+6LLrAJpuMIJ4MG7
# XXZ1OTANBgkqhkiG9w0BAQEFAASCAgB3qo5+0Xy/v2hoN0IB+V6J4db1tn3Yy/yk
# 8/1nzVG7xZQy7JxsSWJZKqe0qmQbkEZxuVTQ9qfJP7mKY97fEz8FtiTeB+447NY3
# e2uH7fsxljOCuA1iOIcL7qPYzSGvNhHdaDEcw54Zji5V4cu25lwYcgS1RAn4S20D
# LeA9eV0Clqkw5XK8DReo0lpotpAsxTrKwam75aMb929YN1ij+p13n/qH7/xjSpFm
# KMhNcobkxZSKSNT3UTmeyNnkcM+VeXpHnsQk0u/PN2AaUCq9XnhlUdbR8P+KAiIA
# 70AYZkIKX2UBSAr+SftDocl9no1JifGclofhjYaCnzOiU2Z662dlJsB+jGEiTfMv
# V4UBsKfW5i3ALSwqmLyyEfRGsdZjLeqhWRt7IfNHn37zpWGxU3j4Q0e2Ar6qk2Uy
# Spz8my51jD6zY1QE+TmhbzbsDSGiqsURUHdJvDzOVmOHpv2A5zffLZy32CAtOP7P
# Zm3jepDE9Zvdpm7jKTnYYJvLjuYIkJ1yR/M/led+QNDlMuk9RLyaG0/bkt5or69U
# ih7kuyKcSStEAuA3yxjRc44PbbZJscsdbbZXCau0be7qunQchoBy2pIFL3lBYe+3
# kmvxnVcbDsWPkBH1qwrV4unZ2Zv0Ae28jMtxyJ8v9lAWjDrpK0HkW+yPmpbRZOU8
# 7SAxAYpZyg==
# SIG # End signature block
