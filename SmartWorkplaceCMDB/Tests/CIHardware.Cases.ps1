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
function Write-HardwareState { $hardwareState | ConvertTo-Json | Set-Content -LiteralPath ($hardwarePath+'.status.json.txt') -Encoding UTF8 }
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
    $manifest=Get-Content -Raw (Join-Path $outputPath 'CIRegistry.manifest.json.txt') | ConvertFrom-Json
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
            missingState { Remove-Item -LiteralPath ($hardwarePath+'.status.json.txt') }
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCvv+zoxZcveY1k
# 96VCCBx2C+RZ9w3EFIX9yLolBnDLD6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCByjN1RcjHlTXllqREsIMJm
# NtyABvtuSPfcirbd5R1TAjANBgkqhkiG9w0BAQEFAASCAYB3tbPKMMm9/kKCKXIW
# 1xsBHrtiXbo3Ea0W+pCQqwi9Kfv1Fl5JllfTmer9wCGXg68meIknaK00w8Zi4mQ6
# OC6SCrQydBBC3wbzj4AROBF0XHnF94cD6A/Lk4Tv7tqAUBPO5VuLa5qmFiAgwsKg
# OYBICP2J1n8gICGTTtiPCZMoKJxMWZdZj8f+oCNNDL4StAcg6YkCbz+nNh0CEy1i
# Kof4tAQSxeCwT4oH9nbJlONrmMWZsYQudTCiSQXs+uW2s6Ff+PWthyzba9nUhjOR
# DQCykH+zUKui7UgKVPqCe+28VQ1kcrAFglZxwPiIWEy0w1fPMUZOC+vzpS+BfdL1
# SvaRZUH4ax+E8zQvcYn7+F8HSZ+5BRoU9pBY4jn3ABZH7ItL1k0cd2Mvvng6mj5d
# 8emRzerVbbbA42vQAptCNObxYROPwCP4YeFrujZVZ36efcGQ+IqPJgXTiiv1hoat
# METQSeHSzQaQAtQwn6UW9xuAv8KETMd5FyJ42hhHH1uTqdQ=
# SIG # End signature block
