<#
.SYNOPSIS
Validates or explicitly collects an optional read-only Intune hardware snapshot.
.DESCRIPTION
Default Graph validation performs no authentication, network call or write.
Collect explicitly enables Graph collection. InputJsonPath selects offline
fixtures; ValidateOnly validates their values without writing. Configurations
are never modified. Existing inventory CSV contracts remain unchanged.
.VERSION
1.1.2
#>
[CmdletBinding(DefaultParameterSetName = 'Validate')]
param(
    [Alias('ProfileKey')][string]$Tenant = 'default',
    [string]$OrganizationKey,
    [string]$EnvironmentKey,
    [string]$TenantKey,
    [string]$TenantId,
    [string]$DataRootPath,
    [string]$DataAllRootPath,
    [string]$LatestOutputRootPath,
    [string]$LogRootPath,
    [string]$GlobalConfigPath,
    [string]$TenantConfigPath,
    [Parameter(ParameterSetName = 'Graph', Mandatory)][switch]$Collect,
    [Parameter(ParameterSetName = 'Fixture', Mandatory)][string]$InputJsonPath,
    [ValidateRange(0, 2147483647)][int]$MaxItems = 0,
    [switch]$NoConfigWrite,
    [switch]$ValidateOnly
)
Set-StrictMode -Version 2.0
$ScriptVersion = '1.1.2'
$ErrorActionPreference = 'Stop'
$fixture = $PSCmdlet.ParameterSetName -eq 'Fixture'
$noWrite = $ValidateOnly -or (-not $fixture -and -not $Collect)
$projectRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
Import-Module (Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1') -Force
Import-Module (Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Graph\SmartWorkplaceCMDB.Graph.psd1') -Force

function Get-HardwareText {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -isnot [string]) { throw 'Hardware text attributes must be strings or null.' }
    return ($Value -replace "`r`n|`n|`r", ' ').Trim()
}

function ConvertTo-HardwareRow {
    param([AllowEmptyCollection()][object[]]$Devices, [string]$CollectedDateTime)
    $seen = @{}
    foreach ($device in $Devices) {
        # The shared getter expects Hashtable/PSObject. Generic Graph dictionaries
        # can expose only a KeyValuePair Contains overload; adapt them locally.
        if ($device -is [System.Collections.IDictionary]) {
            $properties = @{}
            foreach ($key in $device.Keys) { $properties[[string]$key] = $device[$key] }
            $device = $properties
        }
        $id = Get-HardwareText (Get-SmartWorkplaceCMDBGraphObjectValue $device 'id')
        if (-not $id) { throw 'A managed device response is missing id.' }
        if ($seen.ContainsKey($id)) { throw 'Duplicate managed device identifiers were returned.' }
        $seen[$id] = $true
        $row = [ordered]@{
            SourceSystem = 'MicrosoftIntune'
            ManagedDeviceId = $id
            AzureAdDeviceId = Get-HardwareText (Get-SmartWorkplaceCMDBGraphObjectValue $device 'azureADDeviceId')
        }
        $sourceNames = [ordered]@{SerialNumber='serialNumber';Manufacturer='manufacturer';Model='model'}
        foreach ($attribute in $sourceNames.Keys) {
            $value = Get-HardwareText (Get-SmartWorkplaceCMDBGraphObjectValue $device $sourceNames[$attribute])
            $row[$attribute] = $value
            $row[$attribute + 'Status'] = if ($value) { 'Reported' } else { 'Missing' }
        }
        $rawStorage = Get-SmartWorkplaceCMDBGraphObjectValue $device 'totalStorageSpaceInBytes'
        $storage = [long]0
        if ($null -eq $rawStorage) {
            $row.TotalStorageSpaceInBytes = ''
            $row.StorageStatus = 'Missing'
        } else {
            $storageText = [Convert]::ToString($rawStorage, [Globalization.CultureInfo]::InvariantCulture)
            if ($storageText -cnotmatch '^\d+$' -or -not [long]::TryParse($storageText, [ref]$storage)) {
                throw 'Storage capacity must be a nonnegative Int64 byte count or null.'
            }
            $row.TotalStorageSpaceInBytes = $storage.ToString([Globalization.CultureInfo]::InvariantCulture)
            $row.StorageStatus = if ($storage -eq 0) { 'ZeroReported' } else { 'Reported' }
        }
        $row.SourceCollectedDateTime = $CollectedDateTime
        [pscustomobject]$row
    }
}

$bound = @{}
foreach ($key in $PSBoundParameters.Keys) { $bound[$key] = $PSBoundParameters[$key] }
$context = Resolve-SmartWorkplaceCMDBContext -BoundParameters $bound -GlobalConfigPath $GlobalConfigPath -TenantConfigPath $TenantConfigPath -NoConfigWrite:($NoConfigWrite -or $noWrite -or $fixture)
# Resolve without writes first, so invalid fixture input cannot damage prior output.
$paths = Resolve-SmartWorkplaceCMDBCollectionPaths -Paths $context.Paths -Fixture:$fixture -MaxItems $MaxItems -ExplicitDataRoot -NoWrite
$contractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.hardware.tables.json'
$contract = Get-SmartWorkplaceCMDBTableContract -Path $contractPath
if ($contract.channel -cne 'stable' -or $contract.contractVersion -cne '1.0.0') { throw 'Hardware contract must be the frozen V1 stable contract.' }
$table = @($contract.tables | Where-Object name -ceq 'Intune_DeviceHardware.csv')
if ($table.Count -ne 1) { throw 'Exactly one hardware table contract is required.' }
$table = $table[0]
$uri = 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$select={0}&$top=999' -f $contract.graphSelect
$collected = [datetime]::UtcNow.ToString('o')
$rows = @()
if ($fixture) {
    $devices = @(Read-SmartWorkplaceCMDBCollectionFixture -Path $InputJsonPath)
    if ($MaxItems -gt 0) { $devices = @($devices | Select-Object -First $MaxItems) }
    $rows = @(ConvertTo-HardwareRow -Devices $devices -CollectedDateTime $collected)
}
if ($noWrite) {
    [pscustomobject]@{
        Status = 'ValidatedOffline'; Channel = 'stable'; ContractVersion = $contract.contractVersion
        SourceMode = if ($fixture) { 'OfflineJson' } else { 'GraphNotExecuted' }
        DeviceCount = if ($fixture) { $rows.Count } else { $null }
        RequiredGraphPermission = 'DeviceManagementManagedDevices.Read.All'
        RequestUri = $uri; WritesPerformed = $false; AuthenticationValidated = $false
    }
    return
}

if (-not $fixture) {
    $graph = Get-SmartWorkplaceCMDBGraphObjectValue $context.Configuration 'MicrosoftGraph'
    $collection = Get-SmartWorkplaceCMDBGraphObjectValue $context.Configuration 'Collection'
    $cloud = [string](Get-SmartWorkplaceCMDBGraphObjectValue $collection 'Cloud')
    if ($cloud -and $cloud -ne 'Public') { throw 'Hardware collection supports the Public cloud only.' }
    $clientId = [string](Get-SmartWorkplaceCMDBGraphObjectValue $graph 'ClientId')
    $thumbprint = [string](Get-SmartWorkplaceCMDBGraphObjectValue $graph 'CertificateThumbprint')
    Test-SmartWorkplaceCMDBGraphAppOnlyReadiness -TenantId $paths.TenantId -ClientId $clientId -CertificateThumbprint $thumbprint | Out-Null
}
$paths = Resolve-SmartWorkplaceCMDBCollectionPaths -Paths $context.Paths -Fixture:$fixture -MaxItems $MaxItems -ExplicitDataRoot
$latest = Join-Path $paths.LatestOutputRootPath (Join-Path $table.area $table.name)
$executionMode = if ($ValidateOnly) { 'Validate' } elseif ($fixture) { 'Fixture' } else { 'Collect' }
$runtimeContext = Start-SmartWorkplaceCMDBExecutionContext -Context $context -ScriptPath $PSCommandPath -ScriptVersion $ScriptVersion -Mode $executionMode -NoWrite:$ValidateOnly
$executionError = $null
$run = $null
try {
    $run = Start-SmartWorkplaceCMDBSourceCollection -Paths $paths -RawPath @($latest) -Fixture:$fixture -MaxItems $MaxItems -NoWrite:$ValidateOnly
    if (-not $fixture) {
        $devices = @(Invoke-SmartWorkplaceCMDBGraphPagedRequest -TenantId $paths.TenantId -ClientId $clientId -CertificateThumbprint $thumbprint -Uri $uri -RequiredPermission 'DeviceManagementManagedDevices.Read.All' -MaxItems $MaxItems)
        if ($MaxItems -gt 0) { $devices = @($devices | Select-Object -First $MaxItems) }
        $collected = [datetime]::UtcNow.ToString('o')
        $rows = @(ConvertTo-HardwareRow -Devices $devices -CollectedDateTime $collected)
    }
    $stamp = [datetime]::UtcNow
    $history = Join-Path $paths.DataAllRootPath ('Intune\DeviceHardware\{0}\{1}\Intune_DeviceHardware_{2}_{3}.csv' -f $stamp.ToString('yyyy'), $stamp.ToString('MM'), $stamp.ToString('yyyyMMdd-HHmmssfff'), $run.RunId)
    Publish-SmartWorkplaceCMDBSourceCsv `
        -Run $run `
        -InputObject $rows `
        -Columns @($table.columns) `
        -HistoryPath $history `
        -LatestPath $latest `
        -ContractPath $contractPath `
        -ContractTableName 'Intune_DeviceHardware.csv' | Out-Null
    [pscustomobject]@{
        Status = 'Completed'; Channel = 'stable'; ContractVersion = $contract.contractVersion
        SourceMode = if ($fixture) { 'OfflineJson' } else { 'MicrosoftGraphAppOnly' }
        DeviceCount = $rows.Count; CollectedDateTime = $collected
        DuplicateSerialValueCount = @($rows | Where-Object SerialNumberStatus -eq 'Reported' | Group-Object SerialNumber | Where-Object Count -gt 1).Count
        HistoryPath = $history; RawLatestOutputPath = $latest
    }
} catch {
    $executionError = $_
    if ($null -ne $run) { Complete-SmartWorkplaceCMDBSourceCollection -Run $run -Failed }
    throw
} finally {
    Complete-SmartWorkplaceCMDBExecutionContext -RuntimeContext $runtimeContext -ErrorRecord $executionError
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAcKJmIO04u9U20
# d+TU+yNm2q2YtV+08SxJRdfFP8zsjKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINWlb+PsdUA5t8V9SrtNJkIuDmVYBQX+xv4tl4OqCwBSMA0GCSqG
# SIb3DQEBAQUABIIBgJvuymWdka2HnmHXk/ZbqkDqOm4ZGn+7yY7QlyU0Nnag2CB+
# hXA5B6zzJ0g6XTqVV0ciXIu9GrlcvhtVdqU3awZzrt1t+e3SH+8c0JB0jzQZjXbJ
# cbQp+PPZe1Krt2CpOSHueIP7KElUdl8j17Rbh7slvSscEvVcIsPYf28nniYMAjTw
# QYau7zOrEYICjIMiVVGUqGqlO7z0Xlo4NAr8NdZ29LL0N7UuU1MI1zVbsFGzmhs5
# TXtgycdnNeEIpMXUvmquyYUhlGkEpLFzMyVA55jniji/Dn7dsnu3MQzpaC8FlwnI
# +u9IuzxEgJqcnCzvSbnt96Fb1Q6f2iA6oqoOD52E4Xqd/BhALdsdsk4WQBV7C9sO
# Jee+vVk6KFDhVVIpTJKb6PFD0PjywEDuYQQqIFHmY3XIKGrq4ngmYSlxB5yIZrmT
# UPXvMiSiNk+eBQa0ghU8L5F7Le+ZK2MDvS/yJFty9roa91l5gXZgcAlG0xl5nMgT
# mJFHPtjubIRaMk6hf6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxNzI1
# MzFaMC8GCSqGSIb3DQEJBDEiBCD7cpDXGG3cqnzzh7V0oLlDiQb+Pyl7Z4O5cVFM
# S/pWtTANBgkqhkiG9w0BAQEFAASCAgBsfLnj5l311ltyb+Ex4dTnvCbq+VJESXlh
# fL89V9EURvlMQNT/rWhu6sSN+AiOg7L7tThpDhNwuJgjHrzvoOM2n8J34yPUmtyh
# trVTtCBAsi+LRKjSIBNYZbeE0I9leAqivObOKXfu4rU8nTDzJF0JJDvFa5CuloIG
# L7w4oilCpJBNCcFgjZiJdjTjL2ceRTa9lrNNLXF9eIaTlCsFXbEIr0c8rhY6KXNj
# De9IGALi6OAOpYBjwF2IolpsuUeIX6lyGLPNab8xcx+13etdoOlB9dTwMQxUHWrJ
# 4kS5FVDQKfWrlc1b3Y9Lg08LXwvFM/oyPwx/68S96pbZIDsu1Qzt16HD/C5URybz
# NXe8YV0G31xPA30eElqfJNgWeu9vXPYMQFrq2tt3rFrRfdzW2oxBjxqheYVBbFw6
# Le8ZGmVvMAJbj3z2Cltgj5je4BIYmtZa9aNtMgdWEfbAyBDdHJpnutOn61qliDb9
# cwiui+NbU6QrwBWP7sr/DSTcdRLDgamYcahw77Kxj2LnJYAkl9Ff8ilCoSVN5YjV
# gPTTFvVffZKQ5bqR8IY6vAQicQaczJY6/X/RzfZYcnADKkbB9k9GjHlcJYVf8IiJ
# +YRkbQPdKZJhyf3B9nis2Goy/cPmXTFoBiwjfyT38ciHizxwSj2wZtCXuyJlBFFE
# RN3Kdm0xvQ==
# SIG # End signature block
