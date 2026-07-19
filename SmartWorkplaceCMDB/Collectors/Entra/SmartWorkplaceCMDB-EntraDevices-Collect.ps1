<#
.SYNOPSIS
Collects a read-only Microsoft Entra device snapshot for SmartWorkplaceCMDB.

.DESCRIPTION
Uses Microsoft Graph app-only certificate authentication to retrieve base Entra
device identity and state properties. Raw snapshots are written to DATA-ALL and
the latest raw contract is written below DATA-LAST\Raw\Entra. Offline JSON is
supported for safe tests.

.VERSION
0.1.0

.REQUIREMENTS
PowerShell 5.1 or later.
Microsoft.Graph.Authentication for live collection.
Microsoft Graph application permission: Device.Read.All.
#>
[CmdletBinding(DefaultParameterSetName = 'Graph')]
param(
    [Alias('ProfileKey')]
    [string]$Tenant = 'default',
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
    [Parameter(ParameterSetName = 'Fixture', Mandatory)]
    [string]$InputJsonPath,
    [ValidateRange(0, 2147483647)]
    [int]$MaxItems = 0,
    [string]$RawLatestOutputPath,
    [switch]$NoConfigWrite,
    [switch]$ValidateOnly
)

$ScriptVersion = '0.1.0'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Get-SmartWorkplaceCMDBConfigSection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Configuration,
        [Parameter(Mandatory)]
        [string]$Name
    )
    if ($Configuration.Contains($Name) -and
        $Configuration[$Name] -is [System.Collections.IDictionary]) {
        return $Configuration[$Name]
    }
    return [ordered]@{}
}

function Get-SmartWorkplaceCMDBConfigText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Configuration,
        [Parameter(Mandatory)]
        [string]$Name
    )
    if ($Configuration.Contains($Name)) {
        return ([string]$Configuration[$Name]).Trim()
    }
    return ''
}

function ConvertTo-SmartWorkplaceCMDBCleanText {
    [CmdletBinding()]
    param([AllowNull()]$Value)
    if ($null -eq $Value) {
        return ''
    }
    return ([string]$Value -replace "`r`n|`n|`r", ' ').Trim()
}

function Read-SmartWorkplaceCMDBEntraDevicesFixture {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $document = Read-SmartWorkplaceCMDBJsonFile -Path $Path
    $value = Get-SmartWorkplaceCMDBGraphObjectValue -InputObject $document -Name 'value'
    if ($null -ne $value) {
        return @($value)
    }
    if ($document -is [System.Collections.IEnumerable] -and $document -isnot [string]) {
        return @($document)
    }
    throw "Offline Entra devices JSON must be an array or a Microsoft Graph response containing a value array: $Path"
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $scriptRoot)
$coreModulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$graphModulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Graph\SmartWorkplaceCMDB.Graph.psd1'
$rawContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
Import-Module $coreModulePath -Force
Import-Module $graphModulePath -Force

$boundParameterCopy = @{}
foreach ($key in $PSBoundParameters.Keys) {
    $boundParameterCopy[$key] = $PSBoundParameters[$key]
}
$context = Resolve-SmartWorkplaceCMDBContext `
    -BoundParameters $boundParameterCopy `
    -GlobalConfigPath $GlobalConfigPath `
    -TenantConfigPath $TenantConfigPath `
    -NoConfigWrite:($ValidateOnly -or $NoConfigWrite)
$paths = $context.Paths
$rawContract = Get-SmartWorkplaceCMDBTableContract -Path $rawContractPath
$rawTable = @($rawContract.tables | Where-Object name -eq 'Entra_Devices.csv')
if ($rawTable.Count -ne 1) {
    throw 'The raw contract must contain exactly one Entra_Devices.csv definition.'
}
$rawTable = $rawTable[0]

if ([string]::IsNullOrWhiteSpace($RawLatestOutputPath)) {
    $RawLatestOutputPath = Join-Path $paths.LatestOutputRootPath (
        Join-Path ([string]$rawTable.area) ([string]$rawTable.name)
    )
}
$RawLatestOutputPath = [System.IO.Path]::GetFullPath($RawLatestOutputPath)

$graphConfiguration = Get-SmartWorkplaceCMDBConfigSection $context.Configuration 'MicrosoftGraph'
$collectionConfiguration = Get-SmartWorkplaceCMDBConfigSection $context.Configuration 'Collection'
$cloud = Get-SmartWorkplaceCMDBConfigText $collectionConfiguration 'Cloud'
if ([string]::IsNullOrWhiteSpace($cloud)) {
    $cloud = 'Public'
}
if ($PSCmdlet.ParameterSetName -eq 'Graph' -and $cloud -ne 'Public') {
    throw "The initial Entra devices collector supports Collection.Cloud='Public' only. Configured value: '$cloud'."
}

$clientId = Get-SmartWorkplaceCMDBConfigText $graphConfiguration 'ClientId'
$thumbprint = Get-SmartWorkplaceCMDBConfigText $graphConfiguration 'CertificateThumbprint'
$fixtureDevices = $null
$readiness = $null
if ($PSCmdlet.ParameterSetName -eq 'Fixture') {
    $InputJsonPath = [System.IO.Path]::GetFullPath($InputJsonPath)
    $fixtureDevices = @(Read-SmartWorkplaceCMDBEntraDevicesFixture -Path $InputJsonPath)
}
else {
    $readiness = Test-SmartWorkplaceCMDBGraphAppOnlyReadiness `
        -TenantId $paths.TenantId `
        -ClientId $clientId `
        -CertificateThumbprint $thumbprint
}

if ($ValidateOnly) {
    [pscustomobject]@{
        Status                     = 'Valid'
        ScriptVersion              = $ScriptVersion
        SourceMode                 = if ($PSCmdlet.ParameterSetName -eq 'Fixture') { 'OfflineJson' } else { 'MicrosoftGraphAppOnly' }
        FixtureDeviceCount         = if ($null -ne $fixtureDevices) { $fixtureDevices.Count } else { 0 }
        GraphAuthenticationModule  = if ($null -ne $readiness) { $readiness.AuthenticationModule } else { '' }
        GraphAuthenticationVersion = if ($null -ne $readiness) { $readiness.AuthenticationModuleVersion } else { '' }
        RequiredGraphPermission    = 'Device.Read.All'
        RawContractVersion         = [string]$rawContract.contractVersion
        RawLatestOutputPath        = $RawLatestOutputPath
        ProfileKey                 = $paths.ProfileKey
        TenantKey                  = $paths.TenantKey
    } | Format-List
    return
}

$sourceDevices = if ($PSCmdlet.ParameterSetName -eq 'Fixture') {
    @($fixtureDevices)
}
else {
    $select = 'id,deviceId,displayName,operatingSystem,operatingSystemVersion,isManaged,isCompliant,approximateLastSignInDateTime'
    $uri = 'https://graph.microsoft.com/v1.0/devices?$select={0}&$top=999' -f $select
    @(Invoke-SmartWorkplaceCMDBGraphPagedRequest `
        -TenantId $paths.TenantId `
        -ClientId $clientId `
        -CertificateThumbprint $thumbprint `
        -Uri $uri `
        -RequiredPermission 'Device.Read.All' `
        -MaxItems $MaxItems)
}
if ($MaxItems -gt 0) {
    $sourceDevices = @($sourceDevices | Select-Object -First $MaxItems)
}

$collectedDateTime = [datetime]::UtcNow.ToString('o')
$rawRows = @($sourceDevices | ForEach-Object {
    $sourceObjectId = ConvertTo-SmartWorkplaceCMDBCleanText (
        Get-SmartWorkplaceCMDBGraphObjectValue $_ 'id'
    )
    $sourceDeviceId = ConvertTo-SmartWorkplaceCMDBCleanText (
        Get-SmartWorkplaceCMDBGraphObjectValue $_ 'deviceId'
    )
    if ([string]::IsNullOrWhiteSpace($sourceObjectId) -or
        [string]::IsNullOrWhiteSpace($sourceDeviceId)) {
        throw 'A Microsoft Entra device response did not contain both id and deviceId.'
    }
    $isManaged = Get-SmartWorkplaceCMDBGraphObjectValue $_ 'isManaged'
    $isCompliant = Get-SmartWorkplaceCMDBGraphObjectValue $_ 'isCompliant'
    [pscustomobject][ordered]@{
        SourceSystem                = 'MicrosoftEntraID'
        SourceObjectId              = $sourceObjectId
        SourceDeviceId              = $sourceDeviceId
        DeviceName                  = ConvertTo-SmartWorkplaceCMDBCleanText (Get-SmartWorkplaceCMDBGraphObjectValue $_ 'displayName')
        OperatingSystem             = ConvertTo-SmartWorkplaceCMDBCleanText (Get-SmartWorkplaceCMDBGraphObjectValue $_ 'operatingSystem')
        OperatingSystemVersion      = ConvertTo-SmartWorkplaceCMDBCleanText (Get-SmartWorkplaceCMDBGraphObjectValue $_ 'operatingSystemVersion')
        IsManaged                   = if ($null -eq $isManaged) { '' } else { [bool]$isManaged }
        IsCompliant                 = if ($null -eq $isCompliant) { '' } else { [bool]$isCompliant }
        ApproximateLastSignInDateTime = ConvertTo-SmartWorkplaceCMDBCleanText (Get-SmartWorkplaceCMDBGraphObjectValue $_ 'approximateLastSignInDateTime')
        SourceCollectedDateTime     = $collectedDateTime
    }
})

$duplicateDeviceIds = @($rawRows | Group-Object SourceDeviceId | Where-Object Count -gt 1)
if ($duplicateDeviceIds.Count -gt 0) {
    throw "Duplicate Microsoft Entra deviceIds were returned: $($duplicateDeviceIds.Name -join ', ')"
}

$historyTimestamp = [datetime]::UtcNow
$historyFolder = Join-Path $paths.DataAllRootPath (
    'Entra\Devices\{0}\{1}' -f $historyTimestamp.ToString('yyyy'), $historyTimestamp.ToString('MM')
)
$historyPath = Join-Path $historyFolder (
    'Entra_Devices_{0}.csv' -f $historyTimestamp.ToString('yyyyMMdd-HHmmssfff')
)
$exportParameters = @{
    InputObject      = $rawRows
    Columns          = @($rawTable.columns | ForEach-Object { [string]$_ })
    TenantKey        = $paths.TenantKey
    OrganizationKey  = $paths.OrganizationKey
    EnvironmentKey   = $paths.EnvironmentKey
    TenantId         = $paths.TenantId
}
Export-SmartWorkplaceCMDBCsv @exportParameters -Path $historyPath
Export-SmartWorkplaceCMDBCsv @exportParameters -Path $RawLatestOutputPath

$expectedLatestPath = Join-Path $paths.LatestOutputRootPath (
    Join-Path ([string]$rawTable.area) ([string]$rawTable.name)
)
if ([System.IO.Path]::GetFullPath($expectedLatestPath) -eq $RawLatestOutputPath) {
    $results = @(Test-SmartWorkplaceCMDBCsvContract `
        -LatestOutputRootPath $paths.LatestOutputRootPath `
        -ContractPath $rawContractPath)
    $deviceResult = @($results | Where-Object Name -eq 'Entra_Devices.csv')
    if ($deviceResult.Count -ne 1 -or $deviceResult[0].Status -ne 'Valid') {
        throw 'The latest Entra devices raw CSV does not satisfy the raw SmartWorkplaceCMDB contract.'
    }
}

Write-Information (
    "SmartWorkplaceCMDB Entra devices collection completed. Devices={0}; history='{1}'; latest='{2}'." -f
    $rawRows.Count,
    $historyPath,
    $RawLatestOutputPath
) -InformationAction Continue

[pscustomobject]@{
    Status                  = 'Completed'
    ScriptVersion           = $ScriptVersion
    SourceMode              = if ($PSCmdlet.ParameterSetName -eq 'Fixture') { 'OfflineJson' } else { 'MicrosoftGraphAppOnly' }
    DeviceCount             = $rawRows.Count
    CollectedDateTime       = $collectedDateTime
    HistoryPath             = $historyPath
    RawLatestOutputPath     = $RawLatestOutputPath
    RawContractVersion      = [string]$rawContract.contractVersion
    RequiredGraphPermission = 'Device.Read.All'
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCKpJyTsZHyGcKe
# hrk1dpUpRxg+YZrToqo5F2E55SvEjaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIGd82iXzBsZ7C6T10MsLgNkA3xMauiBPrbdNZQk+IAk1MA0GCSqG
# SIb3DQEBAQUABIIBgE/677uZ7mXs0bg2l4d+ZNHDXiJtGcwt5qoFBxWBiZpS/4yB
# 00d5vryRq++FbE0c/CsEj8ek0mhhapiJEtgDccPx2hfx+gRWUDDtjh50Aw7Db9D8
# A/IEP6f0Q/f0oOyVSHK5XEVsEcGogo+YE444JCmBAVeaUhRgKDCCeEtkwKRB4oaZ
# UpOdNstBtHPSEOqd/2vwwkZAm5sraZQS7UIlhtNiViCrK/gtaWGUV4YxUy0XMleW
# gCR1p3MhOHzklWlgmuPn8yBAWMANYua6qy5F7qjHGALzxk9JVqLiXih4/QuygkuI
# u10MB6vIk0AR3SR1e7lAFMDiAk/LfJy6JDgGFAyZlc03IP/0eRklID129CexRga5
# vRGpzQ1XL5/OqO/C6mkHt1IRR+DKc15udIp0xpMQTa72aUIY+/K8i9kGMaHGHQke
# ySTtXd0TjQVURIiZtUSAIAz1gNhMvQ3z5Kxr77bvN5ArnjxajvYKBW9Z+9S+Y8bB
# DiAr0culeqkKHt/7TaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkxNDM3
# MzJaMC8GCSqGSIb3DQEJBDEiBCCqc1JXYNumIEH+Q9yqobKiPs4aTEfyNzMStEb/
# aH2GJDANBgkqhkiG9w0BAQEFAASCAgDG4io2Uxq6zQgr63H6c5/0A8yPDwzfgNEB
# MYXnCLh2CZC0sbQ4kALPrewSwYV2LtpmZiPrzBFpPG6LZs7cAk4m5l8MUHsZlMf4
# qU1RFrWUiVMSzBzjwRWeWBkvnL/mToA1fnMCx7XRcXTM2MFGVZZeW3xCKw5Zq41B
# HeHJAtgauGmvrX+xxMV3xAa80u1GDCWulyYIMwqXACRmysgg3kqZ2to6ICduG52k
# 4KHAZfNyrt8zea8suuG/rnJ/MWzaEGBwU1QiAAuOzr6nsmhfArykXYmLMI95gAjI
# Gl0Qk+Q9u1rR482VUYFxei0tFKfgWGS9bxZF39ZiYpmCn/0YFj0+3dGip5B1DCjP
# I5fgx6up8eTjlX8cAG9fQGRMA4gzc4evFHeWM8Bgtdvqk67Y8LFjbJNyPVNvRLdz
# VtvXTiHXrzQcS2XhNYlAQMp7/0Rxbpx9bUJXbQjrzkAAP0e2M/FFWlhAVK6WNWxM
# J4v8Sd7vIDlVqeJyp1/o7Y1SOIEvcBOCkJguLlIcr10q3xEWtcTyq6mmiLw9mHzO
# mmVGiiTSkka2cR4usK3PdCPZjttEaC3hvbLexRiXGHmOFMu6p2oRWIv51Y8FnG8x
# B8LNOcsUAmy28xKyae7f7RjMfHYWJdWYdlGBv/GrQ1Inl+WoQ1aRGcmCwXFTlCul
# UiuSb/zySQ==
# SIG # End signature block
