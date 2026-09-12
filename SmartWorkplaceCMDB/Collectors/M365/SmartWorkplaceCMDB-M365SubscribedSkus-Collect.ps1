<#
.SYNOPSIS
Collects Microsoft 365 subscribed SKU capacity for SmartWorkplaceCMDB.

.DESCRIPTION
Uses Microsoft Graph app-only certificate authentication to retrieve commercial
subscriptions and license unit capacity. The collector is read-only and writes
history plus the latest raw contract. Offline JSON is supported for safe tests.

.VERSION
1.2.0

.REQUIREMENTS
PowerShell 5.1 or later.
Microsoft.Graph.Authentication for live collection.
Microsoft Graph application permission: LicenseAssignment.Read.All.
#>
[CmdletBinding(DefaultParameterSetName = 'Graph')]
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
    [Parameter(ParameterSetName = 'Fixture', Mandatory)][string]$InputJsonPath,
    [ValidateRange(0, 2147483647)][int]$MaxItems = 0,
    [string]$RawLatestOutputPath,
    [string]$RawServicePlansLatestOutputPath,
    [switch]$NoConfigWrite,
    [switch]$ValidateOnly
)

$ScriptVersion = '1.2.0'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Get-SmartWorkplaceCMDBConfigSection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Configuration,
        [Parameter(Mandatory)][string]$Name
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
        [Parameter(Mandatory)][System.Collections.IDictionary]$Configuration,
        [Parameter(Mandatory)][string]$Name
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

function ConvertTo-SmartWorkplaceCMDBLicenseInteger {
    [CmdletBinding()]
    param(
        [AllowNull()]$Value,
        [Parameter(Mandatory)][string]$FieldName,
        [Parameter(Mandatory)][string]$SubscribedSkuId
    )
    $parsed = 0
    if ($null -eq $Value -or
        -not [int]::TryParse(
            [string]$Value,
            [Globalization.NumberStyles]::Integer,
            [Globalization.CultureInfo]::InvariantCulture,
            [ref]$parsed
        )) {
        throw "$FieldName is missing or invalid for subscribed SKU '$SubscribedSkuId'."
    }
    return $parsed
}

function Read-SmartWorkplaceCMDBM365SubscribedSkusFixture {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    return @(Read-SmartWorkplaceCMDBCollectionFixture -Path $Path)
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
    -NoConfigWrite:($ValidateOnly -or $NoConfigWrite -or $PSCmdlet.ParameterSetName -eq 'Fixture')
$paths = Resolve-SmartWorkplaceCMDBCollectionPaths -Paths $context.Paths -Fixture:($PSCmdlet.ParameterSetName -eq 'Fixture') -MaxItems $MaxItems -ExplicitDataRoot:([bool]$DataRootPath) -NoWrite:$ValidateOnly
$rawContract = Get-SmartWorkplaceCMDBTableContract -Path $rawContractPath
$rawTable = @($rawContract.tables | Where-Object name -eq 'M365_SubscribedSkus.csv')
$servicePlanTable = @($rawContract.tables | Where-Object name -eq 'M365_ServicePlans.csv')
if ($rawTable.Count -ne 1 -or $servicePlanTable.Count -ne 1) {
    throw 'The raw contract must contain exactly one M365_SubscribedSkus.csv and M365_ServicePlans.csv definition.'
}
$rawTable = $rawTable[0]
$servicePlanTable = $servicePlanTable[0]

if ([string]::IsNullOrWhiteSpace($RawLatestOutputPath)) {
    $RawLatestOutputPath = Join-Path $paths.LatestOutputRootPath (
        Join-Path ([string]$rawTable.area) ([string]$rawTable.name)
    )
}
$RawLatestOutputPath = [IO.Path]::GetFullPath($RawLatestOutputPath)
$RawServicePlansLatestOutputPath = if ([string]::IsNullOrWhiteSpace($RawServicePlansLatestOutputPath)) {
    Join-Path $paths.LatestOutputRootPath (Join-Path ([string]$servicePlanTable.area) ([string]$servicePlanTable.name))
}
else { [IO.Path]::GetFullPath($RawServicePlansLatestOutputPath) }
if (($PSCmdlet.ParameterSetName -eq 'Fixture' -or $MaxItems -gt 0) -and
    @(@($RawLatestOutputPath,$RawServicePlansLatestOutputPath) | Where-Object {
        -not $_.StartsWith($paths.LatestOutputRootPath.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)
    }).Count -gt 0) {
    throw 'Raw latest output paths must stay inside the isolated latest output root.'
}
$executionMode = if ($ValidateOnly) { 'Validate' } elseif ($PSCmdlet.ParameterSetName -eq 'Fixture') { 'Fixture' } else { 'Collect' }
$runtimeContext = Start-SmartWorkplaceCMDBExecutionContext -Context $context -ScriptPath $PSCommandPath -ScriptVersion $ScriptVersion -Mode $executionMode -NoWrite:$ValidateOnly
$executionError = $null
$sourceRun = $null
$servicePlanSourceRun = $null
$sourceRunCompleted = $false
$servicePlanSourceRunCompleted = $false
try {
$sourceRun = Start-SmartWorkplaceCMDBSourceCollection -Paths $paths -RawPath @($RawLatestOutputPath) -Fixture:($PSCmdlet.ParameterSetName -eq 'Fixture') -MaxItems $MaxItems -NoWrite:$ValidateOnly
$servicePlanSourceRun = Start-SmartWorkplaceCMDBSourceCollection -Paths $paths -RawPath @($RawServicePlansLatestOutputPath) -Fixture:($PSCmdlet.ParameterSetName -eq 'Fixture') -MaxItems $MaxItems -NoWrite:$ValidateOnly


$graphConfiguration = Get-SmartWorkplaceCMDBConfigSection $context.Configuration 'MicrosoftGraph'
$collectionConfiguration = Get-SmartWorkplaceCMDBConfigSection $context.Configuration 'Collection'
$cloud = Get-SmartWorkplaceCMDBConfigText $collectionConfiguration 'Cloud'
if ([string]::IsNullOrWhiteSpace($cloud)) {
    $cloud = 'Public'
}
if ($PSCmdlet.ParameterSetName -eq 'Graph' -and $cloud -ne 'Public') {
    throw "The initial Microsoft 365 subscribed SKUs collector supports Collection.Cloud='Public' only. Configured value: '$cloud'."
}
$clientId = Get-SmartWorkplaceCMDBConfigText $graphConfiguration 'ClientId'
$thumbprint = Get-SmartWorkplaceCMDBConfigText $graphConfiguration 'CertificateThumbprint'
$fixtureSkus = $null
$readiness = $null
if ($PSCmdlet.ParameterSetName -eq 'Fixture') {
    $InputJsonPath = [IO.Path]::GetFullPath($InputJsonPath)
    $fixtureSkus = @(Read-SmartWorkplaceCMDBM365SubscribedSkusFixture -Path $InputJsonPath)
}
else {
    $readiness = Test-SmartWorkplaceCMDBGraphAppOnlyReadiness `
        -TenantId $paths.TenantId `
        -ClientId $clientId `
        -CertificateThumbprint $thumbprint
}

if ($ValidateOnly) {
    [pscustomobject]@{
        Status = 'Valid'; ScriptVersion = $ScriptVersion
        SourceMode = if ($PSCmdlet.ParameterSetName -eq 'Fixture') { 'OfflineJson' } else { 'MicrosoftGraphAppOnly' }
        FixtureSkuCount = if ($null -ne $fixtureSkus) { $fixtureSkus.Count } else { 0 }
        GraphAuthenticationModule = if ($null -ne $readiness) { $readiness.AuthenticationModule } else { '' }
        GraphAuthenticationVersion = if ($null -ne $readiness) { $readiness.AuthenticationModuleVersion } else { '' }
        RequiredGraphPermission = 'LicenseAssignment.Read.All'
        RawContractVersion = [string]$rawContract.contractVersion
        RawLatestOutputPath = $RawLatestOutputPath
        RawServicePlansLatestOutputPath = $RawServicePlansLatestOutputPath
        ProfileKey = $paths.ProfileKey; TenantKey = $paths.TenantKey
    } | Format-List
    return
}

$sourceSkus = if ($PSCmdlet.ParameterSetName -eq 'Fixture') {
    @($fixtureSkus)
}
else {
    $select = 'id,skuId,skuPartNumber,appliesTo,capabilityStatus,consumedUnits,prepaidUnits,servicePlans'
    $uri = 'https://graph.microsoft.com/v1.0/subscribedSkus?$select={0}' -f $select
    @(Invoke-SmartWorkplaceCMDBGraphPagedRequest `
        -TenantId $paths.TenantId `
        -ClientId $clientId `
        -CertificateThumbprint $thumbprint `
        -Uri $uri `
        -RequiredPermission 'LicenseAssignment.Read.All' `
        -MaxItems $MaxItems)
}
if ($MaxItems -gt 0) {
    $sourceSkus = @($sourceSkus | Select-Object -First $MaxItems)
}

$collectedDateTime = [datetime]::UtcNow.ToString('o')
$rawRows = @($sourceSkus | ForEach-Object {
    $subscribedSkuId = ConvertTo-SmartWorkplaceCMDBCleanText (
        Get-SmartWorkplaceCMDBGraphObjectValue $_ 'id'
    )
    $skuId = ConvertTo-SmartWorkplaceCMDBCleanText (
        Get-SmartWorkplaceCMDBGraphObjectValue $_ 'skuId'
    )
    $skuPartNumber = ConvertTo-SmartWorkplaceCMDBCleanText (
        Get-SmartWorkplaceCMDBGraphObjectValue $_ 'skuPartNumber'
    )
    if ([string]::IsNullOrWhiteSpace($subscribedSkuId) -or
        [string]::IsNullOrWhiteSpace($skuId) -or
        [string]::IsNullOrWhiteSpace($skuPartNumber)) {
        throw 'A subscribed SKU response did not contain id, skuId, and skuPartNumber.'
    }
    $prepaidUnits = Get-SmartWorkplaceCMDBGraphObjectValue $_ 'prepaidUnits'
    [pscustomobject][ordered]@{
        SourceSystem = 'MicrosoftGraph'
        SubscribedSkuId = $subscribedSkuId
        SkuId = $skuId
        SkuPartNumber = $skuPartNumber
        AppliesTo = ConvertTo-SmartWorkplaceCMDBCleanText (Get-SmartWorkplaceCMDBGraphObjectValue $_ 'appliesTo')
        CapabilityStatus = ConvertTo-SmartWorkplaceCMDBCleanText (Get-SmartWorkplaceCMDBGraphObjectValue $_ 'capabilityStatus')
        ConsumedUnits = ConvertTo-SmartWorkplaceCMDBLicenseInteger (Get-SmartWorkplaceCMDBGraphObjectValue $_ 'consumedUnits') 'consumedUnits' $subscribedSkuId
        EnabledUnits = ConvertTo-SmartWorkplaceCMDBLicenseInteger (Get-SmartWorkplaceCMDBGraphObjectValue $prepaidUnits 'enabled') 'prepaidUnits.enabled' $subscribedSkuId
        SuspendedUnits = ConvertTo-SmartWorkplaceCMDBLicenseInteger (Get-SmartWorkplaceCMDBGraphObjectValue $prepaidUnits 'suspended') 'prepaidUnits.suspended' $subscribedSkuId
        WarningUnits = ConvertTo-SmartWorkplaceCMDBLicenseInteger (Get-SmartWorkplaceCMDBGraphObjectValue $prepaidUnits 'warning') 'prepaidUnits.warning' $subscribedSkuId
        LockedOutUnits = ConvertTo-SmartWorkplaceCMDBLicenseInteger (Get-SmartWorkplaceCMDBGraphObjectValue $prepaidUnits 'lockedOut') 'prepaidUnits.lockedOut' $subscribedSkuId
        SourceCollectedDateTime = $collectedDateTime
    }
})
$duplicateIds = @($rawRows | Group-Object SubscribedSkuId | Where-Object Count -gt 1)
if ($duplicateIds.Count -gt 0) {
    throw "Duplicate Microsoft 365 subscribed SKU identifiers were returned: $($duplicateIds.Name -join ', ')"
}

$servicePlanRows = @($sourceSkus | ForEach-Object {
    $skuId = ConvertTo-SmartWorkplaceCMDBCleanText (Get-SmartWorkplaceCMDBGraphObjectValue $_ 'skuId')
    $skuPartNumber = ConvertTo-SmartWorkplaceCMDBCleanText (Get-SmartWorkplaceCMDBGraphObjectValue $_ 'skuPartNumber')
    foreach ($plan in @(Get-SmartWorkplaceCMDBGraphObjectValue $_ 'servicePlans')) {
        $servicePlanId = ConvertTo-SmartWorkplaceCMDBCleanText (Get-SmartWorkplaceCMDBGraphObjectValue $plan 'servicePlanId')
        $servicePlanName = ConvertTo-SmartWorkplaceCMDBCleanText (Get-SmartWorkplaceCMDBGraphObjectValue $plan 'servicePlanName')
        if ([string]::IsNullOrWhiteSpace($skuId) -or [string]::IsNullOrWhiteSpace($servicePlanId) -or
            [string]::IsNullOrWhiteSpace($servicePlanName)) {
            throw 'A subscribed SKU service plan did not contain skuId, servicePlanId, and servicePlanName.'
        }
        [pscustomobject][ordered]@{
            SourceSystem = 'MicrosoftGraph'
            SkuId = $skuId
            SkuPartNumber = $skuPartNumber
            ServicePlanId = $servicePlanId
            ServicePlanName = $servicePlanName
            ProvisioningStatus = ConvertTo-SmartWorkplaceCMDBCleanText (Get-SmartWorkplaceCMDBGraphObjectValue $plan 'provisioningStatus')
            AppliesTo = ConvertTo-SmartWorkplaceCMDBCleanText (Get-SmartWorkplaceCMDBGraphObjectValue $plan 'appliesTo')
            SourceCollectedDateTime = $collectedDateTime
        }
    }
})
$duplicateServicePlans = @($servicePlanRows | Group-Object {
    '{0}|{1}' -f ([string]$_.SkuId).ToLowerInvariant(),([string]$_.ServicePlanId).ToLowerInvariant()
} | Where-Object Count -gt 1)
if ($duplicateServicePlans.Count -gt 0) {
    throw "Duplicate Microsoft 365 service-plan keys were returned: $($duplicateServicePlans.Name -join ', ')"
}

$historyTimestamp = [datetime]::UtcNow
$historyFolder = Join-Path $paths.DataAllRootPath (
    'M365\SubscribedSkus\{0}\{1}' -f $historyTimestamp.ToString('yyyy'), $historyTimestamp.ToString('MM')
)
$historyPath = Join-Path $historyFolder (
    'M365_SubscribedSkus_{0}.csv' -f $historyTimestamp.ToString('yyyyMMdd-HHmmssfff')
)
$servicePlanHistoryPath = Join-Path $historyFolder (
    'M365_ServicePlans_{0}.csv' -f $historyTimestamp.ToString('yyyyMMdd-HHmmssfff')
)
Publish-SmartWorkplaceCMDBSourceCsv `
    -Run $sourceRun `
    -InputObject $rawRows `
    -Columns @($rawTable.columns | ForEach-Object { [string]$_ }) `
    -HistoryPath $historyPath `
    -LatestPath $RawLatestOutputPath `
    -ContractPath $rawContractPath `
    -ContractTableName 'M365_SubscribedSkus.csv' | Out-Null
$sourceRunCompleted = $true
Publish-SmartWorkplaceCMDBSourceCsv `
    -Run $servicePlanSourceRun `
    -InputObject $servicePlanRows `
    -Columns @($servicePlanTable.columns | ForEach-Object { [string]$_ }) `
    -HistoryPath $servicePlanHistoryPath `
    -LatestPath $RawServicePlansLatestOutputPath `
    -ContractPath $rawContractPath `
    -ContractTableName 'M365_ServicePlans.csv' | Out-Null
$servicePlanSourceRunCompleted = $true

Write-Information (
    "SmartWorkplaceCMDB Microsoft 365 subscribed SKUs collection completed. SKUs={0}; ServicePlans={1}; history='{2}'; latest='{3}'." -f
    $rawRows.Count, $servicePlanRows.Count, $historyPath, $RawLatestOutputPath
) -InformationAction Continue
[pscustomobject]@{
    Status = 'Completed'; ScriptVersion = $ScriptVersion; SkuCount = $rawRows.Count
    CollectedDateTime = $collectedDateTime; HistoryPath = $historyPath
    RawLatestOutputPath = $RawLatestOutputPath
    ServicePlanCount = $servicePlanRows.Count
    ServicePlanHistoryPath = $servicePlanHistoryPath
    RawServicePlansLatestOutputPath = $RawServicePlansLatestOutputPath
    RawContractVersion = [string]$rawContract.contractVersion
    RequiredGraphPermission = 'LicenseAssignment.Read.All'
}

} catch {
    $executionError = $_
    if ($null -ne $sourceRun -and -not $sourceRunCompleted) { Complete-SmartWorkplaceCMDBSourceCollection -Run $sourceRun -Failed }
    if ($null -ne $servicePlanSourceRun -and -not $servicePlanSourceRunCompleted) { Complete-SmartWorkplaceCMDBSourceCollection -Run $servicePlanSourceRun -Failed }
    throw
} finally {
    Complete-SmartWorkplaceCMDBExecutionContext -RuntimeContext $runtimeContext -ErrorRecord $executionError
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAumIzb8PFGr9z8
# rFqUElWI3ygTcQuwIZALfDn5mK3ufaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAb+iJzGiy3Y062HpEciIvhuoSR3sM4BE2Ft9VvupZBFMA0GCSqG
# SIb3DQEBAQUABIIBgDSQB0AIxGPB6AhB2f6xDInloJMyd6gfdHnFQNtOQH4Y7jkv
# 5ggpcLPEpSElAQQkiAea1IRJwoBbGMLipj4FUFoDHpjGyvE0PxqPAOjUnWRTB7hR
# 4UNbfZoja+njbHIAeAm/AUy/ZV7fuYWdY5WTm69AwDt7bYFaca0n4AkX/shRZVVR
# sdaSHqAZkvrDSJCw6yYwEBTbhGBAQlUi816VFSfVfHHo+6PNOSCD9tH7rpw61Nlh
# PvPvbcT+hhc5veGUe9YsEMJIQ3EPrbyAlmIHUWtnRtqD6SU9x3o+/0I9HzvF0YrP
# WsjlhrSfhEANzn7VnYcmJk+zCKRjodeumd2/lO4CRgVpTZbd34nzq5FcbakVXt8Q
# QMpDiNyAC1NXJOtRuAbVLzVZUyp8aSYtERp2ITdAcYPidTkJA9VYW7Drv2WqcpQl
# 44HID94RZ4VTqmhH8dT+g9RLl4Cn7vmcR2h7kSUMD7luAfSvcqXIquNXe2WJeCTR
# uabZ0agXVaeZayTjZaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxNzQ0
# MjdaMC8GCSqGSIb3DQEJBDEiBCB3s71D9442JwCrx/u3wJZobbekXDMScXQalO3c
# jZxiYTANBgkqhkiG9w0BAQEFAASCAgAHvvsiSck0fLeRWjAUHlgRa2J9cPeci5vg
# EuDT6QGjo6qtLEPiyog6b+gSVCqqykDyhOWhlDRWcJuveE9n46BjQ5tuz0FIgrKk
# GRqwg5iMDzaUI8oBR+yifHi43Sn3FOv/BkLxk2QwEpn6AANz41qprEYsIIUCY6cp
# q3OzF5MzOLXmNE0KgU8UNzpnh1jCPtjDcMx4VILzLxG+mSZ5SEUxXMegLWyYMtgx
# gbgAKUDWpYoba5XvoFP4ufnZyJB8y3C/0MI1hBWrMaAfqxG0S0AekbFfHafVvpg6
# FqnmkN/IqhCT8WKy8HWnV46sZB9+fbqzLbDpTb3snay41+AWjvhWohnRupmazP3G
# KvlD6z97Urcifa1o3KY1Uwt7CP0jGiblydRxXv6XttU12s+XtAA6CeX/p0WCvoTP
# INbnCS1L3QZPr3SV0N1qwzySGLelSAul+7wsDhAZHdZIYM22zljDDfWTxYZXivsF
# mnOrkCv2pwdPQ6PjGFuaR3rMV612JwO7RdyKPBDa5fuFtZBP0Q4W1athYPnhe1t2
# 5cdQsgOSGaAvJjpkVUkwOdrh0BsaILEQR+5Cwx0IFc/0fvmYF+QY2kv0jQ0uK/Df
# UFSX3Tww0xMjzHh7TfuQKm+xGXiwg6Bl+w84Hmtr9+yJ0KqC/Jwrd1/o7TegX0WV
# UXkblQ2HtA==
# SIG # End signature block
