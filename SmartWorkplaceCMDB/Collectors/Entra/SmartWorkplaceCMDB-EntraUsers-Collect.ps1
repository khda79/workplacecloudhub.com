<#
.SYNOPSIS
Collects a read-only Microsoft Entra user snapshot for SmartWorkplaceCMDB.

.DESCRIPTION
Uses Microsoft Graph app-only certificate authentication to retrieve the base
Entra user identity properties required by SmartWorkplaceCMDB. Raw snapshots
are written to DATA-ALL and the latest raw contract is written below
DATA-LAST\Raw\Entra. An offline JSON input is supported for safe tests.

.VERSION
1.2.1

.REQUIREMENTS
PowerShell 5.1 or later.
Microsoft.Graph.Authentication for live collection.
Microsoft Graph application permissions: User.Read.All and AuditLog.Read.All.
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

$ScriptVersion = '1.2.1'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Get-SmartWorkplaceCMDBObjectValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $InputObject,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }
    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name)) {
            return $InputObject[$Name]
        }
        return $null
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $property) {
        return $property.Value
    }
    return $null
}

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

function Get-SmartWorkplaceCMDBGraphSetting {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$GraphConfiguration,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if ($GraphConfiguration.Contains($Name)) {
        return ([string]$GraphConfiguration[$Name]).Trim()
    }
    return ''
}

function ConvertTo-SmartWorkplaceCMDBCleanText {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return ''
    }
    return ([string]$Value -replace "`r`n|`n|`r", ' ').Trim()
}

function ConvertTo-SmartWorkplaceCMDBDateTimeText {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return '' }
    if ($Value -is [datetimeoffset]) {
        return $Value.ToUniversalTime().ToString('o')
    }
    if ($Value -is [datetime]) {
        return ([datetimeoffset]$Value).ToUniversalTime().ToString('o')
    }
    return ConvertTo-SmartWorkplaceCMDBCleanText $Value
}

function Read-SmartWorkplaceCMDBEntraUsersFixture {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    return @(Read-SmartWorkplaceCMDBCollectionFixture -Path $Path)
}

function Test-SmartWorkplaceCMDBGraphReadiness {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$GraphConfiguration,

        [Parameter(Mandatory)]
        [string]$ResolvedTenantId
    )

    $clientId = Get-SmartWorkplaceCMDBGraphSetting -GraphConfiguration $GraphConfiguration -Name 'ClientId'
    $thumbprint = Get-SmartWorkplaceCMDBGraphSetting -GraphConfiguration $GraphConfiguration -Name 'CertificateThumbprint'
    if ([string]::IsNullOrWhiteSpace($ResolvedTenantId)) {
        throw 'MicrosoftGraph.TenantId is required for live Entra collection.'
    }

    $parsedClientId = [guid]::Empty
    if (-not [guid]::TryParse($clientId, [ref]$parsedClientId)) {
        throw 'MicrosoftGraph.ClientId must contain the application registration GUID for live Entra collection.'
    }
    if ($thumbprint -notmatch '^[a-fA-F0-9]{40,64}$') {
        throw 'MicrosoftGraph.CertificateThumbprint must contain a 40 to 64 character hexadecimal certificate thumbprint.'
    }

    $authenticationModule = Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication' |
        Sort-Object Version -Descending |
        Select-Object -First 1
    if ($null -eq $authenticationModule) {
        throw 'Microsoft.Graph.Authentication is required for live Entra collection.'
    }

    $certificate = Get-ChildItem -Path Cert:\CurrentUser\My |
        Where-Object Thumbprint -eq $thumbprint |
        Select-Object -First 1
    if ($null -eq $certificate) {
        throw 'The configured Microsoft Graph certificate was not found in Cert:\CurrentUser\My.'
    }
    if (-not $certificate.HasPrivateKey) {
        throw 'The configured Microsoft Graph certificate does not have an accessible private key.'
    }
    if ($certificate.NotAfter -le (Get-Date)) {
        throw 'The configured Microsoft Graph certificate is expired.'
    }

    return [pscustomobject]@{
        ClientId                  = $parsedClientId.ToString()
        CertificateThumbprint     = $certificate.Thumbprint
        AuthenticationModule      = $authenticationModule.Name
        AuthenticationModuleVersion = $authenticationModule.Version.ToString()
    }
}

function Get-SmartWorkplaceCMDBEntraUsersFromGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$GraphConfiguration,

        [Parameter(Mandatory)]
        [string]$ResolvedTenantId,

        [ValidateRange(0, 2147483647)]
        [int]$Limit = 0
    )

    $select = 'id,userPrincipalName,displayName,accountEnabled,userType,department,jobTitle,usageLocation,createdDateTime,signInActivity'
    $uri = 'https://graph.microsoft.com/v1.0/users?$select={0}&$top=999' -f $select
    return @(Invoke-SmartWorkplaceCMDBGraphPagedRequest `
            -TenantId $ResolvedTenantId `
            -ClientId ([string](Get-SmartWorkplaceCMDBGraphSetting `
                    -GraphConfiguration $GraphConfiguration `
                    -Name 'ClientId')) `
            -CertificateThumbprint ([string](Get-SmartWorkplaceCMDBGraphSetting `
                    -GraphConfiguration $GraphConfiguration `
                    -Name 'CertificateThumbprint')) `
            -Uri $uri `
            -RequiredPermission 'User.Read.All;AuditLog.Read.All' `
            -MaxItems $Limit)
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $scriptRoot)
$modulePath = Join-Path -Path $projectRoot -ChildPath 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$graphModulePath = Join-Path -Path $projectRoot -ChildPath 'Modules\SmartWorkplaceCMDB.Graph\SmartWorkplaceCMDB.Graph.psd1'
$rawContractPath = Join-Path -Path $projectRoot -ChildPath 'Schema\SmartWorkplaceCMDB.raw.tables.json'

Import-Module $modulePath -Force
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
$rawTable = @($rawContract.tables | Where-Object name -eq 'Entra_Users.csv')
if ($rawTable.Count -ne 1) {
    throw 'The raw contract must contain exactly one Entra_Users.csv definition.'
}
$rawTable = $rawTable[0]

if ([string]::IsNullOrWhiteSpace($RawLatestOutputPath)) {
    $RawLatestOutputPath = Join-Path -Path $paths.LatestOutputRootPath -ChildPath (
        Join-Path -Path ([string]$rawTable.area) -ChildPath ([string]$rawTable.name)
    )
}
$RawLatestOutputPath = [System.IO.Path]::GetFullPath($RawLatestOutputPath)
if (($PSCmdlet.ParameterSetName -eq 'Fixture' -or $MaxItems -gt 0) -and
    -not $RawLatestOutputPath.StartsWith($paths.LatestOutputRootPath.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'RawLatestOutputPath must stay inside the isolated latest output root.'
}
$executionMode = if ($ValidateOnly) { 'Validate' } elseif ($PSCmdlet.ParameterSetName -eq 'Fixture') { 'Fixture' } else { 'Collect' }
$runtimeContext = Start-SmartWorkplaceCMDBExecutionContext -Context $context -ScriptPath $PSCommandPath -ScriptVersion $ScriptVersion -Mode $executionMode -NoWrite:$ValidateOnly
$executionError = $null
$sourceRun = $null
try {
$sourceRun = Start-SmartWorkplaceCMDBSourceCollection -Paths $paths -RawPath @($RawLatestOutputPath) -Fixture:($PSCmdlet.ParameterSetName -eq 'Fixture') -MaxItems $MaxItems -NoWrite:$ValidateOnly


$graphConfiguration = Get-SmartWorkplaceCMDBConfigSection -Configuration $context.Configuration -Name 'MicrosoftGraph'
$collectionConfiguration = Get-SmartWorkplaceCMDBConfigSection -Configuration $context.Configuration -Name 'Collection'
$cloud = Get-SmartWorkplaceCMDBGraphSetting -GraphConfiguration $collectionConfiguration -Name 'Cloud'
if ([string]::IsNullOrWhiteSpace($cloud)) {
    $cloud = 'Public'
}
if ($PSCmdlet.ParameterSetName -eq 'Graph' -and $cloud -ne 'Public') {
    throw "The initial Entra users collector supports Collection.Cloud='Public' only. Configured value: '$cloud'."
}

$fixtureUsers = $null
$readiness = $null
if ($PSCmdlet.ParameterSetName -eq 'Fixture') {
    $InputJsonPath = [System.IO.Path]::GetFullPath($InputJsonPath)
    $fixtureUsers = @(Read-SmartWorkplaceCMDBEntraUsersFixture -Path $InputJsonPath)
}
else {
    $readiness = Test-SmartWorkplaceCMDBGraphReadiness -GraphConfiguration $graphConfiguration -ResolvedTenantId $paths.TenantId
}

if ($ValidateOnly) {
    [pscustomobject]@{
        Status                      = 'Valid'
        ScriptVersion               = $ScriptVersion
        SourceMode                  = if ($PSCmdlet.ParameterSetName -eq 'Fixture') { 'OfflineJson' } else { 'MicrosoftGraphAppOnly' }
        FixtureUserCount            = if ($null -ne $fixtureUsers) { $fixtureUsers.Count } else { 0 }
        GraphAuthenticationModule   = if ($null -ne $readiness) { $readiness.AuthenticationModule } else { '' }
        GraphAuthenticationVersion  = if ($null -ne $readiness) { $readiness.AuthenticationModuleVersion } else { '' }
        RequiredGraphPermission     = 'User.Read.All;AuditLog.Read.All'
        Cloud                       = $cloud
        RawContractVersion          = [string]$rawContract.contractVersion
        RawLatestOutputPath         = $RawLatestOutputPath
        ProfileKey                  = $paths.ProfileKey
        TenantKey                   = $paths.TenantKey
        TenantId                    = $paths.TenantId
    } | Format-List
    return
}

$sourceUsers = if ($PSCmdlet.ParameterSetName -eq 'Fixture') {
    @($fixtureUsers)
}
else {
    @(Get-SmartWorkplaceCMDBEntraUsersFromGraph `
        -GraphConfiguration $graphConfiguration `
        -ResolvedTenantId $paths.TenantId `
        -Limit $MaxItems)
}

if ($MaxItems -gt 0) {
    $sourceUsers = @($sourceUsers | Select-Object -First $MaxItems)
}

$collectedDateTime = [datetime]::UtcNow.ToString('o')
$rawRows = @($sourceUsers | ForEach-Object {
    $sourceUserId = ConvertTo-SmartWorkplaceCMDBCleanText (
        Get-SmartWorkplaceCMDBObjectValue -InputObject $_ -Name 'id'
    )
    if ([string]::IsNullOrWhiteSpace($sourceUserId)) {
        throw 'A Microsoft Entra user response did not contain an id.'
    }

    $accountEnabledValue = Get-SmartWorkplaceCMDBObjectValue -InputObject $_ -Name 'accountEnabled'
    $signInActivity = Get-SmartWorkplaceCMDBObjectValue -InputObject $_ -Name 'signInActivity'
    [pscustomobject][ordered]@{
        SourceSystem           = 'MicrosoftEntraID'
        SourceUserId           = $sourceUserId
        UserPrincipalName      = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue -InputObject $_ -Name 'userPrincipalName'
        )
        DisplayName            = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue -InputObject $_ -Name 'displayName'
        )
        AccountEnabled         = if ($null -eq $accountEnabledValue) { '' } else { [bool]$accountEnabledValue }
        UserType               = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue -InputObject $_ -Name 'userType'
        )
        Department             = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue -InputObject $_ -Name 'department'
        )
        JobTitle               = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue -InputObject $_ -Name 'jobTitle'
        )
        UsageLocation          = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue -InputObject $_ -Name 'usageLocation'
        )
        CreatedDateTime        = ConvertTo-SmartWorkplaceCMDBDateTimeText (
            Get-SmartWorkplaceCMDBObjectValue -InputObject $_ -Name 'createdDateTime'
        )
        LastSignInDateTime     = ConvertTo-SmartWorkplaceCMDBDateTimeText (
            Get-SmartWorkplaceCMDBObjectValue -InputObject $signInActivity -Name 'lastSignInDateTime'
        )
        LastNonInteractiveSignInDateTime = ConvertTo-SmartWorkplaceCMDBDateTimeText (
            Get-SmartWorkplaceCMDBObjectValue -InputObject $signInActivity -Name 'lastNonInteractiveSignInDateTime'
        )
        LastSuccessfulSignInDateTime = ConvertTo-SmartWorkplaceCMDBDateTimeText (
            Get-SmartWorkplaceCMDBObjectValue -InputObject $signInActivity -Name 'lastSuccessfulSignInDateTime'
        )
        SourceCollectedDateTime = $collectedDateTime
    }
})

$duplicateSourceIds = @($rawRows |
    Group-Object SourceUserId |
    Where-Object Count -gt 1 |
    Select-Object -ExpandProperty Name)
if ($duplicateSourceIds.Count -gt 0) {
    throw "Duplicate Microsoft Entra user ids were returned: $($duplicateSourceIds -join ', ')"
}

$historyTimestamp = [datetime]::UtcNow
$historyFolder = Join-Path -Path $paths.DataAllRootPath -ChildPath (
    'Entra\Users\{0}\{1}' -f $historyTimestamp.ToString('yyyy'), $historyTimestamp.ToString('MM')
)
$historyName = 'Entra_Users_{0}.csv' -f $historyTimestamp.ToString('yyyyMMdd-HHmmssfff')
$historyPath = Join-Path -Path $historyFolder -ChildPath $historyName
Publish-SmartWorkplaceCMDBSourceCsv `
    -Run $sourceRun `
    -InputObject $rawRows `
    -Columns @($rawTable.columns | ForEach-Object { [string]$_ }) `
    -HistoryPath $historyPath `
    -LatestPath $RawLatestOutputPath `
    -ContractPath $rawContractPath `
    -ContractTableName 'Entra_Users.csv' | Out-Null

Write-Information (
    "SmartWorkplaceCMDB Entra users collection completed. Users={0}; history='{1}'; latest='{2}'." -f
    $rawRows.Count,
    $historyPath,
    $RawLatestOutputPath
) -InformationAction Continue

[pscustomobject]@{
    Status                  = 'Completed'
    ScriptVersion           = $ScriptVersion
    SourceMode              = if ($PSCmdlet.ParameterSetName -eq 'Fixture') { 'OfflineJson' } else { 'MicrosoftGraphAppOnly' }
    UserCount               = $rawRows.Count
    CollectedDateTime       = $collectedDateTime
    HistoryPath             = $historyPath
    RawLatestOutputPath     = $RawLatestOutputPath
    RawContractVersion      = [string]$rawContract.contractVersion
    RequiredGraphPermission = 'User.Read.All;AuditLog.Read.All'
}

} catch {
    $executionError = $_
    if ($null -ne $sourceRun) { Complete-SmartWorkplaceCMDBSourceCollection -Run $sourceRun -Failed }
    throw
} finally {
    Complete-SmartWorkplaceCMDBExecutionContext -RuntimeContext $runtimeContext -ErrorRecord $executionError
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCMRkDC1q3RvkG+
# Vj99z1qCQU0gwMdXp0CRbPrhNJN8K6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIuuPNzwhiA7SNwf5UVEbqwVm90Pa2OPx6HEIXparFZtMA0GCSqG
# SIb3DQEBAQUABIIBgE+APKSmGWJxLOCBP/uellwQ71X+3giaAW0JM4TvHcPbxJQe
# wCZhh0pyPf6ynX+gdv6uWPX52H67YAWy4zMQzU3NokdiOPsatHO2q+I0Oi1FYn58
# Cj64VV176ZaLD/YaXHXu5jE3jAzGsdbHKW8A1BoBKOr00/TiELstgCyx+gthvgHI
# lQAegkGTbSNq37EjJEv8FYBpUIf3e4PBKf5nutubHIG3948sshNuiuXZymWG8WEg
# robkru/Ee79ZREuwNptuvuN9BwwEO7+qQDW4+ktLKYYk9VAbhoXyZkhjUhzcO1Tf
# PILSUzz2eXguHVJwEdj8mkDa5I/osQxc/55Cb04KFaWibKE1vd+hXong2NaP1U8e
# p5V8HbtJWMb33MLOIchb3FDpl6f+v+3rHONtYDx3WYM3MHl90ASZlAS5YvVK5RdP
# 5dCgMtY9Oayod8li5PmQO3b8XR/7hk3yyTCKXaKzgbslQBFEcGROunzq8e66UA5x
# W8czZMZV5X/XBtkbi6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIyMDI1
# MjFaMC8GCSqGSIb3DQEJBDEiBCDndlXY+rNIMcaZduUk52lbq3YoEG2G2NieRDiT
# s0uOgDANBgkqhkiG9w0BAQEFAASCAgBrSNdXIVMqmlZdGW0zVLJ1YhFFklFAz2lH
# GReAPHzzxIXQaG5AaNW4efDVHMmCFnttl1J8m38+TpSq+8xWgZv5U03YBTLG3H6r
# 1QEf0kbOE8TGbsvmfnx5UnQnsUpeEx4ily10Z1ONIo6WnbxCuwUC5XkI0SOuv/xJ
# +P90airGeB4aD3NuY+AKOmeJnmqswMbis/OydjdWQrIH6+Q95I3N2XB2IpnfMGOF
# Avh+seEwmLGE00jKjXz/H5fsBqXpyLmje0/fgQgvSNMQv1FNnwHbmXqHL5l03j+U
# Or2msCXYHjCQclcGCyJRWBYvFBgkHzCxd+Q16ojEmce5FmP/PEs0+JFKglnQ46zs
# wtIlwJv3gR6NO/H9RHFnxX1KEqzRq7hquPffIP/le08L8/H7CMV4yaHF7ehMUHEN
# Xj3YGUOLUBSReEWWdVesZdzZFpG82vhhOQdrs9DWyNH2sPv5C+whIWHEzbQfN4Sy
# 1r/lmmF5Tn7+6GKcCXHAcDIiHPSrTYz3WZKzcLMD1EZMggk/gzx+eeDN9hqVmCEH
# 8vj3cIHsTKuB0fiBGj+R9ckEFZIoIcqKon06uOg03ktE9HOo0e+in4H/MqxDIUWA
# vpSMNjtyECOCjLWITWEVrfocZOsVZfOcdDFXOw3yGdVrpd9t3WkjvAbMs5L/JIga
# JkHK+JhkXQ==
# SIG # End signature block
