<#
.SYNOPSIS
Collects Exchange Online mailboxes for SmartWorkplaceCMDB.

.DESCRIPTION
Uses unattended certificate authentication and Get-EXOMailbox to collect the
base mailbox inventory. The collector is read-only and writes history plus the
latest raw contract. Offline JSON is supported for safe tests.

.VERSION
0.1.0

.REQUIREMENTS
PowerShell 5.1 or later.
ExchangeOnlineManagement 2.0.3 or later for live collection.
Exchange.ManageAsApp with Exchange read-only recipient RBAC.
#>
[CmdletBinding(DefaultParameterSetName = 'ExchangeOnline')]
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
    [switch]$NoConfigWrite,
    [switch]$ValidateOnly
)

$ScriptVersion = '0.1.0'
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

function Get-SmartWorkplaceCMDBExchangeObjectValue {
    [CmdletBinding()]
    param(
        [AllowNull()]$InputObject,
        [Parameter(Mandatory)][string]$Name
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

function ConvertTo-SmartWorkplaceCMDBExchangeText {
    [CmdletBinding()]
    param([AllowNull()]$Value)
    if ($null -eq $Value) {
        return ''
    }
    return ([string]$Value -replace "`r`n|`n|`r", ' ').Trim()
}

function Read-SmartWorkplaceCMDBExchangeMailboxFixture {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $document = Read-SmartWorkplaceCMDBJsonFile -Path $Path
    $value = Get-SmartWorkplaceCMDBExchangeObjectValue -InputObject $document -Name 'value'
    $hasValueProperty = (
        ($document -is [System.Collections.IDictionary] -and $document.Contains('value')) -or
        ($null -ne $document.PSObject.Properties['value'])
    )
    if ($hasValueProperty) {
        return @($value)
    }
    if ($document -is [System.Collections.IEnumerable] -and
        $document -isnot [string]) {
        return @($document)
    }
    throw "Offline Exchange Online mailbox JSON must be an array or contain a value array: $Path"
}

function Test-SmartWorkplaceCMDBExchangeOnlineReadiness {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AppId,
        [Parameter(Mandatory)][string]$CertificateThumbprint,
        [Parameter(Mandatory)][string]$Organization
    )
    $parsedAppId = [guid]::Empty
    if (-not [guid]::TryParse($AppId, [ref]$parsedAppId)) {
        throw 'MicrosoftGraph.ClientId must contain the application GUID used by Exchange Online.'
    }
    if ($CertificateThumbprint -notmatch '^[a-fA-F0-9]{40,64}$') {
        throw 'MicrosoftGraph.CertificateThumbprint must contain a 40 to 64 character hexadecimal thumbprint.'
    }
    if ($Organization -notmatch '^[a-zA-Z0-9][a-zA-Z0-9.-]*\.onmicrosoft\.com$') {
        throw 'ExchangeOnline.Organization must contain the primary onmicrosoft.com domain.'
    }
    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
        Sort-Object Version -Descending |
        Select-Object -First 1
    if ($null -eq $module -or $module.Version -lt [version]'2.0.3') {
        throw 'ExchangeOnlineManagement 2.0.3 or later is required for live mailbox collection.'
    }
    $certificate = Get-ChildItem -Path Cert:\CurrentUser\My |
        Where-Object Thumbprint -eq $CertificateThumbprint |
        Select-Object -First 1
    if ($null -eq $certificate) {
        throw 'The configured certificate was not found in Cert:\CurrentUser\My.'
    }
    if (-not $certificate.HasPrivateKey) {
        throw 'The configured certificate does not have an accessible private key.'
    }
    $now = Get-Date
    if ($certificate.NotBefore -gt $now -or $certificate.NotAfter -le $now) {
        throw 'The configured certificate is not currently valid.'
    }
    return [pscustomobject]@{
        AppId                    = $parsedAppId.ToString()
        CertificateThumbprint    = $certificate.Thumbprint
        Organization             = $Organization.ToLowerInvariant()
        ExchangeModule           = $module.Name
        ExchangeModuleVersion    = $module.Version.ToString()
    }
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $scriptRoot)
$coreModulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$rawContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
Import-Module $coreModulePath -Force

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
$rawTable = @($rawContract.tables | Where-Object name -eq 'ExchangeOnline_Mailboxes.csv')
if ($rawTable.Count -ne 1) {
    throw 'The raw contract must contain exactly one ExchangeOnline_Mailboxes.csv definition.'
}
$rawTable = $rawTable[0]

if ([string]::IsNullOrWhiteSpace($RawLatestOutputPath)) {
    $RawLatestOutputPath = Join-Path $paths.LatestOutputRootPath (
        Join-Path ([string]$rawTable.area) ([string]$rawTable.name)
    )
}
$RawLatestOutputPath = [IO.Path]::GetFullPath($RawLatestOutputPath)

$fixtureMailboxes = $null
$readiness = $null
if ($PSCmdlet.ParameterSetName -eq 'Fixture') {
    $InputJsonPath = [IO.Path]::GetFullPath($InputJsonPath)
    $fixtureMailboxes = @(Read-SmartWorkplaceCMDBExchangeMailboxFixture -Path $InputJsonPath)
}
else {
    $exchangeConfiguration = Get-SmartWorkplaceCMDBConfigSection $context.Configuration 'ExchangeOnline'
    $graphConfiguration = Get-SmartWorkplaceCMDBConfigSection $context.Configuration 'MicrosoftGraph'
    if ($exchangeConfiguration.Contains('Enabled') -and
        -not [bool]$exchangeConfiguration['Enabled']) {
        throw 'ExchangeOnline.Enabled must be true for live mailbox collection.'
    }
    $organization = Get-SmartWorkplaceCMDBConfigText $exchangeConfiguration 'Organization'
    $appId = Get-SmartWorkplaceCMDBConfigText $graphConfiguration 'ClientId'
    $thumbprint = Get-SmartWorkplaceCMDBConfigText $graphConfiguration 'CertificateThumbprint'
    $readiness = Test-SmartWorkplaceCMDBExchangeOnlineReadiness `
        -AppId $appId `
        -CertificateThumbprint $thumbprint `
        -Organization $organization
}

if ($ValidateOnly) {
    [pscustomobject]@{
        Status                   = 'Valid'
        ScriptVersion            = $ScriptVersion
        SourceMode               = if ($PSCmdlet.ParameterSetName -eq 'Fixture') { 'OfflineJson' } else { 'ExchangeOnlineAppOnly' }
        FixtureMailboxCount      = if ($null -ne $fixtureMailboxes) { $fixtureMailboxes.Count } else { 0 }
        ExchangeModule           = if ($null -ne $readiness) { $readiness.ExchangeModule } else { '' }
        ExchangeModuleVersion    = if ($null -ne $readiness) { $readiness.ExchangeModuleVersion } else { '' }
        RequiredAppPermission    = 'Exchange.ManageAsApp'
        RequiredExchangeRole     = 'View-Only Recipients'
        RawContractVersion       = [string]$rawContract.contractVersion
        RawLatestOutputPath      = $RawLatestOutputPath
        ProfileKey               = $paths.ProfileKey
        TenantKey                = $paths.TenantKey
    } | Format-List
    return
}

$sourceMailboxes = @()
if ($PSCmdlet.ParameterSetName -eq 'Fixture') {
    $sourceMailboxes = @($fixtureMailboxes)
}
else {
    Import-Module ExchangeOnlineManagement -MinimumVersion 2.0.3 -ErrorAction Stop
    $connected = $false
    try {
        Connect-ExchangeOnline `
            -AppId $readiness.AppId `
            -CertificateThumbprint $readiness.CertificateThumbprint `
            -Organization $readiness.Organization `
            -ShowBanner:$false `
            -ErrorAction Stop | Out-Null
        $connected = $true
        $getParameters = @{
            ResultSize = if ($MaxItems -gt 0) { [string]$MaxItems } else { 'Unlimited' }
            Properties = @(
                'ExternalDirectoryObjectId',
                'ExchangeGuid',
                'UserPrincipalName',
                'DisplayName',
                'RecipientTypeDetails',
                'PrimarySmtpAddress',
                'MailboxPlan',
                'ArchiveStatus'
            )
            ErrorAction = 'Stop'
        }
        $sourceMailboxes = @(Get-EXOMailbox @getParameters)
    }
    catch {
        if ($_.Exception.Message -match '403|Forbidden|Unauthorized|not authorized|permission|RBAC') {
            throw 'Exchange Online denied the mailbox query. Verify Exchange.ManageAsApp admin consent and read-only recipient RBAC for the service principal.'
        }
        throw
    }
    finally {
        if ($connected -and
            (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue)) {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue |
                Out-Null
        }
    }
}
if ($MaxItems -gt 0) {
    $sourceMailboxes = @($sourceMailboxes | Select-Object -First $MaxItems)
}

$collectedDateTime = [datetime]::UtcNow.ToString('o')
$rawRows = @($sourceMailboxes | ForEach-Object {
    $externalId = ConvertTo-SmartWorkplaceCMDBExchangeText (
        Get-SmartWorkplaceCMDBExchangeObjectValue $_ 'ExternalDirectoryObjectId'
    )
    $exchangeGuid = ConvertTo-SmartWorkplaceCMDBExchangeText (
        Get-SmartWorkplaceCMDBExchangeObjectValue $_ 'ExchangeGuid'
    )
    $sourceMailboxId = if (-not [string]::IsNullOrWhiteSpace($externalId)) {
        'entra-object:{0}' -f $externalId.ToLowerInvariant()
    }
    elseif (-not [string]::IsNullOrWhiteSpace($exchangeGuid)) {
        'exchange-guid:{0}' -f $exchangeGuid.ToLowerInvariant()
    }
    else {
        throw 'An Exchange Online mailbox did not contain ExternalDirectoryObjectId or ExchangeGuid.'
    }
    [pscustomobject][ordered]@{
        SourceSystem              = 'ExchangeOnline'
        SourceMailboxId           = $sourceMailboxId
        ExternalDirectoryObjectId = $externalId
        ExchangeGuid              = $exchangeGuid
        UserPrincipalName         = ConvertTo-SmartWorkplaceCMDBExchangeText (
            Get-SmartWorkplaceCMDBExchangeObjectValue $_ 'UserPrincipalName'
        )
        DisplayName               = ConvertTo-SmartWorkplaceCMDBExchangeText (
            Get-SmartWorkplaceCMDBExchangeObjectValue $_ 'DisplayName'
        )
        RecipientTypeDetails      = ConvertTo-SmartWorkplaceCMDBExchangeText (
            Get-SmartWorkplaceCMDBExchangeObjectValue $_ 'RecipientTypeDetails'
        )
        PrimarySmtpAddress        = ConvertTo-SmartWorkplaceCMDBExchangeText (
            Get-SmartWorkplaceCMDBExchangeObjectValue $_ 'PrimarySmtpAddress'
        )
        MailboxPlan               = ConvertTo-SmartWorkplaceCMDBExchangeText (
            Get-SmartWorkplaceCMDBExchangeObjectValue $_ 'MailboxPlan'
        )
        ArchiveStatus             = ConvertTo-SmartWorkplaceCMDBExchangeText (
            Get-SmartWorkplaceCMDBExchangeObjectValue $_ 'ArchiveStatus'
        )
        SourceCollectedDateTime   = $collectedDateTime
    }
})
$duplicateIds = @($rawRows |
    Group-Object SourceMailboxId |
    Where-Object Count -gt 1)
if ($duplicateIds.Count -gt 0) {
    throw "Duplicate Exchange Online SourceMailboxId values were returned: $($duplicateIds.Name -join ', ')"
}

$historyTimestamp = [datetime]::UtcNow
$historyFolder = Join-Path $paths.DataAllRootPath (
    'ExchangeOnline\Mailboxes\{0}\{1}' -f
    $historyTimestamp.ToString('yyyy'),
    $historyTimestamp.ToString('MM')
)
$historyPath = Join-Path $historyFolder (
    'ExchangeOnline_Mailboxes_{0}.csv' -f
    $historyTimestamp.ToString('yyyyMMdd-HHmmssfff')
)
$exportParameters = @{
    InputObject     = $rawRows
    Columns         = @($rawTable.columns | ForEach-Object { [string]$_ })
    TenantKey       = $paths.TenantKey
    OrganizationKey = $paths.OrganizationKey
    EnvironmentKey = $paths.EnvironmentKey
    TenantId        = $paths.TenantId
}
Export-SmartWorkplaceCMDBCsv @exportParameters -Path $historyPath
Export-SmartWorkplaceCMDBCsv @exportParameters -Path $RawLatestOutputPath

$expectedLatestPath = Join-Path $paths.LatestOutputRootPath (
    Join-Path ([string]$rawTable.area) ([string]$rawTable.name)
)
if ([IO.Path]::GetFullPath($expectedLatestPath) -eq $RawLatestOutputPath) {
    $results = @(Test-SmartWorkplaceCMDBCsvContract `
        -LatestOutputRootPath $paths.LatestOutputRootPath `
        -ContractPath $rawContractPath)
    $result = @($results | Where-Object Name -eq $rawTable.name)
    if ($result.Count -ne 1 -or $result[0].Status -ne 'Valid') {
        throw 'The exported Exchange Online mailbox CSV does not satisfy its raw contract.'
    }
}

Write-Information (
    "SmartWorkplaceCMDB Exchange Online mailbox collection completed. Mailboxes={0}; RawLatest='{1}'." -f
    $rawRows.Count,
    $RawLatestOutputPath
) -InformationAction Continue

[pscustomobject]@{
    Status                 = 'Completed'
    ScriptVersion          = $ScriptVersion
    MailboxCount           = $rawRows.Count
    RawLatestOutputPath    = $RawLatestOutputPath
    HistoryOutputPath      = $historyPath
    RawContractVersion     = [string]$rawContract.contractVersion
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA0jHNH5j+5QYbK
# h0OUWlcl/GcLbd0+MY6lDYZoXfNZGqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJdfdc279vsGGitscsaitJvKXW5536u7gId69+QhdE90MA0GCSqG
# SIb3DQEBAQUABIIBgByFB9el9L6jSmpAUWPvAy1E+de+J1yYNzOAWYp4Oaer9dcX
# li5mqqWNsHctDurnth56PCrDGMxmUnheGDh2pvI9rNaqF8HjLrTMoJfpxokXliEu
# NjMpl0m/G8BvzDh9FzYt6akbHXUYn2x2R8cyk3QDH9XE9+cSRWlUuk08UxGckf0d
# 0Knz59/8G/88woaLIjg8+UYXXfneUflNKbXZv84yrvHBStHhfTVkqpYzw2WUDqWM
# oFWTOlwCNuFEYkPPDDgbDJOJfDPBN96rL9ZA5anv+71XFCLCykmAOSKqKb4Ki2kV
# bpG3MXu2hQmbcKdnqO/X/71HUQjbHnBeMZsVgeZCx36uiEBOrHrBh7KcghrTeG2x
# W/YNYr5OIpf3iel3bT4ALdWRbN4AWzd53SbYtfTQi4eW2dtGPXHhPVM27/kAQYOH
# 0IPGgE8LUTADqTNOXcfPP6hAM8k6o+K9NY/Aty/I7vVxmUFcQjm+t1wzgSEHx4W0
# dCgwxd8zRPTV25CskaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkxNzEz
# MjNaMC8GCSqGSIb3DQEJBDEiBCASAtAfRtIKb8DZO+h5JxsbxtjToTL1mT5SxpB3
# hlcddDANBgkqhkiG9w0BAQEFAASCAgAQ1EFsD02KQtpoa+gCWfY0zwXBKtcbsti3
# UHbLLCOmjuTY87TUXLTDaDmBvIwAtF4p/mS3y0CCPg3Vw7xPX+j+jgGtdLGGSDO7
# kmZGirpXqvjq7aNvDYYGMuYjSxLP3guXZXdix02/qobQ1GRxgqUwjnys+sJgkEVr
# BQ90x4bkF+jHfN1pOjk6NisO+ZK1JAHKpuvsGQwPicvFQ7krAM3kd3Q+aVzYXlrY
# MuBMvjUGELfsWk+ViiwtAkh5VCP1dDmTcajvtd82B5htuZI819iqJZXMJk/OjUbm
# 4RF2c7awXp+dr9z9GjnlrFB9+hAELp+m7TU88yuvDnS+pY9sjuYp0mqBPS1xfF/f
# 8KI/U0x9Q15MqDDWlDEtvWjC6mONwt8SGzPtQs3qWYHYlBsqPQeSQra3L5G0w0s2
# uwaVXh7DefaFUW/CxqnYyOd+OcITnracIpWdtLBLPvmOp1tJ5z1hgqOOoFZPlzmk
# TNs9A8UZ6LeWVG8tJlxN2sIhl6sQ2CzmIg2RkFEiOjzjiKCgI7BynhD3ij6LRpTb
# TUR/FmNe2pewPifSOgzkSfz/02RwU/4lBJ4GGk3VOyutMf/9QgwZog+4kk6PfqqB
# //btZLK+KOjPi+iPDodOjiYu8rgm0WXZDsL881lOyzLY7TXKAnVuoMPbvjabppa0
# RzlNKLdFOg==
# SIG # End signature block
