<#
.SYNOPSIS
Normalizes Active Directory raw snapshots for SmartWorkplaceCMDB.

.DESCRIPTION
Validates the six Active Directory raw CSV contracts and publishes stable,
tenant-scoped Active Directory source entities and direct group relationships
below DATA-LAST\CMDB\ActiveDirectory. It does not require domain connectivity
or the ActiveDirectory PowerShell module.

.VERSION
1.0.2
#>
[CmdletBinding()]
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
    [switch]$NoConfigWrite,
    [switch]$ValidateOnly
)

$ScriptVersion = '1.0.2'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Test-SmartWorkplaceCMDBExactCsvHeader {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$ExpectedColumns
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{
            Status = 'Missing'
            MissingColumns = ($ExpectedColumns -join ', ')
            UnexpectedColumns = ''
            OrderMatches = $false
        }
    }
    $headerLine = Get-Content -LiteralPath $Path -TotalCount 1 -ErrorAction Stop
    $actual = if ([string]::IsNullOrWhiteSpace($headerLine)) {
        @()
    }
    else {
        @($headerLine.Split(',') | ForEach-Object { $_.Trim().Trim('"') })
    }
    $missing = @($ExpectedColumns | Where-Object { $_ -notin $actual })
    $unexpected = @($actual | Where-Object { $_ -notin $ExpectedColumns })
    $orderMatches = (($actual -join [char]31) -ceq ($ExpectedColumns -join [char]31))
    return [pscustomobject]@{
        Status = if ($missing.Count -eq 0 -and $unexpected.Count -eq 0 -and $orderMatches) {
            'Valid'
        }
        else {
            'Incompatible'
        }
        MissingColumns = ($missing -join ', ')
        UnexpectedColumns = ($unexpected -join ', ')
        OrderMatches = $orderMatches
    }
}

function Assert-SmartWorkplaceCMDBIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Rows,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Identity,
        [Parameter(Mandatory)][string]$Name
    )

    foreach ($row in $Rows) {
        foreach ($identityName in $Identity.Keys) {
            if ([string]$row.$identityName -ne [string]$Identity[$identityName]) {
                throw (
                    "Raw Active Directory $Name identity mismatch for '{0}'. Expected='{1}' Actual='{2}'." -f
                    $identityName, $Identity[$identityName], $row.$identityName
                )
            }
        }
    }
}

function Get-SmartWorkplaceCMDBAdEntityId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)][string]$EntityType,
        [Parameter(Mandatory)][string]$ObjectGuid
    )

    if ([string]::IsNullOrWhiteSpace($ObjectGuid)) {
        return ''
    }
    return '{0}|ad-{1}|{2}' -f
        $TenantKey,
        $EntityType.ToLowerInvariant(),
        $ObjectGuid.ToLowerInvariant()
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $scriptRoot)
$modulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$rawContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
$curatedContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.activedirectory.tables.json'
Import-Module $modulePath -Force

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
$curatedContract = Get-SmartWorkplaceCMDBTableContract -Path $curatedContractPath

$mapping = @(
    @{ Raw = 'ActiveDirectory_Domains.csv'; Curated = 'CMDB_ActiveDirectoryDomains.csv' },
    @{ Raw = 'ActiveDirectory_Users.csv'; Curated = 'CMDB_ActiveDirectoryUsers.csv' },
    @{ Raw = 'ActiveDirectory_Groups.csv'; Curated = 'CMDB_ActiveDirectoryGroups.csv' },
    @{ Raw = 'ActiveDirectory_Computers.csv'; Curated = 'CMDB_ActiveDirectoryComputers.csv' },
    @{ Raw = 'ActiveDirectory_OrganizationalUnits.csv'; Curated = 'CMDB_ActiveDirectoryOrganizationalUnits.csv' },
    @{ Raw = 'ActiveDirectory_GroupMemberships.csv'; Curated = 'CMDB_ActiveDirectoryGroupMemberships.csv' }
)
$resolved = @()
foreach ($entry in $mapping) {
    $rawMatch = @($rawContract.tables | Where-Object name -eq $entry.Raw)
    $curatedMatch = @($curatedContract.tables | Where-Object name -eq $entry.Curated)
    if ($rawMatch.Count -ne 1 -or $curatedMatch.Count -ne 1) {
        throw "Active Directory contracts are missing '$($entry.Raw)' or '$($entry.Curated)'."
    }
    $rawPath = Join-Path $paths.LatestOutputRootPath (
        Join-Path ([string]$rawMatch[0].area) ([string]$rawMatch[0].name)
    )
    $curatedPath = Join-Path $paths.LatestOutputRootPath (
        Join-Path ([string]$curatedMatch[0].area) ([string]$curatedMatch[0].name)
    )
    $rawStatus = Test-SmartWorkplaceCMDBExactCsvHeader `
        -Path $rawPath `
        -ExpectedColumns @($rawMatch[0].columns | ForEach-Object { [string]$_ })
    $curatedStatus = Test-SmartWorkplaceCMDBExactCsvHeader `
        -Path $curatedPath `
        -ExpectedColumns @($curatedMatch[0].columns | ForEach-Object { [string]$_ })
    if ($rawStatus.Status -eq 'Incompatible') {
        throw "Raw Active Directory table '$($entry.Raw)' is incompatible."
    }
    if ($curatedStatus.Status -eq 'Incompatible') {
        throw "Curated Active Directory table '$($entry.Curated)' is incompatible."
    }
    if ($rawStatus.Status -eq 'Missing' -and -not $ValidateOnly) {
        throw "Raw Active Directory table was not found: '$rawPath'."
    }
    $resolved += [pscustomobject]@{
        RawTable = $rawMatch[0]
        CuratedTable = $curatedMatch[0]
        RawPath = $rawPath
        CuratedPath = $curatedPath
        RawStatus = $rawStatus.Status
        CuratedStatus = $curatedStatus.Status
    }
}

if ($ValidateOnly) {
    foreach ($entry in $resolved) {
        if (Test-Path -LiteralPath $entry.RawPath) { Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $entry.RawPath -Paths $paths | Out-Null }
    }
}

if ($ValidateOnly) {
    [pscustomobject]@{
        Status = 'Valid'
        ScriptVersion = $ScriptVersion
        RawContractVersion = [string]$rawContract.contractVersion
        CuratedContractVersion = [string]$curatedContract.contractVersion
        RawValidCount = @($resolved | Where-Object RawStatus -eq 'Valid').Count
        RawMissingCount = @($resolved | Where-Object RawStatus -eq 'Missing').Count
        CuratedValidCount = @($resolved | Where-Object CuratedStatus -eq 'Valid').Count
        CuratedMissingCount = @($resolved | Where-Object CuratedStatus -eq 'Missing').Count
        TenantKey = $paths.TenantKey
    } | Format-List
    return
}

$rawRows = @{}
foreach ($entry in $resolved) {
    $rawRows[[string]$entry.RawTable.name] = @(Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $entry.RawPath -Paths $paths)
}
$identity = [ordered]@{
    TenantKey = $paths.TenantKey
    OrganizationKey = $paths.OrganizationKey
    EnvironmentKey = $paths.EnvironmentKey
    TenantId = $paths.TenantId
}
foreach ($name in $rawRows.Keys) {
    Assert-SmartWorkplaceCMDBIdentity -Rows @($rawRows[$name]) -Identity $identity -Name $name
}

$domainRows = @($rawRows['ActiveDirectory_Domains.csv'] | ForEach-Object {
    [pscustomobject][ordered]@{
        CmdbDomainId = '{0}|ad-domain|{1}' -f
            $paths.TenantKey,
            ([string]$_.DomainDnsRoot).ToLowerInvariant()
        SourceSystem = [string]$_.SourceSystem
        DomainDnsRoot = [string]$_.DomainDnsRoot
        DomainNetBIOSName = [string]$_.DomainNetBIOSName
        ForestName = [string]$_.ForestName
        DistinguishedName = [string]$_.DistinguishedName
        DomainMode = [string]$_.DomainMode
        PDCEmulator = [string]$_.PDCEmulator
        RIDMaster = [string]$_.RIDMaster
        InfrastructureMaster = [string]$_.InfrastructureMaster
        SourceCollectedDateTime = [string]$_.SourceCollectedDateTime
    }
})

$userRows = @($rawRows['ActiveDirectory_Users.csv'] | ForEach-Object {
    [pscustomobject][ordered]@{
        CmdbAdUserId = Get-SmartWorkplaceCMDBAdEntityId $paths.TenantKey 'user' ([string]$_.SourceObjectGuid)
        SourceSystem = [string]$_.SourceSystem
        DomainDnsRoot = [string]$_.DomainDnsRoot
        SourceObjectGuid = [string]$_.SourceObjectGuid
        ObjectSid = [string]$_.ObjectSid
        SamAccountName = [string]$_.SamAccountName
        UserPrincipalName = [string]$_.UserPrincipalName
        DisplayName = [string]$_.DisplayName
        Enabled = [string]$_.Enabled
        Department = [string]$_.Department
        JobTitle = [string]$_.JobTitle
        Mail = [string]$_.Mail
        EmployeeId = [string]$_.EmployeeId
        Country = [string]$_.Country
        Company = [string]$_.Company
        Office = [string]$_.Office
        AccountExpirationDate = [string]$_.AccountExpirationDate
        UserAccountControl = [string]$_.UserAccountControl
        PrimaryGroupId = [string]$_.PrimaryGroupId
        DistinguishedName = [string]$_.DistinguishedName
        OrganizationalUnit = [string]$_.OrganizationalUnit
        ManagerDistinguishedName = [string]$_.ManagerDistinguishedName
        WhenCreated = [string]$_.WhenCreated
        WhenChanged = [string]$_.WhenChanged
        LastLogonDate = [string]$_.LastLogonDate
        PasswordLastSet = [string]$_.PasswordLastSet
        SourceCollectedDateTime = [string]$_.SourceCollectedDateTime
    }
})

$groupRows = @($rawRows['ActiveDirectory_Groups.csv'] | ForEach-Object {
    [pscustomobject][ordered]@{
        CmdbAdGroupId = Get-SmartWorkplaceCMDBAdEntityId $paths.TenantKey 'group' ([string]$_.SourceObjectGuid)
        SourceSystem = [string]$_.SourceSystem
        DomainDnsRoot = [string]$_.DomainDnsRoot
        SourceObjectGuid = [string]$_.SourceObjectGuid
        ObjectSid = [string]$_.ObjectSid
        SamAccountName = [string]$_.SamAccountName
        DisplayName = [string]$_.DisplayName
        GroupCategory = [string]$_.GroupCategory
        GroupScope = [string]$_.GroupScope
        Mail = [string]$_.Mail
        Description = [string]$_.Description
        ManagedByDistinguishedName = [string]$_.ManagedByDistinguishedName
        ProtectedFromAccidentalDeletion = [string]$_.ProtectedFromAccidentalDeletion
        DistinguishedName = [string]$_.DistinguishedName
        OrganizationalUnit = [string]$_.OrganizationalUnit
        WhenCreated = [string]$_.WhenCreated
        WhenChanged = [string]$_.WhenChanged
        SourceCollectedDateTime = [string]$_.SourceCollectedDateTime
    }
})

$computerRows = @($rawRows['ActiveDirectory_Computers.csv'] | ForEach-Object {
    [pscustomobject][ordered]@{
        CmdbAdComputerId = Get-SmartWorkplaceCMDBAdEntityId $paths.TenantKey 'computer' ([string]$_.SourceObjectGuid)
        SourceSystem = [string]$_.SourceSystem
        DomainDnsRoot = [string]$_.DomainDnsRoot
        SourceObjectGuid = [string]$_.SourceObjectGuid
        ObjectSid = [string]$_.ObjectSid
        SamAccountName = [string]$_.SamAccountName
        DeviceName = [string]$_.DeviceName
        DNSHostName = [string]$_.DNSHostName
        Enabled = [string]$_.Enabled
        OperatingSystem = [string]$_.OperatingSystem
        OperatingSystemVersion = [string]$_.OperatingSystemVersion
        IPv4Address = [string]$_.IPv4Address
        CanonicalName = [string]$_.CanonicalName
        DistinguishedName = [string]$_.DistinguishedName
        OrganizationalUnit = [string]$_.OrganizationalUnit
        ManagedByDistinguishedName = [string]$_.ManagedByDistinguishedName
        WhenCreated = [string]$_.WhenCreated
        WhenChanged = [string]$_.WhenChanged
        LastLogonDate = [string]$_.LastLogonDate
        LastLogonTimestamp = [string]$_.LastLogonTimestamp
        PasswordLastSet = [string]$_.PasswordLastSet
        PrimaryGroupId = [string]$_.PrimaryGroupId
        SourceCollectedDateTime = [string]$_.SourceCollectedDateTime
    }
})

$organizationalUnitRows = @($rawRows['ActiveDirectory_OrganizationalUnits.csv'] | ForEach-Object {
    [pscustomobject][ordered]@{
        CmdbAdOrganizationalUnitId = Get-SmartWorkplaceCMDBAdEntityId `
            $paths.TenantKey 'organizational-unit' ([string]$_.SourceObjectGuid)
        SourceSystem = [string]$_.SourceSystem
        DomainDnsRoot = [string]$_.DomainDnsRoot
        SourceObjectGuid = [string]$_.SourceObjectGuid
        Name = [string]$_.Name
        DistinguishedName = [string]$_.DistinguishedName
        ParentDistinguishedName = [string]$_.ParentDistinguishedName
        Description = [string]$_.Description
        ManagedByDistinguishedName = [string]$_.ManagedByDistinguishedName
        ProtectedFromAccidentalDeletion = [string]$_.ProtectedFromAccidentalDeletion
        WhenCreated = [string]$_.WhenCreated
        WhenChanged = [string]$_.WhenChanged
        SourceCollectedDateTime = [string]$_.SourceCollectedDateTime
    }
})

$membershipRows = @($rawRows['ActiveDirectory_GroupMemberships.csv'] | ForEach-Object {
    $groupId = Get-SmartWorkplaceCMDBAdEntityId $paths.TenantKey 'group' ([string]$_.GroupObjectGuid)
    $memberType = switch (([string]$_.MemberObjectClass).ToLowerInvariant()) {
        'user' { 'user' }
        'computer' { 'computer' }
        'group' { 'group' }
        default { 'object' }
    }
    $memberId = Get-SmartWorkplaceCMDBAdEntityId $paths.TenantKey $memberType ([string]$_.MemberObjectGuid)
    [pscustomobject][ordered]@{
        CmdbAdRelationshipId = '{0}|ad-memberof|{1}|{2}' -f
            $paths.TenantKey,
            ([string]$_.GroupObjectGuid).ToLowerInvariant(),
            ([string]$_.MemberObjectGuid).ToLowerInvariant()
        SourceSystem = [string]$_.SourceSystem
        DomainDnsRoot = [string]$_.DomainDnsRoot
        GroupObjectGuid = [string]$_.GroupObjectGuid
        CmdbAdGroupId = $groupId
        MemberObjectGuid = [string]$_.MemberObjectGuid
        MemberObjectSid = [string]$_.MemberObjectSid
        MemberObjectClass = [string]$_.MemberObjectClass
        CmdbAdMemberId = $memberId
        GroupDistinguishedName = [string]$_.GroupDistinguishedName
        MemberDistinguishedName = [string]$_.MemberDistinguishedName
        SourceCollectedDateTime = [string]$_.SourceCollectedDateTime
    }
})

$outputRows = @{
    'CMDB_ActiveDirectoryDomains.csv' = $domainRows
    'CMDB_ActiveDirectoryUsers.csv' = $userRows
    'CMDB_ActiveDirectoryGroups.csv' = $groupRows
    'CMDB_ActiveDirectoryComputers.csv' = $computerRows
    'CMDB_ActiveDirectoryOrganizationalUnits.csv' = $organizationalUnitRows
    'CMDB_ActiveDirectoryGroupMemberships.csv' = $membershipRows
}
$identityExport = @{
    TenantKey = $paths.TenantKey
    OrganizationKey = $paths.OrganizationKey
    EnvironmentKey = $paths.EnvironmentKey
    TenantId = $paths.TenantId
}
foreach ($entry in $resolved) {
    $name = [string]$entry.CuratedTable.name
    Export-SmartWorkplaceCMDBCsv `
        -InputObject @($outputRows[$name]) `
        -Path $entry.CuratedPath `
        -Columns @($entry.CuratedTable.columns | ForEach-Object { [string]$_ }) `
        @identityExport
}

$validation = @(Test-SmartWorkplaceCMDBCsvContract `
        -LatestOutputRootPath $paths.LatestOutputRootPath `
        -ContractPath $curatedContractPath)
if ($validation.Count -ne $mapping.Count -or
    @($validation | Where-Object Status -ne 'Valid').Count -gt 0) {
    throw 'An Active Directory normalized CSV does not satisfy its contract.'
}

Write-Information (
    "SmartWorkplaceCMDB Active Directory normalization completed. Domains={0}; Users={1}; Groups={2}; Computers={3}; OrganizationalUnits={4}; Memberships={5}." -f
    $domainRows.Count, $userRows.Count, $groupRows.Count, $computerRows.Count,
    $organizationalUnitRows.Count, $membershipRows.Count
) -InformationAction Continue

[pscustomobject]@{
    Status = 'Completed'
    ScriptVersion = $ScriptVersion
    DomainCount = $domainRows.Count
    UserCount = $userRows.Count
    GroupCount = $groupRows.Count
    ComputerCount = $computerRows.Count
    OrganizationalUnitCount = $organizationalUnitRows.Count
    GroupMembershipCount = $membershipRows.Count
    CuratedOutputRootPath = Join-Path $paths.LatestOutputRootPath 'CMDB\ActiveDirectory'
    RawContractVersion = [string]$rawContract.contractVersion
    CuratedContractVersion = [string]$curatedContract.contractVersion
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAnY8EZ7gU39tCW
# f9MPf1vOPFPlIUTsG0KgBQdTy6e1TaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIN7SDfzcStPDuN9doyEr0jRPnUJXXvsfjvwcEgl6QxEvMA0GCSqG
# SIb3DQEBAQUABIIBgAWVvfFUCbKkgrMOy7OQrF2FHpy1oQBRp1mNyioHGh2IBg+2
# 08zhoAnvpG1rDNStJYDy1mx+X+gRfER4fIRUIcu6yds6QG2t6KeD+ou9vivyuMsK
# ZO9aMss3DXN3UCrrjGtaN/shXo4+jl+GcywaabLiBwdLAr2sV50C2oaQWDeePJRQ
# Ok3O2eCVmKoiWJjCSSRenAonC8DZTMFxUxzDyVgZXW/SxbdmU8EyatLwKlBvWHg8
# btCt7myZ27W9lWGM6u2DLu0G1gzbQq4cNviYeTi1PicC5RzCKPDfVx0Q0FlE2HKy
# t2fX2O+8uZhZRUaY5bhIKjopCc+px23HTH+4tNdJ6RV/rbHhknYbQr7cZ3J2Py99
# 6yqjq2in3eZMBCKhIeIGeV29hkiu4tPd8Ben66ZYLmLu3ciOiJw+FO0zYNolTzY8
# 7ojkDi+YHabUa2L8DPtuvLpsygSx5pf4bUJ2xLHe5XwYdwNGmVEHGus0XvBKnwss
# J/gukQv3Wo4Vv1F426GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxMzE3
# MTRaMC8GCSqGSIb3DQEJBDEiBCABNH9upVkkKBpr2mwjY6r6SEFPNjp6gok8TUsx
# J0zaKTANBgkqhkiG9w0BAQEFAASCAgBnwI/IKRcIOFW7wvBnJW2Rf6wQZfLbyCLe
# /93Bg6QCw8nVBFi5VYvF++KqyJWdbjw3h32ReYEKu78U4dychY6y5hHO807YGZ6f
# cQaHjoVkLn28iFdFzuK0IR33JBl/iOJaIB0rz3lzD25il0c2wkemelDzF0WdANEl
# b69YazB/yRtHblYO6Y+QiOz2KmAZPkLIbgvo/70MKSkWYFRA+cWWawMWMa4QzLg1
# JOJhQIiepkwjOlhkGTKhTkDsUVE+n0hwQjhh4Jl8RB28gb6/SKTMoT6cx8QjdxwD
# oieuR1RArl0DZPkhp0CDNjxE9jlZ4uN/Xot+AvbvzJVVEL5iWeV/+zPM+COYNdCX
# yIXhXjGM1rW1IBa5vmUcEo5O+UzIjThg9Pe4OwC+YVNll6WgYoQNdAxXoXoo3Whs
# i7V+yF0rPVsMWsFc1mwUBlQoFegpuQK4pxlLRSoRw+vXLEekgMZCu2XuQLhHk9gE
# XQ1AtMN1mJWMisGOrYMUb3r9z8PGO9UvZNEETpbrGgVEH+PMVZcHf3OMeDyYvfJu
# qIKUfiKCaTL5/cC8hmEeSsBR18Mn1KkuvT/6VbfBkYhqKmhc8edAQRX58Stco4b8
# e2wtJKP0AfrOm/yzv4v1nyAlWag3M60vn/9z8PYmlwon5qfBJudp1wb0S4Pkj+x8
# BlxyBchWmQ==
# SIG # End signature block
