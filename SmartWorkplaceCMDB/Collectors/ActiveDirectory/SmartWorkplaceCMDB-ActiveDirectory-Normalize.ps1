<#
.SYNOPSIS
Normalizes Active Directory raw snapshots for SmartWorkplaceCMDB.

.DESCRIPTION
Validates the five Active Directory raw CSV contracts and publishes stable,
tenant-scoped Active Directory source entities and direct group relationships
below DATA-LAST\CMDB\ActiveDirectory. It does not require domain connectivity
or the ActiveDirectory PowerShell module.

.VERSION
0.1.0
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

$ScriptVersion = '0.1.0'
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
    $rawRows[[string]$entry.RawTable.name] = @(Import-Csv -LiteralPath $entry.RawPath)
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
        DistinguishedName = [string]$_.DistinguishedName
        OrganizationalUnit = [string]$_.OrganizationalUnit
        ManagedByDistinguishedName = [string]$_.ManagedByDistinguishedName
        WhenCreated = [string]$_.WhenCreated
        WhenChanged = [string]$_.WhenChanged
        LastLogonDate = [string]$_.LastLogonDate
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
    "SmartWorkplaceCMDB Active Directory normalization completed. Domains={0}; Users={1}; Groups={2}; Computers={3}; Memberships={4}." -f
    $domainRows.Count, $userRows.Count, $groupRows.Count, $computerRows.Count, $membershipRows.Count
) -InformationAction Continue

[pscustomobject]@{
    Status = 'Completed'
    ScriptVersion = $ScriptVersion
    DomainCount = $domainRows.Count
    UserCount = $userRows.Count
    GroupCount = $groupRows.Count
    ComputerCount = $computerRows.Count
    GroupMembershipCount = $membershipRows.Count
    CuratedOutputRootPath = Join-Path $paths.LatestOutputRootPath 'CMDB\ActiveDirectory'
    RawContractVersion = [string]$rawContract.contractVersion
    CuratedContractVersion = [string]$curatedContract.contractVersion
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBzfLWgqvUWp+x/
# b0ZtUDVPc1jh4mlys4lQgYPvVArlKqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFJfKhnU8A/5LYnjp0FeZUAliDeypdMxPVf7H0h9lPmJMA0GCSqG
# SIb3DQEBAQUABIIBgE6Nv/Ntep0dfDySfY63/sQiFvsauGKq8zAkXF1xguxaQldv
# FmA594LXza912p/IfYbxJqPc6CyiWtxgxrSy90Q7bADKf0vq5HvjKLBWaNq73RVv
# PSEZKOsUuHyCsXUuyIEIS7Apk9Ix/SBZVC4rbbsrYiyQlgdUr6igVwoUWHkfFSDE
# DrBybKpDZ0ptb8z6I3JMr82kF1K+PGYGuhcL3v6rw1mg+NzQFtVa6Bs2prIryNjX
# IlDSfdE1zlue67KA0Id6UBOXOQURSdHW0Pon7uh0+cKRUDJ57AejBqSz1Jm2VeyZ
# ZFDRJXngHLJFSUv+cFfBP+OkKSUYnH6YAhhhCIw7pi2i7bQN4i+XFT1Lh7fsJdUz
# s7yZn/I6DW4AKq0TTZGBwqQiYvHqWVeQDGApns8OUzRoIMrSXm7Mj2WotFzGcBRM
# gU35xXNkb4O+WCRhZcdxb5rM685votsZA859+rzaHPzrR/3TQg8XoNQYQGTSiHD6
# 2vfvCia284VF0N56LKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkyMDU3
# MDlaMC8GCSqGSIb3DQEJBDEiBCDXrQ/l24RR8lvbrtXYb6oiXoMZRxUd3cc0Og+m
# 1ES3ozANBgkqhkiG9w0BAQEFAASCAgC0AznQ3fE80B0fFnN7FeNNMq03MR9UHCeM
# vsLB9NUpmjev88DnXfbV07ReGGF7upf+yuD42GTh4jv/iVHdw5rDUlm/9gfsRrWG
# NQdlzxdfzN5PSJho5Hj/VNbvtsa25NGr0Rw84lAna7DpoTTjGRo38q8iSxTbZVfU
# GlQlg4GX5WdF1nz3SSFqxVvdcrnQEg7J6x02LEmItfxGXoAP/l/grz7hrV4RBVbs
# 4Zp7TDUAJc2ZyedFnCERf4gVVsX3iayy1Hdsaeq18fSaONB5qKOTJ27SC2fdKd0b
# 3+SynQK6hefhnfpR0+GOu577u6RxzYTc5wRQo/ww0PkxSmRF0B0gCx97kXGmkv5l
# XkDzmPrWv9RmAap4EtcnoWbq8Eir20v8zxpPVYepiOvp/NMJYlbDs7dBjHyusS5v
# DZRzIexI7ScTTOZs6KPiY+22ZLF3JJwV2D4MzbeHDdH0RQ57IytI09f/0QFuy0Ns
# u7E9WdP+mfveFz4r4lxOTMCFAj7hkbBW1nxqZ1O4ehFClVQ1jp03/+sCQggVVext
# AWHedd1RzKZd5olGJc9F3L+Cpee13YOck+eEzFzHNQeutZVESjeq1UzXX6lwCbkL
# jowKG7WhwaZhvui7xh6bO/sKrFG45oXKBuI+kIAgj1NZlCdxHd0LcrAPTupEMYLb
# ebTH53iiAQ==
# SIG # End signature block
