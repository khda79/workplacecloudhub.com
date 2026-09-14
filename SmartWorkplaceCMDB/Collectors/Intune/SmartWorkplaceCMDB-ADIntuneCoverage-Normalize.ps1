<#
.SYNOPSIS
Builds exact Active Directory workstation to Intune coverage facts.

.DESCRIPTION
Correlates AD computers to Microsoft Entra devices only through the immutable
on-premises security identifier exposed by Microsoft Graph, then correlates
the Entra deviceId to the Intune azureADDeviceId. Device names are never used
as a matching key. The output keeps one row per Windows 7 through Windows 11
AD workstation and distinguishes an Entra-only record from a workstation that
has no exact Entra match.

.VERSION
1.0.0
#>
[CmdletBinding()]
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
    [string]$ADComputerInputPath,
    [string]$EntraDeviceInputPath,
    [string]$IntuneDeviceInputPath,
    [string]$OutputPath,
    [switch]$NoConfigWrite,
    [switch]$ValidateOnly
)

$ScriptVersion = '1.0.0'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Test-SmartWorkplaceCMDBExactHeader {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$Columns
    )
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return 'Missing' }
    $header = Get-Content -LiteralPath $Path -TotalCount 1 -ErrorAction Stop
    $actual = if ([string]::IsNullOrWhiteSpace($header)) { @() } else {
        @($header.Split(',') | ForEach-Object { $_.Trim().Trim('"') })
    }
    if (($actual -join [char]31) -cne ($Columns -join [char]31)) { return 'Incompatible' }
    return 'Valid'
}

function Get-SmartWorkplaceCMDBLatestDateText {
    [CmdletBinding()]
    param([AllowEmptyCollection()][string[]]$Values)
    $dates = New-Object System.Collections.Generic.List[datetimeoffset]
    foreach ($value in @($Values)) {
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        $parsed = [datetimeoffset]::MinValue
        if (-not [datetimeoffset]::TryParse(
                $value,
                [Globalization.CultureInfo]::InvariantCulture,
                [Globalization.DateTimeStyles]::AssumeUniversal,
                [ref]$parsed
            )) { throw "Invalid source collection date '$value'." }
        $dates.Add($parsed.ToUniversalTime())
    }
    if ($dates.Count -eq 0) { throw 'No source collection date is available for an AD workstation.' }
    return (@($dates.ToArray() | Sort-Object -Descending)[0]).ToString(
        'yyyy-MM-ddTHH:mm:ss.fffffffZ',
        [Globalization.CultureInfo]::InvariantCulture
    )
}

function Test-SmartWorkplaceCMDBWindowsWorkstation {
    [CmdletBinding()]
    param([AllowEmptyString()][string]$OperatingSystem)
    $value = $OperatingSystem.Trim()
    return ($value -match '(?i)^Microsoft\s+Windows\s+(7|8(?:\.1)?|10|11)\b|^Windows\s+(7|8(?:\.1)?|10|11)\b') -and
        ($value -notmatch '(?i)server')
}

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$coreModulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$curatedContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json'
$rawContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
$adContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.activedirectory.tables.json'
Import-Module $coreModulePath -Force

$boundParameterCopy = @{}
foreach ($key in $PSBoundParameters.Keys) { $boundParameterCopy[$key] = $PSBoundParameters[$key] }
$context = Resolve-SmartWorkplaceCMDBContext `
    -BoundParameters $boundParameterCopy `
    -GlobalConfigPath $GlobalConfigPath `
    -TenantConfigPath $TenantConfigPath `
    -NoConfigWrite:($ValidateOnly -or $NoConfigWrite)
$paths = $context.Paths
$curatedContract = Get-SmartWorkplaceCMDBTableContract -Path $curatedContractPath
$rawContract = Get-SmartWorkplaceCMDBTableContract -Path $rawContractPath
$adContract = Get-SmartWorkplaceCMDBTableContract -Path $adContractPath

$targetTable = @($curatedContract.tables | Where-Object name -eq 'FactADIntuneCoverage.csv')
$entraTable = @($rawContract.tables | Where-Object name -eq 'Entra_Devices.csv')
$intuneTable = @($rawContract.tables | Where-Object name -eq 'Intune_ManagedDevices.csv')
$adTable = @($adContract.tables | Where-Object name -eq 'CMDB_ActiveDirectoryComputers.csv')
if ($targetTable.Count -ne 1 -or $entraTable.Count -ne 1 -or $intuneTable.Count -ne 1 -or $adTable.Count -ne 1) {
    throw 'The AD to Intune coverage contracts are incomplete or ambiguous.'
}
$targetTable = $targetTable[0]
$entraTable = $entraTable[0]
$intuneTable = $intuneTable[0]
$adTable = $adTable[0]

if ([string]::IsNullOrWhiteSpace($ADComputerInputPath)) {
    $ADComputerInputPath = Join-Path $paths.LatestOutputRootPath (
        Join-Path ([string]$adTable.area) ([string]$adTable.name)
    )
}
if ([string]::IsNullOrWhiteSpace($EntraDeviceInputPath)) {
    $EntraDeviceInputPath = Join-Path $paths.LatestOutputRootPath (
        Join-Path ([string]$entraTable.area) ([string]$entraTable.name)
    )
}
if ([string]::IsNullOrWhiteSpace($IntuneDeviceInputPath)) {
    $IntuneDeviceInputPath = Join-Path $paths.LatestOutputRootPath (
        Join-Path ([string]$intuneTable.area) ([string]$intuneTable.name)
    )
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $paths.LatestOutputRootPath (
        Join-Path ([string]$targetTable.area) ([string]$targetTable.name)
    )
}

$inputs = @(
    [pscustomobject]@{ Name='AD computers'; Path=[IO.Path]::GetFullPath($ADComputerInputPath); Columns=@($adTable.columns | ForEach-Object { [string]$_ }) },
    [pscustomobject]@{ Name='Entra devices'; Path=[IO.Path]::GetFullPath($EntraDeviceInputPath); Columns=@($entraTable.columns | ForEach-Object { [string]$_ }) },
    [pscustomobject]@{ Name='Intune devices'; Path=[IO.Path]::GetFullPath($IntuneDeviceInputPath); Columns=@($intuneTable.columns | ForEach-Object { [string]$_ }) }
)
foreach ($input in $inputs) {
    $input | Add-Member NoteProperty Status (Test-SmartWorkplaceCMDBExactHeader -Path $input.Path -Columns $input.Columns)
    if ($input.Status -eq 'Incompatible') { throw "$($input.Name) input is incompatible: '$($input.Path)'." }
    if ($input.Status -eq 'Missing' -and -not $ValidateOnly) { throw "$($input.Name) input is missing: '$($input.Path)'." }
}
$targetStatus = Test-SmartWorkplaceCMDBExactHeader `
    -Path ([IO.Path]::GetFullPath($OutputPath)) `
    -Columns @($targetTable.columns | ForEach-Object { [string]$_ })
if ($targetStatus -eq 'Incompatible') { throw "Existing FactADIntuneCoverage.csv is incompatible." }

if ($ValidateOnly) {
    [pscustomobject]@{
        Status = 'Valid'; ScriptVersion = $ScriptVersion
        ValidInputCount = @($inputs | Where-Object Status -eq 'Valid').Count
        MissingInputCount = @($inputs | Where-Object Status -eq 'Missing').Count
        TargetStatus = $targetStatus; OutputPath = $OutputPath
        RawContractVersion = [string]$rawContract.contractVersion
        CuratedContractVersion = [string]$curatedContract.contractVersion
    } | Format-List
    return
}

$adRows = @(Import-Csv -LiteralPath $ADComputerInputPath -ErrorAction Stop)
$entraRows = @(Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $EntraDeviceInputPath -Paths $paths -ErrorAction Stop)
$intuneRows = @(Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $IntuneDeviceInputPath -Paths $paths -ErrorAction Stop)
$identityFields = [ordered]@{
    TenantKey=$paths.TenantKey; OrganizationKey=$paths.OrganizationKey
    EnvironmentKey=$paths.EnvironmentKey; TenantId=$paths.TenantId
}
foreach ($set in @($adRows, $entraRows, $intuneRows)) {
    foreach ($row in @($set)) {
        foreach ($name in $identityFields.Keys) {
            if ([string]$row.$name -ne [string]$identityFields[$name]) {
                throw "AD to Intune coverage input identity mismatch for '$name'."
            }
        }
    }
}

$entraBySid = @{}
foreach ($row in $entraRows) {
    $sid = ([string]$row.OnPremisesSecurityIdentifier).Trim()
    if ([string]::IsNullOrWhiteSpace($sid)) { continue }
    $key = $sid.ToLowerInvariant()
    if ($entraBySid.ContainsKey($key)) { throw "Duplicate Entra on-premises security identifier '$sid'." }
    $entraBySid[$key] = $row
}
$intuneByEntraDeviceId = @{}
foreach ($row in $intuneRows) {
    $id = ([string]$row.AzureAdDeviceId).Trim()
    if ([string]::IsNullOrWhiteSpace($id) -or $id -eq '00000000-0000-0000-0000-000000000000') { continue }
    $key = $id.ToLowerInvariant()
    if (-not $intuneByEntraDeviceId.ContainsKey($key)) {
        $intuneByEntraDeviceId[$key] = New-Object System.Collections.Generic.List[object]
    }
    $intuneByEntraDeviceId[$key].Add($row)
}

$coverageRows = New-Object System.Collections.Generic.List[object]
foreach ($ad in @($adRows | Where-Object {
            Test-SmartWorkplaceCMDBWindowsWorkstation ([string]$_.OperatingSystem)
        } | Sort-Object DeviceName, ObjectSid)) {
    $sid = ([string]$ad.ObjectSid).Trim()
    if ([string]::IsNullOrWhiteSpace($sid)) { throw "AD workstation '$($ad.DeviceName)' has no ObjectSid." }
    $entra = if ($entraBySid.ContainsKey($sid.ToLowerInvariant())) { $entraBySid[$sid.ToLowerInvariant()] } else { $null }
    $intune = $null
    if ($null -ne $entra) {
        $deviceIdKey = ([string]$entra.SourceDeviceId).Trim().ToLowerInvariant()
        if ($intuneByEntraDeviceId.ContainsKey($deviceIdKey)) {
            $intune = @($intuneByEntraDeviceId[$deviceIdKey].ToArray() | Sort-Object `
                @{ Expression = { [string]$_.LastSyncDateTime }; Descending = $true }, `
                @{ Expression = { [string]$_.ManagedDeviceId }; Descending = $false })[0]
        }
    }
    $coverageState = if ($null -ne $intune) { 'Managed in Intune' } elseif ($null -ne $entra) { 'Entra only' } else { 'Not found in Entra' }
    $sourceDeviceId = if ($null -ne $entra) { ([string]$entra.SourceDeviceId).Trim() } else { '' }
    $coverageRows.Add([pscustomobject][ordered]@{
        TenantADComputerKey = [string]$ad.CmdbAdComputerId
        CmdbAdComputerId = [string]$ad.CmdbAdComputerId
        DeviceName = [string]$ad.DeviceName
        Enabled = [string]$ad.Enabled
        OperatingSystem = [string]$ad.OperatingSystem
        OperatingSystemVersion = [string]$ad.OperatingSystemVersion
        TenantDeviceKey = if ($sourceDeviceId) { ('{0}|device|{1}' -f $paths.TenantKey, $sourceDeviceId.ToLowerInvariant()) } else { '' }
        EntraDeviceId = $sourceDeviceId
        IntuneManagedDeviceId = if ($null -ne $intune) { [string]$intune.ManagedDeviceId } else { '' }
        CoverageState = $coverageState
        MatchMethod = if ($null -ne $entra) { 'Exact AD ObjectSid -> Entra onPremisesSecurityIdentifier -> deviceId' + $(if ($null -ne $intune) { ' -> Intune azureADDeviceId' } else { '' }) } else { 'No exact Entra SID match' }
        SourceCollectedDateTime = Get-SmartWorkplaceCMDBLatestDateText @(
            [string]$ad.SourceCollectedDateTime,
            $(if ($null -ne $entra) { [string]$entra.SourceCollectedDateTime } else { '' }),
            $(if ($null -ne $intune) { [string]$intune.SourceCollectedDateTime } else { '' })
        )
    })
}

$rows = @($coverageRows.ToArray())
$duplicateKeys = @($rows | Group-Object TenantADComputerKey | Where-Object Count -gt 1)
if ($duplicateKeys.Count -gt 0) { throw 'AD to Intune coverage produced duplicate AD computer keys.' }
$identityExport = @{
    TenantKey=$paths.TenantKey; OrganizationKey=$paths.OrganizationKey
    EnvironmentKey=$paths.EnvironmentKey; TenantId=$paths.TenantId
}
Export-SmartWorkplaceCMDBCsv `
    -InputObject $rows `
    -Path $OutputPath `
    -Columns @($targetTable.columns | ForEach-Object { [string]$_ }) `
    @identityExport
if ((Test-SmartWorkplaceCMDBExactHeader -Path $OutputPath -Columns @($targetTable.columns | ForEach-Object { [string]$_ })) -ne 'Valid') {
    throw 'FactADIntuneCoverage.csv does not satisfy its contract.'
}

$managedCount = @($rows | Where-Object CoverageState -eq 'Managed in Intune').Count
$enabledCount = @($rows | Where-Object Enabled -eq 'True').Count
Write-Information (
    'SmartWorkplaceCMDB exact AD to Intune coverage completed. Workstations={0}; Enabled={1}; Managed={2}; EntraOnly={3}; NoEntraMatch={4}.' -f
    $rows.Count,
    $enabledCount,
    $managedCount,
    @($rows | Where-Object CoverageState -eq 'Entra only').Count,
    @($rows | Where-Object CoverageState -eq 'Not found in Entra').Count
) -InformationAction Continue
[pscustomobject]@{
    Status='Completed'; ScriptVersion=$ScriptVersion; WorkstationCount=$rows.Count
    EnabledWorkstationCount=$enabledCount; ManagedWorkstationCount=$managedCount
    EntraOnlyCount=@($rows | Where-Object CoverageState -eq 'Entra only').Count
    NotFoundInEntraCount=@($rows | Where-Object CoverageState -eq 'Not found in Entra').Count
    OutputPath=$OutputPath; CuratedContractVersion=[string]$curatedContract.contractVersion
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBerc/8FLArDrYn
# ZHhMfCVlyWwRXg4VMTeFuBBdQl0zw6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIXFD/ESut0KZSXEPx1Wc14SEDG7aRuTpOYjQcffOIkfMA0GCSqG
# SIb3DQEBAQUABIIBgKrpWKgd2t2YhiDQ8sZUlTtzNF17wp+7v7MfVbCiAxyvHn4l
# IXAVHY1xL79H3lgx879LJbagxuH6aAsscBDYYG7jykucfX0iTwh+r9ge7lQNQaCy
# H6YVskVxZfrF/Puf3V6nSY0mmF2AANYQF01nlo19VB6fJl+vQoPXrPUgIBW47mFo
# zkHkNh3+78U+KD62ND7BqatBYdACdmuWG7VnWUEF2I43INNppylgQauDHTB4mKGO
# xRIAF7zBIdgt7Ix0aBPzOm1OtvSAhc1J6Btx4ZGLd7OS+t0igNasbNKqvpNzRb8F
# xjMtQitKr3ipdPzKZQqs0ktBuSLu2w5Y6ttB96e3ZG43ddR1VFsK+bvmwE2g2Ib1
# GkXNcOrf+gBuRrWMRqrk9Cq4FgLWQE41+tSpoxXU9mq2NxFruecDo6uVv4A3d7Uc
# Nmd7ZhiQjjcowQX8ijo55urEjYayZ2ixor/NapmoEitAPL1yMyH4VMODcAV3yLHg
# jj7xO8//51Juhasws6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTQwNDI0
# NTRaMC8GCSqGSIb3DQEJBDEiBCA2i8WoXhVa5pMtb9p+rtCqAMI4gzVBrSXt8abb
# K6bhATANBgkqhkiG9w0BAQEFAASCAgCUQkA8xbNLw3WnEdA/+Jfa+kjJbvNUDKpw
# i++hPGXdwRpX0OEfREJM5jsmcg+CopE7I4An9ESIcKMiEkigC+zXQNZytRl+lWeb
# e8XuvZd1EX9qmkjXxVbCXc/vjkB94p205YCGSi0dS0v30EiqYBaC3VkR9dRA8jGW
# HaCXiK4+s+GjKr6tN16XD8Ct2hknplEg+5HTy3OJYHqjVGIAkHHS2a3ILaHWz2pR
# W4/MVislnmPqCS06vl4uMV4piqcmm5FaTQ0enPcZYfTUK+jqeAaO/beE1Q4OTIFO
# oUh6XFfTChLq7Mow8Y/3rMfGDos8IYKXSetWXdlPMo2f3doyVIbsVgasT5rEjqWZ
# r1DOUZpd76gpDY025IxwHa6uu3GgVujTxxUtVv3H/SamWhCUtDL6R7XAPsXAySBk
# 7iP130oJnCQbapmU6VEsx/yo9V/Llq7qIqEvk1gSNa3UYa3jOolV+mItod+MM5/7
# qcH1oDELklExeqGQKXN/hE9rIKs3qTCvfk3/RDctDRSZamZGgn9eGNSlCkrzgMAO
# Pg3uv9gxNcLwAZxTxBfn16Qw5L1c9fwymOUIO0uJ5DE6DsG1w3kegQG9/r5X7FR0
# L70lpb4Zopy0asnEIFsZxfirrcbbQ8PANE8v5a0IJMh3kK624SVZrJqCtMq7J9w/
# 9epOKhYTCw==
# SIG # End signature block
