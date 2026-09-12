<#
.SYNOPSIS
Publishes verified-domain and hybrid identity coverage tables.

.DESCRIPTION
Creates DimVerifiedDomain from the Entra domain snapshot and an aggregate-only
FactHybridIdentityCoverage table from exact, case-insensitive UPN and device
name matches between Active Directory and Entra. No user or device identifier
is exposed in the coverage fact.

.VERSION
1.0.0
#>
[CmdletBinding()]
param(
    [Alias('ProfileKey')][string]$Tenant = 'default',
    [string]$OrganizationKey,[string]$EnvironmentKey,[string]$TenantKey,[string]$TenantId,
    [string]$DataRootPath,[string]$DataAllRootPath,[string]$LatestOutputRootPath,[string]$LogRootPath,
    [string]$GlobalConfigPath,[string]$TenantConfigPath,
    [string]$RawDomainInputPath,[string]$RawActiveDirectoryUsersInputPath,[string]$RawEntraUsersInputPath,
    [string]$RawActiveDirectoryComputersInputPath,[string]$RawEntraDevicesInputPath,
    [switch]$NoConfigWrite,[switch]$ValidateOnly
)

$ScriptVersion = '1.0.0'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Get-SmartWorkplaceCMDBHeaderStatus {
    param([Parameter(Mandatory)][string]$Path,[Parameter(Mandatory)][string[]]$Columns)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return 'Missing' }
    $line = Get-Content -LiteralPath $Path -TotalCount 1
    $actual = if ([string]::IsNullOrWhiteSpace($line)) { @() } else { @($line.Split(',') | ForEach-Object { $_.Trim().Trim('"') }) }
    if (($actual -join [char]31) -ceq ($Columns -join [char]31)) { return 'Valid' }
    return 'Incompatible'
}

function Get-SmartWorkplaceCMDBNormalizedKey {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim().ToLowerInvariant()
}

function Get-SmartWorkplaceCMDBDeviceKey {
    param([Parameter(Mandatory)]$Row,[switch]$ActiveDirectory)
    if ($ActiveDirectory) {
        $dns = Get-SmartWorkplaceCMDBNormalizedKey $Row.DNSHostName
        if (-not [string]::IsNullOrWhiteSpace($dns)) { return $dns }
    }
    return Get-SmartWorkplaceCMDBNormalizedKey $Row.DeviceName
}

function New-SmartWorkplaceCMDBCoverageRow {
    param(
        [Parameter(Mandatory)][string]$EntityType,
        [Parameter(Mandatory)][object[]]$LocalRows,
        [Parameter(Mandatory)][object[]]$CloudRows,
        [Parameter(Mandatory)][scriptblock]$LocalKey,
        [Parameter(Mandatory)][scriptblock]$CloudKey,
        [Parameter(Mandatory)][bool]$LocalAvailable,
        [Parameter(Mandatory)][bool]$CloudAvailable,
        [Parameter(Mandatory)][string]$MatchRule
    )
    $localKeys = @($LocalRows | ForEach-Object { & $LocalKey $_ })
    $cloudKeys = @($CloudRows | ForEach-Object { & $CloudKey $_ })
    $validLocal = @($localKeys | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $validCloud = @($cloudKeys | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $localDistinct = @($validLocal | Sort-Object -Unique)
    $cloudDistinct = @($validCloud | Sort-Object -Unique)
    $cloudSet = @{}; foreach ($key in $cloudDistinct) { $cloudSet[$key] = $true }
    $localSet = @{}; foreach ($key in $localDistinct) { $localSet[$key] = $true }
    $matched = @($localDistinct | Where-Object { $cloudSet.ContainsKey($_) }).Count
    $localOnly = @($localKeys | Where-Object { [string]::IsNullOrWhiteSpace($_) -or -not $cloudSet.ContainsKey($_) }).Count
    $cloudOnly = @($cloudKeys | Where-Object { [string]::IsNullOrWhiteSpace($_) -or -not $localSet.ContainsKey($_) }).Count
    $duplicateLocal = [Math]::Max(0,$validLocal.Count - $localDistinct.Count)
    $duplicateCloud = [Math]::Max(0,$validCloud.Count - $cloudDistinct.Count)
    $rate = if ($localDistinct.Count -gt 0) { [Math]::Round(($matched * 100.0) / $localDistinct.Count,2) } else { 0 }
    [pscustomobject][ordered]@{
        EntityType = $EntityType
        OnPremisesCount = $LocalRows.Count
        CloudCount = $CloudRows.Count
        MatchedCount = $matched
        OnPremisesOnlyCount = $localOnly
        CloudOnlyCount = $cloudOnly
        DuplicateOnPremisesKeyCount = $duplicateLocal
        DuplicateCloudKeyCount = $duplicateCloud
        MatchRatio = $rate.ToString('0.00',[Globalization.CultureInfo]::InvariantCulture)
        MatchMethod = $MatchRule
        SourceCollectedDateTime = [datetime]::UtcNow.ToString('o')
    }
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $scriptRoot)
$modulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$rawContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
$curatedContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json'
Import-Module $modulePath -Force
$boundParameters = @{}; foreach ($key in $PSBoundParameters.Keys) { $boundParameters[$key] = $PSBoundParameters[$key] }
$context = Resolve-SmartWorkplaceCMDBContext -BoundParameters $boundParameters -GlobalConfigPath $GlobalConfigPath -TenantConfigPath $TenantConfigPath -NoConfigWrite:($ValidateOnly -or $NoConfigWrite)
$paths = $context.Paths
$rawContract = Get-SmartWorkplaceCMDBTableContract -Path $rawContractPath
$curatedContract = Get-SmartWorkplaceCMDBTableContract -Path $curatedContractPath
$domainRawTable = @($rawContract.tables | Where-Object name -eq 'Entra_VerifiedDomains.csv')
$adUserTable = @($rawContract.tables | Where-Object name -eq 'ActiveDirectory_Users.csv')
$entraUserTable = @($rawContract.tables | Where-Object name -eq 'Entra_Users.csv')
$adComputerTable = @($rawContract.tables | Where-Object name -eq 'ActiveDirectory_Computers.csv')
$entraDeviceTable = @($rawContract.tables | Where-Object name -eq 'Entra_Devices.csv')
$domainDimTable = @($curatedContract.tables | Where-Object name -eq 'DimVerifiedDomain.csv')
$coverageTable = @($curatedContract.tables | Where-Object name -eq 'FactHybridIdentityCoverage.csv')
foreach ($definition in @($domainRawTable,$adUserTable,$entraUserTable,$adComputerTable,$entraDeviceTable,$domainDimTable,$coverageTable)) {
    if ($definition.Count -ne 1) { throw 'Required verified-domain or hybrid coverage contract definition is missing or duplicated.' }
}
$domainRawTable=$domainRawTable[0];$adUserTable=$adUserTable[0];$entraUserTable=$entraUserTable[0];$adComputerTable=$adComputerTable[0];$entraDeviceTable=$entraDeviceTable[0];$domainDimTable=$domainDimTable[0];$coverageTable=$coverageTable[0]
$pathDefinitions = @(
    @{ Variable='RawDomainInputPath'; Table=$domainRawTable },@{ Variable='RawActiveDirectoryUsersInputPath'; Table=$adUserTable },
    @{ Variable='RawEntraUsersInputPath'; Table=$entraUserTable },@{ Variable='RawActiveDirectoryComputersInputPath'; Table=$adComputerTable },
    @{ Variable='RawEntraDevicesInputPath'; Table=$entraDeviceTable }
)
foreach ($definition in $pathDefinitions) {
    $value = Get-Variable -Name $definition.Variable -ValueOnly
    if ([string]::IsNullOrWhiteSpace($value)) { $value = Join-Path $paths.LatestOutputRootPath (Join-Path ([string]$definition.Table.area) ([string]$definition.Table.name)) }
    Set-Variable -Name $definition.Variable -Value ([IO.Path]::GetFullPath($value))
}
$domainOutputPath = Join-Path $paths.PowerBILatestPath ([string]$domainDimTable.name)
$coverageOutputPath = Join-Path $paths.PowerBILatestPath ([string]$coverageTable.name)
$domainInputStatus = Get-SmartWorkplaceCMDBHeaderStatus $RawDomainInputPath @($domainRawTable.columns | ForEach-Object { [string]$_ })
if ($domainInputStatus -eq 'Incompatible') { throw 'Entra_VerifiedDomains.csv is incompatible with the raw contract.' }
foreach ($target in @(@{Path=$domainOutputPath;Table=$domainDimTable},@{Path=$coverageOutputPath;Table=$coverageTable})) {
    if ((Get-SmartWorkplaceCMDBHeaderStatus $target.Path @($target.Table.columns | ForEach-Object { [string]$_ })) -eq 'Incompatible') { throw "Existing curated table '$([IO.Path]::GetFileName($target.Path))' is incompatible." }
}
if ($ValidateOnly) {
    if ($domainInputStatus -eq 'Valid') { Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $RawDomainInputPath -Paths $paths | Out-Null }
    [pscustomobject]@{ Status='Valid';ScriptVersion=$ScriptVersion;RawDomainInputStatus=$domainInputStatus;DimVerifiedDomainOutputPath=$domainOutputPath;HybridCoverageOutputPath=$coverageOutputPath } | Format-List
    return
}
if ($domainInputStatus -eq 'Missing') { throw 'The raw verified domains CSV was not found.' }
$domainRows = @(Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $RawDomainInputPath -Paths $paths)
$dimRows = @($domainRows | Sort-Object DomainId | ForEach-Object {
    $domainId = Get-SmartWorkplaceCMDBNormalizedKey $_.DomainId
    if ([string]::IsNullOrWhiteSpace($domainId)) { throw 'Verified domain data contains an empty DomainId.' }
    [pscustomobject][ordered]@{
        TenantDomainKey = ('{0}|domain|{1}' -f $paths.TenantKey,$domainId)
        DomainId=$domainId;IsDefault=[string]$_.IsDefault;IsInitial=[string]$_.IsInitial
        AuthenticationType=[string]$_.AuthenticationType;SupportedServices=[string]$_.SupportedServices
        AvailabilityStatus=[string]$_.AvailabilityStatus;SourceCollectedDateTime=[string]$_.SourceCollectedDateTime
    }
})
if (@($dimRows | Group-Object DomainId | Where-Object Count -gt 1).Count -gt 0) { throw 'Verified domain data contains duplicate DomainId values.' }

function Import-SmartWorkplaceCMDBOptionalSource {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @() }
    return @(Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $Path -Paths $paths)
}
$adUsers = @(Import-SmartWorkplaceCMDBOptionalSource $RawActiveDirectoryUsersInputPath)
$entraUsers = @(Import-SmartWorkplaceCMDBOptionalSource $RawEntraUsersInputPath)
$adComputers = @(Import-SmartWorkplaceCMDBOptionalSource $RawActiveDirectoryComputersInputPath)
$entraDevices = @(Import-SmartWorkplaceCMDBOptionalSource $RawEntraDevicesInputPath)
$coverageRows = @(
    New-SmartWorkplaceCMDBCoverageRow -EntityType 'User' -LocalRows $adUsers -CloudRows $entraUsers -LocalKey { param($row) Get-SmartWorkplaceCMDBNormalizedKey $row.UserPrincipalName } -CloudKey { param($row) Get-SmartWorkplaceCMDBNormalizedKey $row.UserPrincipalName } -LocalAvailable:(Test-Path $RawActiveDirectoryUsersInputPath) -CloudAvailable:(Test-Path $RawEntraUsersInputPath) -MatchRule 'Exact normalized UserPrincipalName'
    New-SmartWorkplaceCMDBCoverageRow -EntityType 'Device' -LocalRows $adComputers -CloudRows $entraDevices -LocalKey { param($row) Get-SmartWorkplaceCMDBDeviceKey $row -ActiveDirectory } -CloudKey { param($row) Get-SmartWorkplaceCMDBDeviceKey $row } -LocalAvailable:(Test-Path $RawActiveDirectoryComputersInputPath) -CloudAvailable:(Test-Path $RawEntraDevicesInputPath) -MatchRule 'Exact normalized DNSHostName or DeviceName'
)
$identity = @{TenantKey=$paths.TenantKey;OrganizationKey=$paths.OrganizationKey;EnvironmentKey=$paths.EnvironmentKey;TenantId=$paths.TenantId}
Export-SmartWorkplaceCMDBCsv -InputObject $dimRows -Path $domainOutputPath -Columns @($domainDimTable.columns | ForEach-Object { [string]$_ }) @identity
Export-SmartWorkplaceCMDBCsv -InputObject $coverageRows -Path $coverageOutputPath -Columns @($coverageTable.columns | ForEach-Object { [string]$_ }) @identity
if ((Get-SmartWorkplaceCMDBHeaderStatus $domainOutputPath @($domainDimTable.columns | ForEach-Object { [string]$_ })) -ne 'Valid' -or (Get-SmartWorkplaceCMDBHeaderStatus $coverageOutputPath @($coverageTable.columns | ForEach-Object { [string]$_ })) -ne 'Valid') { throw 'Verified domain or hybrid coverage output failed contract validation.' }
Write-Information ("SmartWorkplaceCMDB tenant identity health normalization completed. Domains={0}; CoverageRows={1}." -f $dimRows.Count,$coverageRows.Count) -InformationAction Continue
[pscustomobject]@{Status='Completed';ScriptVersion=$ScriptVersion;VerifiedDomainCount=$dimRows.Count;CoverageRowCount=$coverageRows.Count;DimVerifiedDomainOutputPath=$domainOutputPath;HybridCoverageOutputPath=$coverageOutputPath}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBFA6A4A/c0NC1x
# dGLZHJetnHFo7KoQ6fvKKYcAsoRUMKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIH4R5sXL/l5WnlWSQvi3si4bzeQ+HXPF93kpk6+F8as+MA0GCSqG
# SIb3DQEBAQUABIIBgIp8ToVIbYyzrkezhLBKc1W8Gg8MHFK1jJdXaaa+nbxjPmN6
# tMh1eaH8qUv/k15hDSJIBXjh3bkfGZmCChHwzBvdda0LA0RDPJuRTmCvcR38BZ9S
# F6fExCEWOM581X+slo6+3/muKezA7hO/pA7aOnGnKzLNVFQs4QeZ5YaksnPkwWOL
# MH3NZXGytqqlkKW1DdKSKe6f/GzjcNigyknGk4V7KXgrGZSbKryvdlsKz/eR7Qh1
# 7A0R+bTgXIUxwjLE1M8DsJcobH+obSRJm8fmK0lkMJgoAGuCybQ/bd4eQXAqzCL9
# ZEy4nHPbeZmeQWu2d/peIbKFqFUJKxbg6oS3ooyk6Zq/DxtKZm7Lgm2N90VAgiXT
# ctgRVnIyW2WG/TPEtWKrWUgJImBQUgCFLeLTVdcGfUkom6zVp1CWuHltMvgSoP6G
# qJHq9aeQA0ZfVkRaodVIAG+ancF4TSsXgAWv5eREbGFrXvGMnG6idgbF0YCc7xih
# wOUkt5KkWLASu2qZ2qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxODA2
# NDVaMC8GCSqGSIb3DQEJBDEiBCCRniXsb7YLkCS9WLnxK/K7SDGaxbbLf7r9AXr/
# y8HzcDANBgkqhkiG9w0BAQEFAASCAgCU/JqDeQU6WyYyzowQYVTjyxqvej3WOIj/
# RnbAVCFNnjZOZW3HvoqegFsCgDphIZ5Q9OlKTJj7dpHm3QLBO1oR3LWJyINVv3FK
# HrNzRojGujjsMG0n6pTDbhkhmFrbRBF0YIq0LCxO3baw+tlwudrWuAhaYRMCTb6O
# l3R/H6yD8qw8UFRMghtFACctYSi2m4iIbLfXGjSfsxIzBZpQxBTEvQOz0B+vzIM1
# /JYUftIy9Jz4MetYO4GcikT7/ctbyTbE76EYzTOvWmL51C9aZKQzIk4FsGe2oiHY
# 5EtTQ1H/gAGrpgukBhehs6W00C0yZXTk1DEcTZBPfaAAFUSbWSdeN0+SiFqZfHgF
# NdgBwpo/ldQy6QDaEPzNVqW6StIjKniQFbd1b+/M605PbwmPpF0cq8IMcCnS0Bop
# iZWUi9cB7t/hehW/EYqI7f9mLSPpbINolSbKt8xLIZn2YA1J5l8o2uI+Hyb6/3bu
# HpGTqOHN6U3zbYYp0IQt363gk4PSx19vi/Jvj/2sb3inUo5hdrcwgh/spWt5z6AO
# bO4gZfEuJOnsdHeyN7MaNYTOkldnYrRS4C+JoLh7VXrz/Gzp22bC5ZMR4R2Am+3h
# QGX5un5jCW7w/RsPmCvWaFbmCUJowyAoJbFPGPynvGUbNxsPLFQd/z4OTuCHbBTO
# USEcfBFC8Q==
# SIG # End signature block
