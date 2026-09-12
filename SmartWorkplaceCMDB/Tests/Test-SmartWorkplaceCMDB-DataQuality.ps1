<#
.SYNOPSIS
Runs offline tests for SmartWorkplaceCMDB data-quality normalization.

.VERSION
1.0.0
#>
[CmdletBinding()]
param()

$ScriptVersion = '1.0.0'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$script:Passed = 0
$script:Failed = 0

function Invoke-SmartWorkplaceCMDBDataQualityTest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$Body
    )
    try {
        & $Body
        $script:Passed++
        Write-Information "PASS $Name" -InformationAction Continue
    }
    catch {
        $script:Failed++
        Write-Information "FAIL $Name - $($_.Exception.Message)" -InformationAction Continue
    }
}

function Assert-SmartWorkplaceCMDBDataQualityTrue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][bool]$Condition,
        [Parameter(Mandatory)][string]$Message
    )
    if (-not $Condition) {
        throw $Message
    }
}

function Assert-SmartWorkplaceCMDBDataQualityThrow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Body,
        [Parameter(Mandatory)][string]$ExpectedText
    )
    try {
        & $Body
    }
    catch {
        if ($_.Exception.Message -notlike "*$ExpectedText*") {
            throw "Unexpected exception: $($_.Exception.Message)"
        }
        return
    }
    throw "Expected exception containing '$ExpectedText'."
}

function Write-SmartWorkplaceCMDBContractFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Table,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Rows
    )
    $columns = @($Table.columns | ForEach-Object { [string]$_ })
    $normalized = @($Rows | ForEach-Object {
        $values = $_
        $row = [ordered]@{}
        foreach ($column in $columns) {
            $row[$column] = if ($values.ContainsKey($column)) {
                [string]$values[$column]
            }
            else {
                ''
            }
        }
        [pscustomobject]$row
    })
    Export-SmartWorkplaceCMDBCsv `
        -InputObject $normalized `
        -Path $Path `
        -Columns $columns | Out-Null
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$normalizer = Join-Path $projectRoot 'Collectors\SmartWorkplaceCMDB-DataQuality-Normalize.ps1'
$modulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$contractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json'
$rawContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
Import-Module $modulePath -Force
$contract = Get-SmartWorkplaceCMDBTableContract -Path $contractPath
$rawContract = Get-SmartWorkplaceCMDBTableContract -Path $rawContractPath
$tables = @{}
foreach ($name in @(
        'CMDB_Users.csv',
        'CMDB_Groups.csv',
        'CMDB_Devices.csv',
        'CMDB_Licenses.csv',
        'CMDB_Mailboxes.csv',
        'FactMailbox.csv'
    )) {
    $tables[$name] = @($contract.tables | Where-Object name -eq $name)[0]
}
$assignmentTable = @($rawContract.tables |
    Where-Object name -eq 'M365_UserLicenseAssignments.csv')[0]

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) (
    'SmartWorkplaceCMDB-DataQuality-Tests-' + [guid]::NewGuid().ToString('N')
)
$sourceRoot = Join-Path $tempRoot 'Sources'
$runtimeRoot = Join-Path $tempRoot 'Runtime'
$identity = @{
    Tenant = 'test'
    OrganizationKey = 'contoso'
    EnvironmentKey = 'prod'
    TenantKey = 'contoso-prod'
    TenantId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
    NoConfigWrite = $true
}
$referenceDate = [datetimeoffset]'2026-07-19T13:00:00Z'
$collectedDate = '2026-07-19T12:00:00.0000000Z'

try {
    New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null
    $paths = @{
        Users = Join-Path $sourceRoot 'CMDB_Users.csv'
        Groups = Join-Path $sourceRoot 'CMDB_Groups.csv'
        Devices = Join-Path $sourceRoot 'CMDB_Devices.csv'
        Licenses = Join-Path $sourceRoot 'CMDB_Licenses.csv'
        Mailboxes = Join-Path $sourceRoot 'CMDB_Mailboxes.csv'
        MailboxFact = Join-Path $sourceRoot 'FactMailbox.csv'
    }

    Write-SmartWorkplaceCMDBContractFile $tables['CMDB_Users.csv'] $paths.Users @(
        @{
            TenantKey = 'contoso-prod'; OrganizationKey = 'contoso'
            EnvironmentKey = 'prod'; TenantId = $identity.TenantId
            CmdbUserId = 'contoso-prod|entra-user|user-1'
            SourceSystem = 'MicrosoftEntraID'; SourceUserId = 'user-1'
            UsageLocation = 'FR'; UsageLocationStatus = 'Reported'
            SourceCollectedDateTime = $collectedDate
        },
        @{
            TenantKey = 'contoso-prod'; OrganizationKey = 'contoso'
            EnvironmentKey = 'prod'; TenantId = $identity.TenantId
            CmdbUserId = 'contoso-prod|entra-user|user-2'
            SourceSystem = 'MicrosoftEntraID'; SourceUserId = 'user-2'
            UsageLocation = 'FR'; UsageLocationStatus = 'Reported'
            SourceCollectedDateTime = $collectedDate
        }
    )
    Write-SmartWorkplaceCMDBContractFile $tables['CMDB_Groups.csv'] $paths.Groups @()
    Write-SmartWorkplaceCMDBContractFile $tables['CMDB_Devices.csv'] $paths.Devices @(
        @{
            TenantKey = 'contoso-prod'; OrganizationKey = 'contoso'
            EnvironmentKey = 'prod'; TenantId = $identity.TenantId
            CmdbDeviceId = 'contoso-prod|device|device-1'
            SourceSystem = 'MicrosoftIntune'; SourceDeviceId = 'device-1'
            PrimaryUserId = 'user-1'
            SourceCollectedDateTime = $collectedDate
        },
        @{
            TenantKey = 'contoso-prod'; OrganizationKey = 'contoso'
            EnvironmentKey = 'prod'; TenantId = $identity.TenantId
            CmdbDeviceId = 'contoso-prod|device|device-2'
            SourceSystem = 'MicrosoftIntune'; SourceDeviceId = 'device-2'
            PrimaryUserId = 'missing-user'
            SourceCollectedDateTime = $collectedDate
        }
    )
    Write-SmartWorkplaceCMDBContractFile $tables['CMDB_Licenses.csv'] $paths.Licenses @()
    Write-SmartWorkplaceCMDBContractFile $tables['CMDB_Mailboxes.csv'] $paths.Mailboxes @(
        @{
            TenantKey = 'contoso-prod'; OrganizationKey = 'contoso'
            EnvironmentKey = 'prod'; TenantId = $identity.TenantId
            CmdbMailboxId = 'contoso-prod|mailbox|mailbox-1'
            SourceSystem = 'ExchangeOnline'
            SourceCollectedDateTime = $collectedDate
        },
        @{
            TenantKey = 'contoso-prod'; OrganizationKey = 'contoso'
            EnvironmentKey = 'prod'; TenantId = $identity.TenantId
            CmdbMailboxId = 'contoso-prod|mailbox|mailbox-2'
            SourceSystem = 'ExchangeOnline'
            SourceCollectedDateTime = $collectedDate
        }
    )
    Write-SmartWorkplaceCMDBContractFile $tables['FactMailbox.csv'] $paths.MailboxFact @(
        @{
            TenantKey = 'contoso-prod'; OrganizationKey = 'contoso'
            EnvironmentKey = 'prod'; TenantId = $identity.TenantId
            TenantMailboxKey = 'contoso-prod|mailbox|mailbox-1'
            CmdbMailboxId = 'contoso-prod|mailbox|mailbox-1'
            CmdbUserId = 'contoso-prod|entra-user|user-1'
            SourceSystem = 'ExchangeOnline'
        },
        @{
            TenantKey = 'contoso-prod'; OrganizationKey = 'contoso'
            EnvironmentKey = 'prod'; TenantId = $identity.TenantId
            TenantMailboxKey = 'contoso-prod|mailbox|mailbox-2'
            CmdbMailboxId = 'contoso-prod|mailbox|mailbox-2'
            CmdbUserId = ''
            SourceSystem = 'ExchangeOnline'
        }
    )

    $sourceParameters = @{
        UserInputPath = $paths.Users
        GroupInputPath = $paths.Groups
        DeviceInputPath = $paths.Devices
        LicenseInputPath = $paths.Licenses
        MailboxInputPath = $paths.Mailboxes
        MailboxFactInputPath = $paths.MailboxFact
    }

    Invoke-SmartWorkplaceCMDBDataQualityTest 'ValidateOnly stays read-only' {
        $validateRoot = Join-Path $tempRoot 'Validate'
        & $normalizer @identity @sourceParameters `
            -DataRootPath $validateRoot `
            -ReferenceDateTime $referenceDate `
            -ValidateOnly | Out-Null
        Assert-SmartWorkplaceCMDBDataQualityTrue `
            (-not (Test-Path $validateRoot)) `
            'ValidateOnly created data-quality output.'
    }

    Invoke-SmartWorkplaceCMDBDataQualityTest 'Publish reference findings' {
        $script:Normalization = & $normalizer @identity @sourceParameters `
            -DataRootPath $runtimeRoot `
            -ReferenceDateTime $referenceDate
        $rows = @(Import-Csv $script:Normalization.CmdbOutputPath)
        Assert-SmartWorkplaceCMDBDataQualityTrue `
            ($script:Normalization.FindingCount -eq 3 -and
                $script:Normalization.WarningCount -eq 3 -and
                $rows.Count -eq 3) `
            'Reference finding counts are invalid.'
    }

    Invoke-SmartWorkplaceCMDBDataQualityTest 'Map orphan, country, and mailbox findings' {
        $rows = @(Import-Csv $script:Normalization.CmdbOutputPath)
        Assert-SmartWorkplaceCMDBDataQualityTrue `
            (@($rows |
                    Where-Object FindingType -eq 'OrphanPrimaryUserReference').Count -eq 1 -and
                @($rows |
                    Where-Object FindingType -eq 'DeviceCountryUnknown').Count -eq 1 -and
                @($rows |
                    Where-Object FindingType -eq 'UnlinkedMailbox').Count -eq 1) `
            'Reference finding types are invalid.'
    }

    Invoke-SmartWorkplaceCMDBDataQualityTest 'Publish matching Power BI facts' {
        $cmdbRows = @(Import-Csv $script:Normalization.CmdbOutputPath)
        $factRows = @(Import-Csv $script:Normalization.FactOutputPath)
        Assert-SmartWorkplaceCMDBDataQualityTrue `
            ($factRows.Count -eq $cmdbRows.Count -and
                @($factRows |
                    Where-Object {
                        $_.TenantFindingKey -ne $_.FindingId
                    }).Count -eq 0) `
            'FactDataQuality does not match CMDB findings.'
    }

    foreach ($scenario in @(
        @{Name='Discovery without external ID stays visible as information';Type='DiscoveryMailbox';FactType='DiscoveryMailbox';ExternalId='';Linked=$false;Severity='Information';Finding='TechnicalMailboxWithoutUser'},
        @{Name='Discovery with unresolved external ID remains warning';Type='DiscoveryMailbox';FactType='DiscoveryMailbox';ExternalId='missing-user';Linked=$false;Severity='Warning';Finding='UnlinkedMailbox'},
        @{Name='Shared mailbox without user remains warning';Type='SharedMailbox';FactType='SharedMailbox';ExternalId='';Linked=$false;Severity='Warning';Finding='UnlinkedMailbox'},
        @{Name='Room mailbox without user remains warning';Type='RoomMailbox';FactType='RoomMailbox';ExternalId='';Linked=$false;Severity='Warning';Finding='UnlinkedMailbox'},
        @{Name='Unknown mailbox type remains warning';Type='';FactType='';ExternalId='';Linked=$false;Severity='Warning';Finding='UnlinkedMailbox'},
        @{Name='Conflicting mailbox types cannot downgrade warning';Type='UserMailbox';FactType='DiscoveryMailbox';ExternalId='';Linked=$false;Severity='Warning';Finding='UnlinkedMailbox'},
        @{Name='Linked discovery does not emit a finding';Type='DiscoveryMailbox';FactType='DiscoveryMailbox';ExternalId='user-1';Linked=$true;Severity='';Finding=''},
        @{Name='Discovery classification tolerates case and whitespace';Type=' discoverymailbox ';FactType='DiscoveryMailbox';ExternalId=' ';Linked=$false;Severity='Information';Finding='TechnicalMailboxWithoutUser'}
    )) {
        Invoke-SmartWorkplaceCMDBDataQualityTest $scenario.Name {
            $mailRows = @(Import-Csv $paths.Mailboxes)
            $factRows = @(Import-Csv $paths.MailboxFact)
            $mailRows[1].RecipientTypeDetails = $scenario.Type
            $mailRows[1].ExternalDirectoryObjectId = $scenario.ExternalId
            $factRows[1].RecipientTypeDetails = $scenario.FactType
            if ($scenario.Linked) { $factRows[1].CmdbUserId = 'contoso-prod|entra-user|user-1' }
            $caseRoot = Join-Path $tempRoot ([guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $caseRoot | Out-Null
            $mailRows | Export-Csv (Join-Path $caseRoot 'Mailboxes.csv') -NoTypeInformation -Encoding UTF8
            $factRows | Export-Csv (Join-Path $caseRoot 'Fact.csv') -NoTypeInformation -Encoding UTF8
            $caseParams = @{} + $sourceParameters
            $caseParams.MailboxInputPath = Join-Path $caseRoot 'Mailboxes.csv'
            $caseParams.MailboxFactInputPath = Join-Path $caseRoot 'Fact.csv'
            $result = & $normalizer @identity @caseParams -DataRootPath (Join-Path $caseRoot 'Output') -ReferenceDateTime $referenceDate
            $rows = @(Import-Csv $result.CmdbOutputPath | Where-Object EntityType -eq Mailbox)
            if ($scenario.Linked) {
                Assert-SmartWorkplaceCMDBDataQualityTrue ($rows.Count -eq 0) 'Linked mailbox emitted an unlinked finding.'
            } else {
                Assert-SmartWorkplaceCMDBDataQualityTrue ($rows.Count -eq 1 -and $rows[0].Severity -eq $scenario.Severity -and $rows[0].FindingType -eq $scenario.Finding) 'Mailbox classification does not match the source evidence.'
                $baseline = @(Import-Csv $script:Normalization.CmdbOutputPath | Where-Object EntityType -eq Mailbox)[0]
                Assert-SmartWorkplaceCMDBDataQualityTrue ($rows[0].FindingId -eq $baseline.FindingId) 'Classification changed the stable finding key.'
                $facts = @(Import-Csv $result.FactOutputPath | Where-Object EntityType -eq Mailbox)
                Assert-SmartWorkplaceCMDBDataQualityTrue ($facts.Count -eq 1 -and $facts[0].Severity -eq $scenario.Severity -and $facts[0].FindingType -eq $scenario.Finding) 'Power BI fact lost the classified finding.'
            }
        }
    }

    Invoke-SmartWorkplaceCMDBDataQualityTest 'Report deterministic coverage gaps' {
        $caseRoot = Join-Path $tempRoot 'CoverageGaps'
        New-Item -ItemType Directory -Path $caseRoot | Out-Null
        $users = @(Import-Csv $paths.Users)
        $devices = @(Import-Csv $paths.Devices)
        $users[0].UsageLocation = ''
        $users[0].UsageLocationStatus = 'Not Reported'
        $devices[0].PrimaryUserId = ''
        $devices[0].SourceDeviceId = ''
        $userPath = Join-Path $caseRoot 'Users.csv'
        $devicePath = Join-Path $caseRoot 'Devices.csv'
        $users | Export-Csv $userPath -NoTypeInformation -Encoding UTF8
        $devices | Export-Csv $devicePath -NoTypeInformation -Encoding UTF8
        $caseParams = @{} + $sourceParameters
        $caseParams.UserInputPath = $userPath
        $caseParams.DeviceInputPath = $devicePath
        $result = & $normalizer @identity @caseParams `
            -DataRootPath (Join-Path $caseRoot 'Output') `
            -ReferenceDateTime $referenceDate
        $rows = @(Import-Csv $result.CmdbOutputPath)
        foreach ($findingType in @(
                'UserCountryUnknown',
                'DeviceWithoutPrimaryUser',
                'DeviceCountryUnknown',
                'MissingSourceIdentity'
            )) {
            Assert-SmartWorkplaceCMDBDataQualityTrue `
                (@($rows | Where-Object FindingType -eq $findingType).Count -ge 1) `
                "Coverage finding '$findingType' was not reported."
        }
    }

    Invoke-SmartWorkplaceCMDBDataQualityTest 'Report observed license assignment errors' {
        $caseRoot = Join-Path $tempRoot 'LicenseAssignmentError'
        New-Item -ItemType Directory -Path $caseRoot | Out-Null
        $assignmentPath = Join-Path $caseRoot 'Assignments.csv'
        Write-SmartWorkplaceCMDBContractFile $assignmentTable $assignmentPath @(
            @{
                TenantKey = 'contoso-prod'; OrganizationKey = 'contoso'
                EnvironmentKey = 'prod'; TenantId = $identity.TenantId
                SourceSystem = 'MicrosoftEntraID'
                RawAssignmentKey = 'assignment-1'; SourceUserId = 'user-1'
                SkuId = 'sku-1'; AssignmentState = 'Error'
                AssignmentError = 'CountViolation'
                SourceCollectedDateTime = $collectedDate
            },
            @{
                TenantKey = 'contoso-prod'; OrganizationKey = 'contoso'
                EnvironmentKey = 'prod'; TenantId = $identity.TenantId
                SourceSystem = 'MicrosoftEntraID'
                RawAssignmentKey = 'assignment-2'; SourceUserId = 'user-2'
                SkuId = 'sku-2'; AssignmentState = 'Active'
                AssignmentError = ''
                SourceCollectedDateTime = $collectedDate
            }
        )
        $caseParams = @{} + $sourceParameters
        $caseParams.UserLicenseAssignmentInputPath = $assignmentPath
        $result = & $normalizer @identity @caseParams `
            -DataRootPath (Join-Path $caseRoot 'Output') `
            -ReferenceDateTime $referenceDate
        $rows = @(Import-Csv $result.CmdbOutputPath |
            Where-Object FindingType -eq 'ObservedLicenseAssignmentError')
        Assert-SmartWorkplaceCMDBDataQualityTrue `
            ($rows.Count -eq 1 -and
                $rows[0].EntityId -eq 'contoso-prod|entra-user|user-1') `
            'License assignment error mapping is invalid.'
    }

    Invoke-SmartWorkplaceCMDBDataQualityTest 'Apply critical freshness threshold' {
        $staleRoot = Join-Path $tempRoot 'Stale'
        $staleReference = $referenceDate.AddHours(200)
        $result = & $normalizer @identity @sourceParameters `
            -DataRootPath $staleRoot `
            -ReferenceDateTime $staleReference
        $rows = @(Import-Csv $result.CmdbOutputPath)
        Assert-SmartWorkplaceCMDBDataQualityTrue `
            (@($rows |
                    Where-Object FindingType -eq 'StaleDataset').Count -eq 3 -and
                @($rows |
                    Where-Object {
                        $_.FindingType -eq 'StaleDataset' -and
                        $_.Severity -eq 'Critical'
                    }).Count -eq 3) `
            'Stale dataset findings are invalid.'
    }

    Invoke-SmartWorkplaceCMDBDataQualityTest 'Report missing collection dates' {
        $users = @(Import-Csv $paths.Users)
        $users[0].SourceCollectedDateTime = ''
        $missingDatePath = Join-Path $tempRoot 'MissingDateUsers.csv'
        $users | Export-Csv $missingDatePath -NoTypeInformation -Encoding UTF8
        $missingRoot = Join-Path $tempRoot 'MissingDate'
        $missingParameters = @{} + $sourceParameters
        $missingParameters['UserInputPath'] = $missingDatePath
        $result = & $normalizer @identity @missingParameters `
            -DataRootPath $missingRoot `
            -ReferenceDateTime $referenceDate
        $rows = @(Import-Csv $result.CmdbOutputPath)
        Assert-SmartWorkplaceCMDBDataQualityTrue `
            (@($rows |
                    Where-Object FindingType -eq 'MissingSourceCollectedDateTime').Count -eq 1) `
            'Missing source collection date was not reported.'
    }

    Invoke-SmartWorkplaceCMDBDataQualityTest 'Reject tenant identity mismatch' {
        $users = @(Import-Csv $paths.Users)
        $users[0].TenantKey = 'wrong-prod'
        $brokenPath = Join-Path $tempRoot 'BrokenUsers.csv'
        $users | Export-Csv $brokenPath -NoTypeInformation -Encoding UTF8
        $brokenParameters = @{} + $sourceParameters
        $brokenParameters['UserInputPath'] = $brokenPath
        Assert-SmartWorkplaceCMDBDataQualityThrow {
            & $normalizer @identity @brokenParameters `
                -DataRootPath (Join-Path $tempRoot 'Broken') `
                -ReferenceDateTime $referenceDate | Out-Null
        } 'tenant identity mismatch'
    }

    Invoke-SmartWorkplaceCMDBDataQualityTest 'Validate contracts and stable keys' {
        $contractResults = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath (Join-Path $runtimeRoot 'DATA-LAST') `
            -ContractPath $contractPath)
        $targets = @($contractResults |
            Where-Object Name -in @(
                'CMDB_DataQuality.csv',
                'FactDataQuality.csv'
            ))
        $rows = @(Import-Csv $script:Normalization.CmdbOutputPath)
        Assert-SmartWorkplaceCMDBDataQualityTrue `
            ($targets.Count -eq 2 -and
                @($targets | Where-Object Status -ne 'Valid').Count -eq 0 -and
                @($rows |
                    Group-Object FindingId |
                    Where-Object Count -gt 1).Count -eq 0) `
            'Data-quality contracts or stable keys are invalid.'
    }
}
finally {
    $resolved = [IO.Path]::GetFullPath($tempRoot)
    $tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    if ($resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and
        (Test-Path $resolved)) {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

Write-Information (
    "SmartWorkplaceCMDB data-quality tests completed. Version={0}; Passed={1}; Failed={2}" -f
    $ScriptVersion,
    $script:Passed,
    $script:Failed
) -InformationAction Continue
if ($script:Failed -gt 0) {
    exit 1
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAwGvGcTInBPXWs
# qILKB9z90ZYhzC/p5WvIBTrvFVesiqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILlbuXwdJdWaJ0O6AORsaufBayWILgxfpYtrNUwULxigMA0GCSqG
# SIb3DQEBAQUABIIBgHL7mtDQR4R/OP+SA2ti7fGwPcOOT60UmgGXwvtHzLuLc0u2
# uwWqpvx3F4L+ZwMTJYk/L5rlazO6E3GcCihNtCcwVsOY1jU6JPfDU8Wh2HYaaT2o
# u+HHGMnxTnBkQ0KktXVgRSClptBemHmmwNKsM0SZSP9ldEWSPnpnJPnsj0/0MMVk
# 8ho5qzZh1PmEr/+o1g6zxdUFyJxBab4pCL/ghMK5DCbU6VHvQxQ/CHaM+AQf/Had
# T9SYjPqkIoUYB0Vu6uqEDBvn9WMmMLDVYBI1g4GcsI5HNUD4jHvbMbBeJvRVIhde
# V56lJStbF9dTyg7oOZWLLLp5ka0bkKd+mA27OBM+FxzCcKAyo2uIfLEHVvY7MyPI
# Kf/O+GVhM02B7szkj7B6lrpsOKapoQS5JDDXTBN+izMsIJFs4TPD+uQXEzwWGT+l
# O/GdtQJrVhPFaIO4zi9L4XmqvW16fjzchjccVXBQEyhGvzSW+Shg6zxy12LibAAr
# 6HVub6DU31lRYSMFkaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxMzI1
# MjRaMC8GCSqGSIb3DQEJBDEiBCAMSLCVulpVjb/K4YC00lyrJwNLDENPzw/s4VTE
# BOl8IzANBgkqhkiG9w0BAQEFAASCAgAVUA1Fwy4rV9tUs1Wu8vpjXbrUNPQMrYKu
# wpfRqlZDsfaGFf9QPqDBmt21amhhV6nLlWPIzShxbGLMpCFShIwEnoPr1GMHDAFp
# GYgp6siWmrzePSSiNCwM5uYoly4r8LCFL59Jy+6bZ4N4L7hEm67fNWUFiuuuL/+B
# PpxSviBt0AHztyUeV2ApXAW/yjRYNTv+QlDhOzhWTPWR1yCeqojpEAWCcfMSO6N4
# xO4MTZgQwZ8PTjn+ySr56djsJWtYMOKhxXsuN9B/rE34U2kxLBvmtyO9m8GmRNM6
# Dj04djaXW7VOg4NK/IPzZu16+jVEoOs27T3X7UcTDL6c1FwQczxgwiznklWj/cay
# nB3xDa8sU9CzHRej3bOVBy9FtOGZ35dQRhJIcbe7hjiHd7QY/JKTb2WJ1fNt0q3U
# Xc2Sx/xsCuhO7jZ1LJdvTl9umCwWYjzIwyF22VzQT0D2Ofy2CNkDjXsTB7Hgy3nq
# 8SlJijZMezkSejLimx2ryyerOLI6E7pO9IQKjks5G4GPtBHHRiayCzkWYvB/atvl
# xRr2DW9p0EfX7AU8e12SZNH7uGYFsFKqHYTT4IpkwpeGqrl9U3uVab4sWikCHCrU
# eySWaG6KVEkwFh1msGG1LsbFdu1NHlAwlgEkxIQouy0l7+rf9PwWPwDayZDu9EYe
# 0Ryzv3ZSeg==
# SIG # End signature block
