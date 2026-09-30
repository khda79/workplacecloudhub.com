<#
.SYNOPSIS
Runs offline tests for the SmartWorkplaceCMDB Active Directory collector.

.VERSION
1.0.5
#>
[CmdletBinding()]
param()

$ScriptVersion = '1.0.5'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Assert-SmartWorkplaceCMDBAdTrue {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw $Message
    }
}

function Invoke-SmartWorkplaceCMDBAdTest {
    param([string]$Name, [scriptblock]$Test)
    try {
        & $Test
        $script:Passed++
        Write-Information "[PASS] $Name" -InformationAction Continue
    }
    catch {
        $script:Failed++
        Write-Error "[FAIL] $Name - $($_.Exception.Message)`n$($_.ScriptStackTrace)" -ErrorAction Continue
    }
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot
$collector = Join-Path $projectRoot 'Collectors\ActiveDirectory\SmartWorkplaceCMDB-ActiveDirectory-Collect.ps1'
$normalizer = Join-Path $projectRoot 'Collectors\ActiveDirectory\SmartWorkplaceCMDB-ActiveDirectory-Normalize.ps1'
$orchestrator = Join-Path $projectRoot 'Orchestration\SmartWorkplaceCMDB-Orchestrator.ps1'
$fixture = Join-Path $scriptRoot 'Fixtures\ActiveDirectory.sample.json'
$rawContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
$curatedContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.activedirectory.tables.json'
$coreModulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
Import-Module $coreModulePath -Force

$tokens = $null
$parseErrors = $null
$collectorAst = [System.Management.Automation.Language.Parser]::ParseFile(
    $collector,
    [ref]$tokens,
    [ref]$parseErrors
)
if ($parseErrors.Count -gt 0) {
    throw 'The Active Directory collector cannot be parsed for helper tests.'
}
foreach ($functionName in @(
        'Add-SmartWorkplaceCMDBActiveDirectoryDomainContext',
        'Get-SmartWorkplaceCMDBActiveDirectoryMemberDomainContext',
        'Get-SmartWorkplaceCMDBActiveDirectoryRangedMember',
        'Test-SmartWorkplaceCMDBActiveDirectoryObjectNotFoundError',
        'Test-SmartWorkplaceCMDBTransientActiveDirectoryError',
        'Get-SmartWorkplaceCMDBActiveDirectoryRetryServer',
        'Invoke-SmartWorkplaceCMDBActiveDirectoryDomainOperation',
        'Invoke-SmartWorkplaceCMDBActiveDirectoryDomainCollection',
        'Get-SmartWorkplaceCMDBActiveDirectoryLiveData'
    )) {
    $functionAst = @($collectorAst.FindAll({
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                $node.Name -eq $functionName
            }, $true) | Select-Object -First 1)
    if ($functionAst.Count -ne 1) {
        throw "The Active Directory helper '$functionName' was not found."
    }
    Invoke-Expression $functionAst[0].Extent.Text
}

$tempBase = [IO.Path]::GetTempPath()
$tempRoot = Join-Path $tempBase (
    'SmartWorkplaceCMDB-ActiveDirectory-Tests-{0}' -f [guid]::NewGuid().ToString('N')
)
$runtimeRoot = Join-Path $tempRoot 'Runtime'
$identity = @{
    Tenant = 'test'
    OrganizationKey = 'contoso'
    EnvironmentKey = 'prod'
    TenantKey = 'contoso-prod'
    TenantId = '00000000-0000-0000-0000-000000000001'
    NoConfigWrite = $true
}
$script:Passed = 0
$script:Failed = 0

try {
    Invoke-SmartWorkplaceCMDBAdTest 'Page every bulk AD object query explicitly' {
        $source = Get-Content -LiteralPath $collector -Raw
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $source -match 'ResultPageSize\s*=\s*500' -and
                $source -match 'Get-ADUser\s+-Filter\s+\*\s+@common' -and
                $source -match 'Get-ADGroup\s+-Filter\s+\*\s+@common' -and
                $source -match 'Get-ADComputer\s+-Filter\s+\*\s+@common' -and
                $source -match 'Get-ADOrganizationalUnit\s+-Filter\s+\*\s+@common' -and
                $source -match "'PrimaryGroupID'" -and
                $source -notmatch 'Get-ADGroupMember' -and
                $source -match 'New-SmartWorkplaceCMDBActiveDirectoryLdapConnection' -and
                $source -match '\$ldapConnectionFactory\s*=\s*\$\{function:New-SmartWorkplaceCMDBActiveDirectoryLdapConnection\}' -and
                $source -match '\$rangedMemberReader\s*=\s*\$\{function:Get-SmartWorkplaceCMDBActiveDirectoryRangedMember\}' -and
                $source -match '&\s+\$ldapConnectionFactory' -and
                $source -match '&\s+\$rangedMemberReader' -and
                $source -match '-DomainContext\s+\$memberDomainContext' -and
                $source -match '-Connection\s+\$domainConnection'
            ) `
            -Message 'One or more bulk Active Directory queries are not explicitly paged.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Retrieve more than five thousand group members by LDAP ranges' {
        $largeMembership = @(0..5104 | ForEach-Object {
                'CN=Member-{0},OU=People,DC=example,DC=invalid' -f $_
            })
        $rangeCalls = New-Object System.Collections.Generic.List[string]
        $reader = {
            param([string]$AttributeName, [int]$RangeStart, [int]$RangeEnd)
            $rangeCalls.Add($AttributeName)
            if ($AttributeName -ne ('member;range={0}-{1}' -f $RangeStart, $RangeEnd)) {
                throw 'Unexpected requested range name.'
            }
            $last = [Math]::Min($RangeEnd, $largeMembership.Count - 1)
            $name = if ($last -eq ($largeMembership.Count - 1)) {
                'member;range={0}-*' -f $RangeStart
            }
            else {
                'member;range={0}-{1}' -f $RangeStart, $last
            }
            [pscustomobject]@{
                Name = $name
                Values = @($largeMembership[$RangeStart..$last])
            }
        }.GetNewClosure()
        $members = @(Get-SmartWorkplaceCMDBActiveDirectoryRangedMember `
                -Server 'dc.example.invalid' `
                -GroupDistinguishedName 'CN=Large,DC=example,DC=invalid' `
                -RangeReader $reader)
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $members.Count -eq 5105 -and
                $rangeCalls.Count -eq 6 -and
                $members[0] -eq $largeMembership[0] -and
                $members[-1] -eq $largeMembership[-1]
            ) `
            -Message 'LDAP range retrieval did not return the complete large group.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Reject a repeated LDAP range without looping' {
        $reader = {
            param([string]$AttributeName, [int]$RangeStart, [int]$RangeEnd)
            [pscustomobject]@{
                Name = 'member;range=0-999'
                Values = @('CN=One,DC=example,DC=invalid')
            }
        }
        $rejected = $false
        try {
            Get-SmartWorkplaceCMDBActiveDirectoryRangedMember `
                -Server 'dc.example.invalid' `
                -GroupDistinguishedName 'CN=Broken,DC=example,DC=invalid' `
                -RangeReader $reader | Out-Null
        }
        catch {
            $rejected = $_.Exception.Message -match 'while .* was requested|made no progress'
        }
        Assert-SmartWorkplaceCMDBAdTrue $rejected `
            'A repeated LDAP range was not rejected.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Accept an empty group as a complete ranged result' {
        $reader = {
            param([string]$AttributeName, [int]$RangeStart, [int]$RangeEnd)
            [pscustomobject]@{ Name = ''; Values = @() }
        }
        $members = @(Get-SmartWorkplaceCMDBActiveDirectoryRangedMember `
                -Server 'dc.example.invalid' `
                -GroupDistinguishedName 'CN=Empty,DC=example,DC=invalid' `
                -RangeReader $reader)
        Assert-SmartWorkplaceCMDBAdTrue ($members.Count -eq 0) `
            'An empty LDAP group did not produce an empty result.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Route an unresolved member to its own forest domain' {
        $esContext = [pscustomobject]@{
            Domain = [pscustomobject]@{
                DNSRoot = 'es.example.invalid'
                DistinguishedName = 'DC=es,DC=example,DC=invalid'
            }
            Server = 'dc.es.example.invalid'
        }
        $ptContext = [pscustomobject]@{
            Domain = [pscustomobject]@{
                DNSRoot = 'pt.example.invalid'
                DistinguishedName = 'DC=pt,DC=example,DC=invalid'
            }
            Server = 'dc.pt.example.invalid'
        }
        $inventory = @(
            [pscustomobject]@{ Context = $esContext },
            [pscustomobject]@{ Context = $ptContext }
        )
        $selected = Get-SmartWorkplaceCMDBActiveDirectoryMemberDomainContext `
            -MemberDistinguishedName (
                'CN=Cross Domain User,OU=Users,DC=pt,DC=example,DC=invalid'
            ) `
            -DomainInventory $inventory `
            -DefaultDomainContext $esContext
        $fallback = Get-SmartWorkplaceCMDBActiveDirectoryMemberDomainContext `
            -MemberDistinguishedName 'CN=Unknown,DC=external,DC=invalid' `
            -DomainInventory $inventory `
            -DefaultDomainContext $esContext
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                [object]::ReferenceEquals($selected, $ptContext) -and
                [object]::ReferenceEquals($fallback, $esContext)
            ) `
            -Message 'Cross-domain member resolution selected the wrong domain context.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Page a large domain and preserve primary group semantics' {
        $queryPageSizes = New-Object System.Collections.Generic.List[int]
        $counters = [pscustomobject]@{ Connections = 0; Ranges = 0 }
        $fakeUsers = @(0..5104 | ForEach-Object {
                [pscustomobject]@{
                    ObjectGUID = [guid]('00000000-0000-0000-0001-{0:D12}' -f $_)
                    ObjectSID = 'S-1-5-21-1-2-3-{0}' -f (1000 + $_)
                    DistinguishedName = 'CN=Member-{0},OU=People,DC=example,DC=invalid' -f $_
                    PrimaryGroupID = 513
                }
            })
        $fakeGroup = [pscustomobject]@{
            ObjectGUID = [guid]'00000000-0000-0000-0002-000000000513'
            ObjectSID = 'S-1-5-21-1-2-3-513'
            DistinguishedName = 'CN=Domain Users,CN=Users,DC=example,DC=invalid'
        }
        function global:Get-ADUser {
            param($Filter, $Server, $ErrorAction, $ResultPageSize, $Properties,
                $SearchBase, $ResultSetSize)
            $queryPageSizes.Add([int]$ResultPageSize)
            return $fakeUsers
        }
        function global:Get-ADGroup {
            param($Filter, $Server, $ErrorAction, $ResultPageSize, $Properties,
                $SearchBase, $ResultSetSize)
            $queryPageSizes.Add([int]$ResultPageSize)
            return @($fakeGroup)
        }
        function global:Get-ADComputer {
            param($Filter, $Server, $ErrorAction, $ResultPageSize, $Properties,
                $SearchBase, $ResultSetSize)
            $queryPageSizes.Add([int]$ResultPageSize)
            return @()
        }
        function global:Get-ADOrganizationalUnit {
            param($Filter, $Server, $ErrorAction, $ResultPageSize, $Properties,
                $SearchBase, $ResultSetSize)
            $queryPageSizes.Add([int]$ResultPageSize)
            return @()
        }
        function global:New-SmartWorkplaceCMDBActiveDirectoryLdapConnection {
            param([string]$Server)
            $counters.Connections++
            return [System.IO.MemoryStream]::new()
        }
        function global:Get-SmartWorkplaceCMDBActiveDirectoryRangedMember {
            param([string]$Server, [string]$GroupDistinguishedName,
                [object]$Connection)
            $counters.Ranges++
            return @($fakeUsers.DistinguishedName)
        }
        function global:Get-ADObject { throw 'The complete forest lookup should resolve every test member.' }

        $readiness = [pscustomobject]@{
            Domains = @([pscustomobject]@{
                    Domain = [pscustomobject]@{
                        DNSRoot = 'example.invalid'
                        NetBIOSName = 'EXAMPLE'
                    }
                    Server = 'dc.example.invalid'
                })
        }
        $result = Get-SmartWorkplaceCMDBActiveDirectoryLiveData `
            -Readiness $readiness -CollectMemberships $true -Limit 0 -RetryCount 0
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $result.users.Count -eq 5105 -and
                $result.groups.Count -eq 1 -and
                $result.groupMemberships.Count -eq 5105 -and
                $queryPageSizes.Count -eq 4 -and
                @($queryPageSizes | Where-Object { $_ -ne 500 }).Count -eq 0 -and
                $counters.Connections -eq 1 -and
                $counters.Ranges -eq 1
            ) `
            -Message 'Large-domain paging, connection reuse, or primary-group deduplication failed.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Classify only Active Directory object disappearance errors' {
        $missingError = [System.Management.Automation.ErrorRecord]::new(
            [System.Exception]::new(
                '0000208D: NameErr: DSID-0310028D, problem 2001 (NO_OBJECT)'
            ),
            'ActiveDirectoryObjectMissing',
            [System.Management.Automation.ErrorCategory]::ObjectNotFound,
            $null
        )
        $authorizationError = [System.Management.Automation.ErrorRecord]::new(
            [System.UnauthorizedAccessException]::new('Access is denied.'),
            'ActiveDirectoryAccessDenied',
            [System.Management.Automation.ErrorCategory]::PermissionDenied,
            $null
        )
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                (Test-SmartWorkplaceCMDBActiveDirectoryObjectNotFoundError `
                    $missingError) -and
                -not (Test-SmartWorkplaceCMDBActiveDirectoryObjectNotFoundError `
                    $authorizationError)
            ) `
            -Message 'The LDAP no-such-object classifier accepted an unrelated error.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Recover a moved group and exclude a deleted group from one snapshot' {
        $rangeCalls = New-Object System.Collections.Generic.List[string]
        $fakeUser = [pscustomobject]@{
            ObjectGUID = [guid]'00000000-0000-0000-0003-000000000001'
            ObjectSID = 'S-1-5-21-1-2-3-1101'
            DistinguishedName = 'CN=Member,OU=People,DC=example,DC=invalid'
            PrimaryGroupID = $null
        }
        $movedGroup = [pscustomobject]@{
            ObjectGUID = [guid]'00000000-0000-0000-0004-000000000001'
            ObjectSID = 'S-1-5-21-1-2-3-2101'
            DistinguishedName = 'CN=Moved,OU=Old,DC=example,DC=invalid'
        }
        $deletedGroup = [pscustomobject]@{
            ObjectGUID = [guid]'00000000-0000-0000-0004-000000000002'
            ObjectSID = 'S-1-5-21-1-2-3-2102'
            DistinguishedName = 'CN=Deleted,OU=Old,DC=example,DC=invalid'
        }
        $movedCurrentDn = 'CN=Moved,OU=Current,DC=example,DC=invalid'
        function global:Get-ADUser {
            param($Filter, $Server, $ErrorAction, $ResultPageSize, $Properties,
                $SearchBase, $ResultSetSize)
            return @($fakeUser)
        }
        function global:Get-ADGroup {
            param($Filter, $Identity, $Server, $ErrorAction, $ResultPageSize,
                $Properties, $SearchBase, $ResultSetSize)
            if ($PSBoundParameters.ContainsKey('Identity')) {
                if ([string]$Identity -eq [string]$movedGroup.ObjectGUID) {
                    return [pscustomobject]@{
                        ObjectGUID = $movedGroup.ObjectGUID
                        DistinguishedName = $movedCurrentDn
                    }
                }
                throw [System.Exception]::new(
                    "Cannot find an object with identity '$Identity'."
                )
            }
            return @($movedGroup, $deletedGroup)
        }
        function global:Get-ADComputer {
            param($Filter, $Server, $ErrorAction, $ResultPageSize, $Properties,
                $SearchBase, $ResultSetSize)
            return @()
        }
        function global:Get-ADOrganizationalUnit {
            param($Filter, $Server, $ErrorAction, $ResultPageSize, $Properties,
                $SearchBase, $ResultSetSize)
            return @()
        }
        function global:New-SmartWorkplaceCMDBActiveDirectoryLdapConnection {
            param([string]$Server)
            return [System.IO.MemoryStream]::new()
        }
        function global:Get-SmartWorkplaceCMDBActiveDirectoryRangedMember {
            param([string]$Server, [string]$GroupDistinguishedName,
                [object]$Connection)
            $rangeCalls.Add($GroupDistinguishedName)
            if ($GroupDistinguishedName -eq $movedCurrentDn) {
                return @($fakeUser.DistinguishedName)
            }
            throw [System.Exception]::new(
                '0000208D: NameErr: DSID-0310028D, problem 2001 (NO_OBJECT)'
            )
        }
        function global:Get-ADObject {
            throw 'The known member should resolve from the domain inventory.'
        }

        $readiness = [pscustomobject]@{
            Domains = @([pscustomobject]@{
                    Domain = [pscustomobject]@{
                        DNSRoot = 'example.invalid'
                        NetBIOSName = 'EXAMPLE'
                    }
                    Server = 'dc.example.invalid'
                })
        }
        $warnings = @()
        $result = Get-SmartWorkplaceCMDBActiveDirectoryLiveData `
            -Readiness $readiness `
            -CollectMemberships $true `
            -Limit 0 `
            -RetryCount 0 `
            -WarningVariable warnings
        $warningText = $warnings -join "`n"
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $result.groups.Count -eq 1 -and
                $result.groups[0].ObjectGUID -eq $movedGroup.ObjectGUID -and
                $result.groups[0].DistinguishedName -eq $movedCurrentDn -and
                $result.groupMemberships.Count -eq 1 -and
                $result.groupMemberships[0].GroupDistinguishedName -eq
                    $movedCurrentDn -and
                $result.staleGroupCount -eq 1 -and
                $rangeCalls.Count -eq 3 -and
                $rangeCalls.Contains($movedCurrentDn) -and
                $warningText -match 'moved or was renamed' -and
                $warningText -match 'disappeared during membership collection'
            ) `
            -Message 'A concurrent group move or deletion produced an incomplete current snapshot.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Retry only a transient domain failure on a newly selected ADWS server' {
        $attempts = [pscustomobject]@{ Count = 0; Sleeps = 0; Selections = 0 }
        $domainContext = [pscustomobject]@{
            Domain = [pscustomobject]@{ DNSRoot = 'example.invalid' }
            Server = 'dc01.example.invalid'
        }
        $action = {
            param([string]$SelectedServer)
            $attempts.Count++
            if ($attempts.Count -eq 1) {
                throw [System.IO.IOException]::new('The LDAP server is unavailable.')
            }
            return $SelectedServer
        }.GetNewClosure()
        $selector = {
            param([string]$DomainDnsRoot, [string]$CurrentServer)
            $attempts.Selections++
            return 'dc02.example.invalid'
        }.GetNewClosure()
        $sleep = {
            param([int]$Seconds)
            $attempts.Sleeps++
        }.GetNewClosure()
        $result = Invoke-SmartWorkplaceCMDBActiveDirectoryDomainOperation `
            -DomainContext $domainContext `
            -OperationName 'test inventory' `
            -Action $action `
            -RetryCount 3 `
            -RetryDelaysSeconds @(0, 0, 0) `
            -ServerSelector $selector `
            -SleepAction $sleep
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $result.Value -eq 'dc02.example.invalid' -and
                $result.Server -eq 'dc02.example.invalid' -and
                $result.RetryCount -eq 1 -and
                $attempts.Count -eq 2 -and
                $attempts.Selections -eq 1 -and
                $attempts.Sleeps -eq 1
            ) `
            -Message 'Transient retry did not select and use the replacement ADWS server.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Fail fast on an Active Directory authorization error' {
        $attempts = [pscustomobject]@{ Count = 0; Selections = 0 }
        $domainContext = [pscustomobject]@{
            Domain = [pscustomobject]@{ DNSRoot = 'example.invalid' }
            Server = 'dc01.example.invalid'
        }
        $action = {
            param([string]$SelectedServer)
            $attempts.Count++
            throw [System.UnauthorizedAccessException]::new('Access is denied.')
        }.GetNewClosure()
        $selector = {
            param([string]$DomainDnsRoot, [string]$CurrentServer)
            $attempts.Selections++
            return 'dc02.example.invalid'
        }.GetNewClosure()
        $failed = $false
        try {
            Invoke-SmartWorkplaceCMDBActiveDirectoryDomainOperation `
                -DomainContext $domainContext `
                -OperationName 'test inventory' `
                -Action $action `
                -RetryCount 3 `
                -RetryDelaysSeconds @(0, 0, 0) `
                -ServerSelector $selector `
                -SleepAction { param([int]$Seconds) } | Out-Null
        }
        catch {
            $failed = $_.Exception.Message -match 'Access is denied'
        }
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition ($failed -and $attempts.Count -eq 1 -and $attempts.Selections -eq 0) `
            -Message 'An authorization error was retried instead of failing immediately.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Validate fixture mode without AD connectivity' {
        $result = & $collector @identity `
            -DataRootPath $runtimeRoot `
            -InputJsonPath $fixture `
            -ValidateOnly
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition ((($result | Out-String) -match 'OfflineJson') -and
                (($result | Out-String) -match 'DomainCount\s*:\s*2')) `
            -Message 'Fixture validation did not report OfflineJson mode.'
    }

    $script:Collection = $null
    Invoke-SmartWorkplaceCMDBAdTest 'Collect six Active Directory raw tables offline' {
        $script:Collection = & $collector @identity `
            -DataRootPath $runtimeRoot `
            -InputJsonPath $fixture `
            -IncludeGroupMemberships
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $script:Collection.DomainCount -eq 2 -and
                $script:Collection.UserCount -eq 3 -and
                $script:Collection.GroupCount -eq 2 -and
                $script:Collection.ComputerCount -eq 3 -and
                $script:Collection.OrganizationalUnitCount -eq 3 -and
                $script:Collection.GroupMembershipCount -eq 3
            ) `
            -Message 'Unexpected Active Directory fixture row counts.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Validate all Active Directory raw contracts' {
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
                -LatestOutputRootPath (Join-Path $runtimeRoot 'DATA-LAST') `
                -ContractPath $rawContractPath |
                Where-Object Name -like 'ActiveDirectory_*.csv')
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $results.Count -eq 6 -and
                @($results | Where-Object Status -ne 'Valid').Count -eq 0
            ) `
            -Message 'One or more Active Directory raw contracts are invalid.'
    }

    $script:Normalization = $null
    Invoke-SmartWorkplaceCMDBAdTest 'Normalize Active Directory source tables without AD module' {
        $script:Normalization = & $normalizer @identity -DataRootPath $runtimeRoot
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $script:Normalization.DomainCount -eq 2 -and
                $script:Normalization.UserCount -eq 3 -and
                $script:Normalization.GroupCount -eq 2 -and
                $script:Normalization.ComputerCount -eq 3 -and
                $script:Normalization.OrganizationalUnitCount -eq 3 -and
                $script:Normalization.GroupMembershipCount -eq 3
            ) `
            -Message 'Unexpected normalized Active Directory row counts.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Validate all normalized Active Directory contracts' {
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
                -LatestOutputRootPath (Join-Path $runtimeRoot 'DATA-LAST') `
                -ContractPath $curatedContractPath)
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $results.Count -eq 6 -and
                @($results | Where-Object Status -ne 'Valid').Count -eq 0
            ) `
            -Message 'One or more normalized Active Directory contracts are invalid.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Create stable tenant-scoped Active Directory keys' {
        $userPath = Join-Path $script:Normalization.CuratedOutputRootPath 'CMDB_ActiveDirectoryUsers.csv'
        $membershipPath = Join-Path $script:Normalization.CuratedOutputRootPath 'CMDB_ActiveDirectoryGroupMemberships.csv'
        $ouPath = Join-Path $script:Normalization.CuratedOutputRootPath 'CMDB_ActiveDirectoryOrganizationalUnits.csv'
        $users = @(Import-Csv -LiteralPath $userPath)
        $memberships = @(Import-Csv -LiteralPath $membershipPath)
        $organizationalUnits = @(Import-Csv -LiteralPath $ouPath)
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $users[0].CmdbAdUserId -match '^contoso-prod\\|ad-user\\|' -and
                $memberships[0].CmdbAdGroupId -match '^contoso-prod\\|ad-group\\|' -and
                $memberships[0].CmdbAdMemberId -match '^contoso-prod\\|ad-(user|computer)\\|' -and
                $organizationalUnits[0].CmdbAdOrganizationalUnitId -match '^contoso-prod\\|ad-organizational-unit\\|' -and
                @($users | Where-Object DomainDnsRoot -eq 'child.example.invalid').Count -eq 1 -and
                @($memberships | Where-Object DomainDnsRoot -eq 'child.example.invalid').Count -eq 1
            ) `
            -Message 'Active Directory normalized keys are not stable and tenant-scoped.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Preserve the complete last-valid AD snapshot after a rejected collection' {
        $rawUserPath = Join-Path $runtimeRoot 'DATA-LAST\Raw\ActiveDirectory\ActiveDirectory_Users.csv'
        $rawStatusPath = $rawUserPath + '.status.json.txt'
        $csvHash = (Get-FileHash -LiteralPath $rawUserPath -Algorithm SHA256).Hash
        $statusHash = (Get-FileHash -LiteralPath $rawStatusPath -Algorithm SHA256).Hash
        $badFixture = Join-Path $tempRoot 'ActiveDirectory.invalid.json'
        $badDocument = Get-Content -LiteralPath $fixture -Raw | ConvertFrom-Json
        $badDocument.organizationalUnits = @(
            $badDocument.organizationalUnits + $badDocument.organizationalUnits[0]
        )
        $badDocument | ConvertTo-Json -Depth 12 |
            Set-Content -LiteralPath $badFixture -Encoding UTF8
        $rejected = $false
        try {
            & $collector @identity `
                -DataRootPath $runtimeRoot `
                -InputJsonPath $badFixture `
                -IncludeGroupMemberships | Out-Null
        }
        catch {
            $rejected = $_.Exception.Message -match 'Duplicate values detected'
        }
        $state = Get-Content -LiteralPath $rawStatusPath -Raw | ConvertFrom-Json
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $rejected -and
                (Get-FileHash -LiteralPath $rawUserPath -Algorithm SHA256).Hash -eq $csvHash -and
                (Get-FileHash -LiteralPath $rawStatusPath -Algorithm SHA256).Hash -eq $statusHash -and
                $state.Status -eq 'Completed'
            ) `
            -Message 'A rejected AD collection damaged the last-valid CSV or its completed evidence.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Honor MaxItems in an isolated output root' {
        $boundedRoot = Join-Path $tempRoot 'Bounded'
        $result = & $collector @identity `
            -DataRootPath $boundedRoot `
            -InputJsonPath $fixture `
            -IncludeGroupMemberships `
            -MaxItems 1
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $result.UserCount -eq 1 -and
                $result.ComputerCount -eq 1 -and
                $result.GroupMembershipCount -eq 1
            ) `
            -Message 'MaxItems did not bound the Active Directory fixture output.'
    }

    if ($PSVersionTable.PSVersion.Major -ge 7) {
        Invoke-SmartWorkplaceCMDBAdTest 'Run the Active Directory orchestrator pipeline offline' {
            $orchestratedRoot = Join-Path $tempRoot 'Orchestrated'
            $result = & $orchestrator @identity `
                -DataRootPath $orchestratedRoot `
                -Pipeline ActiveDirectory `
                -FixtureRootPath (Split-Path -Parent $fixture)
            Assert-SmartWorkplaceCMDBAdTrue `
                -Condition ($result.StepCount -eq 2 -and $result.FailedStepCount -eq 0) `
                -Message 'The offline Active Directory orchestrator pipeline did not complete both steps.'
        }
    }
    else {
        Write-Information '[SKIP] Active Directory orchestrator pipeline requires PowerShell 7.' `
            -InformationAction Continue
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Validate centralized Active Directory launchers' {
        $launcherRoot = Join-Path $projectRoot 'Launchers\ActiveDirectory'
        $collectLauncher = Join-Path $launcherRoot 'Start-SmartWorkplaceCMDB-ActiveDirectory-Collect.cmd'
        $validateLauncher = Join-Path $launcherRoot 'Start-SmartWorkplaceCMDB-ActiveDirectory-Validate.cmd'
        $collectText = Get-Content -LiteralPath $collectLauncher -Raw
        $validateText = Get-Content -LiteralPath $validateLauncher -Raw
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $collectText -match '-Pipeline ActiveDirectory' -and
                $collectText -match '-Collect' -and
                $validateText -match '-Pipeline ActiveDirectory' -and
                $validateText -match '-ValidateOnly' -and
                $collectText -match '%\*' -and
                $validateText -match '%\*'
            ) `
            -Message 'The centralized Active Directory launchers are incomplete.'
    }
}
finally {
    $resolvedTempRoot = [IO.Path]::GetFullPath($tempRoot)
    if ($resolvedTempRoot.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $resolvedTempRoot) -like 'SmartWorkplaceCMDB-ActiveDirectory-Tests-*' -and
        (Test-Path -LiteralPath $resolvedTempRoot)) {
        Remove-Item -LiteralPath $resolvedTempRoot -Recurse -Force
    }
}

Write-Information (
    "SmartWorkplaceCMDB Active Directory tests completed. Version={0}; Passed={1}; Failed={2}" -f
    $ScriptVersion, $script:Passed, $script:Failed
) -InformationAction Continue
if ($script:Failed -gt 0) {
    exit 1
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCqJ7V2ESScV0n7
# e1r/FBpuyk8HJIWmKpgj8XE6LpCMNqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCJjmZAFpySKkbJ0p2Wp8uN
# +wI8EtRZaNdtc5Si3X66NjANBgkqhkiG9w0BAQEFAASCAYAgdVOOixAD/a42AC3K
# 37jwRzBkZX/ITWrSoRjudlNzNn/Mwp8LW0nbRHHUsAmTT1K9JJI+1W1MKicW1EH8
# Mih9AO8fD1pltVSd9l+M9ZtGm+QVP1n73gcA/W9LSch8q0CyQthDGWGyCku39rbD
# iPJWeES7I20CFQz8bBAo6LxzdJ8zXQsVfJKnz6POghCCJ3fZ6HD4V/x3XEiH5oMG
# qhARGB5mHMjSAZVVE2p9GNN3/wguGlrYbVKjO2iv78qRvTSCYsThyESv/s0G1BdN
# 6mRp8t2Uixj9oSSr13/VoVjDqxQYdKDrVlCjvIl0XJzXxlET1rkDHPAIld5CK4Kj
# 2saTkllSTu4q4mWYTwsg0BTjQ+JyK6f8zAwd15GFXTIrvzYobFmPzeKMqx9YxU+E
# DjwK5dDYf8zaH5ww1Tu0sv/F0LSrvSpNF31BMXfPX05QyvgaleODwD4UJlmYAYtM
# OAm3XgZb6W7MGGE/PMtlz8CWQ+6+Bbexy5nCxoGTnUTDTJg=
# SIG # End signature block
