<#
.SYNOPSIS
Collects a read-only Active Directory snapshot for SmartWorkplaceCMDB.

.DESCRIPTION
Collects domain metadata, users, groups, computers, and optionally direct group
memberships by using the Windows ActiveDirectory module. It runs on the same
SmartWorkplaceCMDB collection host as the cloud collectors; that host must also
have network access to a domain controller. Outputs are written to the shared
tenant DATA-ALL and DATA-LAST locations. Offline JSON input is supported for
development and tests on machines that cannot reach Active Directory.

.VERSION
1.0.1

.REQUIREMENTS
PowerShell 7 on the SmartWorkplaceCMDB collection host.
The RSAT ActiveDirectory module for live collection.
Read access to Active Directory.
#>
[CmdletBinding(DefaultParameterSetName = 'Live')]
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
    [string]$Server,
    [string]$SearchBase,
    [switch]$IncludeGroupMemberships,
    [Parameter(ParameterSetName = 'Fixture', Mandatory)]
    [string]$InputJsonPath,
    [ValidateRange(0, 2147483647)]
    [int]$MaxItems = 0,
    [switch]$NoConfigWrite,
    [switch]$ValidateOnly
)

$ScriptVersion = '1.0.1'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Get-SmartWorkplaceCMDBObjectValue {
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

function Get-SmartWorkplaceCMDBSetting {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Configuration,
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()]$DefaultValue = ''
    )

    if ($Configuration.Contains($Name)) {
        return $Configuration[$Name]
    }
    return $DefaultValue
}

function ConvertTo-SmartWorkplaceCMDBCleanText {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) {
        return ''
    }
    return ([string]$Value -replace "`r`n|`n|`r", ' ').Trim()
}

function ConvertTo-SmartWorkplaceCMDBGuidText {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return ''
    }
    if ($Value -is [guid]) {
        return $Value.ToString()
    }
    $parsed = [guid]::Empty
    if ([guid]::TryParse([string]$Value, [ref]$parsed)) {
        return $parsed.ToString()
    }
    throw "Invalid Active Directory object GUID: '$Value'."
}

function ConvertTo-SmartWorkplaceCMDBSidText {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) {
        return ''
    }
    if ($Value -is [System.Security.Principal.SecurityIdentifier]) {
        return $Value.Value
    }
    $valueProperty = $Value.PSObject.Properties['Value']
    if ($null -ne $valueProperty) {
        return ConvertTo-SmartWorkplaceCMDBCleanText $valueProperty.Value
    }
    return ConvertTo-SmartWorkplaceCMDBCleanText $Value
}

function ConvertTo-SmartWorkplaceCMDBDateText {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return ''
    }
    if ($Value -is [datetimeoffset]) {
        return $Value.ToUniversalTime().ToString('o')
    }
    if ($Value -is [datetime]) {
        return ([datetimeoffset]$Value).ToUniversalTime().ToString('o')
    }
    $parsed = [datetimeoffset]::MinValue
    if ([datetimeoffset]::TryParse(
            [string]$Value,
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::AssumeUniversal,
            [ref]$parsed
        )) {
        return $parsed.ToUniversalTime().ToString('o')
    }
    throw "Invalid Active Directory date value: '$Value'."
}

function Get-SmartWorkplaceCMDBOrganizationalUnit {
    [CmdletBinding()]
    param([AllowNull()]$DistinguishedName)

    $text = ConvertTo-SmartWorkplaceCMDBCleanText $DistinguishedName
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }
    $parts = @($text -split '(?<!\\),')
    if ($parts.Count -le 1) {
        return ''
    }
    return (($parts | Select-Object -Skip 1) -join ',')
}

function Read-SmartWorkplaceCMDBActiveDirectoryFixture {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $document = Read-SmartWorkplaceCMDBJsonFile -Path $Path
    foreach ($name in @('users', 'groups', 'computers', 'groupMemberships')) {
        if ($null -eq (Get-SmartWorkplaceCMDBObjectValue $document $name)) {
            throw "Offline Active Directory JSON is missing '$name': $Path"
        }
    }
    $domains = @(Get-SmartWorkplaceCMDBObjectValue $document 'domains')
    $legacyDomain = Get-SmartWorkplaceCMDBObjectValue $document 'domain'
    if ($domains.Count -eq 0 -and $null -eq $legacyDomain) {
        throw "Offline Active Directory JSON must contain 'domains' or legacy 'domain': $Path"
    }
    return $document
}

function Add-SmartWorkplaceCMDBActiveDirectoryDomainContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$InputObject,
        [Parameter(Mandatory)][string]$DomainDnsRoot,
        [Parameter(Mandatory)][string]$DomainNetBIOSName
    )

    $InputObject | Add-Member -NotePropertyName DomainDnsRoot `
        -NotePropertyValue $DomainDnsRoot -Force
    $InputObject | Add-Member -NotePropertyName DomainNetBIOSName `
        -NotePropertyValue $DomainNetBIOSName -Force
    return $InputObject
}

function Get-SmartWorkplaceCMDBActiveDirectoryDomainValue {
    [CmdletBinding()]
    param(
        [AllowNull()]$InputObject,
        [Parameter(Mandatory)][string]$Name,
        [string]$DefaultValue
    )

    $value = ConvertTo-SmartWorkplaceCMDBCleanText (
        Get-SmartWorkplaceCMDBObjectValue $InputObject $Name
    )
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $DefaultValue
    }
    return $value
}
function Get-SmartWorkplaceCMDBActiveDirectoryReadiness {
    [CmdletBinding()]
    param(
        [string]$PreferredServer,
        [bool]$ForestWide
    )

    $module = Get-Module -ListAvailable -Name ActiveDirectory |
        Sort-Object Version -Descending |
        Select-Object -First 1
    if ($null -eq $module) {
        throw 'The RSAT ActiveDirectory PowerShell module is required for live collection.'
    }

    Import-Module ActiveDirectory -ErrorAction Stop
    $bootstrapParameters = @{}
    if (-not [string]::IsNullOrWhiteSpace($PreferredServer)) {
        $bootstrapParameters['Server'] = $PreferredServer
    }
    $bootstrapDomain = Get-ADDomain @bootstrapParameters -ErrorAction Stop
    $forest = Get-ADForest @bootstrapParameters -ErrorAction Stop
    $domainNames = if ($ForestWide) {
        @($forest.Domains | Sort-Object -Unique)
    }
    else {
        @([string]$bootstrapDomain.DNSRoot)
    }
    if ($domainNames.Count -eq 0) {
        throw 'Active Directory forest discovery returned no domains.'
    }

    $domainContexts = New-Object System.Collections.Generic.List[object]
    foreach ($domainName in $domainNames) {
        $domainServer = ''
        if (-not [string]::IsNullOrWhiteSpace($PreferredServer) -and
            $domainName -ieq [string]$bootstrapDomain.DNSRoot) {
            $domainServer = $PreferredServer
        }
        elseif ($domainName -ieq [string]$bootstrapDomain.DNSRoot -and
            -not [string]::IsNullOrWhiteSpace([string]$bootstrapDomain.PDCEmulator)) {
            $domainServer = [string]$bootstrapDomain.PDCEmulator
        }
        else {
            $domainController = Get-ADDomainController -Discover `
                -DomainName $domainName `
                -Service ADWS `
                -ErrorAction Stop
            $domainServer = [string]$domainController.HostName
        }
        if ([string]::IsNullOrWhiteSpace($domainServer)) {
            throw "No Active Directory Web Services domain controller was discovered for '$domainName'."
        }
        $domain = Get-ADDomain -Identity $domainName `
            -Server $domainServer `
            -ErrorAction Stop
        $domainContexts.Add([pscustomobject]@{
            Domain = $domain
            Server = $domainServer
        })
    }

    return [pscustomobject]@{
        ModuleName    = $module.Name
        ModuleVersion = $module.Version.ToString()
        ForestName    = [string]$forest.Name
        ForestWide    = $ForestWide
        Domains       = @($domainContexts.ToArray())
        Server        = if ([string]::IsNullOrWhiteSpace($PreferredServer)) {
            [string]$bootstrapDomain.PDCEmulator
        }
        else {
            $PreferredServer
        }
    }
}

function New-SmartWorkplaceCMDBActiveDirectoryLdapConnection {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Server)

    Add-Type -AssemblyName System.DirectoryServices.Protocols
    $identifier = [System.DirectoryServices.Protocols.LdapDirectoryIdentifier]::new(
        $Server,
        389,
        $false,
        $false
    )
    $connection = [System.DirectoryServices.Protocols.LdapConnection]::new($identifier)
    try {
        $connection.AuthType = [System.DirectoryServices.Protocols.AuthType]::Negotiate
        $connection.Timeout = [timespan]::FromMinutes(2)
        $connection.SessionOptions.ProtocolVersion = 3
        $connection.SessionOptions.Signing = $true
        $connection.SessionOptions.Sealing = $true
        $connection.Bind()
        return $connection
    }
    catch {
        $connection.Dispose()
        throw
    }
}

function Get-SmartWorkplaceCMDBActiveDirectoryRangedMember {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$GroupDistinguishedName,
        [ValidateRange(1, 1000)][int]$RangeSize = 1000,
        [object]$Connection,
        [scriptblock]$RangeReader
    )

    $ownsConnection = $false
    if ($null -eq $RangeReader) {
        if ($null -eq $Connection) {
            $Connection = New-SmartWorkplaceCMDBActiveDirectoryLdapConnection -Server $Server
            $ownsConnection = $true
        }

        $RangeReader = {
            param([string]$AttributeName, [int]$RangeStart, [int]$RangeEnd)
            $request = [System.DirectoryServices.Protocols.SearchRequest]::new(
                $GroupDistinguishedName,
                '(objectClass=group)',
                [System.DirectoryServices.Protocols.SearchScope]::Base,
                [string[]]@($AttributeName)
            )
            $response = [System.DirectoryServices.Protocols.SearchResponse](
                $Connection.SendRequest($request)
            )
            if ($response.Entries.Count -ne 1) {
                throw "LDAP range retrieval did not return the requested group '$GroupDistinguishedName'."
            }
            $entry = $response.Entries[0]
            $returnedName = @($entry.Attributes.AttributeNames | Where-Object {
                    [string]$_ -ieq 'member' -or
                    [string]$_ -imatch ('^member;range={0}-' -f $RangeStart)
                } | Select-Object -First 1)
            if ($returnedName.Count -eq 0) {
                return [pscustomobject]@{ Name = ''; Values = @() }
            }
            $name = [string]$returnedName[0]
            return [pscustomobject]@{
                Name = $name
                Values = @($entry.Attributes[$name].GetValues([string]))
            }
        }.GetNewClosure()
    }

    $values = New-Object System.Collections.Generic.List[string]
    $seen = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $start = 0
    try {
        while ($true) {
            $end = $start + $RangeSize - 1
            $requestedName = 'member;range={0}-{1}' -f $start, $end
            $result = & $RangeReader $requestedName $start $end
            if ($null -eq $result) {
                throw "LDAP range retrieval returned no response for '$requestedName'."
            }
            $returnedName = [string]$result.Name
            $returnedValues = @($result.Values)
            if ([string]::IsNullOrWhiteSpace($returnedName)) {
                if ($start -eq 0 -and $returnedValues.Count -eq 0) { break }
                throw "LDAP range retrieval omitted '$requestedName' before the final range."
            }

            $isFinal = $false
            $nextStart = -1
            if ($returnedName -ieq 'member') {
                if ($start -ne 0) {
                    throw "LDAP range retrieval returned an unscoped member attribute after '$requestedName'."
                }
                $isFinal = $true
            }
            elseif ($returnedName -imatch '^member;range=(\d+)-(\d+|\*)$') {
                $returnedStart = [int]$Matches[1]
                if ($returnedStart -ne $start) {
                    throw "LDAP range retrieval returned '$returnedName' while '$requestedName' was requested."
                }
                if ($Matches[2] -eq '*') {
                    $isFinal = $true
                }
                else {
                    $nextStart = [int]$Matches[2] + 1
                    if ($nextStart -le $start) {
                        throw "LDAP range retrieval made no progress after '$returnedName'."
                    }
                }
            }
            else {
                throw "LDAP range retrieval returned an unexpected attribute '$returnedName'."
            }

            foreach ($value in $returnedValues) {
                $text = [string]$value
                if (-not [string]::IsNullOrWhiteSpace($text) -and $seen.Add($text)) {
                    $values.Add($text)
                }
            }
            if ($isFinal) { break }
            $start = $nextStart
        }
        return @($values.ToArray())
    }
    finally {
        if ($ownsConnection -and $null -ne $Connection) { $Connection.Dispose() }
    }
}

function Get-SmartWorkplaceCMDBActiveDirectoryLiveData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Readiness,
        [string]$PreferredSearchBase,
        [bool]$CollectMemberships,
        [int]$Limit
    )

    if (-not [string]::IsNullOrWhiteSpace($PreferredSearchBase) -and
        @($Readiness.Domains).Count -gt 1) {
        throw 'ActiveDirectory.SearchBase can only be used when ForestWide is false. Leave SearchBase empty for forest-wide collection.'
    }

    $users = New-Object System.Collections.Generic.List[object]
    $groups = New-Object System.Collections.Generic.List[object]
    $computers = New-Object System.Collections.Generic.List[object]
    $memberships = New-Object System.Collections.Generic.List[object]
    $domainInventories = New-Object System.Collections.Generic.List[object]

    $domainIndex = 0
    foreach ($domainContext in @($Readiness.Domains)) {
        $domainIndex++
        $domain = $domainContext.Domain
        $domainDnsRoot = [string]$domain.DNSRoot
        $domainNetBIOSName = [string]$domain.NetBIOSName
        Write-Information (
            "Active Directory domain [{0}/{1}] DNS='{2}' Server='{3}'." -f
            $domainIndex, @($Readiness.Domains).Count, $domainDnsRoot, $domainContext.Server
        ) -InformationAction Continue

        $common = @{
            Server = $domainContext.Server
            ErrorAction = 'Stop'
            ResultPageSize = 500
        }
        if (-not [string]::IsNullOrWhiteSpace($PreferredSearchBase)) {
            $common['SearchBase'] = $PreferredSearchBase
        }
        if ($Limit -gt 0) {
            $common['ResultSetSize'] = $Limit
        }

        $domainUsers = @(Get-ADUser -Filter * @common -Properties @(
                'ObjectGUID', 'ObjectSID', 'UserPrincipalName', 'DisplayName',
                'Enabled', 'Department', 'Title', 'Mail', 'EmployeeID',
                'DistinguishedName', 'Manager', 'WhenCreated', 'WhenChanged',
                'LastLogonDate', 'PasswordLastSet', 'PrimaryGroupID'
            ))
        $domainGroups = @(Get-ADGroup -Filter * @common -Properties @(
                'ObjectGUID', 'ObjectSID', 'SamAccountName', 'DisplayName',
                'GroupCategory', 'GroupScope', 'Mail', 'DistinguishedName',
                'WhenCreated', 'WhenChanged'
            ))
        $domainComputers = @(Get-ADComputer -Filter * @common -Properties @(
                'ObjectGUID', 'ObjectSID', 'SamAccountName', 'Name',
                'DNSHostName', 'Enabled', 'OperatingSystem',
                'OperatingSystemVersion', 'IPv4Address', 'DistinguishedName',
                'ManagedBy', 'WhenCreated', 'WhenChanged', 'LastLogonDate',
                'PrimaryGroupID'
            ))

        foreach ($item in $domainUsers) {
            $users.Add((Add-SmartWorkplaceCMDBActiveDirectoryDomainContext `
                    $item $domainDnsRoot $domainNetBIOSName))
        }
        foreach ($item in $domainGroups) {
            $groups.Add((Add-SmartWorkplaceCMDBActiveDirectoryDomainContext `
                    $item $domainDnsRoot $domainNetBIOSName))
        }
        foreach ($item in $domainComputers) {
            $computers.Add((Add-SmartWorkplaceCMDBActiveDirectoryDomainContext `
                    $item $domainDnsRoot $domainNetBIOSName))
        }

        $domainInventories.Add([pscustomobject]@{
            Context = $domainContext
            DomainDnsRoot = $domainDnsRoot
            DomainNetBIOSName = $domainNetBIOSName
            Users = $domainUsers
            Groups = $domainGroups
            Computers = $domainComputers
        })
    }

    if ($CollectMemberships) {
        $principalByDistinguishedName = [System.Collections.Generic.Dictionary[string, object]]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )
        $primaryMembersByDomainAndRid = @{}
        foreach ($inventory in $domainInventories) {
            foreach ($set in @(
                    @{ Items = $inventory.Users; Class = 'user' },
                    @{ Items = $inventory.Groups; Class = 'group' },
                    @{ Items = $inventory.Computers; Class = 'computer' }
                )) {
                foreach ($item in @($set.Items)) {
                    $dn = [string]$item.DistinguishedName
                    if (-not [string]::IsNullOrWhiteSpace($dn)) {
                        $principalByDistinguishedName[$dn] = [pscustomobject]@{
                            ObjectGUID = $item.ObjectGUID
                            SID = $item.ObjectSID
                            DistinguishedName = $dn
                            ObjectClass = $set.Class
                        }
                    }
                    if (-not [string]::IsNullOrWhiteSpace($dn) -and
                        $set.Class -in @('user', 'computer') -and
                        $null -ne $item.PrimaryGroupID) {
                        $primaryKey = '{0}|{1}' -f (
                            [string]$inventory.DomainDnsRoot
                        ).ToLowerInvariant(), [string]$item.PrimaryGroupID
                        if (-not $primaryMembersByDomainAndRid.ContainsKey($primaryKey)) {
                            $primaryMembersByDomainAndRid[$primaryKey] = New-Object System.Collections.Generic.List[object]
                        }
                        $primaryMembersByDomainAndRid[$primaryKey].Add(
                            $principalByDistinguishedName[$dn]
                        )
                    }
                }
            }
        }

        foreach ($inventory in $domainInventories) {
            $domainContext = $inventory.Context
            $domainDnsRoot = [string]$inventory.DomainDnsRoot
            $domainNetBIOSName = [string]$inventory.DomainNetBIOSName
            $domainGroups = @($inventory.Groups)
            $started = [datetimeoffset]::UtcNow
            $domainConnection = New-SmartWorkplaceCMDBActiveDirectoryLdapConnection `
                -Server $domainContext.Server
            try {
                for ($index = 0; $index -lt $domainGroups.Count; $index++) {
                    $group = $domainGroups[$index]
                    if ($index -eq 0 -or (($index + 1) % 25) -eq 0 -or
                        $index -eq ($domainGroups.Count - 1)) {
                        $elapsed = [datetimeoffset]::UtcNow - $started
                        $eta = if ($index -gt 0) {
                            [timespan]::FromSeconds(
                                ($elapsed.TotalSeconds / $index) * ($domainGroups.Count - $index)
                            ).ToString('hh\:mm\:ss')
                        }
                        else {
                            'estimating'
                        }
                        Write-Information (
                            "Active Directory memberships '{0}' [{1}/{2}] elapsed={3}; ETA={4}" -f
                            $domainDnsRoot, ($index + 1), $domainGroups.Count,
                            $elapsed.ToString('hh\:mm\:ss'), $eta
                        ) -InformationAction Continue
                    }

                    try {
                        $memberDistinguishedNames = @(Get-SmartWorkplaceCMDBActiveDirectoryRangedMember `
                                -Server $domainContext.Server `
                                -GroupDistinguishedName $group.DistinguishedName `
                                -Connection $domainConnection)
                    }
                    catch {
                        throw "Active Directory membership retrieval failed for domain '$domainDnsRoot', group '$($group.DistinguishedName)', server '$($domainContext.Server)': $($_.Exception.Message)"
                    }
                    $groupMembers = New-Object System.Collections.Generic.List[object]
                    foreach ($memberDn in $memberDistinguishedNames) {
                        $member = $null
                        if ($principalByDistinguishedName.ContainsKey($memberDn)) {
                            $member = $principalByDistinguishedName[$memberDn]
                        }
                        else {
                            $resolved = Get-ADObject -Identity $memberDn `
                                -Server $domainContext.Server -ErrorAction Stop `
                                -Properties @('ObjectGUID', 'ObjectSID', 'ObjectClass', 'DistinguishedName')
                            if ($null -ne $resolved.ObjectSID) {
                                $member = [pscustomobject]@{
                                    ObjectGUID = $resolved.ObjectGUID
                                    SID = $resolved.ObjectSID
                                    DistinguishedName = $resolved.DistinguishedName
                                    ObjectClass = $resolved.ObjectClass
                                }
                                $principalByDistinguishedName[$memberDn] = $member
                            }
                        }
                        if ($null -ne $member) { $groupMembers.Add($member) }
                    }

                    $groupSid = [string]$group.ObjectSID
                    if ($groupSid -match '-(\d+)$') {
                        $primaryKey = '{0}|{1}' -f $domainDnsRoot.ToLowerInvariant(), $Matches[1]
                        if ($primaryMembersByDomainAndRid.ContainsKey($primaryKey)) {
                            foreach ($member in $primaryMembersByDomainAndRid[$primaryKey]) {
                                $groupMembers.Add($member)
                            }
                        }
                    }

                    $seenMemberGuid = [System.Collections.Generic.HashSet[string]]::new(
                        [System.StringComparer]::OrdinalIgnoreCase
                    )
                    foreach ($member in $groupMembers) {
                        $memberGuid = [string]$member.ObjectGUID
                        if ([string]::IsNullOrWhiteSpace($memberGuid) -or
                            -not $seenMemberGuid.Add($memberGuid)) {
                            continue
                        }
                        $memberships.Add([pscustomobject]@{
                            DomainDnsRoot          = $domainDnsRoot
                            DomainNetBIOSName      = $domainNetBIOSName
                            GroupObjectGuid        = $group.ObjectGUID
                            GroupDistinguishedName = $group.DistinguishedName
                            MemberObjectGuid       = $member.ObjectGUID
                            MemberObjectSid        = $member.SID
                            MemberDistinguishedName = $member.DistinguishedName
                            MemberObjectClass      = $member.ObjectClass
                        })
                    }
                }
            }
            finally {
                $domainConnection.Dispose()
            }
        }
    }

    return [pscustomobject]@{
        domains          = @($Readiness.Domains | ForEach-Object { $_.Domain })
        users            = @($users.ToArray())
        groups           = @($groups.ToArray())
        computers        = @($computers.ToArray())
        groupMemberships = @($memberships.ToArray())
    }
}


$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $scriptRoot)
$modulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$rawContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
Import-Module $modulePath -Force

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

$adConfiguration = Get-SmartWorkplaceCMDBConfigSection `
    -Configuration $context.Configuration `
    -Name 'ActiveDirectory'
$activeDirectoryEnabled = [bool](
    Get-SmartWorkplaceCMDBSetting $adConfiguration 'Enabled' $false
)
$forestWide = [bool](
    Get-SmartWorkplaceCMDBSetting $adConfiguration 'ForestWide' $true
)
if ([string]::IsNullOrWhiteSpace($Server)) {
    $Server = [string](Get-SmartWorkplaceCMDBSetting $adConfiguration 'Server' '')
}
if ([string]::IsNullOrWhiteSpace($SearchBase)) {
    $SearchBase = [string](Get-SmartWorkplaceCMDBSetting $adConfiguration 'SearchBase' '')
}
if (-not $PSBoundParameters.ContainsKey('IncludeGroupMemberships')) {
    $IncludeGroupMemberships = [bool](
        Get-SmartWorkplaceCMDBSetting $adConfiguration 'IncludeGroupMemberships' $true
    )
}
if ($PSCmdlet.ParameterSetName -eq 'Live' -and $forestWide -and
    -not [string]::IsNullOrWhiteSpace($SearchBase)) {
    throw 'ActiveDirectory.SearchBase requires ForestWide=false. Leave SearchBase empty to collect every forest domain.'
}

$rawContract = Get-SmartWorkplaceCMDBTableContract -Path $rawContractPath
$tableNames = @(
    'ActiveDirectory_Domains.csv',
    'ActiveDirectory_Users.csv',
    'ActiveDirectory_Groups.csv',
    'ActiveDirectory_Computers.csv',
    'ActiveDirectory_GroupMemberships.csv'
)
$tables = @{}
foreach ($name in $tableNames) {
    $match = @($rawContract.tables | Where-Object name -eq $name)
    if ($match.Count -ne 1) {
        throw "The raw contract must contain exactly one '$name' definition."
    }
    $tables[$name] = $match[0]
}

$sourcePaths = @($tableNames | ForEach-Object { Join-Path $paths.LatestOutputRootPath (Join-Path $tables[$_].area $_) })
$sourceRun = Start-SmartWorkplaceCMDBSourceCollection -Paths $paths -RawPath $sourcePaths -Fixture:($PSCmdlet.ParameterSetName -eq 'Fixture') -MaxItems $MaxItems -Scoped:([bool]$SearchBase -or -not $forestWide) -NoWrite:$ValidateOnly
try {
$fixture = $null
$readiness = $null
if ($PSCmdlet.ParameterSetName -eq 'Fixture') {
    $InputJsonPath = [IO.Path]::GetFullPath($InputJsonPath)
    $fixture = Read-SmartWorkplaceCMDBActiveDirectoryFixture -Path $InputJsonPath
}
else {
    if (-not $activeDirectoryEnabled) {
        throw 'ActiveDirectory.Enabled must be true in the tenant-local configuration for live collection.'
    }
    $readiness = Get-SmartWorkplaceCMDBActiveDirectoryReadiness `
        -PreferredServer $Server `
        -ForestWide $forestWide
    $Server = $readiness.Server
}

$validationDomains = if ($null -ne $fixture) {
    $fixtureDomains = @(Get-SmartWorkplaceCMDBObjectValue $fixture 'domains')
    if ($fixtureDomains.Count -gt 0) {
        @($fixtureDomains)
    }
    else {
        @(Get-SmartWorkplaceCMDBObjectValue $fixture 'domain')
    }
}
else {
    @($readiness.Domains | ForEach-Object { $_.Domain })
}

if ($ValidateOnly) {
    [pscustomobject]@{
        Status                    = 'Valid'
        ScriptVersion             = $ScriptVersion
        SourceMode                = if ($null -ne $fixture) { 'OfflineJson' } else { 'ActiveDirectoryModule' }
        ActiveDirectoryModule     = if ($null -ne $readiness) { $readiness.ModuleName } else { '' }
        ActiveDirectoryVersion    = if ($null -ne $readiness) { $readiness.ModuleVersion } else { '' }
        Server                    = $Server
        SearchBase                = $SearchBase
        ForestWide               = $forestWide
        DomainCount              = @($validationDomains).Count
        Domains                  = (@($validationDomains | ForEach-Object { [string]$_.DNSRoot }) -join ';')
        IncludeGroupMemberships   = [bool]$IncludeGroupMemberships
        RawContractVersion        = [string]$rawContract.contractVersion
        LatestOutputRootPath      = $paths.LatestOutputRootPath
        ProfileKey                = $paths.ProfileKey
        TenantKey                 = $paths.TenantKey
    } | Format-List
    return
}

$source = if ($null -ne $fixture) {
    $fixture
}
else {
    Get-SmartWorkplaceCMDBActiveDirectoryLiveData `
        -Readiness $readiness `
        -PreferredSearchBase $SearchBase `
        -CollectMemberships ([bool]$IncludeGroupMemberships) `
        -Limit $MaxItems
}

$sourceDomains = @(Get-SmartWorkplaceCMDBObjectValue $source 'domains')
if ($sourceDomains.Count -eq 0) {
    $legacyDomain = Get-SmartWorkplaceCMDBObjectValue $source 'domain'
    if ($null -ne $legacyDomain) {
        $sourceDomains = @($legacyDomain)
    }
}
if ($sourceDomains.Count -eq 0) {
    throw 'Active Directory collection returned no domain metadata.'
}
$defaultDomain = $sourceDomains[0]
$defaultDomainDnsRoot = ConvertTo-SmartWorkplaceCMDBCleanText (
    Get-SmartWorkplaceCMDBObjectValue $defaultDomain 'DNSRoot'
)
$defaultDomainNetBiosName = ConvertTo-SmartWorkplaceCMDBCleanText (
    Get-SmartWorkplaceCMDBObjectValue $defaultDomain 'NetBIOSName'
)
$collectedDateTime = [datetime]::UtcNow.ToString('o')

$domainRows = @($sourceDomains | ForEach-Object {
    [pscustomobject][ordered]@{
        SourceSystem            = 'ActiveDirectory'
        DomainDnsRoot           = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'DNSRoot'
        )
        DomainNetBIOSName       = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'NetBIOSName'
        )
        ForestName              = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'Forest'
        )
        DistinguishedName       = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'DistinguishedName'
        )
        DomainMode              = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'DomainMode'
        )
        PDCEmulator             = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'PDCEmulator'
        )
        RIDMaster               = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'RIDMaster'
        )
        InfrastructureMaster    = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'InfrastructureMaster'
        )
        SourceCollectedDateTime = $collectedDateTime
    }
})
$duplicateDomains = @($domainRows | Group-Object DomainDnsRoot | Where-Object Count -gt 1)
if ($duplicateDomains.Count -gt 0 -or
    @($domainRows | Where-Object { [string]::IsNullOrWhiteSpace($_.DomainDnsRoot) }).Count -gt 0) {
    throw 'Active Directory domain metadata contains an empty or duplicate DNS root.'
}
$users = @(Get-SmartWorkplaceCMDBObjectValue $source 'users')
if ($MaxItems -gt 0) {
    $users = @($users | Select-Object -First $MaxItems)
}
$userRows = @($users | ForEach-Object {
    $rowDomainDnsRoot = Get-SmartWorkplaceCMDBActiveDirectoryDomainValue `
        $_ 'DomainDnsRoot' $defaultDomainDnsRoot
    $rowDomainNetBiosName = Get-SmartWorkplaceCMDBActiveDirectoryDomainValue `
        $_ 'DomainNetBIOSName' $defaultDomainNetBiosName
    $objectGuid = ConvertTo-SmartWorkplaceCMDBGuidText (
        Get-SmartWorkplaceCMDBObjectValue $_ 'ObjectGUID'
    )
    if ([string]::IsNullOrWhiteSpace($objectGuid)) {
        throw 'An Active Directory user does not contain ObjectGUID.'
    }
    $distinguishedName = ConvertTo-SmartWorkplaceCMDBCleanText (
        Get-SmartWorkplaceCMDBObjectValue $_ 'DistinguishedName'
    )
    [pscustomobject][ordered]@{
        SourceSystem             = 'ActiveDirectory'
        DomainDnsRoot            = $rowDomainDnsRoot
        DomainNetBIOSName        = $rowDomainNetBiosName
        SourceObjectGuid         = $objectGuid
        ObjectSid                = ConvertTo-SmartWorkplaceCMDBSidText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'ObjectSID'
        )
        SamAccountName           = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'SamAccountName'
        )
        UserPrincipalName        = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'UserPrincipalName'
        )
        DisplayName              = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'DisplayName'
        )
        Enabled                  = [string](Get-SmartWorkplaceCMDBObjectValue $_ 'Enabled')
        Department               = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'Department'
        )
        JobTitle                 = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'Title'
        )
        Mail                     = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'Mail'
        )
        EmployeeId               = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'EmployeeID'
        )
        DistinguishedName        = $distinguishedName
        OrganizationalUnit       = Get-SmartWorkplaceCMDBOrganizationalUnit $distinguishedName
        ManagerDistinguishedName = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'Manager'
        )
        WhenCreated              = ConvertTo-SmartWorkplaceCMDBDateText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'WhenCreated'
        )
        WhenChanged              = ConvertTo-SmartWorkplaceCMDBDateText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'WhenChanged'
        )
        LastLogonDate            = ConvertTo-SmartWorkplaceCMDBDateText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'LastLogonDate'
        )
        PasswordLastSet          = ConvertTo-SmartWorkplaceCMDBDateText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'PasswordLastSet'
        )
        SourceCollectedDateTime  = $collectedDateTime
    }
})

$groups = @(Get-SmartWorkplaceCMDBObjectValue $source 'groups')
if ($MaxItems -gt 0) {
    $groups = @($groups | Select-Object -First $MaxItems)
}
$groupRows = @($groups | ForEach-Object {
    $rowDomainDnsRoot = Get-SmartWorkplaceCMDBActiveDirectoryDomainValue `
        $_ 'DomainDnsRoot' $defaultDomainDnsRoot
    $rowDomainNetBiosName = Get-SmartWorkplaceCMDBActiveDirectoryDomainValue `
        $_ 'DomainNetBIOSName' $defaultDomainNetBiosName
    $objectGuid = ConvertTo-SmartWorkplaceCMDBGuidText (
        Get-SmartWorkplaceCMDBObjectValue $_ 'ObjectGUID'
    )
    if ([string]::IsNullOrWhiteSpace($objectGuid)) {
        throw 'An Active Directory group does not contain ObjectGUID.'
    }
    $distinguishedName = ConvertTo-SmartWorkplaceCMDBCleanText (
        Get-SmartWorkplaceCMDBObjectValue $_ 'DistinguishedName'
    )
    [pscustomobject][ordered]@{
        SourceSystem            = 'ActiveDirectory'
        DomainDnsRoot           = $rowDomainDnsRoot
        DomainNetBIOSName       = $rowDomainNetBiosName
        SourceObjectGuid        = $objectGuid
        ObjectSid               = ConvertTo-SmartWorkplaceCMDBSidText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'ObjectSID'
        )
        SamAccountName          = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'SamAccountName'
        )
        DisplayName             = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'DisplayName'
        )
        GroupCategory           = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'GroupCategory'
        )
        GroupScope              = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'GroupScope'
        )
        Mail                    = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'Mail'
        )
        DistinguishedName       = $distinguishedName
        OrganizationalUnit      = Get-SmartWorkplaceCMDBOrganizationalUnit $distinguishedName
        WhenCreated             = ConvertTo-SmartWorkplaceCMDBDateText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'WhenCreated'
        )
        WhenChanged             = ConvertTo-SmartWorkplaceCMDBDateText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'WhenChanged'
        )
        SourceCollectedDateTime = $collectedDateTime
    }
})

$computers = @(Get-SmartWorkplaceCMDBObjectValue $source 'computers')
if ($MaxItems -gt 0) {
    $computers = @($computers | Select-Object -First $MaxItems)
}
$computerRows = @($computers | ForEach-Object {
    $rowDomainDnsRoot = Get-SmartWorkplaceCMDBActiveDirectoryDomainValue `
        $_ 'DomainDnsRoot' $defaultDomainDnsRoot
    $rowDomainNetBiosName = Get-SmartWorkplaceCMDBActiveDirectoryDomainValue `
        $_ 'DomainNetBIOSName' $defaultDomainNetBiosName
    $objectGuid = ConvertTo-SmartWorkplaceCMDBGuidText (
        Get-SmartWorkplaceCMDBObjectValue $_ 'ObjectGUID'
    )
    if ([string]::IsNullOrWhiteSpace($objectGuid)) {
        throw 'An Active Directory computer does not contain ObjectGUID.'
    }
    $distinguishedName = ConvertTo-SmartWorkplaceCMDBCleanText (
        Get-SmartWorkplaceCMDBObjectValue $_ 'DistinguishedName'
    )
    [pscustomobject][ordered]@{
        SourceSystem            = 'ActiveDirectory'
        DomainDnsRoot           = $rowDomainDnsRoot
        DomainNetBIOSName       = $rowDomainNetBiosName
        SourceObjectGuid        = $objectGuid
        ObjectSid               = ConvertTo-SmartWorkplaceCMDBSidText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'ObjectSID'
        )
        SamAccountName          = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'SamAccountName'
        )
        DeviceName              = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'Name'
        )
        DNSHostName             = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'DNSHostName'
        )
        Enabled                 = [string](Get-SmartWorkplaceCMDBObjectValue $_ 'Enabled')
        OperatingSystem         = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'OperatingSystem'
        )
        OperatingSystemVersion  = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'OperatingSystemVersion'
        )
        IPv4Address             = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'IPv4Address'
        )
        DistinguishedName       = $distinguishedName
        OrganizationalUnit      = Get-SmartWorkplaceCMDBOrganizationalUnit $distinguishedName
        ManagedByDistinguishedName = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'ManagedBy'
        )
        WhenCreated             = ConvertTo-SmartWorkplaceCMDBDateText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'WhenCreated'
        )
        WhenChanged             = ConvertTo-SmartWorkplaceCMDBDateText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'WhenChanged'
        )
        LastLogonDate           = ConvertTo-SmartWorkplaceCMDBDateText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'LastLogonDate'
        )
        SourceCollectedDateTime = $collectedDateTime
    }
})

$memberships = if ([bool]$IncludeGroupMemberships) {
    @(Get-SmartWorkplaceCMDBObjectValue $source 'groupMemberships')
}
else {
    @()
}
if ($MaxItems -gt 0) {
    $memberships = @($memberships | Select-Object -First $MaxItems)
}
$membershipRows = @($memberships | ForEach-Object {
    $rowDomainDnsRoot = Get-SmartWorkplaceCMDBActiveDirectoryDomainValue `
        $_ 'DomainDnsRoot' $defaultDomainDnsRoot
    $groupGuid = ConvertTo-SmartWorkplaceCMDBGuidText (
        Get-SmartWorkplaceCMDBObjectValue $_ 'GroupObjectGuid'
    )
    $memberGuid = ConvertTo-SmartWorkplaceCMDBGuidText (
        Get-SmartWorkplaceCMDBObjectValue $_ 'MemberObjectGuid'
    )
    if ([string]::IsNullOrWhiteSpace($groupGuid) -or
        [string]::IsNullOrWhiteSpace($memberGuid)) {
        throw 'An Active Directory group membership does not contain both object GUIDs.'
    }
    [pscustomobject][ordered]@{
        SourceSystem              = 'ActiveDirectory'
        DomainDnsRoot             = $rowDomainDnsRoot
        RelationshipKey           = '{0}|{1}' -f $groupGuid.ToLowerInvariant(), $memberGuid.ToLowerInvariant()
        GroupObjectGuid           = $groupGuid
        GroupDistinguishedName    = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'GroupDistinguishedName'
        )
        MemberObjectGuid          = $memberGuid
        MemberObjectSid           = ConvertTo-SmartWorkplaceCMDBSidText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'MemberObjectSid'
        )
        MemberDistinguishedName   = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'MemberDistinguishedName'
        )
        MemberObjectClass         = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'MemberObjectClass'
        )
        SourceCollectedDateTime   = $collectedDateTime
    }
})

foreach ($set in @(
        @{ Name = 'ActiveDirectory_Users.csv'; Rows = $userRows; Key = 'SourceObjectGuid' },
        @{ Name = 'ActiveDirectory_Groups.csv'; Rows = $groupRows; Key = 'SourceObjectGuid' },
        @{ Name = 'ActiveDirectory_Computers.csv'; Rows = $computerRows; Key = 'SourceObjectGuid' },
        @{ Name = 'ActiveDirectory_GroupMemberships.csv'; Rows = $membershipRows; Key = 'RelationshipKey' }
    )) {
    $duplicates = @($set.Rows | Group-Object -Property $set.Key | Where-Object Count -gt 1)
    if ($duplicates.Count -gt 0) {
        throw "Duplicate values detected in '$($set.Name)' for '$($set.Key)': $($duplicates.Name -join ', ')"
    }
}

$sets = @(
    @{ Name = 'ActiveDirectory_Domains.csv'; Rows = $domainRows; HistoryArea = 'Domains' },
    @{ Name = 'ActiveDirectory_Users.csv'; Rows = $userRows; HistoryArea = 'Users' },
    @{ Name = 'ActiveDirectory_Groups.csv'; Rows = $groupRows; HistoryArea = 'Groups' },
    @{ Name = 'ActiveDirectory_Computers.csv'; Rows = $computerRows; HistoryArea = 'Computers' },
    @{ Name = 'ActiveDirectory_GroupMemberships.csv'; Rows = $membershipRows; HistoryArea = 'GroupMemberships' }
)
$timestamp = [datetime]::UtcNow
$outputPaths = [ordered]@{}
foreach ($set in $sets) {
    $table = $tables[$set.Name]
    $latestPath = Join-Path $paths.LatestOutputRootPath (
        Join-Path ([string]$table.area) ([string]$table.name)
    )
    $historyFolder = Join-Path $paths.DataAllRootPath (
        'ActiveDirectory\{0}\{1}\{2}' -f
        $set.HistoryArea, $timestamp.ToString('yyyy'), $timestamp.ToString('MM')
    )
    $historyName = '{0}_{1}.csv' -f
        [IO.Path]::GetFileNameWithoutExtension($set.Name),
        $timestamp.ToString('yyyyMMdd-HHmmssfff')
    $historyPath = Join-Path $historyFolder $historyName
    $export = @{
        InputObject     = @($set.Rows)
        Columns         = @($table.columns | ForEach-Object { [string]$_ })
        TenantKey       = $paths.TenantKey
        OrganizationKey = $paths.OrganizationKey
        EnvironmentKey  = $paths.EnvironmentKey
        TenantId        = $paths.TenantId
    }
    Export-SmartWorkplaceCMDBCsv @export -Path $historyPath
    Export-SmartWorkplaceCMDBCsv @export -Path $latestPath
    $outputPaths[$set.Name] = $latestPath
}

$contractResults = @(Test-SmartWorkplaceCMDBCsvContract `
        -LatestOutputRootPath $paths.LatestOutputRootPath `
        -ContractPath $rawContractPath)
$adResults = @($contractResults | Where-Object Name -in $tableNames)
if ($adResults.Count -ne $tableNames.Count -or
    @($adResults | Where-Object Status -ne 'Valid').Count -gt 0) {
    throw 'An Active Directory raw CSV does not satisfy the SmartWorkplaceCMDB raw contract.'
}

$notCollected = if (-not $IncludeGroupMemberships) { @($sourcePaths | Where-Object { [IO.Path]::GetFileName($_) -eq 'ActiveDirectory_GroupMemberships.csv' }) } else { @() }
Complete-SmartWorkplaceCMDBSourceCollection -Run $sourceRun -NotCollectedPath $notCollected

Write-Information (
    "SmartWorkplaceCMDB Active Directory collection completed. Domains={0}; Users={1}; Groups={2}; Computers={3}; Memberships={4}." -f
    $domainRows.Count, $userRows.Count, $groupRows.Count, $computerRows.Count, $membershipRows.Count
) -InformationAction Continue

[pscustomobject]@{
    Status                    = 'Completed'
    ScriptVersion             = $ScriptVersion
    SourceMode                = if ($null -ne $fixture) { 'OfflineJson' } else { 'ActiveDirectoryModule' }
    DomainCount               = $domainRows.Count
    UserCount                 = $userRows.Count
    GroupCount                = $groupRows.Count
    ComputerCount             = $computerRows.Count
    GroupMembershipCount      = $membershipRows.Count
    IncludeGroupMemberships   = [bool]$IncludeGroupMemberships
    RawContractVersion        = [string]$rawContract.contractVersion
    LatestOutputRootPath      = $paths.LatestOutputRootPath
}

} catch {
    Complete-SmartWorkplaceCMDBSourceCollection -Run $sourceRun -Failed
    throw
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDE5gMpUHZPlijg
# /zPPmYpFX5q+8Pme+23iTfkVR5KJIqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEYV8mzgrD4cZwevWve/2DYcAWN7kjTcucTX403RtFvgMA0GCSqG
# SIb3DQEBAQUABIIBgH2k5LQyH2DMb2j6QhBQYD+JjNwmSUpPTr85XDWdih0Jo+Sp
# iEzCZWMuzPsJSDjYsAHFVBIdMQF973SelZm3CkSBbCr40iebl03umU9RubdgIL9/
# /El5t/Sz6JqqC4wIQHthIO9soGy/xhpdRma2S3xFE7+2HqDYmabZNklwYJifP5g5
# 1q2Ry171niyY2sNbOmKZDi/TaI0g3CUeFByGuVjqUFWa2vkMGlUhZvn0pxBq03Lr
# 2xB+ZTSsLi8XQ+I2XBFMC+viv8ape1K8d38HtUCxHFfIq/c3TEOFjaSaHzLQENv4
# dOJgDJTapbTD9lnGRvkNBNwR4p5CL0ApAra+oiTlnT3ffq6zwU5ofFzjx3J5+gUd
# el8vo6sEq4L3wBtOxkD9i+AIX8DUkxXvhJjbBRx/3XS7i1t9VeVYoPW+PfccY2Ol
# dkYK+gBXa/ZQGLUjNBCidyHOcXKJUEePIyqgLKtBEnFKr8RQvOzMaqUvOT3Luiax
# LYbVsV85su1hGhD5BKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxMDM5
# MTJaMC8GCSqGSIb3DQEJBDEiBCACIYyQBjFRXc6Rxaqp22itHqgHnniKNsxgL96L
# dPtbjDANBgkqhkiG9w0BAQEFAASCAgAJsidAMwdkjzU7EcOnwiLhhJn3xbBcsfqE
# K4S5UVQaSq2P4hFMBr2OBt2/5I3bgQut1yjbHSLj6Mm3ZQXkgLWT/g2vqy15VnQE
# 8xarV7hxas9B06IUHyYBgaprEgt1vu4xIjCCPouIWSgo4cjdMYcbFHbuj4rn1Qp6
# H502P/uGUu2IqE5PZIPIEPr4TGkiUyJXy7mazmXKOkmZOk4Y3AXN8e1xMS22guiL
# u4ZpSoXdLp7c7boNJlgvsAicSvw4RqtT4GksBOlPL9KnkB+DVrGuLl6Wvwj0uVcF
# hfIJzMd3X/nVHxQBYECwraQpItixGbopT9/8Yv9cjQMolDANgikZinMlRox/lqUt
# nxjriEt3wJRJDYKI44qlfckYojOlQN40vgl+Po4iv91bKK32/t2yFZ5FfrZMH2F8
# nFUdF/gh36gwGUxgBSx1VsNV8vWsJSBAwHXvl19CUsZNXOHZRoRes/D+vXyKueC0
# f/OEeeThUThSVeFzF1IxRG8CrAUnERK8BTQwy/ciUKIjseaoKRopeo6mFAiLJWxt
# 1FsHzW36dvuGGOQ50cWv71dVVzVurG9gfcS1xt4TmA6UFPiU5JMgjEaLincixPan
# SoA4xtkSg953kkUehcW8JDxl0zN0WqDQ2/Es9KTCDCcdftguHXl9mTIvhhWdds5l
# aJwgZYtISA==
# SIG # End signature block
