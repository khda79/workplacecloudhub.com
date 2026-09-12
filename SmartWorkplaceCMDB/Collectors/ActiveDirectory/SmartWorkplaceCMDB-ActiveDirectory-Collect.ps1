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
1.0.2

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
    [string[]]$TargetDomains = @(),
    [int]$DomainRetryCount = -1,
    [int[]]$DomainRetryDelaysSeconds = @(),
    [switch]$IncludeGroupMemberships,
    [Parameter(ParameterSetName = 'Fixture', Mandatory)]
    [string]$InputJsonPath,
    [ValidateRange(0, 2147483647)]
    [int]$MaxItems = 0,
    [switch]$NoConfigWrite,
    [switch]$ValidateOnly
)

$ScriptVersion = '1.0.2'
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
        [bool]$ForestWide,
        [string[]]$TargetDomains = @()
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
    $forestDomainNames = if ($ForestWide) {
        @($forest.Domains | Sort-Object -Unique)
    }
    else {
        @([string]$bootstrapDomain.DNSRoot)
    }
    if ($forestDomainNames.Count -eq 0) {
        throw 'Active Directory forest discovery returned no domains.'
    }

    $normalizedTargets = @($TargetDomains | ForEach-Object {
            ([string]$_).Trim().ToLowerInvariant()
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique)
    if ($normalizedTargets.Count -gt 0 -and -not $ForestWide) {
        throw 'ActiveDirectory.TargetDomains requires ForestWide=true.'
    }
    $knownDomains = @($forestDomainNames | ForEach-Object {
            ([string]$_).Trim().ToLowerInvariant()
        })
    $unknownTargets = @($normalizedTargets | Where-Object { $_ -notin $knownDomains })
    if ($unknownTargets.Count -gt 0) {
        throw "ActiveDirectory.TargetDomains contains domains outside the discovered forest: $($unknownTargets -join ', ')."
    }
    $domainNames = if ($normalizedTargets.Count -gt 0) {
        @($forestDomainNames | Where-Object {
                ([string]$_).Trim().ToLowerInvariant() -in $normalizedTargets
            })
    }
    else {
        @($forestDomainNames)
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

function Test-SmartWorkplaceCMDBTransientActiveDirectoryError {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ErrorRecord)

    $exception = $ErrorRecord.Exception
    while ($null -ne $exception) {
        $typeName = $exception.GetType().FullName
        $message = [string]$exception.Message
        if ($typeName -eq 'System.DirectoryServices.Protocols.LdapException') {
            if ($message -imatch 'access is denied|insufficient access|unauthorized|authentication|invalid credentials') {
                return $false
            }
            $ldapErrorCode = [int]$exception.ErrorCode
            if ($ldapErrorCode -in @(51, 52, 81, 85, 91) -or
                $message -imatch (
                    'server.*(unavailable|down|not operational)|' +
                    'LDAP.*unavailable|operation.*timed out|timeout|' +
                    'invalid enumeration context|server busy|connection.*(closed|reset)'
                )) {
                return $true
            }
            return $false
        }
        if ($typeName -in @(
                'Microsoft.ActiveDirectory.Management.ADServerDownException',
                'System.TimeoutException',
                'System.Net.Sockets.SocketException',
                'System.IO.IOException'
            ) -or $message -imatch (
                'server.*(unavailable|down|not operational)|' +
                'LDAP.*unavailable|operation.*timed out|timeout|' +
                'invalid enumeration context|server busy|connection.*(closed|reset)'
            )) {
            if ($message -imatch 'access is denied|insufficient access|unauthorized|authentication') {
                return $false
            }
            return $true
        }
        if ($exception -is [System.UnauthorizedAccessException] -or
            $message -imatch 'access is denied|insufficient access|unauthorized|authentication') {
            return $false
        }
        $exception = $exception.InnerException
    }
    return $false
}

function Get-SmartWorkplaceCMDBActiveDirectoryRetryServer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DomainDnsRoot,
        [Parameter(Mandatory)][string]$CurrentServer
    )

    $fallback = ''
    for ($discoveryAttempt = 0; $discoveryAttempt -lt 3; $discoveryAttempt++) {
        $controller = Get-ADDomainController -Discover `
            -DomainName $DomainDnsRoot `
            -Service ADWS `
            -ForceDiscover `
            -ErrorAction Stop
        $candidate = [string]$controller.HostName
        if (-not [string]::IsNullOrWhiteSpace($candidate)) {
            $fallback = $candidate
            if ($candidate -ine $CurrentServer) {
                return $candidate
            }
        }
    }
    if ([string]::IsNullOrWhiteSpace($fallback)) {
        throw "No replacement Active Directory Web Services domain controller was discovered for '$DomainDnsRoot'."
    }
    Write-Warning (
        "ADWS rediscovery for '$DomainDnsRoot' returned the current server '$CurrentServer'; the retry will revalidate it."
    )
    return $fallback
}

function Invoke-SmartWorkplaceCMDBActiveDirectoryDomainOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$DomainContext,
        [Parameter(Mandatory)][string]$OperationName,
        [Parameter(Mandatory)][scriptblock]$Action,
        [ValidateRange(0, 20)][int]$RetryCount = 3,
        [int[]]$RetryDelaysSeconds = @(5, 15, 30),
        [scriptblock]$SleepAction,
        [scriptblock]$ServerSelector
    )

    if ($null -eq $SleepAction) {
        $SleepAction = { param([int]$Seconds) Start-Sleep -Seconds $Seconds }
    }
    if ($null -eq $ServerSelector) {
        $ServerSelector = {
            param([string]$DomainDnsRoot, [string]$CurrentServer)
            Get-SmartWorkplaceCMDBActiveDirectoryRetryServer `
                -DomainDnsRoot $DomainDnsRoot `
                -CurrentServer $CurrentServer
        }
    }
    if ($RetryCount -gt 0 -and $RetryDelaysSeconds.Count -eq 0) {
        throw 'ActiveDirectory.DomainRetryDelaysSeconds must contain at least one delay when retries are enabled.'
    }
    if (@($RetryDelaysSeconds | Where-Object { $_ -lt 0 }).Count -gt 0) {
        throw 'ActiveDirectory.DomainRetryDelaysSeconds cannot contain a negative delay.'
    }

    $domainDnsRoot = [string]$DomainContext.Domain.DNSRoot
    $server = [string]$DomainContext.Server
    for ($attempt = 0; $attempt -le $RetryCount; $attempt++) {
        try {
            $value = & $Action $server
            return [pscustomobject]@{
                Value = $value
                Server = $server
                RetryCount = $attempt
            }
        }
        catch {
            if ($attempt -ge $RetryCount -or
                -not (Test-SmartWorkplaceCMDBTransientActiveDirectoryError $_)) {
                throw
            }
            $delayIndex = [Math]::Min($attempt, $RetryDelaysSeconds.Count - 1)
            $delay = [int]$RetryDelaysSeconds[$delayIndex]
            $replacement = & $ServerSelector $domainDnsRoot $server
            if (-not [string]::IsNullOrWhiteSpace([string]$replacement)) {
                $server = [string]$replacement
            }
            Write-Warning (
                "Transient Active Directory failure during {0} for '{1}'. Retry {2}/{3} on '{4}' in {5}s: {6}" -f
                $OperationName, $domainDnsRoot, ($attempt + 1), $RetryCount,
                $server, $delay, $_.Exception.Message
            )
            & $SleepAction $delay
        }
    }
}

function Invoke-SmartWorkplaceCMDBActiveDirectoryDomainCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$DomainContext,
        [string]$PreferredSearchBase,
        [int]$Limit,
        [ValidateRange(0, 20)][int]$RetryCount = 3,
        [int[]]$RetryDelaysSeconds = @(5, 15, 30),
        [scriptblock]$SleepAction,
        [scriptblock]$ServerSelector
    )

    $domain = $DomainContext.Domain
    $domainDnsRoot = [string]$domain.DNSRoot
    $domainNetBIOSName = [string]$domain.NetBIOSName
    $action = {
        param([string]$SelectedServer)
        $common = @{
            Server = $SelectedServer
            ErrorAction = 'Stop'
            ResultPageSize = 500
        }
        if (-not [string]::IsNullOrWhiteSpace($PreferredSearchBase)) {
            $common['SearchBase'] = $PreferredSearchBase
        }
        if ($Limit -gt 0) {
            $common['ResultSetSize'] = $Limit
        }
        [pscustomobject]@{
            Users = @(Get-ADUser -Filter * @common -Properties @(
                    'ObjectGUID', 'ObjectSID', 'SamAccountName', 'UserPrincipalName',
                    'DisplayName', 'Enabled', 'Department', 'Title', 'Mail',
                    'EmployeeID', 'Country', 'Company', 'Office',
                    'AccountExpirationDate', 'UserAccountControl', 'DistinguishedName',
                    'Manager', 'WhenCreated', 'WhenChanged', 'LastLogonDate',
                    'PasswordLastSet', 'PrimaryGroupID'
                ))
            Groups = @(Get-ADGroup -Filter * @common -Properties @(
                    'ObjectGUID', 'ObjectSID', 'SamAccountName', 'DisplayName',
                    'GroupCategory', 'GroupScope', 'Mail', 'Description',
                    'ManagedBy', 'ProtectedFromAccidentalDeletion',
                    'DistinguishedName', 'WhenCreated', 'WhenChanged'
                ))
            Computers = @(Get-ADComputer -Filter * @common -Properties @(
                    'ObjectGUID', 'ObjectSID', 'SamAccountName', 'Name',
                    'DNSHostName', 'Enabled', 'OperatingSystem',
                    'OperatingSystemVersion', 'IPv4Address', 'CanonicalName',
                    'DistinguishedName', 'ManagedBy', 'WhenCreated', 'WhenChanged',
                    'LastLogonDate', 'LastLogonTimestamp', 'PasswordLastSet',
                    'PrimaryGroupID'
                ))
            OrganizationalUnits = @(Get-ADOrganizationalUnit -Filter * @common -Properties @(
                    'ObjectGUID', 'Name', 'DistinguishedName', 'Description',
                    'ManagedBy', 'ProtectedFromAccidentalDeletion',
                    'WhenCreated', 'WhenChanged'
                ))
        }
    }.GetNewClosure()
    $operation = Invoke-SmartWorkplaceCMDBActiveDirectoryDomainOperation `
        -DomainContext $DomainContext `
        -OperationName 'object inventory' `
        -Action $action `
        -RetryCount $RetryCount `
        -RetryDelaysSeconds $RetryDelaysSeconds `
        -SleepAction $SleepAction `
        -ServerSelector $ServerSelector

    return [pscustomobject]@{
        Context = [pscustomobject]@{ Domain = $domain; Server = $operation.Server }
        DomainDnsRoot = $domainDnsRoot
        DomainNetBIOSName = $domainNetBIOSName
        Users = @($operation.Value.Users)
        Groups = @($operation.Value.Groups)
        Computers = @($operation.Value.Computers)
        OrganizationalUnits = @($operation.Value.OrganizationalUnits)
        RetryCount = [int]$operation.RetryCount
    }
}

function Get-SmartWorkplaceCMDBActiveDirectoryLiveData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Readiness,
        [string]$PreferredSearchBase,
        [bool]$CollectMemberships,
        [int]$Limit,
        [ValidateRange(0, 20)][int]$RetryCount = 3,
        [int[]]$RetryDelaysSeconds = @(5, 15, 30)
    )

    if (-not [string]::IsNullOrWhiteSpace($PreferredSearchBase) -and
        @($Readiness.Domains).Count -gt 1) {
        throw 'ActiveDirectory.SearchBase can only be used when ForestWide is false. Leave SearchBase empty for forest-wide collection.'
    }

    $users = New-Object System.Collections.Generic.List[object]
    $groups = New-Object System.Collections.Generic.List[object]
    $computers = New-Object System.Collections.Generic.List[object]
    $organizationalUnits = New-Object System.Collections.Generic.List[object]
    $memberships = New-Object System.Collections.Generic.List[object]
    $domainInventories = New-Object System.Collections.Generic.List[object]
    $totalRetryCount = 0

    $domainIndex = 0
    foreach ($domainContext in @($Readiness.Domains)) {
        $domainIndex++
        $domainDnsRoot = [string]$domainContext.Domain.DNSRoot
        Write-Information (
            "Active Directory domain [{0}/{1}] DNS='{2}' Server='{3}'." -f
            $domainIndex, @($Readiness.Domains).Count, $domainDnsRoot, $domainContext.Server
        ) -InformationAction Continue

        $inventory = Invoke-SmartWorkplaceCMDBActiveDirectoryDomainCollection `
            -DomainContext $domainContext `
            -PreferredSearchBase $PreferredSearchBase `
            -Limit $Limit `
            -RetryCount $RetryCount `
            -RetryDelaysSeconds $RetryDelaysSeconds
        $totalRetryCount += $inventory.RetryCount
        foreach ($item in $inventory.Users) {
            $users.Add((Add-SmartWorkplaceCMDBActiveDirectoryDomainContext `
                    $item $inventory.DomainDnsRoot $inventory.DomainNetBIOSName))
        }
        foreach ($item in $inventory.Groups) {
            $groups.Add((Add-SmartWorkplaceCMDBActiveDirectoryDomainContext `
                    $item $inventory.DomainDnsRoot $inventory.DomainNetBIOSName))
        }
        foreach ($item in $inventory.Computers) {
            $computers.Add((Add-SmartWorkplaceCMDBActiveDirectoryDomainContext `
                    $item $inventory.DomainDnsRoot $inventory.DomainNetBIOSName))
        }
        foreach ($item in $inventory.OrganizationalUnits) {
            $organizationalUnits.Add((Add-SmartWorkplaceCMDBActiveDirectoryDomainContext `
                    $item $inventory.DomainDnsRoot $inventory.DomainNetBIOSName))
        }
        $domainInventories.Add($inventory)
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
            $membershipAction = {
                param([string]$SelectedServer)
                $started = [datetimeoffset]::UtcNow
                $directMemberships = New-Object System.Collections.Generic.List[object]
                $domainConnection = New-SmartWorkplaceCMDBActiveDirectoryLdapConnection `
                    -Server $SelectedServer
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
                        $directMemberships.Add([pscustomobject]@{
                                Group = $group
                                MemberDistinguishedNames = @(
                                    Get-SmartWorkplaceCMDBActiveDirectoryRangedMember `
                                        -Server $SelectedServer `
                                        -GroupDistinguishedName $group.DistinguishedName `
                                        -Connection $domainConnection
                                )
                            })
                    }
                    return @($directMemberships.ToArray())
                }
                finally {
                    $domainConnection.Dispose()
                }
            }.GetNewClosure()
            $membershipOperation = Invoke-SmartWorkplaceCMDBActiveDirectoryDomainOperation `
                -DomainContext $domainContext `
                -OperationName 'group membership inventory' `
                -Action $membershipAction `
                -RetryCount $RetryCount `
                -RetryDelaysSeconds $RetryDelaysSeconds
            $totalRetryCount += $membershipOperation.RetryCount
            $domainContext.Server = $membershipOperation.Server

            foreach ($directMembership in @($membershipOperation.Value)) {
                    $group = $directMembership.Group
                    $memberDistinguishedNames = @($directMembership.MemberDistinguishedNames)
                    $groupMembers = New-Object System.Collections.Generic.List[object]
                    foreach ($memberDn in $memberDistinguishedNames) {
                        $member = $null
                        if ($principalByDistinguishedName.ContainsKey($memberDn)) {
                            $member = $principalByDistinguishedName[$memberDn]
                        }
                        else {
                            $resolveAction = {
                                param([string]$SelectedServer)
                                Get-ADObject -Identity $memberDn `
                                    -Server $SelectedServer `
                                    -ErrorAction Stop `
                                    -Properties @(
                                        'ObjectGUID', 'ObjectSID', 'ObjectClass',
                                        'DistinguishedName'
                                    )
                            }.GetNewClosure()
                            $resolveOperation = Invoke-SmartWorkplaceCMDBActiveDirectoryDomainOperation `
                                -DomainContext $domainContext `
                                -OperationName "member resolution '$memberDn'" `
                                -Action $resolveAction `
                                -RetryCount $RetryCount `
                                -RetryDelaysSeconds $RetryDelaysSeconds
                            $totalRetryCount += $resolveOperation.RetryCount
                            $domainContext.Server = $resolveOperation.Server
                            $resolved = $resolveOperation.Value
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
    }

    return [pscustomobject]@{
        domains          = @($Readiness.Domains | ForEach-Object { $_.Domain })
        users            = @($users.ToArray())
        groups           = @($groups.ToArray())
        computers        = @($computers.ToArray())
        organizationalUnits = @($organizationalUnits.ToArray())
        groupMemberships = @($memberships.ToArray())
        retryCount       = $totalRetryCount
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
if (-not $PSBoundParameters.ContainsKey('TargetDomains')) {
    $targetDomainSetting = if ($adConfiguration.Contains('TargetDomains')) {
        $adConfiguration['TargetDomains']
    }
    else {
        ''
    }
    $TargetDomains = @(if ($targetDomainSetting -is [string]) {
        [string]$targetDomainSetting -split '[,;\s]+' | Where-Object {
                -not [string]::IsNullOrWhiteSpace($_)
            }
    }
    else {
        $targetDomainSetting | ForEach-Object { [string]$_ }
    })
}
if ($DomainRetryCount -lt 0) {
    $DomainRetryCount = [int](
        Get-SmartWorkplaceCMDBSetting $adConfiguration 'DomainRetryCount' 3
    )
}
if (-not $PSBoundParameters.ContainsKey('DomainRetryDelaysSeconds')) {
    $retryDelaySetting = if ($adConfiguration.Contains('DomainRetryDelaysSeconds')) {
        $adConfiguration['DomainRetryDelaysSeconds']
    }
    else {
        '5,15,30'
    }
    $retryDelayValues = if ($retryDelaySetting -is [string]) {
        @([string]$retryDelaySetting -split '[,;\s]+' | Where-Object {
                -not [string]::IsNullOrWhiteSpace($_)
            })
    }
    else {
        @($retryDelaySetting)
    }
    $DomainRetryDelaysSeconds = @($retryDelayValues | ForEach-Object { [int]$_ })
}
$domainParallelThrottleLimit = [int](
    Get-SmartWorkplaceCMDBSetting $adConfiguration 'DomainParallelThrottleLimit' 1
)
if ($DomainRetryCount -lt 0 -or $DomainRetryCount -gt 20) {
    throw 'ActiveDirectory.DomainRetryCount must be between 0 and 20.'
}
if ($domainParallelThrottleLimit -ne 1) {
    throw 'ActiveDirectory.DomainParallelThrottleLimit must remain 1 in SmartWorkplaceCMDB 1.0.2.'
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
if ($PSCmdlet.ParameterSetName -eq 'Live' -and $TargetDomains.Count -gt 0 -and
    -not $forestWide) {
    throw 'ActiveDirectory.TargetDomains requires ForestWide=true.'
}

$rawContract = Get-SmartWorkplaceCMDBTableContract -Path $rawContractPath
$tableNames = @(
    'ActiveDirectory_Domains.csv',
    'ActiveDirectory_Users.csv',
    'ActiveDirectory_Groups.csv',
    'ActiveDirectory_Computers.csv',
    'ActiveDirectory_OrganizationalUnits.csv',
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
$sourceRun = Start-SmartWorkplaceCMDBSourceCollection -Paths $paths -RawPath $sourcePaths -Fixture:($PSCmdlet.ParameterSetName -eq 'Fixture') -MaxItems $MaxItems -Scoped:([bool]$SearchBase -or -not $forestWide -or $TargetDomains.Count -gt 0) -NoWrite:$ValidateOnly
$sourceRunCompleted = $false
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
        -ForestWide $forestWide `
        -TargetDomains $TargetDomains
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
        TargetDomains             = ($TargetDomains -join ';')
        DomainRetryCount          = $DomainRetryCount
        DomainRetryDelaysSeconds  = ($DomainRetryDelaysSeconds -join ';')
        DomainParallelThrottleLimit = $domainParallelThrottleLimit
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
        -Limit $MaxItems `
        -RetryCount $DomainRetryCount `
        -RetryDelaysSeconds $DomainRetryDelaysSeconds
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
        Country                  = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'Country'
        )
        Company                  = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'Company'
        )
        Office                   = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'Office'
        )
        AccountExpirationDate    = ConvertTo-SmartWorkplaceCMDBDateText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'AccountExpirationDate'
        )
        UserAccountControl       = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'UserAccountControl'
        )
        PrimaryGroupId           = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'PrimaryGroupID'
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
        Description             = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'Description'
        )
        ManagedByDistinguishedName = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'ManagedBy'
        )
        ProtectedFromAccidentalDeletion = [string](
            Get-SmartWorkplaceCMDBObjectValue $_ 'ProtectedFromAccidentalDeletion'
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
        CanonicalName           = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'CanonicalName'
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
        LastLogonTimestamp      = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'LastLogonTimestamp'
        )
        PasswordLastSet         = ConvertTo-SmartWorkplaceCMDBDateText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'PasswordLastSet'
        )
        PrimaryGroupId          = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'PrimaryGroupID'
        )
        SourceCollectedDateTime = $collectedDateTime
    }
})

$organizationalUnitSource = Get-SmartWorkplaceCMDBObjectValue $source 'organizationalUnits'
$organizationalUnits = if ($null -eq $organizationalUnitSource) {
    @()
}
else {
    @($organizationalUnitSource)
}
if ($MaxItems -gt 0) {
    $organizationalUnits = @($organizationalUnits | Select-Object -First $MaxItems)
}
$organizationalUnitRows = @($organizationalUnits | ForEach-Object {
    $rowDomainDnsRoot = Get-SmartWorkplaceCMDBActiveDirectoryDomainValue `
        $_ 'DomainDnsRoot' $defaultDomainDnsRoot
    $rowDomainNetBiosName = Get-SmartWorkplaceCMDBActiveDirectoryDomainValue `
        $_ 'DomainNetBIOSName' $defaultDomainNetBiosName
    $objectGuid = ConvertTo-SmartWorkplaceCMDBGuidText (
        Get-SmartWorkplaceCMDBObjectValue $_ 'ObjectGUID'
    )
    if ([string]::IsNullOrWhiteSpace($objectGuid)) {
        throw 'An Active Directory organizational unit does not contain ObjectGUID.'
    }
    $distinguishedName = ConvertTo-SmartWorkplaceCMDBCleanText (
        Get-SmartWorkplaceCMDBObjectValue $_ 'DistinguishedName'
    )
    [pscustomobject][ordered]@{
        SourceSystem = 'ActiveDirectory'
        DomainDnsRoot = $rowDomainDnsRoot
        DomainNetBIOSName = $rowDomainNetBiosName
        SourceObjectGuid = $objectGuid
        Name = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'Name'
        )
        DistinguishedName = $distinguishedName
        ParentDistinguishedName = Get-SmartWorkplaceCMDBOrganizationalUnit $distinguishedName
        Description = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'Description'
        )
        ManagedByDistinguishedName = ConvertTo-SmartWorkplaceCMDBCleanText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'ManagedBy'
        )
        ProtectedFromAccidentalDeletion = [string](
            Get-SmartWorkplaceCMDBObjectValue $_ 'ProtectedFromAccidentalDeletion'
        )
        WhenCreated = ConvertTo-SmartWorkplaceCMDBDateText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'WhenCreated'
        )
        WhenChanged = ConvertTo-SmartWorkplaceCMDBDateText (
            Get-SmartWorkplaceCMDBObjectValue $_ 'WhenChanged'
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
        @{ Name = 'ActiveDirectory_OrganizationalUnits.csv'; Rows = $organizationalUnitRows; Key = 'SourceObjectGuid' },
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
    @{ Name = 'ActiveDirectory_OrganizationalUnits.csv'; Rows = $organizationalUnitRows; HistoryArea = 'OrganizationalUnits' },
    @{ Name = 'ActiveDirectory_GroupMemberships.csv'; Rows = $membershipRows; HistoryArea = 'GroupMemberships' }
)
$timestamp = [datetime]::UtcNow
$outputPaths = [ordered]@{}
$transactionRoot = Join-Path $paths.DataRootPath (
    '.staging\ActiveDirectory\{0}' -f $sourceRun.RunId
)
$stagedLatestRoot = Join-Path $transactionRoot 'DATA-LAST'
$backupRoot = Join-Path $transactionRoot 'backups'
$promotions = New-Object System.Collections.Generic.List[object]
$completedPromotions = New-Object System.Collections.Generic.List[object]
$notCollected = if (-not $IncludeGroupMemberships) { @($sourcePaths | Where-Object { [IO.Path]::GetFileName($_) -eq 'ActiveDirectory_GroupMemberships.csv' }) } else { @() }
try {
    foreach ($set in $sets) {
        $table = $tables[$set.Name]
        $stagedPath = Join-Path $stagedLatestRoot (
            Join-Path ([string]$table.area) ([string]$table.name)
        )
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
        Export-SmartWorkplaceCMDBCsv @export -Path $stagedPath
        $promotions.Add([pscustomobject]@{
                Source = $stagedPath
                Destination = $historyPath
                Backup = Join-Path $backupRoot ('history-' + $historyName)
            })
        $promotions.Add([pscustomobject]@{
                Source = $stagedPath
                Destination = $latestPath
                Backup = Join-Path $backupRoot ('latest-' + $set.Name)
            })
        $outputPaths[$set.Name] = $latestPath
    }

    $stagedContractResults = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath $stagedLatestRoot `
            -ContractPath $rawContractPath)
    $stagedAdResults = @($stagedContractResults | Where-Object Name -in $tableNames)
    if ($stagedAdResults.Count -ne $tableNames.Count -or
        @($stagedAdResults | Where-Object Status -ne 'Valid').Count -gt 0) {
        throw 'A staged Active Directory raw CSV does not satisfy the SmartWorkplaceCMDB raw contract.'
    }

    foreach ($promotion in $promotions) {
        $destinationFolder = Split-Path $promotion.Destination -Parent
        New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
        $hadPrevious = Test-Path -LiteralPath $promotion.Destination -PathType Leaf
        if ($hadPrevious) {
            New-Item -ItemType Directory -Path (Split-Path $promotion.Backup -Parent) -Force | Out-Null
            Copy-Item -LiteralPath $promotion.Destination -Destination $promotion.Backup -Force
        }
        $candidate = $promotion.Destination + '.candidate.' + $sourceRun.RunId
        Copy-Item -LiteralPath $promotion.Source -Destination $candidate -Force
        Move-Item -LiteralPath $candidate -Destination $promotion.Destination -Force
        $completedPromotions.Add([pscustomobject]@{
                Destination = $promotion.Destination
                Backup = $promotion.Backup
                HadPrevious = $hadPrevious
            })
    }
    $contractResults = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath $paths.LatestOutputRootPath `
            -ContractPath $rawContractPath)
    $adResults = @($contractResults | Where-Object Name -in $tableNames)
    if ($adResults.Count -ne $tableNames.Count -or
        @($adResults | Where-Object Status -ne 'Valid').Count -gt 0) {
        throw 'An Active Directory raw CSV does not satisfy the SmartWorkplaceCMDB raw contract after promotion.'
    }
    Complete-SmartWorkplaceCMDBSourceCollection -Run $sourceRun -NotCollectedPath $notCollected
    $sourceRunCompleted = $true
}
catch {
    for ($promotionIndex = $completedPromotions.Count - 1;
        $promotionIndex -ge 0;
        $promotionIndex--) {
        $promotion = $completedPromotions[$promotionIndex]
        if ($promotion.HadPrevious -and
            (Test-Path -LiteralPath $promotion.Backup -PathType Leaf)) {
            Copy-Item -LiteralPath $promotion.Backup `
                -Destination $promotion.Destination -Force
        }
        elseif (Test-Path -LiteralPath $promotion.Destination -PathType Leaf) {
            Remove-Item -LiteralPath $promotion.Destination -Force
        }
    }
    throw
}
finally {
    if (Test-Path -LiteralPath $transactionRoot) {
        Remove-Item -LiteralPath $transactionRoot -Recurse -Force
    }
}

Write-Information (
    "SmartWorkplaceCMDB Active Directory collection completed. Domains={0}; Users={1}; Groups={2}; Computers={3}; OrganizationalUnits={4}; Memberships={5}." -f
    $domainRows.Count, $userRows.Count, $groupRows.Count, $computerRows.Count,
    $organizationalUnitRows.Count, $membershipRows.Count
) -InformationAction Continue

[pscustomobject]@{
    Status                    = 'Completed'
    ScriptVersion             = $ScriptVersion
    SourceMode                = if ($null -ne $fixture) { 'OfflineJson' } else { 'ActiveDirectoryModule' }
    DomainCount               = $domainRows.Count
    UserCount                 = $userRows.Count
    GroupCount                = $groupRows.Count
    ComputerCount             = $computerRows.Count
    OrganizationalUnitCount  = $organizationalUnitRows.Count
    GroupMembershipCount      = $membershipRows.Count
    RetryCount                = [int](Get-SmartWorkplaceCMDBObjectValue $source 'retryCount')
    IncludeGroupMemberships   = [bool]$IncludeGroupMemberships
    RawContractVersion        = [string]$rawContract.contractVersion
    LatestOutputRootPath      = $paths.LatestOutputRootPath
}

} catch {
    if (-not $sourceRunCompleted) {
        Complete-SmartWorkplaceCMDBSourceCollection -Run $sourceRun -Failed
    }
    throw
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCvgiXoWY8iVckS
# KTxlgiQwIpv/ai4vtak3y7pHybEZBKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICM6uBJvIBkZEkG3opCaDgT9yB4+uW/xW962Udb/SYYXMA0GCSqG
# SIb3DQEBAQUABIIBgAIdLeV9dpny8YMpOASP+S31IXr2UJLXyOBnZWUhoYa6wCVX
# NeEcx68A4h2SzdzvX51iiAxA7vOD2FdoCBoXYBJFyo9ZdEe17gwBIyRGoaptEeL7
# 46sUOiPdAaROSeXwZMvwjahQQWvWyslJ4Fv16qtULnPuG3Nk1YIfhm6yfu3uGwcx
# snc1aND9DrNTlSglCfKsqNmoTidl2tV4F1BGGSdZruRehLW1lyeN++2qAR13Cnjd
# WY6yImQ0RhEjoBKUlDVdDIxeemmNN8obEdV1PfTO1qinwibq9Ryqnae2jWzlBnkS
# arFWQWAjzNkgfHyg7iasBtfUdkUC9dzGYxpoq02YOIljRhZIdfTFTNs8DvxOI4Vf
# nc8W7mNDTxxE77bYv2GSJ0Qh/uSQWYAuLWDQzE6hmJrr2M99yICXTe6ibK5Ygcr8
# rlz74rUydGawT7AHe2gqCiyIH1mDOGxCtS7bRx1+vaqyE2PKN/1pews7s+31zZoc
# +LMbAx2l7D12+w9gUKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxMzE3
# MTNaMC8GCSqGSIb3DQEJBDEiBCDWZeCMszhhKs2bmAyL6TtlfYk68DyxAN25KUbi
# mCfp1jANBgkqhkiG9w0BAQEFAASCAgANth5fGzy9dvVzzhkyGqGvOyoxQUX4lM9Z
# 5wP0BV2Un9hHc+hny47xGfXtNNyj+CJS7Sr/1Srsc+uGQIDD+wZ+UWMe05dJz4f0
# 4LZveSJnCG1CMM+k2awZW0JvudgGzNE7l65D9ByimWAkoJbY9xI8tLhmRYHM9oRO
# HU7fxW8ZzrvJtqyWevCV/H+dg0ZRyZD98a3aNYYK+EHq/f4qRSlIT1kNg4QJXn/4
# 2ZgB2xf16+LH1rVMUsf3shCD6JWH2HlIabeM4HzOVIa6qqxF8VeZojez2ipstDQe
# uvW2sTmCvJgzijuWmM5BCKdb1493Txkv9fhhlzQqLnT8vklb/jttWsSvEKYihHlN
# ABUYcPXYHLniU6zkZrNoMx0oodfnLAaHO2aQDri/kbKjkNvp6x7rvLTfo5qztBYL
# VXJGGJfmF+1ecWGAXZx24hylrY9jOEH7cCLPlP7KG9utokhC5EkjhAuQ5jMB8luT
# tyyUIRPAiqUj5/tBCBmd/+KHuzTKTh12M1TxXnNrbN+UI4070t6L6XOWysE+7O7p
# oFkrAoWOQM06T9or3rPwOfybBw8Nr05U9fVNIeqTSbyOohEdfa9PQmWrHYEtkkcY
# Yod4WqXqfvtQz+TKos8DdTrOhZyiZbxlBnS/P8M4a14YlmQBhGIW/k0Yi56zptqc
# 7cLGkRA06w==
# SIG # End signature block
