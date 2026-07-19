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
0.2.0

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

$ScriptVersion = '0.2.0'
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

        $common = @{ Server = $domainContext.Server; ErrorAction = 'Stop' }
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
                'LastLogonDate', 'PasswordLastSet'
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
                'ManagedBy', 'WhenCreated', 'WhenChanged', 'LastLogonDate'
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

        if ($CollectMemberships) {
            $started = [datetimeoffset]::UtcNow
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

                $members = @(Get-ADGroupMember -Identity $group.DistinguishedName `
                        -Server $domainContext.Server -ErrorAction Stop)
                foreach ($member in $members) {
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
    -NoConfigWrite:($ValidateOnly -or $NoConfigWrite)
$paths = $context.Paths

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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAJ425mFvqwJIm6
# /AwDvf73Oz2gW+4V8ahdvLnx304t46CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHFO6gDmeaoMZEeXnORDCMsD0y7WptBs9FebhIEGkIs7MA0GCSqG
# SIb3DQEBAQUABIIBgHEuvxcGf15Pu/xndrYo6iqDJRLwCHDvnSmHpI5xPkGRbP98
# 0cqNcR6RWCXu9M3kguVk5SVNyXbhZB57HOVWPKnZnUUB2ejo0gAP5OzTRkKNioOj
# XMBnqS4+pN1OD/j4nW79SD5Czb8B0hpvRr482xHcNCRjHwENznqes8Eb7NkBV+n5
# dWA2ug/cwC5rkU9LVNp0k1xB7pbS7Oedo0EFXXaXZvDBLivXJVOaebyqAz3Nq6Rb
# QNPHsL/a/WyhHbguDOmNKOwUtuYzBHZfTxqne9TS0Tfsq+gBV56A++tiR+1ZLXuQ
# rkZXTac4oE7vIxdmC5iBnWQAgFBw/YvqXr29e6bJPjhzZiRNK87uwO7mkHn60KXG
# tLqnrnq6VDQyCIKxX9Sjf8TkzncwxHq5hLE3FzPhDw56RY94evWaj1RoIpBGuMPb
# 8zrMOu1jqmYu8AXF4OAUVqTKEThylnbjGd5WI/oJFKEdO8Wo26AsdCiHMJ9VoAZp
# Kws2G9Lwm07I/T/rIKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkyMjM2
# MTVaMC8GCSqGSIb3DQEJBDEiBCCXY5ppotl7F4CBlSdvXJ/HmJfruZgeUkDdJn96
# jl7KpTANBgkqhkiG9w0BAQEFAASCAgBWVYyWl9klhhurnjnKJh/wMc9sBhWFpovM
# lULbAHvJqtSpQL9HGsH1mF6FxgBFhWYCcOr2Y/optAa2cXHKf0jDAdDYEoad6a0c
# TFAPmgMp1a1VOReKDabb3LPquKbQDF83vBcPD4UGm2IoBpDCWU4F6lIKqPJPNPwQ
# oSBYJq8Ikm2+VraYAgzOgKinmMI3Gqq4hIPi3o4RO5GAKohYUAQpqag2kPr/J459
# vpeUY33l//7TzMldWjs/9S2G/JxONAn7Fe3xaHrvx9WqhgscQR4yTvqVpteXIIiM
# IRhep9pxVkHnbt5l946WeG3ClnusP4sswN3BXGEyepvBgpLK49Ujg5HZacV5r22v
# rodGAAWuTShOTz97EsZdvmQI8wp4aO5/uSTWOEruUdxRCY/lvZ38Wy2x2B38iKAV
# dx34nhe2gMU2fw2oPR1z6/DlbVEQ6Ty7jkbOf5hQZzQqYQkcJH62mA8jtjqqMCN6
# Ex1GMx6FYaHadhtjrXijNAb5bHSTBVuKr/KWHhjWY0aNlBafelMaGCPjo0hj5kDN
# NpTYXPbMSe7MZDGSyJnLi6rIIAd4SsNENftXHP9CQrMArBuyJ1IQ0zaoej7VAHUc
# ba2CbL0KmmBNcMYRU+aucgS7clPFGrEHGSYgRM4DO7upwL8Bd60qKR5corTpN7dC
# T9J6uS+ifA==
# SIG # End signature block
