<#
.SYNOPSIS
    Inventories SharePoint Online permissions to a dedicated CSV file.

.DESCRIPTION
    Uses PnP.PowerShell with interactive, device login, or certificate authentication.
    SiteUrl and WebUrlsFile inputs are treated as roots: descendant subsites
    are included by default for permission inventory.
    The output columns intentionally match SmartM365-SharePointSource-PermissionInventory.ps1 as closely
    as possible so both inventories can be compared.

.VERSION
    1.1.6
#>

[CmdletBinding(DefaultParameterSetName = 'WebUrlsFile')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Site')]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true, ParameterSetName = 'WebUrlsFile')]
    [ValidateNotNullOrEmpty()]
    [string]$WebUrlsFile,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath,

    [ValidateNotNullOrEmpty()]
    [string]$ErrorPath,

    [ValidateNotNullOrEmpty()]
    [string]$LogPath,

    [string]$ClientId = $env:SPO_INVENTORY_CLIENT_ID,

    [string]$Tenant = $env:SPO_INVENTORY_TENANT,

    [string]$TenantId = $env:SPO_INVENTORY_TENANT_ID,

    [string]$Thumbprint = $env:SPO_INVENTORY_THUMBPRINT,

    [switch]$Interactive,

    [switch]$DeviceLogin,

    [switch]$ForceAuthentication,

    [switch]$IncludeHiddenLists,

    [switch]$IncludeSystemLists,

    [switch]$DocumentLibrariesOnly,

    [switch]$IncludeItemPermissions = $true,

    [int]$PageSize = 2000,

    [ValidateRange(1, [int]::MaxValue)]
    [int]$ItemProgressInterval = 500
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Launchers\SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
$script:ConsoleLifecycleContext = Start-SmartM365MigrationConsoleLifecycle -ScriptPath $PSCommandPath
$script:ConsoleLifecycleFailure = $null
$script:ConsoleLifecycleStatus = 'CANCELLED'
try {

function Get-ConsoleTimestamp {
    return (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
}

function Write-ConsoleMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [System.ConsoleColor]$ForegroundColor
    )

    $line = "{0} {1}" -f (Get-ConsoleTimestamp), $Message
    if ($PSBoundParameters.ContainsKey('ForegroundColor')) {
        Microsoft.PowerShell.Utility\Write-Host $line -ForegroundColor $ForegroundColor
    }
    else {
        Microsoft.PowerShell.Utility\Write-Host $line
    }
}

function Write-ConsoleWarning {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-ConsoleMessage -Message ("WARNING: {0}" -f $Message) -ForegroundColor Yellow
}

function Add-TimestampToLogFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }

    $timestampPrefixPattern = '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} '
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $temporaryPath = "{0}.timestamp.tmp" -f $Path

    Get-Content -LiteralPath $Path | ForEach-Object {
        if ($_ -match $timestampPrefixPattern) {
            $_
        }
        else {
            "{0} {1}" -f $timestamp, $_
        }
    } | Set-Content -LiteralPath $temporaryPath -Encoding UTF8 -WhatIf:$false

    Move-Item -LiteralPath $temporaryPath -Destination $Path -Force -WhatIf:$false
}

function Stop-TimestampedTranscript {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Stop-Transcript | Out-Null
    $script:TranscriptStarted = $false
    Add-TimestampToLogFile -Path $Path
}

$PermissionColumns = @(
    'SiteCollectionUrl',
    'WebUrl',
    'WebTitle',
    'AssociatedMemberGroup',
    'AssociatedOwnerGroup',
    'AssociatedVisitorGroup',
    'ObjectScope',
    'ObjectUrl',
    'ObjectServerRelativeUrl',
    'ObjectTitle',
    'ObjectId',
    'ParentObjectUrl',
    'ListTitle',
    'ListUrl',
    'ListBaseTemplate',
    'ListBaseType',
    'IsDocumentLibrary',
    'ItemId',
    'ItemFileSystemObjectType',
    'HasUniqueRoleAssignments',
    'InheritedFrom',
    'PrincipalType',
    'PrincipalName',
    'PrincipalLoginName',
    'PrincipalId',
    'PrincipalMemberCount',
    'PrincipalUserMemberCount',
    'PrincipalDomainGroupMemberCount',
    'PrincipalMemberLoginNames',
    'PrincipalMemberDisplayNames',
    'PrincipalMemberLookupStatus',
    'PermissionLevels',
    'IsLimitedAccessOnly'
)

function ConvertTo-SafeFileName {
    param(
        [string]$Name,
        [string]$Fallback = 'SharePointOnline'
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        $Name = $Fallback
    }

    $invalidCharacters = [regex]::Escape((-join [System.IO.Path]::GetInvalidFileNameChars()))
    $safeName = [regex]::Replace($Name.Trim(), "[$invalidCharacters]+", '-')
    $safeName = [regex]::Replace($safeName, '\s+', ' ')
    $safeName = $safeName.Trim(" .-".ToCharArray())

    if ([string]::IsNullOrWhiteSpace($safeName)) {
        $safeName = $Fallback
    }

    if ($safeName.Length -gt 80) {
        $safeName = $safeName.Substring(0, 80).Trim(" .-".ToCharArray())
    }

    return $safeName
}

function Format-InventoryDuration {
    param(
        [TimeSpan]$Elapsed
    )

    if ($Elapsed.Days -gt 0) {
        return ("{0}d {1:00}:{2:00}:{3:00}" -f $Elapsed.Days, $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds)
    }

    return ("{0:00}:{1:00}:{2:00}" -f $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds)
}

function Write-ItemPermissionHeartbeat {
    param(
        [string]$WebUrl,
        [string]$ListTitle,
        [int]$ProcessedItems,
        [int]$UniquePermissionItems,
        [int]$ExportedRows,
        [TimeSpan]$Elapsed,
        [string]$LastItem
    )

    Write-ConsoleMessage -Message ("  Item heartbeat: web='{0}' library='{1}' processed={2}; unique={3}; permission rows exported={4}; elapsed={5}; last='{6}'" -f `
            $WebUrl, `
            $ListTitle, `
            $ProcessedItems, `
            $UniquePermissionItems, `
            $ExportedRows, `
            (Format-InventoryDuration -Elapsed $Elapsed), `
            $LastItem) -ForegroundColor DarkCyan
}

function Write-PermissionHeaderOnlyCsv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $header = ($PermissionColumns | ForEach-Object { '"{0}"' -f ($_ -replace '"', '""') }) -join ';'
    Set-Content -LiteralPath $Path -Value $header -Encoding UTF8
}

function Export-PermissionRows {
    param(
        [object[]]$Rows,
        [string]$CsvPath
    )

    if (-not $Rows -or $Rows.Count -eq 0) {
        return
    }

    if ($script:CsvCreated) {
        $Rows | Export-Csv -Delimiter ';' -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Append
    }
    else {
        $Rows | Export-Csv -Delimiter ';' -Path $CsvPath -NoTypeInformation -Encoding UTF8
        $script:CsvCreated = $true
    }
}

function Write-InventoryError {
    param(
        [string]$Scope,
        [string]$Url,
        [string]$Name,
        [string]$Message,
        [string]$ItemId = '',
        [string]$ItemUrl = ''
    )

    if ([string]::IsNullOrWhiteSpace($script:ErrorPath)) {
        return
    }

    $row = [pscustomobject]@{
        Time    = Get-Date
        Scope   = $Scope
        Url     = $Url
        Name    = $Name
        Message = $Message
        ItemId  = $ItemId
        ItemUrl = $ItemUrl
    }

    if ($script:ErrorCsvCreated) {
        $row | Export-Csv -Delimiter ';' -Path $script:ErrorPath -NoTypeInformation -Encoding UTF8 -Append
    }
    else {
        $row | Export-Csv -Delimiter ';' -Path $script:ErrorPath -NoTypeInformation -Encoding UTF8
        $script:ErrorCsvCreated = $true
    }
}

function Import-RequiredModule {
    $module = Get-Module -ListAvailable -Name 'PnP.PowerShell' |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($null -eq $module) {
        throw "Required module 'PnP.PowerShell' is not installed."
    }

    Import-Module PnP.PowerShell -ErrorAction Stop
    Write-ConsoleMessage -Message ("PnP.PowerShell version: {0}" -f $module.Version)
    Write-ConsoleMessage -Message ("PnP.PowerShell path: {0}" -f $module.Path)
}

function Connect-ToSPOWeb {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    Write-ConsoleMessage -Message ("Connecting to: {0}" -f $Url)

    $parameters = @{
        Url         = $Url
        ErrorAction = 'Stop'
        ReturnConnection = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($ClientId)) {
        $parameters.ClientId = $ClientId
    }

    if (-not [string]::IsNullOrWhiteSpace($Tenant)) {
        $parameters.Tenant = $Tenant
    }

    if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
        $parameters.Tenant = $TenantId
    }

    if ($Interactive) {
        $parameters.Interactive = $true
    }
    elseif ($DeviceLogin) {
        $parameters.DeviceLogin = $true
    }
    elseif (-not [string]::IsNullOrWhiteSpace($Thumbprint)) {
        $parameters.Thumbprint = $Thumbprint
    }
    else {
        $parameters.Interactive = $true
    }

    if ($ForceAuthentication -and -not $script:ForceAuthenticationUsed -and -not $parameters.ContainsKey('Thumbprint')) {
        try {
            Disconnect-PnPOnline -ClearPersistedLogin -ErrorAction SilentlyContinue
            Write-ConsoleMessage -Message "Cleared persisted PnP login before forced authentication."
        }
        catch {
            Write-ConsoleWarning -Message ("Could not clear persisted PnP login: {0}" -f $_.Exception.Message)
        }

        $parameters.ForceAuthentication = $true
        $script:ForceAuthenticationUsed = $true
    }

    $connection = Connect-PnPOnline @parameters
    $script:SPOPermissionConnection = $connection
    Write-SPOConnectionIdentity -Connection $connection -Url $Url
    $tokenContext = '{0}|{1}|{2}|{3}|{4}|{5}|{6}' -f ([uri]$Url).Host, $Tenant, $TenantId, $ClientId, $Thumbprint, [bool]$DeviceLogin, $script:SPOConnectedAccount
    if ($script:TokenSummaryContexts.Add($tokenContext)) { Write-PnPTokenSummary -Connection $connection }
}

function Write-SPOConnectionIdentity {
    param(
        $Connection,
        [string]$Url
    )

    $script:SPOConnectedAccount = $Url
    try {
        $context = Get-PnPContext -Connection $Connection
        $context.Load($context.Web.CurrentUser)
        $context.ExecuteQuery()

        $currentUser = $context.Web.CurrentUser
        $script:SPOConnectedAccount = [string]$currentUser.LoginName
        $loginName = if ($currentUser.LoginName) { $currentUser.LoginName } else { '<unknown>' }
        $email = if ($currentUser.Email) { $currentUser.Email } else { '<no email>' }
        $title = if ($currentUser.Title) { $currentUser.Title } else { '<no title>' }

        Write-ConsoleMessage -Message ("Connected account for {0}: {1} | {2} | {3}" -f $Url, $loginName, $email, $title) -ForegroundColor Cyan
    }
    catch {
        Write-ConsoleWarning -Message ("Could not determine connected account for '{0}': {1}" -f $Url, $_.Exception.Message)
    }
}

function ConvertFrom-Base64Url {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $base64 = $Value.Replace('-', '+').Replace('_', '/')
    while ($base64.Length % 4 -ne 0) {
        $base64 += '='
    }

    return [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($base64))
}

function Write-PnPTokenSummary {
    param(
        $Connection
    )

    try {
        $sharePointScopes = @(Get-PnPAccessToken -ResourceTypeName SharePoint -ListPermissionScopes -Connection $Connection -ErrorAction Stop)
        if ($sharePointScopes.Count -gt 0) {
            Write-ConsoleMessage -Message ("PnP SharePoint token scopes: {0}" -f ($sharePointScopes -join ' ')) -ForegroundColor DarkCyan
        }
        else {
            Write-ConsoleMessage -Message "PnP SharePoint token scopes: <none returned>" -ForegroundColor DarkYellow
        }
    }
    catch {
        Write-ConsoleWarning -Message ("Could not list SharePoint token scopes: {0}" -f $_.Exception.Message)
    }

    try {
        if ($null -eq $Connection -or [string]::IsNullOrWhiteSpace($Connection.AccessToken)) {
            Write-ConsoleMessage -Message "PnP token details: <not available>" -ForegroundColor DarkYellow
            return
        }

        $parts = $Connection.AccessToken.Split('.')
        if ($parts.Count -lt 2) {
            Write-ConsoleMessage -Message "PnP token details: <unrecognized token format>" -ForegroundColor DarkYellow
            return
        }

        $payload = ConvertFrom-Json -InputObject (ConvertFrom-Base64Url -Value $parts[1])
        $appId = if ($payload.appid) { $payload.appid } elseif ($payload.azp) { $payload.azp } else { '<none>' }
        $scopes = if ($payload.scp) { $payload.scp } else { '<none>' }
        $roles = if ($payload.roles) { ($payload.roles -join ' ') } else { '<none>' }

        Write-ConsoleMessage -Message ("PnP token app id: {0}" -f $appId) -ForegroundColor DarkCyan
        Write-ConsoleMessage -Message ("PnP token scopes: {0}" -f $scopes) -ForegroundColor DarkCyan
        Write-ConsoleMessage -Message ("PnP token roles: {0}" -f $roles) -ForegroundColor DarkCyan
    }
    catch {
        Write-ConsoleWarning -Message ("Could not decode PnP token details: {0}" -f $_.Exception.Message)
    }
}

function Get-SiteCollectionUrlFromWebUrl {
    param(
        [string]$WebUrl
    )

    $uri = [System.Uri]$WebUrl
    $segments = @($uri.AbsolutePath.Trim('/').Split('/') | Where-Object { $_ })
    if ($segments.Count -ge 2 -and $segments[0].Equals('sites', [System.StringComparison]::OrdinalIgnoreCase)) {
        return ("{0}://{1}/sites/{2}" -f $uri.Scheme, $uri.Authority, $segments[1])
    }

    return ("{0}://{1}" -f $uri.Scheme, $uri.Authority)
}

function ConvertTo-AbsoluteSharePointUrl {
    param(
        [string]$WebUrl,
        [string]$ServerRelativeUrl
    )

    if ([string]::IsNullOrWhiteSpace($ServerRelativeUrl)) {
        return $null
    }

    if ($ServerRelativeUrl -match '^https?://') {
        return $ServerRelativeUrl
    }

    if (-not $ServerRelativeUrl.StartsWith('/')) {
        $ServerRelativeUrl = "/$ServerRelativeUrl"
    }

    $uri = [System.Uri]$WebUrl
    return ("{0}://{1}{2}" -f $uri.Scheme, $uri.Authority, $ServerRelativeUrl)
}

function Test-SystemList {
    param(
        $List
    )

    $rootFolder = Get-PnPProperty -ClientObject $List -Property RootFolder -Connection $script:SPOPermissionConnection
    $systemUrls = @(
        '_catalogs/masterpage',
        '_catalogs/wp',
        '_catalogs/lt',
        'Style Library',
        'FormServerTemplates',
        'PreservationHoldLibrary',
        'Site Collection Documents',
        'Site Collection Images'
    )

    foreach ($url in $systemUrls) {
        if ($rootFolder.ServerRelativeUrl -like "*/$url*" -or $rootFolder.Name -eq $url) {
            return $true
        }
    }

    return $false
}

function Get-PrincipalInfo {
    param(
        $Member
    )

    $principalType = $Member.PrincipalType
    $name = $Member.Title
    $loginName = $Member.LoginName
    $id = $Member.Id

    [pscustomobject]@{
        PrincipalType      = $principalType
        PrincipalName      = $name
        PrincipalLoginName = $loginName
        PrincipalId        = $id
    }
}

function Get-PnPGroupTitle {
    param(
        $Group
    )

    if ($null -eq $Group) {
        return ''
    }

    # A single-property request returns its value, not the original group object.
    return [string](Get-PnPProperty -ClientObject $Group -Property Title -ErrorAction Stop -Connection $script:SPOPermissionConnection)
}

function Get-AssociatedWebGroupNames {
    param(
        $Web
    )

    $key = [string]$Web.Url
    if ($script:AssociatedWebGroupCache.ContainsKey($key)) { return $script:AssociatedWebGroupCache[$key] }
    $names = [ordered]@{}
    foreach ($property in @('AssociatedMemberGroup', 'AssociatedOwnerGroup', 'AssociatedVisitorGroup')) {
        $names[$property] = ''
        try {
            $names[$property] = Get-PnPGroupTitle -Group (Get-PnPProperty -ClientObject $Web -Property $property -ErrorAction Stop -Connection $script:SPOPermissionConnection)
        }
        catch {
            Write-ConsoleWarning -Message ("Failed to read {0} for '{1}': {2}" -f $property, $key, $_.Exception.Message)
            Write-InventoryError -Scope 'AssociatedGroup' -Url $key -Name $property -Message $_.Exception.Message
        }
    }
    $result = [pscustomobject]$names
    $script:AssociatedWebGroupCache[$key] = $result
    return $result
}

$script:AssociatedWebGroupCache = @{}
$script:TokenSummaryContexts = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$script:SPOInheritanceLoad = $null
$script:SPOInheritanceItemType = $null
$script:SharePointGroupMembershipCache = @{}

function Join-PrincipalMemberValues {
    param(
        [object[]]$Values
    )

    if (-not $Values -or $Values.Count -eq 0) {
        return ''
    }

    return (($Values | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique) -join ' || ')
}

function Get-EmptyPrincipalMembershipInfo {
    param(
        [string]$Status = ''
    )

    [pscustomobject]@{
        PrincipalMemberCount            = ''
        PrincipalUserMemberCount        = ''
        PrincipalDomainGroupMemberCount = ''
        PrincipalMemberLoginNames       = ''
        PrincipalMemberDisplayNames     = ''
        PrincipalMemberLookupStatus     = $Status
    }
}

function Get-PrincipalMembershipInfo {
    param(
        $Member,
        [string]$PrincipalType
    )

    if ($PrincipalType -ne 'SharePointGroup') {
        return Get-EmptyPrincipalMembershipInfo
    }

    $groupIdentity = if (-not [string]::IsNullOrWhiteSpace([string]$Member.Title)) { [string]$Member.Title } else { [string]$Member.LoginName }
    if ([string]::IsNullOrWhiteSpace($groupIdentity)) {
        return Get-EmptyPrincipalMembershipInfo -Status 'Skipped: group identity is empty'
    }

    $cacheKey = if ($null -ne $Member.Id) { [string]$Member.Id } else { $groupIdentity.ToLowerInvariant() }
    if ($script:SharePointGroupMembershipCache.ContainsKey($cacheKey)) {
        return $script:SharePointGroupMembershipCache[$cacheKey]
    }

    try {
        $members = @(Get-PnPGroupMember -Group $groupIdentity -ErrorAction Stop -Connection $script:SPOPermissionConnection)
        $loginNames = @($members | ForEach-Object { $_.LoginName })
        $displayNames = @($members | ForEach-Object { $_.Title })
        $domainGroupMembers = @($members | Where-Object { [string]$_.PrincipalType -match 'SecurityGroup|DistributionList|SharePointGroup' })
        $userMembers = @($members | Where-Object { [string]$_.PrincipalType -eq 'User' })

        $info = [pscustomobject]@{
            PrincipalMemberCount            = $members.Count
            PrincipalUserMemberCount        = $userMembers.Count
            PrincipalDomainGroupMemberCount = $domainGroupMembers.Count
            PrincipalMemberLoginNames       = Join-PrincipalMemberValues -Values $loginNames
            PrincipalMemberDisplayNames     = Join-PrincipalMemberValues -Values $displayNames
            PrincipalMemberLookupStatus     = 'OK'
        }
    }
    catch {
        $info = Get-EmptyPrincipalMembershipInfo -Status ("Failed: {0}" -f $_.Exception.Message)
    }

    $script:SharePointGroupMembershipCache[$cacheKey] = $info
    return $info
}
function Get-RoleAssignmentRows {
    param(
        [string]$SiteCollectionUrl,
        [string]$WebUrl,
        [string]$WebTitle,
        [string]$AssociatedMemberGroup,
        [string]$AssociatedOwnerGroup,
        [string]$AssociatedVisitorGroup,
        [string]$ObjectScope,
        [string]$ObjectUrl,
        [string]$ObjectServerRelativeUrl,
        [string]$ObjectTitle,
        [string]$ObjectId,
        [string]$ParentObjectUrl,
        [string]$ListTitle,
        [string]$ListUrl,
        [object]$ListBaseTemplate,
        [string]$ListBaseType,
        [object]$IsDocumentLibrary,
        [object]$ItemId,
        [string]$ItemFileSystemObjectType,
        [bool]$HasUniqueRoleAssignments,
        [string]$InheritedFrom,
        $RoleAssignments
    )

    $assignmentIndex = 0
    foreach ($roleAssignment in $RoleAssignments) {
        $assignmentIndex++
        try {
            $label = "role assignment #{0} on '{1}' ({2})" -f $assignmentIndex, $ObjectUrl, (Get-SPORoleAssignmentIdentity $roleAssignment)
            $properties = Invoke-SPORead -Label $label -Operation {
                Get-PnPProperty -ClientObject $roleAssignment -Property Member,RoleDefinitionBindings -Connection $script:SPOPermissionConnection -ErrorAction Stop | Out-Null
                # Multiple-property loads populate the client object without returning values.
                [pscustomobject]@{ Member = $roleAssignment.Member; Bindings = @($roleAssignment.RoleDefinitionBindings) }
            }
            $member = $properties.Member
            $bindings = @($properties.Bindings)
            $permissionLevels = @($bindings | ForEach-Object { $_.Name })
            if ($permissionLevels.Count -eq 0) {
                continue
            }

            $principal = Get-PrincipalInfo -Member $member
            $membership = Get-PrincipalMembershipInfo -Member $member -PrincipalType $principal.PrincipalType
            [pscustomobject]@{
                SiteCollectionUrl        = $SiteCollectionUrl
                WebUrl                   = $WebUrl
                WebTitle                 = $WebTitle
                AssociatedMemberGroup    = $AssociatedMemberGroup
                AssociatedOwnerGroup     = $AssociatedOwnerGroup
                AssociatedVisitorGroup   = $AssociatedVisitorGroup
                ObjectScope              = $ObjectScope
                ObjectUrl                = $ObjectUrl
                ObjectServerRelativeUrl  = $ObjectServerRelativeUrl
                ObjectTitle              = $ObjectTitle
                ObjectId                 = $ObjectId
                ParentObjectUrl          = $ParentObjectUrl
                ListTitle                = $ListTitle
                ListUrl                  = $ListUrl
                ListBaseTemplate        = $ListBaseTemplate
                ListBaseType            = $ListBaseType
                IsDocumentLibrary       = $IsDocumentLibrary
                ItemId                   = $ItemId
                ItemFileSystemObjectType = $ItemFileSystemObjectType
                HasUniqueRoleAssignments = $HasUniqueRoleAssignments
                InheritedFrom            = $InheritedFrom
                PrincipalType            = $principal.PrincipalType
                PrincipalName            = $principal.PrincipalName
                PrincipalLoginName       = $principal.PrincipalLoginName
                PrincipalId              = $principal.PrincipalId
                PrincipalMemberCount         = $membership.PrincipalMemberCount
                PrincipalUserMemberCount     = $membership.PrincipalUserMemberCount
                PrincipalDomainGroupMemberCount = $membership.PrincipalDomainGroupMemberCount
                PrincipalMemberLoginNames    = $membership.PrincipalMemberLoginNames
                PrincipalMemberDisplayNames  = $membership.PrincipalMemberDisplayNames
                PrincipalMemberLookupStatus  = $membership.PrincipalMemberLookupStatus
                PermissionLevels         = ($permissionLevels -join '|')
                IsLimitedAccessOnly      = ($permissionLevels.Count -eq 1 -and $permissionLevels[0] -eq 'Limited Access')
            }
        }
        catch {
            $details = "Assignment #{0}; {1}; {2}" -f $assignmentIndex, (Get-SPORoleAssignmentIdentity $roleAssignment), (Get-SPOExceptionDetails $_.Exception)
            Write-ConsoleWarning -Message ("Failed to read role assignment on '{0}': {1}" -f $ObjectUrl, $details)
            Write-InventoryError -Scope "$ObjectScope RoleAssignment" -Url $ObjectUrl -Name $ObjectTitle -ItemId ([string]$ItemId) -ItemUrl $(if ($ObjectScope -eq 'Item') { $ObjectUrl } else { '' }) -Message $details
        }
    }
}

function Export-ItemPermissionInventory {
    param(
        $Web,
        $List,
        [string]$CsvPath,
        [int]$ProgressInterval
    )

    $associatedGroups = Get-AssociatedWebGroupNames -Web $Web
    $rootFolder = Get-PnPProperty -ClientObject $List -Property RootFolder -Connection $script:SPOPermissionConnection
    $listUrl = ConvertTo-AbsoluteSharePointUrl -WebUrl $Web.Url -ServerRelativeUrl $rootFolder.ServerRelativeUrl
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    $state = @{ Processed=0; Unique=0; Exported=0; LastItem='<none>' }
    $script:SPOPermissionPageState = $state
    $script:SPOPermissionPageWeb = $Web
    $script:SPOPermissionPageList = $List
    $script:SPOPermissionPageCsv = $CsvPath
    $script:SPOPermissionPageInterval = $ProgressInterval
    $script:SPOPermissionPageStopwatch = $stopwatch
    try {
        # Process the callback before the next page is fetched; do not retain the whole library.
        Get-PnPListItem -List $List -PageSize $PageSize -Fields 'FileRef','FileLeafRef','FSObjType','UniqueId','ID' -Connection $script:SPOPermissionConnection -ScriptBlock {
            param($PageItems)
            Export-SPOPermissionPage -PageItems $PageItems -Web $script:SPOPermissionPageWeb -List $script:SPOPermissionPageList -CsvPath $script:SPOPermissionPageCsv -State $script:SPOPermissionPageState -ProgressInterval $script:SPOPermissionPageInterval -Stopwatch $script:SPOPermissionPageStopwatch
        } -ErrorAction Stop | Out-Null
    }
    catch {
        Write-ConsoleWarning -Message ("Failed to enumerate items for list '{0}' in web '{1}': {2}" -f $List.Title, $Web.Url, $_.Exception.Message)
        Write-InventoryError -Scope 'ListItems' -Url $listUrl -Name $List.Title -Message $_.Exception.Message
    }
    finally { $stopwatch.Stop() }
    Write-ItemPermissionHeartbeat -WebUrl $Web.Url -ListTitle $List.Title -ProcessedItems $state.Processed -UniquePermissionItems $state.Unique -ExportedRows $state.Exported -Elapsed $stopwatch.Elapsed -LastItem $state.LastItem
}

function Export-SPOPermissionPage {
    param($PageItems, $Web, $List, [string]$CsvPath, [hashtable]$State, [int]$ProgressInterval, $Stopwatch)
    $associatedGroups = Get-AssociatedWebGroupNames -Web $Web
    $rootFolder = Get-PnPProperty -ClientObject $List -Property RootFolder -Connection $script:SPOPermissionConnection
    $listUrl = ConvertTo-AbsoluteSharePointUrl -WebUrl $Web.Url -ServerRelativeUrl $rootFolder.ServerRelativeUrl
    $inheritance = Get-SPOPageInheritance -PageItems $PageItems -Web $Web -List $List -ListUrl $listUrl
    foreach ($item in $PageItems) {
        $State.Processed++
        $State.LastItem = "ID $($item.Id)"
        try {
            if ($State.Processed % $ProgressInterval -eq 0) {
                Write-ItemPermissionHeartbeat -WebUrl $Web.Url -ListTitle $List.Title -ProcessedItems $State.Processed -UniquePermissionItems $State.Unique -ExportedRows $State.Exported -Elapsed $Stopwatch.Elapsed -LastItem $State.LastItem
            }
            if (-not $inheritance.ContainsKey([int]$item.Id) -or -not $inheritance[[int]$item.Id]) { continue }
            $State.Unique++
            $roleAssignments = Get-PnPProperty -ClientObject $item -Property RoleAssignments -Connection $script:SPOPermissionConnection
            $serverRelativeUrl = [string]$item.FieldValues.FileRef
            $absoluteUrl = ConvertTo-AbsoluteSharePointUrl -WebUrl $Web.Url -ServerRelativeUrl $serverRelativeUrl
            $fsObjType = if ([string]$item.FieldValues.FSObjType -eq '1') { 'Folder' } else { 'FileOrItem' }
            $itemName = [string]$item.FieldValues.FileLeafRef
            if ([string]::IsNullOrWhiteSpace($itemName)) {
                $itemName = [string]$item.Id
            }

            $rows = @(Get-RoleAssignmentRows `
                    -SiteCollectionUrl (Get-SiteCollectionUrlFromWebUrl -WebUrl $Web.Url) `
                    -WebUrl $Web.Url `
                    -WebTitle $Web.Title `
                    -AssociatedMemberGroup $associatedGroups.AssociatedMemberGroup `
                    -AssociatedOwnerGroup $associatedGroups.AssociatedOwnerGroup `
                    -AssociatedVisitorGroup $associatedGroups.AssociatedVisitorGroup `
                    -ObjectScope 'Item' `
                    -ObjectUrl $absoluteUrl `
                    -ObjectServerRelativeUrl $serverRelativeUrl `
                    -ObjectTitle $itemName `
                    -ObjectId ([string]$item.FieldValues.UniqueId) `
                    -ParentObjectUrl $listUrl `
                    -ListTitle $List.Title `
                    -ListUrl $listUrl `
                    -ListBaseTemplate ([int]$List.BaseTemplate) `
                    -ListBaseType $(if ([string]$List.BaseType -eq 'DocumentLibrary') { 'DocumentLibrary' } else { '' }) `
                    -IsDocumentLibrary ([string]$List.BaseType -eq 'DocumentLibrary') `
                    -ItemId $item.Id `
                    -ItemFileSystemObjectType $fsObjType `
                    -HasUniqueRoleAssignments $true `
                    -InheritedFrom '' `
                    -RoleAssignments $roleAssignments)

            Export-PermissionRows -Rows $rows -CsvPath $CsvPath
            $State.Exported += $rows.Count
        }
        catch {
            Write-SPOItemError -Item $item -Web $Web -List $List -ListUrl $listUrl -Message $_.Exception.Message
        }
    }
}

function Add-SPOInheritanceRead {
    param($Context, $Item)
    # Queue only HasUniqueRoleAssignments through the supported CSOM Load<T> API.
    $itemType = $Item.GetType()
    if (-not $script:SPOInheritanceLoad -or $script:SPOInheritanceItemType -ne $itemType) {
        $parameter = [System.Linq.Expressions.Expression]::Parameter($itemType, 'item')
        $property = [System.Linq.Expressions.Expression]::Property($parameter, 'HasUniqueRoleAssignments')
        $boxed = [System.Linq.Expressions.Expression]::Convert($property, [object])
        $delegateType = [System.Func`2].MakeGenericType($itemType, [object])
        $lambda = [System.Linq.Expressions.Expression]::Lambda($delegateType, $boxed, [System.Linq.Expressions.ParameterExpression[]]@($parameter))
        $expressionType = [System.Linq.Expressions.Expression`1].MakeGenericType($delegateType)
        $script:SPOInheritanceSelectors = [Array]::CreateInstance($expressionType, 1)
        $script:SPOInheritanceSelectors.SetValue($lambda, 0)
        $method = @($Context.GetType().GetMethods() | Where-Object { $_.Name -eq 'Load' -and $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 2 })[0]
        $script:SPOInheritanceLoad = $method.MakeGenericMethod($itemType)
        $script:SPOInheritanceItemType = $itemType
    }
    [void]$script:SPOInheritanceLoad.Invoke($Context, [object[]]@($Item, $script:SPOInheritanceSelectors))
}

function Write-SPOItemError {
    param($Item, $Web, $List, [string]$ListUrl, [string]$Message)
    $itemUrl = ConvertTo-AbsoluteSharePointUrl -WebUrl $Web.Url -ServerRelativeUrl ([string]$Item.FieldValues.FileRef)
    if ($Message -match '(?i)item does not exist|does not exist.*item') {
        try {
            $probe = Get-PnPListItem -List $List -Id $Item.Id -Fields 'FileRef' -Connection $script:SPOPermissionConnection -ErrorAction Stop
            $Message += $(if ($null -eq $probe) { ' Existence check: no item returned.' } else { ' Existence check: item still exists; permission read failed.' })
        }
        catch { $Message += ' Existence check failed: ' + $_.Exception.Message }
    }
    Write-ConsoleWarning -Message ("Failed item permission read: list='{0}'; ID={1}; path='{2}'; {3}" -f $List.Title, $Item.Id, $itemUrl, $Message)
    Write-InventoryError -Scope 'Item' -Url $ListUrl -Name $List.Title -ItemId ([string]$Item.Id) -ItemUrl $itemUrl -Message $Message
}

function Get-SPOPageInheritance {
    param($PageItems, $Web, $List, [string]$ListUrl)
    $context = Get-PnPContext -Connection $script:SPOPermissionConnection
    $result = @{}
    # A 2,000-item page can exceed SPO's 2 MiB CSOM request limit. Keep the
    # enumeration page size, but load inheritance in smaller independent requests.
    $items = @($PageItems)
    $batchSize = 100
    for ($offset = 0; $offset -lt $items.Count; $offset += $batchSize) {
        $last = [Math]::Min($offset + $batchSize - 1, $items.Count - 1)
        $batch = @($items[$offset..$last])
        $label = "inheritance batch in {0} ({1} items; IDs {2}-{3})" -f $List.Title, $batch.Count, $batch[0].Id, $batch[-1].Id
        try {
            Invoke-SPORead -Label $label -Operation {
                foreach ($item in $batch) { Add-SPOInheritanceRead -Context $context -Item $item }
                $context.ExecuteQuery()
            } | Out-Null
            foreach ($item in $batch) { $result[[int]$item.Id] = [bool]$item.HasUniqueRoleAssignments }
        }
        catch {
            if ($_.Exception.Message -notmatch '(?i)item does not exist|does not exist.*item') {
                throw ("Failed {0}: {1}" -f $label, $_.Exception.Message)
            }
            # Isolate missing reads in this batch; continue collecting later batches.
            foreach ($item in $batch) {
                try { $result[[int]$item.Id] = [bool](Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments -ErrorAction Stop -Connection $script:SPOPermissionConnection) }
                catch { Write-SPOItemError -Item $item -Web $Web -List $List -ListUrl $ListUrl -Message $_.Exception.Message }
            }
        }
    }
    return $result
}

function Export-WebPermissionInventory {
    param(
        [string]$WebUrl,
        [string]$CsvPath
    )

    Connect-ToSPOWeb -Url $WebUrl
    $web = Invoke-SPORead -Label 'web' -Operation { Get-PnPWeb -Includes Title,Url,ServerRelativeUrl,HasUniqueRoleAssignments -ErrorAction Stop -Connection $script:SPOPermissionConnection }
    $siteCollectionUrl = Get-SiteCollectionUrlFromWebUrl -WebUrl $web.Url

    Write-ConsoleMessage -Message ("Web: {0}" -f $web.Url)
    $associatedGroups = Get-AssociatedWebGroupNames -Web $web

    try {
        $webRoleAssignments = Get-PnPProperty -ClientObject $web -Property RoleAssignments -Connection $script:SPOPermissionConnection
        $rows = @(Get-RoleAssignmentRows `
                -SiteCollectionUrl $siteCollectionUrl `
                -WebUrl $web.Url `
                -WebTitle $web.Title `
                -AssociatedMemberGroup $associatedGroups.AssociatedMemberGroup `
                -AssociatedOwnerGroup $associatedGroups.AssociatedOwnerGroup `
                -AssociatedVisitorGroup $associatedGroups.AssociatedVisitorGroup `
                -ObjectScope 'Web' `
                -ObjectUrl $web.Url `
                -ObjectServerRelativeUrl $web.ServerRelativeUrl `
                -ObjectTitle $web.Title `
                -ObjectId ([string]$web.Id) `
                -ParentObjectUrl '' `
                -ListTitle '' `
                -ListUrl '' `
                -ListBaseTemplate '' `
                -ListBaseType '' `
                -IsDocumentLibrary '' `
                -ItemId $null `
                -ItemFileSystemObjectType '' `
                -HasUniqueRoleAssignments $web.HasUniqueRoleAssignments `
                -InheritedFrom $(if ($web.HasUniqueRoleAssignments) { '' } else { 'ParentWeb' }) `
                -RoleAssignments $webRoleAssignments)
        Export-PermissionRows -Rows $rows -CsvPath $CsvPath
    }
    catch {
        Write-ConsoleWarning -Message ("Failed to inventory web permissions '{0}': {1}" -f $web.Url, $_.Exception.Message)
        Write-InventoryError -Scope 'Web' -Url $web.Url -Name $web.Title -Message $_.Exception.Message
    }

    try {
        $lists = Invoke-SPORead -Label 'lists' -Operation { Get-PnPList -Includes Title,Hidden,BaseTemplate,BaseType,Id,RootFolder,HasUniqueRoleAssignments -ErrorAction Stop -Connection $script:SPOPermissionConnection }
    }
    catch {
        Write-ConsoleWarning -Message ("Failed to enumerate lists for web '{0}': {1}" -f $web.Url, $_.Exception.Message)
        Write-InventoryError -Scope 'WebLists' -Url $web.Url -Name $web.Title -Message $_.Exception.Message
        return
    }

    foreach ($list in $lists) {
        $listTitle = '<unknown>'
        $listUrl = $web.Url

        try {
            $rootFolder = Get-PnPProperty -ClientObject $list -Property RootFolder -Connection $script:SPOPermissionConnection
            $listTitle = $list.Title
            $listUrl = ConvertTo-AbsoluteSharePointUrl -WebUrl $web.Url -ServerRelativeUrl $rootFolder.ServerRelativeUrl

            if (-not $IncludeHiddenLists -and $list.Hidden) {
                Write-ConsoleMessage -Message ("  Skipping list/library '{0}' ({1}): hidden; use -IncludeHiddenLists to include it." -f $listTitle, $listUrl)
                continue
            }

            if (-not $IncludeSystemLists -and (Test-SystemList -List $list)) {
                Write-ConsoleMessage -Message ("  Skipping list/library '{0}' ({1}): system; use -IncludeSystemLists to include it." -f $listTitle, $listUrl)
                continue
            }

            if ($DocumentLibrariesOnly -and [string]$list.BaseType -ne 'DocumentLibrary') {
                Write-ConsoleMessage -Message ("  Skipping list '{0}' ({1}): not a document library; -DocumentLibrariesOnly is enabled." -f $listTitle, $listUrl)
                continue
            }

            $isDocumentLibrary = ([string]$list.BaseType -eq 'DocumentLibrary')
            $listBaseType = if ($isDocumentLibrary) { 'DocumentLibrary' } else { '' }
            $roleAssignments = Get-PnPProperty -ClientObject $list -Property RoleAssignments -Connection $script:SPOPermissionConnection
            $rows = @(Get-RoleAssignmentRows `
                    -SiteCollectionUrl $siteCollectionUrl `
                    -WebUrl $web.Url `
                    -WebTitle $web.Title `
                    -AssociatedMemberGroup $associatedGroups.AssociatedMemberGroup `
                    -AssociatedOwnerGroup $associatedGroups.AssociatedOwnerGroup `
                    -AssociatedVisitorGroup $associatedGroups.AssociatedVisitorGroup `
                    -ObjectScope 'List' `
                    -ObjectUrl $listUrl `
                    -ObjectServerRelativeUrl $rootFolder.ServerRelativeUrl `
                    -ObjectTitle $list.Title `
                    -ObjectId ([string]$list.Id) `
                    -ParentObjectUrl $web.Url `
                    -ListTitle $list.Title `
                    -ListUrl $listUrl `
                    -ListBaseTemplate ([int]$list.BaseTemplate) `
                    -ListBaseType $listBaseType `
                    -IsDocumentLibrary $isDocumentLibrary `
                    -ItemId $null `
                    -ItemFileSystemObjectType '' `
                    -HasUniqueRoleAssignments $list.HasUniqueRoleAssignments `
                    -InheritedFrom $(if ($list.HasUniqueRoleAssignments) { '' } else { 'Web' }) `
                    -RoleAssignments $roleAssignments)
            Export-PermissionRows -Rows $rows -CsvPath $CsvPath

            if ($IncludeItemPermissions) {
                Write-ConsoleMessage -Message ("  Item permissions: {0}" -f $list.Title)
                Export-ItemPermissionInventory -Web $web -List $list -CsvPath $CsvPath -ProgressInterval $ItemProgressInterval
            }
        }
        catch {
            Write-ConsoleWarning -Message ("Failed to inventory list permissions '{0}' in web '{1}': {2}" -f $listTitle, $web.Url, $_.Exception.Message)
            Write-InventoryError -Scope 'List' -Url $listUrl -Name $listTitle -Message $_.Exception.Message
        }
    }
}

function Get-DefaultOutputPath {
    $scriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $targetName = if ($PSCmdlet.ParameterSetName -eq 'Site') {
        ConvertTo-SafeFileName -Name ([System.Uri]$SiteUrl).Segments[-1].Trim('/')
    }
    else {
        ConvertTo-SafeFileName -Name ([System.IO.Path]::GetFileNameWithoutExtension($WebUrlsFile))
    }

    $targetDirectory = Join-Path -Path $scriptDirectory -ChildPath $targetName
    return Join-Path -Path $targetDirectory -ChildPath ("SPO-PermissionInventory-{0}-{1:yyyyMMdd-HHmmss}.csv" -f $targetName, (Get-Date))
}

function Get-WebUrlsFromFile {
    param(
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Web URLs file not found: $Path"
    }

    Get-Content -LiteralPath $Path |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -and -not $_.StartsWith('#') } |
        Sort-Object -Unique
}


function Write-SPOReadRetry { param([string]$Message) Write-ConsoleWarning -Message $Message }

function Get-SPOExceptionDetails {
    param([System.Exception]$Exception)
    $messages = New-Object 'System.Collections.Generic.List[string]'
    while ($null -ne $Exception) {
        $messages.Add(("{0}: {1}" -f $Exception.GetType().FullName, ($Exception.Message -replace '\s+', ' ').Trim()))
        $Exception = $Exception.InnerException
    }
    return ($messages -join ' --> ')
}

function Get-SPORoleAssignmentIdentity {
    param($RoleAssignment)
    $principalId = '<unavailable>'
    $principalName = '<unavailable>'
    # These getters inspect loaded values only; missing CSOM properties may throw.
    try { $principalId = [string]$RoleAssignment.PrincipalId } catch { }
    try {
        $member = $RoleAssignment.Member
        if ($null -ne $member) {
            $principalId = [string]$member.Id
            $principalName = [string]$member.LoginName
            if ([string]::IsNullOrWhiteSpace($principalName)) { $principalName = [string]$member.Title }
        }
    } catch { }
    return ("principal ID={0}; name='{1}'" -f $principalId, ($principalName -replace '\s+', ' '))
}

function Invoke-SPORead {
    param([scriptblock]$Operation, [string]$Label, [int]$Attempts = 3)
    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            # Buffer each attempt so a partial response is never emitted twice.
            $result = @(& $Operation)
            return $result
        }
        catch {
            $message = $_.Exception.ToString()
            $permanent = $message -match '(?i)access (is )?denied|unauthori[sz]ed|forbidden|\((401|403)\)|\b(HTTP|status( code)?)[: =]+(401|403)\b|item does not exist|cannot access a disposed object'
            $transient = -not $permanent -and $message -match '(?i)HttpClient.Timeout|timed? out|timeout|\b429\b|\b503\b|TooManyRequests|temporarily unavailable|connection.*(closed|reset)|error while copying content to a stream|response ended prematurely|unable to read data from the transport connection'
            if (-not $transient -or $attempt -eq $Attempts) { throw }
            $delay = if ($attempt -eq 1) { 5 } else { 15 }
            Write-SPOReadRetry -Message ("Retrying {0} after a transient read failure ({1}/{2}); waiting {3}s: {4}" -f $Label, $attempt, $Attempts, $delay, (Get-SPOExceptionDetails $_.Exception))
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-WebUrlsFromSite {
    param(
        [string]$Url
    )

    $webUrls = New-Object System.Collections.Generic.List[string]
    $pendingWebUrls = New-Object System.Collections.Generic.Queue[string]
    $seenWebUrls = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    $pendingWebUrls.Enqueue($Url)

    while ($pendingWebUrls.Count -gt 0) {
        $currentWebUrl = $pendingWebUrls.Dequeue()
        if (-not $seenWebUrls.Add($currentWebUrl)) {
            continue
        }

        try {
            Connect-ToSPOWeb -Url $currentWebUrl
            $web = Invoke-SPORead -Label 'web' -Operation { Get-PnPWeb -Includes Title,Url,ServerRelativeUrl -ErrorAction Stop -Connection $script:SPOPermissionConnection }
            $webUrls.Add($web.Url)

            try {
                $subWebs = @(Invoke-SPORead -Label 'subsites' -Operation { Get-PnPSubWeb -Includes Title,Url,ServerRelativeUrl -ErrorAction Stop -Connection $script:SPOPermissionConnection })
                Write-ConsoleMessage -Message ("  Subsites found under {0}: {1}" -f $web.Url, $subWebs.Count)

                foreach ($subWeb in $subWebs) {
                    if ($subWeb.Url -and -not $seenWebUrls.Contains($subWeb.Url)) {
                        $pendingWebUrls.Enqueue($subWeb.Url)
                    }
                }
            }
            catch {
                Write-ConsoleWarning -Message ("Failed to enumerate immediate subsites for web '{0}': {1}" -f $web.Url, $_.Exception.Message)
                Write-InventoryError -Scope 'SubsiteEnumeration' -Url $web.Url -Name $web.Title -Message $_.Exception.Message
            }
        }
        catch {
            Write-ConsoleWarning -Message ("Failed to connect to or read web '{0}': {1}" -f $currentWebUrl, $_.Exception.Message)
            Write-InventoryError -Scope 'Web' -Url $currentWebUrl -Name $currentWebUrl -Message $_.Exception.Message
        }
    }

    $webUrls
}

Import-RequiredModule

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Get-DefaultOutputPath
}

if ([string]::IsNullOrWhiteSpace($ErrorPath)) {
    $errorBaseDirectory = Split-Path -Path $OutputPath -Parent
    $ErrorPath = Join-Path -Path $errorBaseDirectory -ChildPath ("{0}-Errors.csv" -f [System.IO.Path]::GetFileNameWithoutExtension($OutputPath))
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $logBaseDirectory = Split-Path -Path $OutputPath -Parent
    $LogPath = Join-Path -Path (Join-Path -Path $logBaseDirectory -ChildPath 'logs') -ChildPath ("{0}-Run.log" -f [System.IO.Path]::GetFileNameWithoutExtension($OutputPath))
}

$outputDirectory = Split-Path -Path $OutputPath -Parent
if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

$TempOutputPath = "{0}.tmp" -f $OutputPath

foreach ($path in @($ErrorPath, $LogPath)) {
    $directory = Split-Path -Path $path -Parent
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -Path $directory -ItemType Directory -Force | Out-Null
    }
}

foreach ($path in @($OutputPath, $TempOutputPath, $ErrorPath, $LogPath)) {
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

$script:ErrorPath = $ErrorPath
$script:CsvCreated = $false
$script:ErrorCsvCreated = $false
$script:TranscriptStarted = $false
$script:InventoryStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$script:ForceAuthenticationUsed = $false

try {
    Start-Transcript -Path $LogPath -Force -WhatIf:$false | Out-Null
    $script:TranscriptStarted = $true
    Write-ConsoleMessage -Message ("Permission inventory output: {0}" -f $OutputPath)
    Write-ConsoleMessage -Message ("Temporary permission inventory output: {0}" -f $TempOutputPath)
    Write-ConsoleMessage -Message ("Error output: {0}" -f $ErrorPath)
    Write-ConsoleMessage -Message ("Run log: {0}" -f $LogPath)
    $permissionScope = if ($DocumentLibrariesOnly) { 'DocumentLibrariesOnly' } else { 'AllListsAndLibraries' }
    $authMode = if ($Interactive) { 'Interactive' } elseif ($DeviceLogin) { 'DeviceLogin' } elseif (-not [string]::IsNullOrWhiteSpace($Thumbprint)) { 'Certificate' } else { 'Interactive' }
    $inputScope = if ($PSCmdlet.ParameterSetName -eq 'Site') { $SiteUrl } else { $WebUrlsFile }
    Write-ConsoleMessage -Message ("Permission scan options: Scope={0}; IncludeItemPermissions={1}; IncludeHiddenLists={2}; IncludeSystemLists={3}; PageSize={4}; ItemProgressInterval={5}; AuthMode={6}; ForceAuthentication={7}; ParameterSet={8}; Input={9}" -f $permissionScope, [bool]$IncludeItemPermissions, [bool]$IncludeHiddenLists, [bool]$IncludeSystemLists, $PageSize, $ItemProgressInterval, $authMode, [bool]$ForceAuthentication, $PSCmdlet.ParameterSetName, $inputScope)
}
catch {
    Write-ConsoleWarning -Message ("Could not start transcript log '{0}': {1}" -f $LogPath, $_.Exception.Message)
}

try {
    $rootWebUrls = @(if ($PSCmdlet.ParameterSetName -eq 'Site') { $SiteUrl } else { Get-WebUrlsFromFile -Path $WebUrlsFile })
    Write-ConsoleMessage -Message ("Root web URLs loaded: {0}; descendant subsites are included." -f $rootWebUrls.Count)

    $expandedWebUrls = New-Object System.Collections.Generic.List[string]
    $seenExpandedWebUrls = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($rootWebUrl in $rootWebUrls) {
        $rootExpandedWebUrls = @(Get-WebUrlsFromSite -Url $rootWebUrl)
        Write-ConsoleMessage -Message ("Web URLs expanded from root '{0}': {1}" -f $rootWebUrl, $rootExpandedWebUrls.Count)
        foreach ($expandedWebUrl in $rootExpandedWebUrls) {
            if ($seenExpandedWebUrls.Add($expandedWebUrl)) {
                $expandedWebUrls.Add($expandedWebUrl)
            }
        }
    }

    $webUrls = @($expandedWebUrls)
    Write-ConsoleMessage -Message ("Web URLs to inventory after expansion: {0}" -f $webUrls.Count)

    foreach ($webUrl in $webUrls) {
        try {
            Export-WebPermissionInventory -WebUrl $webUrl -CsvPath $TempOutputPath
        }
        catch {
            Write-ConsoleWarning -Message ("Failed to inventory web permissions '{0}': {1}" -f $webUrl, $_.Exception.Message)
            Write-InventoryError -Scope 'Web' -Url $webUrl -Name $webUrl -Message $_.Exception.Message
        }
    }

    if (-not $script:CsvCreated) {
        Write-PermissionHeaderOnlyCsv -Path $TempOutputPath
        Write-ConsoleWarning -Message "No permission rows were exported. The final permission CSV contains headers only."
    }

    if ($script:ErrorCsvCreated) {
        throw ("Permission inventory errors were recorded. Final CSV was not published. Error details: {0}" -f $ErrorPath)
    }

    Move-Item -LiteralPath $TempOutputPath -Destination $OutputPath -Force
    $script:InventoryStopwatch.Stop()
    Write-ConsoleMessage -Message ("Permission inventory completed: {0}" -f $OutputPath)
    Write-ConsoleMessage -Message ("Scan duration: {0}" -f (Format-InventoryDuration -Elapsed $script:InventoryStopwatch.Elapsed))
}
catch {
    if ($script:InventoryStopwatch -and $script:InventoryStopwatch.IsRunning) {
        $script:InventoryStopwatch.Stop()
    }
    Write-ConsoleMessage -Message ("ERROR: {0}" -f $_) -ForegroundColor Red
    if ($script:InventoryStopwatch) {
        Write-ConsoleWarning -Message ("Scan duration before failure: {0}" -f (Format-InventoryDuration -Elapsed $script:InventoryStopwatch.Elapsed))
    }
    if (Test-Path -LiteralPath $TempOutputPath) {
        Write-ConsoleWarning -Message ("Permission inventory failed before final CSV publication. Partial temporary CSV kept: {0}" -f $TempOutputPath)
    }
    if (Test-Path -LiteralPath $OutputPath) {
        Remove-Item -LiteralPath $OutputPath -Force
        Write-ConsoleWarning -Message ("Removed incomplete final permission CSV: {0}" -f $OutputPath)
    }
    $script:ConsoleLifecycleStatus = 'FAILED'
    exit 1
}
finally {
    if ($script:TranscriptStarted) {
        Stop-TimestampedTranscript -Path $LogPath
    }
}
$script:ConsoleLifecycleStatus = 'SUCCESS'
}
catch {
    $script:ConsoleLifecycleFailure = $_
    throw
}
finally {
    Complete-SmartM365MigrationConsoleLifecycle -Context $script:ConsoleLifecycleContext -Failure $script:ConsoleLifecycleFailure -Status $script:ConsoleLifecycleStatus
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCARCDnU5/ZH9mu
# AWlXZ/SUVP1caVBSXSIEQfX+lUhjcqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJXDGLmzDKVzE9AW0Uw+ofWipQkUq2op99x8FHntce1xMA0GCSqG
# SIb3DQEBAQUABIIBgGTgsRraLolUtHVLj5iuyclFkbCgTQleROTyyLYsFuTRQzqe
# 1EpAXmEta8XAUf+QNmob2TPJMpRoKAM4LT0MDsL8Di0P8rNiuJp6/4w6HD9+DznG
# c9/nXYc8Vr0I4A6xhPiW9lXMcKx5u08JacanRS81WllR2mRUbSlnlbWabbkdc5xJ
# K1GU/6pJs4qdGU5dMo5Qd3TkmG9BW+DFt/PMDtwiROEtIYATeJ6JeBIfYqra70Uf
# DL5b7zMyZ2QtEU/Vpm/ttPSdDyetCx+fpPhPkn7Agieh+ePaEzlghn3+f+XYlz1g
# Ymw9+f+hNFkP/t50J+31oYLmYPSrY6s2gfJAE1BBjDA6AwRUcsRNBcq8VG61cNmS
# FkjZC+LrumKmfcWQNVff1Q1x515pwzJ5+MR8Hkg4qD2Pi451WIs0dTf/HnFYQJ7g
# QZ2Ip2U9+MwT7QOZGyzivW4JrHi0r1VBkF3tFyJegEt7LrwnDjkBHvDLSwGsgTsc
# X1ySMe06zZM1gfGzyKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDgxMzEw
# MjBaMC8GCSqGSIb3DQEJBDEiBCDSKJQxgdXOmki56XBAFJHtZkZYui1feekGGXDb
# 1O9GpDANBgkqhkiG9w0BAQEFAASCAgBHWRtlK/Ss2DB6qi5S7VuuG7Ho7dJGk526
# MrqU/cWT/2zQDcNj5jhfS+BQ0triPyj9GBgFfLjNc5oGjHMmsqZ9X339pWMp18jr
# W75aqTkBG4ToYlZuYzGyYJSTT0LRXjrFw1JJjz4spjY2O3UTf9PrBuk4PkQq/ggQ
# 7ed5CRGrksymlLCH1l8iBsD1TXogc/U6Mki2jF/e+ulbWY526+FFlpUh3Ulbi1cY
# YQiK6+oNPUdN5Yy354zDBnwjpSD29BZPUNSuXJ3/CysdnsVx/lSST3r1FbKxkWn0
# oYXgdVj2lraYD0J0hI6fV8qKR+pI7nxhYpez7nS5MDEyUo1n2zpPttt1GA1TGprk
# q2zB95AiGboOs2b9Zb72ddz8ZUOYeQqsSJNOvV/PM458iBhvxIodblYUyDunHc+I
# OWhe9Tlp4I/Y30sqW1kbR7HIaMVvgrAatW5b24Lley4xLTCQP2hyJZpCHH6wNhKH
# Vtujngzb7X5YG/rwtKiZxWi5A5eMedsY7fkAqGGgPVefT7iy2I+wmShraoVnn7bJ
# iLQOkPdi1ShBmTdTdxrLaYfWa5+XWDVO3kHx3J7zOuMBHlvPzxqHNfsYPrc/gBGn
# SX4QF7rF7Eq1d0CddG/5fSym7WxgxnXIzIqrIZVM5/LHAGI0kfvF+KyQiyMp+uu2
# ACaEEsJ2PA==
# SIG # End signature block
