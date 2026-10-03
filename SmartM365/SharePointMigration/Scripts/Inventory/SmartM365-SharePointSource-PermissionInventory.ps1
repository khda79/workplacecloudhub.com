<#
.SYNOPSIS
    Inventories SharePoint 2019 permissions to a dedicated CSV file.

.DESCRIPTION
    Run this script on a SharePoint 2019 server from the SharePoint Management
    Shell, or from a PowerShell console started with farm administrator rights.

    The script exports role assignments for webs and lists/libraries. Item and
    folder permissions are optional because they can create very large outputs.
    When -IncludeItemPermissions is used, only items and folders with unique
    When -UseSiteUrlFilter/-SiteUrlsFile is used, each listed site URL is
    treated as a root and its descendant subsites are included by default.

    permissions are exported.

.EXAMPLE
    .\SmartM365-SharePointSource-PermissionInventory.ps1 -WebApplicationUrl "https://intranet"

.EXAMPLE
    .\SmartM365-SharePointSource-PermissionInventory.ps1 -WebApplicationUrl "https://intranet" -OutputPath "C:\Temp\SP2019PermissionScans"

.EXAMPLE
    .\SmartM365-SharePointSource-PermissionInventory.ps1 -WebApplicationUrl "https://intranet" -UseSiteUrlFilter -SiteUrlsFile "C:\Temp\sp2019-site-urls.txt"

.EXAMPLE
    .\SmartM365-SharePointSource-PermissionInventory.ps1 -WebUrl "https://intranet/sites/finance" -DocumentLibrariesOnly -IncludeItemPermissions

.VERSION
    1.1.1
#>

[CmdletBinding(DefaultParameterSetName = 'WebApplication')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'WebApplication')]
    [ValidateNotNullOrEmpty()]
    [string]$WebApplicationUrl,

    [Parameter(Mandatory = $true, ParameterSetName = 'Site')]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true, ParameterSetName = 'Web')]
    [ValidateNotNullOrEmpty()]
    [string]$WebUrl,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath,

    [ValidateNotNullOrEmpty()]
    [string]$ErrorPath,

    [ValidateNotNullOrEmpty()]
    [string]$LogPath,

    [ValidateNotNullOrEmpty()]
    [string]$SiteUrlsFile,

    [switch]$UseSiteUrlFilter,

    [switch]$IncludeHiddenLists,

    [switch]$IncludeSystemLists,

    [switch]$DocumentLibrariesOnly,

    [switch]$IncludeItemPermissions = $true,

    [int]$RowLimit = 2000,

    [ValidateRange(1, [int]::MaxValue)]
    [int]$ItemProgressInterval = 500
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Launchers\SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
$script:ConsoleLifecycleContext = Start-SmartM365MigrationConsoleLifecycle -ScriptPath $PSCommandPath
$script:ConsoleLifecycleFailure = $null
$script:ConsoleLifecycleStatus = 'SUCCESS'
try {

function Get-ConsoleTimestamp {
    return (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
}

function Write-Host {
    param(
        [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
        [object[]]$Object,

        [ConsoleColor]$ForegroundColor,

        [switch]$NoNewline
    )

    $message = if ($Object) { ($Object -join ' ') } else { '' }
    $line = if ($message -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} ') { $message } else { "{0} {1}" -f (Get-ConsoleTimestamp), $message }
    $parameters = @{ Object = $line }
    if ($PSBoundParameters.ContainsKey('ForegroundColor')) { $parameters.ForegroundColor = $ForegroundColor }
    if ($NoNewline) { $parameters.NoNewline = $true }
    Microsoft.PowerShell.Utility\Write-Host @parameters
}

function Write-Warning {
    param(
        [Parameter(Position = 0)]
        [string]$Message
    )

    $line = if ($Message -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} ') { $Message } else { "{0} WARNING: {1}" -f (Get-ConsoleTimestamp), $Message }
    Microsoft.PowerShell.Utility\Write-Host $line -ForegroundColor Yellow
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

function Add-SharePointSnapIn {
    if (-not (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
        Add-PSSnapin Microsoft.SharePoint.PowerShell
    }
}

function ConvertTo-SafeFileName {
    param(
        [string]$Name,
        [string]$Fallback = 'SharePoint2019'
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

    Write-Host ("{0}   Item heartbeat: web='{1}' library='{2}' processed={3}; unique={4}; permission rows exported={5}; elapsed={6}; last='{7}'" -f `
            (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), `
            $WebUrl, `
            $ListTitle, `
            $ProcessedItems, `
            $UniquePermissionItems, `
            $ExportedRows, `
            (Format-InventoryDuration -Elapsed $Elapsed), `
            $LastItem) -ForegroundColor DarkCyan
}

function Format-InventoryDuration {
    param(
        [TimeSpan]$Elapsed
    )

    if ($Elapsed.Days -gt 0) {
        return ("{0}d {1:00}:{2:00}:{3:00}" -f $Elapsed.Days, $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds)
    }

    return ("{0:00}:{1:00}:{2:00}" -f [int]$Elapsed.TotalHours, $Elapsed.Minutes, $Elapsed.Seconds)
}

function Write-PermissionHeaderOnlyCsv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $header = ($PermissionColumns | ForEach-Object { '"{0}"' -f ($_ -replace '"', '""') }) -join ';'
    Set-Content -LiteralPath $Path -Value $header -Encoding UTF8
}

function Get-NameFromUrl {
    param(
        [string]$Url
    )

    try {
        $uri = New-Object System.Uri($Url)
        $path = $uri.AbsolutePath.Trim('/')
        if ($path) {
            $segments = @($path.Split('/') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($segments.Count -gt 0) {
                return $segments[$segments.Count - 1]
            }
        }

        return $uri.Host
    }
    catch {
        return $Url
    }
}

function Get-NormalizedUrlPath {
    param(
        [string]$Url
    )

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return $null
    }

    $cleanUrl = $Url.Trim()
    try {
        $path = (New-Object System.Uri($cleanUrl)).AbsolutePath
    }
    catch {
        $path = $cleanUrl
    }

    $path = ($path -replace '[?#].*$', '').Trim()
    if (-not $path.StartsWith('/')) {
        $path = "/$path"
    }

    $path = $path.TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($path)) {
        $path = '/'
    }

    return $path.ToLowerInvariant()
}

function Import-SiteUrlFilter {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Site URL filter file not found: $Path"
    }

    $filterPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    Get-Content -LiteralPath $Path |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -and -not $_.StartsWith('#') } |
        ForEach-Object {
            $normalizedPath = Get-NormalizedUrlPath -Url $_
            if ($normalizedPath) {
                [void]$filterPaths.Add($normalizedPath)
            }
        }

    return $filterPaths
}

function Test-PathMatchesSiteUrlFilter {
    param(
        [string]$WebPath
    )

    if ([string]::IsNullOrWhiteSpace($WebPath)) {
        return $false
    }

    foreach ($filterPath in $script:SiteUrlFilterPaths) {
        $normalizedFilterPath = ([string]$filterPath).TrimEnd('/')
        if ([string]::IsNullOrWhiteSpace($normalizedFilterPath)) {
            $normalizedFilterPath = '/'
        }

        if ($normalizedFilterPath -eq '/') {
            return $true
        }

        if ($WebPath.Equals($normalizedFilterPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }

        if ($WebPath.StartsWith(($normalizedFilterPath + '/'), [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Test-WebMatchesSiteUrlFilter {
    param(
        [Microsoft.SharePoint.SPWeb]$Web
    )

    if (-not $script:UseSiteUrlFilter) {
        return $true
    }

    $webPath = Get-NormalizedUrlPath -Url $Web.Url
    return (Test-PathMatchesSiteUrlFilter -WebPath $webPath)
}

function Test-SystemList {
    param(
        [Microsoft.SharePoint.SPList]$List
    )

    $systemUrls = @(
        '_catalogs/masterpage',
        '_catalogs/wp',
        '_catalogs/lt',
        'Style Library',
        'SiteAssets',
        'SitePages',
        'FormServerTemplates'
    )

    foreach ($url in $systemUrls) {
        if ($List.RootFolder.Url -like "$url*" -or $List.RootFolder.ServerRelativeUrl -like "*/$url*") {
            return $true
        }
    }

    return $false
}

function ConvertTo-AbsoluteSharePointUrl {
    param(
        [Microsoft.SharePoint.SPWeb]$Web,
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

    $webUri = New-Object System.Uri($Web.Url)
    return ("{0}://{1}{2}" -f $webUri.Scheme, $webUri.Authority, $ServerRelativeUrl)
}

function Get-PrincipalInfo {
    param(
        $Member
    )

    $principalType = $Member.GetType().Name
    $name = $Member.Name
    $loginName = $null
    $id = $null

    if ($Member -is [Microsoft.SharePoint.SPGroup]) {
        $principalType = 'SharePointGroup'
        $loginName = $Member.LoginName
        $id = $Member.ID
    }
    elseif ($Member -is [Microsoft.SharePoint.SPUser]) {
        if ($Member.IsDomainGroup) {
            $principalType = 'DomainGroup'
        }
        else {
            $principalType = 'User'
        }

        $loginName = $Member.LoginName
        $id = $Member.ID
    }

    [pscustomobject]@{
        PrincipalType      = $principalType
        PrincipalName      = $name
        PrincipalLoginName = $loginName
        PrincipalId        = $id
    }
}

function Get-AssociatedWebGroupNames {
    param(
        [Microsoft.SharePoint.SPWeb]$Web
    )

    $memberGroup = ''
    $ownerGroup = ''
    $visitorGroup = ''

    try {
        if ($null -ne $Web.AssociatedMemberGroup) {
            $memberGroup = [string]$Web.AssociatedMemberGroup.Name
        }
    }
    catch {}

    try {
        if ($null -ne $Web.AssociatedOwnerGroup) {
            $ownerGroup = [string]$Web.AssociatedOwnerGroup.Name
        }
    }
    catch {}

    try {
        if ($null -ne $Web.AssociatedVisitorGroup) {
            $visitorGroup = [string]$Web.AssociatedVisitorGroup.Name
        }
    }
    catch {}

    [pscustomobject]@{
        AssociatedMemberGroup  = $memberGroup
        AssociatedOwnerGroup   = $ownerGroup
        AssociatedVisitorGroup = $visitorGroup
    }
}

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
        $Member
    )

    if (-not ($Member -is [Microsoft.SharePoint.SPGroup])) {
        return Get-EmptyPrincipalMembershipInfo
    }

    $cacheKey = if ($null -ne $Member.ID) { [string]$Member.ID } else { [string]$Member.LoginName }
    if ($script:SharePointGroupMembershipCache.ContainsKey($cacheKey)) {
        return $script:SharePointGroupMembershipCache[$cacheKey]
    }

    try {
        $members = @($Member.Users)
        $loginNames = @($members | ForEach-Object { $_.LoginName })
        $displayNames = @($members | ForEach-Object { $_.Name })
        $domainGroupMembers = @($members | Where-Object { $_.IsDomainGroup })
        $userMembers = @($members | Where-Object { -not $_.IsDomainGroup })

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

    foreach ($roleAssignment in $RoleAssignments) {
        $permissionLevels = @($roleAssignment.RoleDefinitionBindings | ForEach-Object { $_.Name })
        if ($permissionLevels.Count -eq 0) {
            continue
        }

        $principal = Get-PrincipalInfo -Member $roleAssignment.Member
        $membership = Get-PrincipalMembershipInfo -Member $roleAssignment.Member
        [pscustomobject]@{
            SiteCollectionUrl       = $SiteCollectionUrl
            WebUrl                  = $WebUrl
            WebTitle                = $WebTitle
            AssociatedMemberGroup   = $AssociatedMemberGroup
            AssociatedOwnerGroup    = $AssociatedOwnerGroup
            AssociatedVisitorGroup  = $AssociatedVisitorGroup
            ObjectScope             = $ObjectScope
            ObjectUrl               = $ObjectUrl
            ObjectServerRelativeUrl = $ObjectServerRelativeUrl
            ObjectTitle             = $ObjectTitle
            ObjectId                = $ObjectId
            ParentObjectUrl         = $ParentObjectUrl
            ListTitle               = $ListTitle
            ListUrl                 = $ListUrl
            ListBaseTemplate        = $ListBaseTemplate
            ListBaseType            = $ListBaseType
            IsDocumentLibrary       = $IsDocumentLibrary
            ItemId                  = $ItemId
            ItemFileSystemObjectType = $ItemFileSystemObjectType
            HasUniqueRoleAssignments = $HasUniqueRoleAssignments
            InheritedFrom           = $InheritedFrom
            PrincipalType           = $principal.PrincipalType
            PrincipalName           = $principal.PrincipalName
            PrincipalLoginName      = $principal.PrincipalLoginName
            PrincipalId             = $principal.PrincipalId
            PrincipalMemberCount        = $membership.PrincipalMemberCount
            PrincipalUserMemberCount    = $membership.PrincipalUserMemberCount
            PrincipalDomainGroupMemberCount = $membership.PrincipalDomainGroupMemberCount
            PrincipalMemberLoginNames   = $membership.PrincipalMemberLoginNames
            PrincipalMemberDisplayNames = $membership.PrincipalMemberDisplayNames
            PrincipalMemberLookupStatus = $membership.PrincipalMemberLookupStatus
            PermissionLevels        = ($permissionLevels -join '|')
            IsLimitedAccessOnly     = ($permissionLevels.Count -eq 1 -and $permissionLevels[0] -eq 'Limited Access')
        }
    }
}

function Export-PermissionRows {
    param(
        [object[]]$Rows,
        [string]$CsvPath,
        [switch]$Append
    )

    if (-not $Rows -or $Rows.Count -eq 0) {
        return
    }

    if ($Append -or $script:CsvCreated) {
        $Rows | Export-Csv -Delimiter ';' -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Append
    }
    else {
        $Rows | Export-Csv -Delimiter ';' -Path $CsvPath -NoTypeInformation -Encoding UTF8
        $script:CsvCreated = $true
    }
}

function Get-ItemServerRelativeUrl {
    param(
        [Microsoft.SharePoint.SPListItem]$Item
    )

    if ($Item.File -and $Item.File.ServerRelativeUrl) {
        return $Item.File.ServerRelativeUrl
    }

    if ($Item.Folder -and $Item.Folder.ServerRelativeUrl) {
        return $Item.Folder.ServerRelativeUrl
    }

    if ($Item.Url) {
        return $Item.Url
    }

    return $null
}

function Export-ItemPermissionInventory {
    param(
        [Microsoft.SharePoint.SPWeb]$Web,
        [Microsoft.SharePoint.SPList]$List,
        [string]$CsvPath,
        [int]$BatchSize,
        [int]$ProgressInterval
    )

    $associatedGroups = Get-AssociatedWebGroupNames -Web $Web
    $query = New-Object Microsoft.SharePoint.SPQuery
    $query.ViewAttributes = "Scope='RecursiveAll'"
    $query.RowLimit = [uint32]$BatchSize

    $processedItems = 0
    $uniquePermissionItems = 0
    $exportedRows = 0
    $lastItem = '<none>'
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    do {
        try {
            $items = $List.GetItems($query)
            $query.ListItemCollectionPosition = $items.ListItemCollectionPosition
        }
        catch {
            $listUrl = ConvertTo-AbsoluteSharePointUrl -Web $Web -ServerRelativeUrl $List.RootFolder.ServerRelativeUrl
            Write-Warning ("Failed to enumerate items for list '{0}' in web '{1}': {2}" -f $List.Title, $Web.Url, $_.Exception.Message)
            Stop-InventoryAfterError -Scope 'ListItems' -Url $listUrl -Name $List.Title -Message $_.Exception.Message
        }

        foreach ($item in $items) {
            $processedItems++
            try {
                $lastItem = if ($item.ID) { "ID $($item.ID)" } else { '<unknown>' }
                if ($processedItems % $ProgressInterval -eq 0) {
                    Write-ItemPermissionHeartbeat `
                        -WebUrl $Web.Url `
                        -ListTitle $List.Title `
                        -ProcessedItems $processedItems `
                        -UniquePermissionItems $uniquePermissionItems `
                        -ExportedRows $exportedRows `
                        -Elapsed $stopwatch.Elapsed `
                        -LastItem $lastItem
                }

                if (-not $item.HasUniqueRoleAssignments) {
                    continue
                }

                $uniquePermissionItems++
                $serverRelativeUrl = Get-ItemServerRelativeUrl -Item $item
                $absoluteUrl = ConvertTo-AbsoluteSharePointUrl -Web $Web -ServerRelativeUrl $serverRelativeUrl
                $fsObjType = if ([string]$item['FSObjType'] -eq '1') { 'Folder' } else { 'FileOrItem' }
                $listUrl = ConvertTo-AbsoluteSharePointUrl -Web $Web -ServerRelativeUrl $List.RootFolder.ServerRelativeUrl
                $rows = @(Get-RoleAssignmentRows `
                        -SiteCollectionUrl $Web.Site.Url `
                        -WebUrl $Web.Url `
                        -WebTitle $Web.Title `
                        -AssociatedMemberGroup $associatedGroups.AssociatedMemberGroup `
                        -AssociatedOwnerGroup $associatedGroups.AssociatedOwnerGroup `
                        -AssociatedVisitorGroup $associatedGroups.AssociatedVisitorGroup `
                        -ObjectScope 'Item' `
                        -ObjectUrl $absoluteUrl `
                        -ObjectServerRelativeUrl $serverRelativeUrl `
                        -ObjectTitle ([string]$item.Name) `
                        -ObjectId ([string]$item.UniqueId) `
                        -ParentObjectUrl $listUrl `
                        -ListTitle $List.Title `
                        -ListUrl $listUrl `
                        -ListBaseTemplate ([int]$List.BaseTemplate) `
                        -ListBaseType ([string]$List.BaseType) `
                        -IsDocumentLibrary ($List.BaseType -eq [Microsoft.SharePoint.SPBaseType]::DocumentLibrary) `
                        -ItemId $item.ID `
                        -ItemFileSystemObjectType $fsObjType `
                        -HasUniqueRoleAssignments $true `
                        -InheritedFrom '' `
                        -RoleAssignments $item.RoleAssignments)

                Export-PermissionRows -Rows $rows -CsvPath $CsvPath
                $exportedRows += $rows.Count
            }
            catch {
                $itemUrl = $null
                $itemName = $null
                try {
                    $itemUrl = ConvertTo-AbsoluteSharePointUrl -Web $Web -ServerRelativeUrl (Get-ItemServerRelativeUrl -Item $item)
                    $itemName = [string]$item.Name
                }
                catch {
                    $itemUrl = $Web.Url
                    $itemName = $List.Title
                }

                Write-Warning ("Failed to inventory item permissions in list '{0}' in web '{1}': {2}" -f $List.Title, $Web.Url, $_.Exception.Message)
                Stop-InventoryAfterError -Scope 'Item' -Url $itemUrl -Name $itemName -Message $_.Exception.Message
            }
        }
    }
    while ($null -ne $query.ListItemCollectionPosition)

    $stopwatch.Stop()
    Write-ItemPermissionHeartbeat `
        -WebUrl $Web.Url `
        -ListTitle $List.Title `
        -ProcessedItems $processedItems `
        -UniquePermissionItems $uniquePermissionItems `
        -ExportedRows $exportedRows `
        -Elapsed $stopwatch.Elapsed `
        -LastItem $lastItem
}

function Export-WebPermissionInventory {
    param(
        [Microsoft.SharePoint.SPWeb]$Web,
        [string]$CsvPath,
        [switch]$IncludeHidden,
        [switch]$IncludeSystem,
        [switch]$LibrariesOnly,
        [switch]$IncludeItems,
        [int]$BatchSize
    )

    Write-Host ("Web: {0}" -f $Web.Url)
    $associatedGroups = Get-AssociatedWebGroupNames -Web $Web

    try {
        $webHasUniqueRoleAssignments = $Web.HasUniqueRoleAssignments
        $webRows = @(Get-RoleAssignmentRows `
                -SiteCollectionUrl $Web.Site.Url `
                -WebUrl $Web.Url `
                -WebTitle $Web.Title `
                -AssociatedMemberGroup $associatedGroups.AssociatedMemberGroup `
                -AssociatedOwnerGroup $associatedGroups.AssociatedOwnerGroup `
                -AssociatedVisitorGroup $associatedGroups.AssociatedVisitorGroup `
                -ObjectScope 'Web' `
                -ObjectUrl $Web.Url `
                -ObjectServerRelativeUrl $Web.ServerRelativeUrl `
                -ObjectTitle $Web.Title `
                -ObjectId ([string]$Web.ID) `
                -ParentObjectUrl '' `
                -ListTitle '' `
                -ListUrl '' `
                -ListBaseTemplate '' `
                -ListBaseType '' `
                -IsDocumentLibrary '' `
                -ItemId $null `
                -ItemFileSystemObjectType '' `
                -HasUniqueRoleAssignments $webHasUniqueRoleAssignments `
                -InheritedFrom $(if ($webHasUniqueRoleAssignments) { '' } else { 'ParentWeb' }) `
                -RoleAssignments $Web.RoleAssignments)
        Export-PermissionRows -Rows $webRows -CsvPath $CsvPath
    }
    catch {
        Write-Warning ("Failed to inventory web permissions '{0}': {1}" -f $Web.Url, $_.Exception.Message)
        Stop-InventoryAfterError -Scope 'Web' -Url $Web.Url -Name $Web.Title -Message $_.Exception.Message
    }

    try {
        $lists = @($Web.Lists)
    }
    catch {
        Write-Warning ("Failed to enumerate lists for web '{0}': {1}" -f $Web.Url, $_.Exception.Message)
        Stop-InventoryAfterError -Scope 'WebLists' -Url $Web.Url -Name $Web.Title -Message $_.Exception.Message
    }

    foreach ($list in $lists) {
        $listTitle = '<unknown>'
        $listUrl = $Web.Url

        try {
            $listTitle = $list.Title
            $listUrl = ConvertTo-AbsoluteSharePointUrl -Web $Web -ServerRelativeUrl $list.RootFolder.ServerRelativeUrl

            if (-not $IncludeHidden -and $list.Hidden) {
                continue
            }

            if (-not $IncludeSystem -and (Test-SystemList -List $list)) {
                continue
            }

            if ($LibrariesOnly -and $list.BaseTemplate -ne [Microsoft.SharePoint.SPListTemplateType]::DocumentLibrary) {
                continue
            }

            $listHasUniqueRoleAssignments = $list.HasUniqueRoleAssignments
            $isDocumentLibrary = ($list.BaseType -eq [Microsoft.SharePoint.SPBaseType]::DocumentLibrary)
            $inheritedFrom = if ($listHasUniqueRoleAssignments) { '' } else { 'Web' }
            $listRows = @(Get-RoleAssignmentRows `
                    -SiteCollectionUrl $Web.Site.Url `
                    -WebUrl $Web.Url `
                    -WebTitle $Web.Title `
                    -AssociatedMemberGroup $associatedGroups.AssociatedMemberGroup `
                    -AssociatedOwnerGroup $associatedGroups.AssociatedOwnerGroup `
                    -AssociatedVisitorGroup $associatedGroups.AssociatedVisitorGroup `
                    -ObjectScope 'List' `
                    -ObjectUrl $listUrl `
                    -ObjectServerRelativeUrl $list.RootFolder.ServerRelativeUrl `
                    -ObjectTitle $list.Title `
                    -ObjectId ([string]$list.ID) `
                    -ParentObjectUrl $Web.Url `
                    -ListTitle $list.Title `
                    -ListUrl $listUrl `
                    -ListBaseTemplate ([int]$list.BaseTemplate) `
                    -ListBaseType ([string]$list.BaseType) `
                    -IsDocumentLibrary $isDocumentLibrary `
                    -ItemId $null `
                    -ItemFileSystemObjectType '' `
                    -HasUniqueRoleAssignments $listHasUniqueRoleAssignments `
                    -InheritedFrom $inheritedFrom `
                    -RoleAssignments $list.RoleAssignments)
            Export-PermissionRows -Rows $listRows -CsvPath $CsvPath

            if ($IncludeItems) {
                Write-Host ("  Item permissions: {0}" -f $list.Title)
                Export-ItemPermissionInventory -Web $Web -List $list -CsvPath $CsvPath -BatchSize $BatchSize -ProgressInterval $ItemProgressInterval
            }
        }
        catch {
            Write-Warning ("Failed to inventory list permissions '{0}' in web '{1}': {2}" -f $listTitle, $Web.Url, $_.Exception.Message)
            Stop-InventoryAfterError -Scope 'List' -Url $listUrl -Name $listTitle -Message $_.Exception.Message
        }
    }
}

function Export-SitePermissionInventory {
    param(
        [Microsoft.SharePoint.SPSite]$Site,
        [string]$CsvPath,
        [switch]$IncludeHidden,
        [switch]$IncludeSystem,
        [switch]$LibrariesOnly,
        [switch]$IncludeItems,
        [int]$BatchSize
    )

    Write-Host ("Site collection: {0}" -f $Site.Url)
    foreach ($web in $Site.AllWebs) {
        try {
            if (Test-WebMatchesSiteUrlFilter -Web $web) {
                Export-WebPermissionInventory `
                    -Web $web `
                    -CsvPath $CsvPath `
                    -IncludeHidden:$IncludeHidden `
                    -IncludeSystem:$IncludeSystem `
                    -LibrariesOnly:$LibrariesOnly `
                    -IncludeItems:$IncludeItems `
                    -BatchSize $BatchSize
            }
            else {
                Write-Host ("Skipping web not in filter: {0}" -f $web.Url)
            }
        }
        finally {
            $web.Dispose()
        }
    }
}

function Get-TargetSafeName {
    param(
        [string]$ParameterSetName,
        [string]$WebApplicationUrl,
        [string]$SiteUrl,
        [string]$WebUrl
    )

    $targetName = $null

    switch ($ParameterSetName) {
        'WebApplication' {
            try {
                $targetName = (Get-SPWebApplication -Identity $WebApplicationUrl).Name
            }
            catch {
                $targetName = Get-NameFromUrl -Url $WebApplicationUrl
            }
        }

        'Site' {
            $targetName = Get-NameFromUrl -Url $SiteUrl
        }

        'Web' {
            $targetName = Get-NameFromUrl -Url $WebUrl
        }
    }

    return (ConvertTo-SafeFileName -Name $targetName)
}

function Get-DefaultOutputPath {
    param(
        [string]$ParameterSetName,
        [string]$WebApplicationUrl,
        [string]$SiteUrl,
        [string]$WebUrl,
        [string]$BaseDirectory
    )

    if ([string]::IsNullOrWhiteSpace($BaseDirectory)) {
        $BaseDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    }

    $safeTargetName = Get-TargetSafeName `
        -ParameterSetName $ParameterSetName `
        -WebApplicationUrl $WebApplicationUrl `
        -SiteUrl $SiteUrl `
        -WebUrl $WebUrl

    if ([string]::IsNullOrWhiteSpace($safeTargetName)) {
        $safeTargetName = 'SharePoint'
    }

    $safeTargetName = ConvertTo-SafeFileName -Name $safeTargetName
    $targetDirectory = Join-Path -Path $BaseDirectory -ChildPath $safeTargetName
    return Join-Path -Path $targetDirectory -ChildPath ("SP2019-PermissionInventory-{0}-{1:yyyyMMdd-HHmmss}.csv" -f $safeTargetName, (Get-Date))
}

function Test-OutputPathIsDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (Test-Path -LiteralPath $Path -PathType Container) {
        return $true
    }

    if ($Path.EndsWith('\') -or $Path.EndsWith('/')) {
        return $true
    }

    $extension = [System.IO.Path]::GetExtension($Path)
    return [string]::IsNullOrWhiteSpace($extension)
}

function Write-InventoryError {
    param(
        [string]$Scope,
        [string]$Url,
        [string]$Name,
        [string]$Message
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
    }

    if ($script:ErrorCsvCreated) {
        $row | Export-Csv -Delimiter ';' -Path $script:ErrorPath -NoTypeInformation -Encoding UTF8 -Append
    }
    else {
        $row | Export-Csv -Delimiter ';' -Path $script:ErrorPath -NoTypeInformation -Encoding UTF8
        $script:ErrorCsvCreated = $true
    }
}

function Stop-InventoryAfterError {
    param(
        [string]$Scope,
        [string]$Url,
        [string]$Name,
        [string]$Message
    )

    Write-InventoryError -Scope $Scope -Url $Url -Name $Name -Message $Message
    throw ("Permission inventory stopped after {0} error at '{1}'. Final CSV was not published. Details: {2}" -f $Scope, $Url, $Message)
}

Add-SharePointSnapIn

$script:UseSiteUrlFilter = $UseSiteUrlFilter -or (-not [string]::IsNullOrWhiteSpace($SiteUrlsFile))
$script:SiteUrlFilterPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

if ($script:UseSiteUrlFilter) {
    if ([string]::IsNullOrWhiteSpace($SiteUrlsFile)) {
        throw "UseSiteUrlFilter requires -SiteUrlsFile."
    }

    $script:SiteUrlFilterPaths = Import-SiteUrlFilter -Path $SiteUrlsFile
    Write-Host ("Site URL filter enabled: {0} root paths loaded from {1}; descendant subsites are included." -f $script:SiteUrlFilterPaths.Count, $SiteUrlsFile)
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Get-DefaultOutputPath `
        -ParameterSetName $PSCmdlet.ParameterSetName `
        -WebApplicationUrl $WebApplicationUrl `
        -SiteUrl $SiteUrl `
        -WebUrl $WebUrl
}
elseif (Test-OutputPathIsDirectory -Path $OutputPath) {
    $OutputPath = Get-DefaultOutputPath `
        -ParameterSetName $PSCmdlet.ParameterSetName `
        -WebApplicationUrl $WebApplicationUrl `
        -SiteUrl $SiteUrl `
        -WebUrl $WebUrl `
        -BaseDirectory $OutputPath
}

if ([string]::IsNullOrWhiteSpace($ErrorPath)) {
    $errorBaseDirectory = Split-Path -Path $OutputPath -Parent
    if ([string]::IsNullOrWhiteSpace($errorBaseDirectory)) {
        $errorBaseDirectory = (Get-Location).Path
    }

    $ErrorPath = Join-Path -Path $errorBaseDirectory -ChildPath ("{0}-Errors.csv" -f [System.IO.Path]::GetFileNameWithoutExtension($OutputPath))
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $logBaseDirectory = Split-Path -Path $OutputPath -Parent
    if ([string]::IsNullOrWhiteSpace($logBaseDirectory)) {
        $logBaseDirectory = (Get-Location).Path
    }

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

try {
    Start-Transcript -Path $LogPath -Force -WhatIf:$false | Out-Null
    $script:TranscriptStarted = $true
    Write-Host ("Permission inventory output: {0}" -f $OutputPath)
    Write-Host ("Temporary permission inventory output: {0}" -f $TempOutputPath)
    Write-Host ("Error output: {0}" -f $ErrorPath)
    Write-Host ("Run log: {0}" -f $LogPath)
    $permissionScope = if ($DocumentLibrariesOnly) { 'DocumentLibrariesOnly' } else { 'AllListsAndLibraries' }
    $siteUrlsFileLabel = if ([string]::IsNullOrWhiteSpace($SiteUrlsFile)) { '<none>' } else { $SiteUrlsFile }
    Write-Host ("Permission scan options: Scope={0}; IncludeItemPermissions={1}; IncludeHiddenLists={2}; IncludeSystemLists={3}; RowLimit={4}; ItemProgressInterval={5}; ParameterSet={6}; UseSiteUrlFilter={7}; SiteUrlsFile={8}" -f $permissionScope, [bool]$IncludeItemPermissions, [bool]$IncludeHiddenLists, [bool]$IncludeSystemLists, $RowLimit, $ItemProgressInterval, $PSCmdlet.ParameterSetName, [bool]$script:UseSiteUrlFilter, $siteUrlsFileLabel)
}
catch {
    Write-Warning ("Could not start transcript log '{0}': {1}" -f $LogPath, $_.Exception.Message)
}

try {
    switch ($PSCmdlet.ParameterSetName) {
        'WebApplication' {
            $sites = Get-SPSite -WebApplication $WebApplicationUrl -Limit All
            foreach ($site in $sites) {
                try {
                    Export-SitePermissionInventory `
                        -Site $site `
                        -CsvPath $TempOutputPath `
                        -IncludeHidden:$IncludeHiddenLists `
                        -IncludeSystem:$IncludeSystemLists `
                        -LibrariesOnly:$DocumentLibrariesOnly `
                        -IncludeItems:$IncludeItemPermissions `
                        -BatchSize $RowLimit
                }
                catch {
                    Write-Warning ("Failed to inventory site collection permissions '{0}': {1}" -f $site.Url, $_.Exception.Message)
                    Stop-InventoryAfterError -Scope 'SiteCollection' -Url $site.Url -Name $site.Url -Message $_.Exception.Message
                }
                finally {
                    $site.Dispose()
                }
            }
        }

        'Site' {
            $site = Get-SPSite -Identity $SiteUrl
            try {
                Export-SitePermissionInventory `
                    -Site $site `
                    -CsvPath $TempOutputPath `
                    -IncludeHidden:$IncludeHiddenLists `
                    -IncludeSystem:$IncludeSystemLists `
                    -LibrariesOnly:$DocumentLibrariesOnly `
                    -IncludeItems:$IncludeItemPermissions `
                    -BatchSize $RowLimit
            }
            finally {
                $site.Dispose()
            }
        }

        'Web' {
            $site = New-Object Microsoft.SharePoint.SPSite($WebUrl)
            try {
                $web = $site.OpenWeb()
                try {
                    if (Test-WebMatchesSiteUrlFilter -Web $web) {
                        Export-WebPermissionInventory `
                            -Web $web `
                            -CsvPath $TempOutputPath `
                            -IncludeHidden:$IncludeHiddenLists `
                            -IncludeSystem:$IncludeSystemLists `
                            -LibrariesOnly:$DocumentLibrariesOnly `
                            -IncludeItems:$IncludeItemPermissions `
                            -BatchSize $RowLimit
                    }
                    else {
                        Write-Host ("Skipping web not in filter: {0}" -f $web.Url)
                    }
                }
                finally {
                    $web.Dispose()
                }
            }
            finally {
                $site.Dispose()
            }
        }
    }

    if (-not $script:CsvCreated) {
        Write-PermissionHeaderOnlyCsv -Path $TempOutputPath
        Write-Warning "No permission rows were exported. The final permission CSV contains headers only."
    }

    if ($script:ErrorCsvCreated) {
        throw ("Permission inventory errors were recorded. Final CSV was not published. Error details: {0}" -f $ErrorPath)
    }

    Move-Item -LiteralPath $TempOutputPath -Destination $OutputPath -Force
    $script:InventoryStopwatch.Stop()
    Write-Host ("Permission inventory completed: {0}" -f $OutputPath)
    Write-Host ("Scan duration: {0}" -f (Format-InventoryDuration -Elapsed $script:InventoryStopwatch.Elapsed))
    if ($script:ErrorCsvCreated) {
        Write-Warning ("Some permissions could not be inventoried. Error details: {0}" -f $ErrorPath)
    }
}
catch {
    if ($script:InventoryStopwatch -and $script:InventoryStopwatch.IsRunning) {
        $script:InventoryStopwatch.Stop()
    }
    Write-Error $_
    if ($script:InventoryStopwatch) {
        Write-Warning ("Scan duration before failure: {0}" -f (Format-InventoryDuration -Elapsed $script:InventoryStopwatch.Elapsed))
    }
    if (Test-Path -LiteralPath $TempOutputPath) {
        Write-Warning ("Permission inventory failed before final CSV publication. Partial temporary CSV kept: {0}" -f $TempOutputPath)
    }
    if (Test-Path -LiteralPath $OutputPath) {
        Remove-Item -LiteralPath $OutputPath -Force
        Write-Warning ("Removed incomplete final permission CSV: {0}" -f $OutputPath)
    }
    $script:ConsoleLifecycleStatus = 'FAILED'
    exit 1
}
finally {
    if ($script:TranscriptStarted) {
        Stop-TimestampedTranscript -Path $LogPath
    }
}
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCevGMn+1D5KDZf
# JBtOnwSXlylYZNkw1Llk11uABqWyOqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGrzHKbdblsOAKHpfEV+WNee6EsK3uvyLoxvtePPfLgwMA0GCSqG
# SIb3DQEBAQUABIIBgAR736xz3RLzdzbHAxL6kKm+I6sAb8id6ma9Uf908usWEH6q
# NO5KLL/2BHjDkGMhd1kSJ0cR0RLtO8yhRHvChFJf2Wz9AkhRMaSRsxuYTgTlD0IV
# nHtqL71ZLwXUEPYOOcXW/3hzft5WdOFJvSHpHYBvQkkUKRZlmYDVGvSud82AYwcT
# LzuuXzN+uAd5mw5IYT0tBK7iuNYl3ZPZQerKrES+Iqa9lwsRCTilqbYiigoN0a1c
# Yv81wF4Pw1TRsvhUp8BGLUXF5LzRr5XaVN9aiK5DlxPWkGHCeF3JdReW49pZN5Vs
# K1qQLkK/c1Ed+SRWCkDvS2V+Al6nxF0fzaCgW/70kUq8HJYGFWOKhAI6MguyRXRQ
# U2I9p3p6DXlxaPsCJsWA2DzjrsHOqloqWXjT/a0E5smjFmsB354dobw3L9kG5U/G
# OPM2NGAYDYzvwh/Cei2tPrsO6IereCQ3W498y8CoB+p6WmAG77cIHZUBHSpP2KfA
# DjH2wKSnuSsL0VZnKqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxNDA4
# MDFaMC8GCSqGSIb3DQEJBDEiBCAO1xVzEi/FUiMXBtWdrhMG14dBYM1sKeZcWQCs
# p62U0zANBgkqhkiG9w0BAQEFAASCAgAUDODzfIu+0kuRdNAkMvEiuoNy0/E1JbcW
# 1xFHXJ0GDuh9pUp9OcGKxRrMR5AqVV/QgiLCy16N8Yvgx0bJsrn9SA4KcRRxuROo
# 3q5EBIWZ12nj8ZAEwDhXWBE9bzP7QbWuX5h91U69pvMaAEjeTFHEabjwM2kD0HMw
# sjhoGAbtPhcOyVYEuoAn2kJF3EU0Iiya59xnCN+gpAMBoMEdZqK6wJ2mEit++zdI
# ETcUvBsy1TB6Pd7WzsCujxySoNOrwzAfBWwiirvh+DRENdmCFx8SxUFOjTGQD8Cb
# obiKjPV5x/J5N74E1gzqkO2+WX69t79dP+Mk4JmNYQngPZ3slnYyVRit93qHNzhz
# bJlUXnQTZnPKpvsxQVXowCa3LL9lu4tniSFvuIaE3+8oOJlutI9Fic0X4H+u1+Oe
# 6j96bGIl8C38N1gnKbiuApjDYIC81Fv1LCx4StrvGTrUf1vYMEqd9DG/MQSdRPSl
# CzkSxH4gLEG2iuWu3bZ7JmS3YYOoxFqFHYsyaLGixF16PHYOenu9Wwoltv7++TTt
# W7bQ5l6Ad0a9UmZQuMelW/L4eu32wNyIZ2zge5qkzC4XxfOvGZ/ctCem8esaGcJE
# p0IRGSpXxvMsCk9L4gMp9VZbQ38M0c2yXy5Yjj5+A24bGKiMpkDKdzWSqrTYiHFm
# WEgEfLSsNQ==
# SIG # End signature block
