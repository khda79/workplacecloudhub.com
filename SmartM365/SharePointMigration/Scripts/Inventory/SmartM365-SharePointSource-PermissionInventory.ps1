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
    1.1.0
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
    exit 1
}
finally {
    if ($script:TranscriptStarted) {
        Stop-TimestampedTranscript -Path $LogPath
    }
}








# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCr9G8WpbCk8Dys
# iM2Exc9TZmPG0tBzuJsdzkS3+smaGqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPFrmn0dkGVkyfyNMDm48dEFZJJSAbLNJmDe0vJuL0iJMA0GCSqG
# SIb3DQEBAQUABIIBgCAJQWmOhGpTUubsXtCadX73BB/WT7NIk0HMO1jP/IJ1i/rA
# vIVArAW1uO5vh/q9ZkUeuy+gZUI4Qmqq80u6OJ7+f+t9QX5lTvzdeqKBrcrqbGyw
# GIR4BWf7llFZYXRAXuNaTU0U7sK4FISbdofg1y/G8C7j02ruYJo1REuIASpzfwle
# YDLH5mky6ws/1Uddj8ux2Fst+9FFpAvsnVWNxdugCSefk78HuMqvtbQRvSUfYMUS
# GTlsDA4EQRnZV2JXJsF0RS4AUfxtho4dYF3zR0sZuutWnVwEn5z1oLlbmaWPRMzJ
# kagJqcDxly+rrN/bZLoTNNO2Hk9SR9D+1AUZG2vFSFdUH/YG2YwLXqIcIqFq31ys
# UL50Rqtr3hUJIVQcLX8wzT+hiLxN5gtEjFa2oaOo1Nnc4yo0BZgZNfKInxDVHfeP
# F52bkqHCHSPUxklgObjR3RyVd/VHWogMg1xROwuTZatsicEeV1RxYzNDFoSUujiL
# 1I6YWSkLCTYGvLOCcaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODUw
# MTRaMC8GCSqGSIb3DQEJBDEiBCAoMWeM01oPxGuZsxS0hiEK/yc9YwSd/8U6Sghx
# xlg2ojANBgkqhkiG9w0BAQEFAASCAgCHzAci5TpLFOIP/weTPJoCBTsmkHIInvy8
# PMhW4122/fP0Uhw1Ft2VxpkBFEbcchQnMRbq+HmISLdERrlRGDndQ2WBoVe8p2yE
# m8LCIqJnuzZkezkAj0KcooMiTJr2LxUfG+t3f4G4UCVgsynMGj3Xwf2M3A8LFWIQ
# Sc3Tw46A9EEGwiN/AbQFqgjCzXfBAOuAayUTmAaktXbBG9iHsCBvy3Ow6UdOezy9
# 7yf49W6lNbJeSmtQhpR8evcx6wT7Tugbs/tNxk1U5ynOouRIVj0hwCGhPZhsnIrV
# 3SSOOyz4MU6q3aSP7EXTQb9tuOWDwZ+yGZrrgISp8OsPRWoDX8md4ULK20qaQqK4
# zzlGit2cLfOlW0g3MRCB0ayoV2c7ib5xHxtDUpzNHH3+Ysq8h5x5flOrw6ERkyNI
# QtbE10nJf8VWOy30mmD1weMV86s9xWIuaeMQ5GxFOCRUPJzrviPGFxym7q71h/IH
# P2G5ljbpkgyd6Vushl8nyr3IKFclTfqY1x62eYbl4w42bm/3qMOPCRQtwStK7wrR
# B4y2f5TESzcqBZlkeLraQqzvr63/4qC1sXRCyV6FLQsdBMFr+NXWbxcH8DUZfSgz
# IAYsMzqrLM8tmqH4hkEfA1BDZrlnOtzs5nlHfk8R6AHkAEIABVsP0fdPg3Kggnsc
# P9/ilaJQgw==
# SIG # End signature block
