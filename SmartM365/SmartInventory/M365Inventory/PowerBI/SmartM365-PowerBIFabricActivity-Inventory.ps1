<#
.SYNOPSIS
Exports read-only Power BI and Microsoft Fabric activity events.

.DESCRIPTION
Calls the official Power BI admin Activity Events REST API one UTC day at a
time, follows continuationUri or continuationToken pagination, enforces the
200 requests-per-hour limit, and publishes detailed and per-principal CSV
files for SmartInventory and SmartFinOps.

.VERSION
1.0.0

.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core 1.0.41+; MicrosoftPowerBIMgmt.Profile 1.2.1111+.
    Interactive: a signed-in Fabric Administrator with delegated Tenant.Read.All
    or Tenant.ReadWrite.All.
    App-only: certificate authentication and the Fabric tenant setting
    "Service principals can access read-only admin APIs" scoped to a security
    group containing the service principal. Do not add Power BI application
    permissions that require admin consent.
    Conditional: Sites.Selected write is required only when SharePoint upload
    is enabled.

.NOTES
Author: https://github.com/khda79/workplacecloudhub.com
API: GET https://api.powerbi.com/v1.0/myorg/admin/activityevents
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [ValidateRange(1, 28)]
    [int]$LookbackDays = 28,
    [Nullable[datetime]]$FromDate,
    [Nullable[datetime]]$ToDate,
    [string]$OutputPath,
    [string]$LatestCsvFolderPath,
    [switch]$InteractiveAuth,
    [switch]$ValidateOnly,
    [ValidateRange(0, 2147483647)]
    [int]$MaxItems = 0
)

$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.0.0'
$TaskName = "SmartM365-PowerBIFabricActivity-Inventory v$ScriptVersion"
$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$script:PowerBIActivityApiUri = 'https://api.powerbi.com/v1.0/myorg/admin/activityevents'
$script:PowerBITokenAudience = 'https://analysis.windows.net/powerbi/api'
$script:PowerBIResourceAppId = '00000009-0000-0000-c000-000000000000'
$script:PowerBIRequestLimitPerHour = 200
$script:PowerBIRequestTimesUtc = New-Object 'System.Collections.Generic.List[datetime]'
$script:PowerBIConnected = $false

if ($PSVersionTable.PSVersion.Major -lt 7) {
    throw "This script requires PowerShell 7 or later. Current version: $($PSVersionTable.PSVersion)"
}

if ($MaxItems -gt 0) {
    $global:SmartM365MaxItems = $MaxItems
    $global:SmartM365TestMaxItems = $MaxItems
    $global:SmartM365IsMaxItemsRun = $true
}

$tenantContextPath = & {
    $directory = $PSScriptRoot
    while ($directory) {
        $candidates = @(
            (Join-Path -Path $directory -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $directory -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($candidate in $candidates) {
            if (Test-Path -LiteralPath $candidate) { return $candidate }
        }
        $parent = Split-Path -Path $directory -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $directory) { break }
        $directory = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}

. $tenantContextPath
$script:SmartM365EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

$scriptConfigPath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-PowerBIFabricActivity-Inventory.local.json'
$script:PowerBIActivityConfig = Read-SmartM365JsonConfig -Path $scriptConfigPath -Required

function Test-SmartM365ConfigFallbackValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return $true }
    if ($Value -isnot [string]) { return $false }
    return ([string]::IsNullOrWhiteSpace($Value) -or $Value.Trim() -in @('__USE_GLOBAL__', 'USE_GLOBAL'))
}

function Get-PowerBIActivityRawConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()]$DefaultValue = $null
    )

    $scriptProperty = $script:PowerBIActivityConfig.PSObject.Properties[$Name]
    if ($null -ne $scriptProperty -and -not (Test-SmartM365ConfigFallbackValue -Value $scriptProperty.Value)) {
        return $scriptProperty.Value
    }

    $tenantProperty = $script:SmartM365EffectiveConfig.PSObject.Properties[$Name]
    if ($null -ne $tenantProperty -and -not (Test-SmartM365ConfigFallbackValue -Value $tenantProperty.Value)) {
        return $tenantProperty.Value
    }

    return $DefaultValue
}

function Resolve-PowerBIActivityConfigValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    $resolved = [string]$Value
    for ($index = 0; $index -lt 10; $index++) {
        $configMatches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($configMatches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $configMatches) {
            $replacement = Get-PowerBIActivityRawConfigValue -Name $match.Groups['Name'].Value
            if ($null -eq $replacement) { continue }
            $replacement = Resolve-PowerBIActivityConfigValue -Value $replacement
            if ($null -eq $replacement) { continue }
            $resolved = $resolved.Replace($match.Value, [string]$replacement)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    return $resolved
}

function Get-PowerBIActivityConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()]$DefaultValue = $null
    )

    return Resolve-PowerBIActivityConfigValue -Value (Get-PowerBIActivityRawConfigValue -Name $Name -DefaultValue $DefaultValue)
}

$smartM365RootPath = [string](Get-PowerBIActivityConfigValue -Name 'SmartM365RootPath' -DefaultValue '')
$coreModulePath = Join-Path -Path $smartM365RootPath -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
Import-Module -Name $coreModulePath -MinimumVersion '1.0.41' -Force -ErrorAction Stop

$global:SmartM365ExecutionStartTime = Get-Date
$global:SmartM365ExecutionSummaryWritten = $false
$global:SmartM365ScriptName = $TaskName
$global:RetentionMaxCSV = [int](Get-PowerBIActivityConfigValue -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-PowerBIActivityConfigValue -Name 'RetentionMaxLogs' -DefaultValue 30)
$global:EnableSharePointUpload = if ($ValidateOnly) { $false } else { [bool](Get-PowerBIActivityConfigValue -Name 'EnableSharePointUpload' -DefaultValue $false) }
$global:EnableTeamsNotifications = if ($ValidateOnly) { $false } else { [bool](Get-PowerBIActivityConfigValue -Name 'EnableTeamsNotifications' -DefaultValue $false) }
$global:SharePointSiteHostname = Get-PowerBIActivityConfigValue -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-PowerBIActivityConfigValue -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-PowerBIActivityConfigValue -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-PowerBIActivityConfigValue -Name 'SharePointTargetFolderPath' -DefaultValue ''

$detailColumns = @(
    'CreationTime',
    'Activity',
    'Operation',
    'Workload',
    'UserId',
    'UserPrincipalName',
    'UserType',
    'PrincipalType',
    'IsSuccess',
    'WorkspaceId',
    'WorkspaceName',
    'ItemName',
    'ReportName',
    'DatasetName',
    'DataflowName',
    'CapacityId',
    'CapacityName'
)

$userActivityColumns = @(
    'UserId',
    'UserPrincipalName',
    'UserType',
    'PrincipalType',
    'LastActivityDate',
    'ActiveDays',
    'TotalActivityCount',
    'ViewReportCount',
    'ViewDashboardCount',
    'CreateOrPublishCount',
    'RefreshCount',
    'ExportCount',
    'DistinctWorkspaceCount',
    'DistinctReportCount',
    'DistinctDatasetCount',
    'HasRecentActivity'
)

function Resolve-PowerBIActivityDateRange {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$RequestedLookbackDays,
        [Nullable[datetime]]$RequestedFromDate,
        [Nullable[datetime]]$RequestedToDate,
        [datetime]$UtcNow = [datetime]::UtcNow
    )

    $utcToday = [datetime]::SpecifyKind($UtcNow.Date, [DateTimeKind]::Utc)
    $latestCompleteDate = $utcToday.AddDays(-1)
    $earliestAvailableDate = $utcToday.AddDays(-28)
    $hasFrom = $null -ne $RequestedFromDate
    $hasTo = $null -ne $RequestedToDate

    if ($hasTo) {
        $effectiveToDate = [datetime]::SpecifyKind(([datetime]$RequestedToDate).Date, [DateTimeKind]::Utc)
    }
    else {
        $effectiveToDate = $latestCompleteDate
    }

    if ($hasFrom) {
        $effectiveFromDate = [datetime]::SpecifyKind(([datetime]$RequestedFromDate).Date, [DateTimeKind]::Utc)
    }
    else {
        $effectiveFromDate = $effectiveToDate.AddDays(-($RequestedLookbackDays - 1))
    }

    if ($effectiveFromDate -gt $effectiveToDate) {
        throw "FromDate must be earlier than or equal to ToDate. FromDate=$($effectiveFromDate.ToString('yyyy-MM-dd')); ToDate=$($effectiveToDate.ToString('yyyy-MM-dd'))."
    }
    if ($effectiveFromDate -lt $earliestAvailableDate) {
        throw "FromDate is outside the Power BI Activity Events 28-day retention window. Earliest allowed UTC date: $($earliestAvailableDate.ToString('yyyy-MM-dd'))."
    }
    if ($effectiveToDate -gt $latestCompleteDate) {
        throw "ToDate must not be later than the previous complete UTC day. Latest allowed UTC date: $($latestCompleteDate.ToString('yyyy-MM-dd'))."
    }

    $dayCount = [int](($effectiveToDate - $effectiveFromDate).TotalDays) + 1
    if ($dayCount -lt 1 -or $dayCount -gt 28) {
        throw "The requested range must contain between 1 and 28 UTC days. Requested days: $dayCount."
    }

    return [pscustomobject]@{
        FromDate = $effectiveFromDate
        ToDate = $effectiveToDate
        DayCount = $dayCount
        EarliestAvailableDate = $earliestAvailableDate
        LatestCompleteDate = $latestCompleteDate
    }
}

function New-PowerBIActivityDailyWindows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$RangeFromDate,
        [Parameter(Mandatory)][datetime]$RangeToDate
    )

    $windows = New-Object 'System.Collections.Generic.List[object]'
    for ($date = $RangeToDate.Date; $date -ge $RangeFromDate.Date; $date = $date.AddDays(-1)) {
        $startUtc = [datetime]::SpecifyKind($date.Date, [DateTimeKind]::Utc)
        $endUtc = $startUtc.AddDays(1).AddMilliseconds(-1)
        $windows.Add([pscustomobject]@{
            Date = $startUtc.Date
            StartUtc = $startUtc
            EndUtc = $endUtc
        }) | Out-Null
    }
    return $windows.ToArray()
}

function New-PowerBIActivityInitialUri {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$StartUtc,
        [Parameter(Mandatory)][datetime]$EndUtc
    )

    if ($StartUtc.Kind -ne [DateTimeKind]::Utc -or $EndUtc.Kind -ne [DateTimeKind]::Utc) {
        throw 'Activity Events request boundaries must use DateTimeKind.Utc.'
    }
    if ($StartUtc.Date -ne $EndUtc.Date) {
        throw 'Each Activity Events request must stay within one UTC day.'
    }

    $startText = [uri]::EscapeDataString(("'{0}'" -f $StartUtc.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')))
    $endText = [uri]::EscapeDataString(("'{0}'" -f $EndUtc.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')))
    return '{0}?startDateTime={1}&endDateTime={2}' -f $script:PowerBIActivityApiUri, $startText, $endText
}

function Get-PowerBIActivityNextUri {
    [CmdletBinding()]
    param([AllowNull()]$Response)

    if ($null -eq $Response) { return $null }
    $continuationUri = [string](Get-ActivityPropertyValue -InputObject $Response -Names @('continuationUri'))
    if (-not [string]::IsNullOrWhiteSpace($continuationUri)) {
        $parsedUri = $null
        if (-not [uri]::TryCreate($continuationUri, [UriKind]::Absolute, [ref]$parsedUri)) {
            throw 'Power BI returned an invalid continuationUri.'
        }
        if ($parsedUri.Scheme -ne 'https' -or $parsedUri.Host -ne 'api.powerbi.com' -or
            $parsedUri.AbsolutePath.TrimEnd('/') -ne '/v1.0/myorg/admin/activityevents') {
            throw "Power BI returned an unexpected continuationUri target: $($parsedUri.GetLeftPart([UriPartial]::Path))."
        }
        return $parsedUri.AbsoluteUri
    }

    $continuationToken = [string](Get-ActivityPropertyValue -InputObject $Response -Names @('continuationToken'))
    if ([string]::IsNullOrWhiteSpace($continuationToken)) { return $null }
    $encodedToken = if ($continuationToken -match '%[0-9A-Fa-f]{2}') {
        $continuationToken
    }
    else {
        [uri]::EscapeDataString($continuationToken)
    }
    return '{0}?continuationToken={1}' -f $script:PowerBIActivityApiUri, $encodedToken
}

function Get-ActivityPropertyValue {
    [CmdletBinding()]
    param(
        [AllowNull()]$InputObject,
        [Parameter(Mandatory)][string[]]$Names
    )

    if ($null -eq $InputObject) { return $null }
    foreach ($name in $Names) {
        $property = $InputObject.PSObject.Properties[$name]
        if ($null -ne $property -and $null -ne $property.Value) { return $property.Value }
    }
    return $null
}

function ConvertTo-ActivityText {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value.Trim() }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [System.Collections.IDictionary]) {
        return (@($Value | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) -join ';')
    }
    return ([string]$Value).Trim()
}

function Get-ActivityCollectionNames {
    [CmdletBinding()]
    param(
        [AllowNull()]$InputObject,
        [Parameter(Mandatory)][string[]]$CollectionNames,
        [Parameter(Mandatory)][string[]]$ItemPropertyNames
    )

    $names = New-Object 'System.Collections.Generic.List[string]'
    $collection = Get-ActivityPropertyValue -InputObject $InputObject -Names $CollectionNames
    foreach ($item in @($collection)) {
        $name = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $item -Names $ItemPropertyNames)
        if (-not [string]::IsNullOrWhiteSpace($name) -and -not $names.Contains($name)) {
            $names.Add($name) | Out-Null
        }
    }
    return ($names.ToArray() -join ';')
}

function ConvertTo-PowerBIActivityUtcTimestamp {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    $text = ConvertTo-ActivityText -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return [pscustomobject]@{ Text = ''; Value = $null }
    }

    $parsed = [datetimeoffset]::MinValue
    $styles = [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
    if ([datetimeoffset]::TryParse($text, [Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsed)) {
        return [pscustomobject]@{
            Text = $parsed.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
            Value = $parsed.UtcDateTime
        }
    }
    return [pscustomobject]@{ Text = $text; Value = $null }
}

function Get-PowerBIActivityPrincipalType {
    [CmdletBinding()]
    param(
        [AllowNull()]$UserType,
        [string]$UserId,
        [string]$UserPrincipalName,
        [AllowNull()]$ActivityEvent
    )

    $userText = ('{0}|{1}' -f $UserId, $UserPrincipalName)
    $normalizedType = (ConvertTo-ActivityText -Value $UserType).Replace(' ', '').ToLowerInvariant()
    $mappedType = switch ($normalizedType) {
        '0' { 'User' }
        'regular' { 'User' }
        '2' { 'User' }
        'admin' { 'User' }
        '3' { 'MicrosoftDatacenterAdmin' }
        'dcadmin' { 'MicrosoftDatacenterAdmin' }
        '4' { 'System' }
        'system' { 'System' }
        '5' { 'Application' }
        'application' { 'Application' }
        '6' { 'ServicePrincipal' }
        'serviceprincipal' { 'ServicePrincipal' }
        '7' { 'CustomPolicy' }
        'custompolicy' { 'CustomPolicy' }
        '8' { 'SystemPolicy' }
        'systempolicy' { 'SystemPolicy' }
        '9' { 'PartnerTechnician' }
        'partnertechnician' { 'PartnerTechnician' }
        '10' { 'Guest' }
        'guest' { 'Guest' }
        '11' { 'Agent' }
        'agent' { 'Agent' }
        default { '' }
    }

    if ($mappedType -in @('User', '') -and $userText -match '(?i)#EXT#') { return 'Guest' }
    if (-not [string]::IsNullOrWhiteSpace($mappedType)) { return $mappedType }

    $appIdentifier = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('AppId', 'ApplicationId', 'ClientId'))
    if (-not [string]::IsNullOrWhiteSpace($appIdentifier) -and [string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        return 'Application'
    }
    if (-not [string]::IsNullOrWhiteSpace($UserPrincipalName) -or $UserId -match '@') { return 'User' }
    return 'Unknown'
}

function ConvertTo-PowerBIActivityDetail {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ActivityEvent)

    $creationTime = ConvertTo-PowerBIActivityUtcTimestamp -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('CreationTime'))
    $activity = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('Activity'))
    $operation = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('Operation'))
    if ([string]::IsNullOrWhiteSpace($activity)) { $activity = $operation }
    if ([string]::IsNullOrWhiteSpace($operation)) { $operation = $activity }

    $userType = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('UserType'))
    $userId = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('UserId'))
    $userPrincipalName = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('UserPrincipalName'))
    if ([string]::IsNullOrWhiteSpace($userPrincipalName) -and $userId -match '@') { $userPrincipalName = $userId }

    $principalType = Get-PowerBIActivityPrincipalType -UserType $userType -UserId $userId -UserPrincipalName $userPrincipalName -ActivityEvent $ActivityEvent
    if ([string]::IsNullOrWhiteSpace($userId) -and $principalType -in @('Application', 'ServicePrincipal')) {
        $userId = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('UserKey', 'AppId', 'ApplicationId', 'ClientId'))
    }

    $workspaceId = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('WorkspaceId', 'WorkSpaceId'))
    $workspaceName = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('WorkspaceName', 'WorkSpaceName'))
    $reportName = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('ReportName'))
    if ([string]::IsNullOrWhiteSpace($reportName)) {
        $reportName = Get-ActivityCollectionNames -InputObject $ActivityEvent -CollectionNames @('Reports') -ItemPropertyNames @('ReportName', 'Name')
    }
    $datasetName = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('DatasetName', 'SemanticModelName'))
    if ([string]::IsNullOrWhiteSpace($datasetName)) {
        $datasetName = Get-ActivityCollectionNames -InputObject $ActivityEvent -CollectionNames @('Datasets', 'SemanticModels') -ItemPropertyNames @('DatasetName', 'SemanticModelName', 'Name')
    }
    $dataflowName = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('DataflowName'))
    if ([string]::IsNullOrWhiteSpace($dataflowName)) {
        $dataflowName = Get-ActivityCollectionNames -InputObject $ActivityEvent -CollectionNames @('Dataflows') -ItemPropertyNames @('DataflowName', 'Name')
    }
    $itemName = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('ItemName', 'ArtifactName'))
    if ([string]::IsNullOrWhiteSpace($itemName)) {
        $itemName = @($reportName, $datasetName, $dataflowName, (ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('ObjectId')))) |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -First 1
    }

    $successValue = Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('IsSuccess')
    $isSuccess = if ($successValue -is [bool]) {
        $successValue
    }
    elseif ((ConvertTo-ActivityText -Value $successValue) -match '^(?i:true|false)$') {
        [bool]::Parse((ConvertTo-ActivityText -Value $successValue))
    }
    else {
        ConvertTo-ActivityText -Value $successValue
    }

    $reportId = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('ReportId'))
    $datasetId = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('DatasetId', 'SemanticModelId'))
    $eventNames = @($activity, $operation)
    if ([string]::IsNullOrWhiteSpace($reportId) -and $eventNames -contains 'ViewReport') { $reportId = $reportName }
    if ([string]::IsNullOrWhiteSpace($datasetId)) { $datasetId = $datasetName }

    return [pscustomobject][ordered]@{
        CreationTime = $creationTime.Text
        Activity = $activity
        Operation = $operation
        Workload = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('Workload'))
        UserId = $userId
        UserPrincipalName = $userPrincipalName
        UserType = $userType
        PrincipalType = $principalType
        IsSuccess = $isSuccess
        WorkspaceId = $workspaceId
        WorkspaceName = $workspaceName
        ItemName = $itemName
        ReportName = $reportName
        DatasetName = $datasetName
        DataflowName = $dataflowName
        CapacityId = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('CapacityId'))
        CapacityName = ConvertTo-ActivityText -Value (Get-ActivityPropertyValue -InputObject $ActivityEvent -Names @('CapacityName'))
        __CreationTimeUtc = $creationTime.Value
        __WorkspaceKey = if (-not [string]::IsNullOrWhiteSpace($workspaceId)) { $workspaceId } else { $workspaceName }
        __ReportKey = if (-not [string]::IsNullOrWhiteSpace($reportId)) { $reportId } else { $reportName }
        __DatasetKey = $datasetId
    }
}

function Test-PowerBIActivityName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Detail,
        [Parameter(Mandatory)][string]$Pattern
    )

    foreach ($name in @([string]$Detail.Activity, [string]$Detail.Operation)) {
        if (-not [string]::IsNullOrWhiteSpace($name) -and $name -match $Pattern) { return $true }
    }
    return $false
}

function ConvertTo-PowerBIUserActivitySummary {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Details)

    $states = @{}
    foreach ($detail in $Details) {
        $key = '{0}|{1}|{2}' -f ([string]$detail.UserId).ToLowerInvariant(), ([string]$detail.UserPrincipalName).ToLowerInvariant(), ([string]$detail.PrincipalType).ToLowerInvariant()
        if (-not $states.ContainsKey($key)) {
            $states[$key] = [pscustomobject]@{
                UserId = [string]$detail.UserId
                UserPrincipalName = [string]$detail.UserPrincipalName
                UserType = [string]$detail.UserType
                PrincipalType = [string]$detail.PrincipalType
                LastActivityDate = $null
                ActiveDays = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
                TotalActivityCount = 0
                ViewReportCount = 0
                ViewDashboardCount = 0
                CreateOrPublishCount = 0
                RefreshCount = 0
                ExportCount = 0
                Workspaces = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
                Reports = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
                Datasets = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
            }
        }

        $state = $states[$key]
        $state.TotalActivityCount++
        if ($null -ne $detail.__CreationTimeUtc) {
            $eventDate = [datetime]$detail.__CreationTimeUtc
            if ($null -eq $state.LastActivityDate -or $eventDate -gt $state.LastActivityDate) { $state.LastActivityDate = $eventDate }
            [void]$state.ActiveDays.Add($eventDate.ToString('yyyy-MM-dd'))
        }
        if (Test-PowerBIActivityName -Detail $detail -Pattern '^(?i:ViewReport)$') { $state.ViewReportCount++ }
        if (Test-PowerBIActivityName -Detail $detail -Pattern '^(?i:ViewDashboard)$') { $state.ViewDashboardCount++ }
        if (Test-PowerBIActivityName -Detail $detail -Pattern '^(?i:(Create|Publish|Upload|Import|Save))') { $state.CreateOrPublishCount++ }
        if (Test-PowerBIActivityName -Detail $detail -Pattern '(?i:Refresh)') { $state.RefreshCount++ }
        if (Test-PowerBIActivityName -Detail $detail -Pattern '^(?i:(Export|Download))') { $state.ExportCount++ }
        if (-not [string]::IsNullOrWhiteSpace([string]$detail.__WorkspaceKey)) { [void]$state.Workspaces.Add([string]$detail.__WorkspaceKey) }
        if (-not [string]::IsNullOrWhiteSpace([string]$detail.__ReportKey)) { [void]$state.Reports.Add([string]$detail.__ReportKey) }
        if (-not [string]::IsNullOrWhiteSpace([string]$detail.__DatasetKey)) { [void]$state.Datasets.Add([string]$detail.__DatasetKey) }
    }

    foreach ($state in @($states.Values | Sort-Object PrincipalType, UserPrincipalName, UserId)) {
        [pscustomobject][ordered]@{
            UserId = $state.UserId
            UserPrincipalName = $state.UserPrincipalName
            UserType = $state.UserType
            PrincipalType = $state.PrincipalType
            LastActivityDate = if ($null -ne $state.LastActivityDate) { ([datetime]$state.LastActivityDate).ToString('yyyy-MM-ddTHH:mm:ss.fffZ') } else { '' }
            ActiveDays = $state.ActiveDays.Count
            TotalActivityCount = $state.TotalActivityCount
            ViewReportCount = $state.ViewReportCount
            ViewDashboardCount = $state.ViewDashboardCount
            CreateOrPublishCount = $state.CreateOrPublishCount
            RefreshCount = $state.RefreshCount
            ExportCount = $state.ExportCount
            DistinctWorkspaceCount = $state.Workspaces.Count
            DistinctReportCount = $state.Reports.Count
            DistinctDatasetCount = $state.Datasets.Count
            HasRecentActivity = ($state.TotalActivityCount -gt 0)
        }
    }
}

function ConvertFrom-PowerBIJwtPayload {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Authorization)

    $token = ($Authorization -replace '^(?i:Bearer)\s+', '').Trim()
    $segments = $token.Split('.')
    if ($segments.Count -lt 2) { throw 'The Power BI access token is not a valid JWT.' }
    $payload = $segments[1].Replace('-', '+').Replace('_', '/')
    switch ($payload.Length % 4) {
        2 { $payload += '==' }
        3 { $payload += '=' }
        1 { throw 'The Power BI access token has invalid base64url padding.' }
    }
    $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload))
    return $json | ConvertFrom-Json -ErrorAction Stop
}

function Assert-PowerBIAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Headers,
        [Parameter(Mandatory)][bool]$IsInteractive
    )

    $authorization = [string]$Headers['Authorization']
    if ([string]::IsNullOrWhiteSpace($authorization)) { throw 'Get-PowerBIAccessToken did not return an Authorization header.' }
    $claims = ConvertFrom-PowerBIJwtPayload -Authorization $authorization
    $audience = [string]$claims.aud
    if ($audience -notin @($script:PowerBITokenAudience, $script:PowerBIResourceAppId)) {
        throw "The access token audience is '$audience', not the Power BI resource '$($script:PowerBITokenAudience)'. Microsoft Graph tokens are not accepted."
    }

    $configuredTenantId = [string](Get-PowerBIActivityConfigValue -Name 'TenantId' -DefaultValue '')
    $tenantGuid = [guid]::Empty
    if ([guid]::TryParse($configuredTenantId, [ref]$tenantGuid) -and [string]$claims.tid -ne $tenantGuid.Guid) {
        throw "The Power BI token tenant does not match the selected SmartM365 tenant profile '$Tenant'."
    }

    if ($IsInteractive) {
        $scopes = @(([string]$claims.scp -split '\s+') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($scopes -notcontains 'Tenant.Read.All' -and $scopes -notcontains 'Tenant.ReadWrite.All') {
            throw 'The delegated Power BI token is missing Tenant.Read.All or Tenant.ReadWrite.All.'
        }
    }
    else {
        $roles = @($claims.roles | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        if ($roles.Count -gt 0) {
            throw "The service principal token contains Power BI application roles ($($roles -join ', ')). Remove admin-consent-required Power BI permissions from the app and use the Fabric tenant setting for read-only admin APIs."
        }
    }
    return $claims
}

function Connect-PowerBIActivitySession {
    [CmdletBinding()]
    param([switch]$Force)

    if ($script:PowerBIConnected -and -not $Force) { return }
    if (Get-Command -Name Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue) {
        try { Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue | Out-Null } catch {}
    }

    $tenantId = [string](Get-PowerBIActivityConfigValue -Name 'TenantId' -DefaultValue '')
    if ([string]::IsNullOrWhiteSpace($tenantId)) { throw "TenantId is missing from the effective tenant configuration for '$Tenant'." }

    if ($InteractiveAuth) {
        WriteLog -Message 'Connecting interactively to Power BI. The signed-in account must be a Fabric Administrator.' -Level 'INFO'
        Connect-PowerBIServiceAccount -Tenant $tenantId -ErrorAction Stop | Out-Null
    }
    else {
        $appId = [string](Get-PowerBIActivityConfigValue -Name 'AppId' -DefaultValue '')
        $thumbprint = [string](Get-PowerBIActivityConfigValue -Name 'Thumbprint' -DefaultValue (Get-PowerBIActivityConfigValue -Name 'Thumb' -DefaultValue ''))
        if ([string]::IsNullOrWhiteSpace($appId) -or [string]::IsNullOrWhiteSpace($thumbprint)) {
            throw "App-only Power BI authentication requires AppId, TenantId, and Thumbprint in the selected local SmartM365 tenant configuration. No secret is supported."
        }
        WriteLog -Message 'Connecting to Power BI with the SmartM365 service principal certificate.' -Level 'INFO'
        Connect-PowerBIServiceAccount -ServicePrincipal -ApplicationId $appId -CertificateThumbprint $thumbprint -Tenant $tenantId -ErrorAction Stop | Out-Null
    }

    $headers = Get-PowerBIAccessToken -ErrorAction Stop
    $claims = Assert-PowerBIAccessToken -Headers $headers -IsInteractive ([bool]$InteractiveAuth)
    $script:PowerBIConnected = $true
    WriteLog -Message ("Power BI authentication validated. Mode={0}; Audience={1}; TenantClaimPresent={2}" -f $(if ($InteractiveAuth) { 'Interactive Fabric Administrator' } else { 'Service principal certificate' }), [string]$claims.aud, (-not [string]::IsNullOrWhiteSpace([string]$claims.tid))) -Level 'SUCCESS'
}

function Get-PowerBIActivityAuthorizationHeaders {
    [CmdletBinding()]
    param()

    if (-not $script:PowerBIConnected) { Connect-PowerBIActivitySession }
    $headers = Get-PowerBIAccessToken -ErrorAction Stop
    $null = Assert-PowerBIAccessToken -Headers $headers -IsInteractive ([bool]$InteractiveAuth)
    return $headers
}

function Get-PowerBIRateLimitDelaySeconds {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][datetime[]]$RequestTimesUtc,
        [datetime]$UtcNow = [datetime]::UtcNow,
        [int]$RequestLimit = 200
    )

    $recent = @($RequestTimesUtc | Where-Object { $_ -gt $UtcNow.AddHours(-1) } | Sort-Object)
    if ($recent.Count -lt $RequestLimit) { return 0 }
    return [int][Math]::Max(1, [Math]::Ceiling(($recent[0].AddHours(1).AddSeconds(1) - $UtcNow).TotalSeconds))
}

function Wait-PowerBIActivityRequestSlot {
    [CmdletBinding()]
    param()

    while ($true) {
        $now = [datetime]::UtcNow
        for ($index = $script:PowerBIRequestTimesUtc.Count - 1; $index -ge 0; $index--) {
            if ($script:PowerBIRequestTimesUtc[$index] -le $now.AddHours(-1)) {
                $script:PowerBIRequestTimesUtc.RemoveAt($index)
            }
        }
        $delaySeconds = Get-PowerBIRateLimitDelaySeconds -RequestTimesUtc $script:PowerBIRequestTimesUtc.ToArray() -UtcNow $now -RequestLimit $script:PowerBIRequestLimitPerHour
        if ($delaySeconds -le 0) {
            $script:PowerBIRequestTimesUtc.Add($now) | Out-Null
            return
        }
        WriteLog -Message ("Power BI hourly request limit reached. Waiting {0} second(s) before the next request." -f $delaySeconds) -Level 'WARNING'
        Start-Sleep -Seconds $delaySeconds
    }
}

function Get-PowerBIRetryAfterSeconds {
    [CmdletBinding()]
    param(
        [AllowNull()]$Headers,
        [int]$DefaultSeconds = 60
    )

    if ($null -eq $Headers) { return $DefaultSeconds }
    $retryAfter = [string]$Headers['Retry-After']
    $seconds = 0
    if ([int]::TryParse($retryAfter, [ref]$seconds)) { return [Math]::Max(1, $seconds) }
    $retryDate = [datetimeoffset]::MinValue
    if ([datetimeoffset]::TryParse($retryAfter, [ref]$retryDate)) {
        return [int][Math]::Max(1, [Math]::Ceiling(($retryDate.UtcDateTime - [datetime]::UtcNow).TotalSeconds))
    }
    return $DefaultSeconds
}

function Get-PowerBIErrorText {
    [CmdletBinding()]
    param([AllowNull()]$Response)

    if ($null -eq $Response) { return 'Empty response body.' }
    try { return ($Response | ConvertTo-Json -Depth 12 -Compress) } catch { return [string]$Response }
}

function Invoke-PowerBIActivityApiRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [int]$MaxRetryCount = 8,
        [int]$TimeoutSeconds = 120
    )

    $reauthenticated = $false
    for ($attempt = 1; $attempt -le $MaxRetryCount; $attempt++) {
        Wait-PowerBIActivityRequestSlot
        $headers = Get-PowerBIActivityAuthorizationHeaders
        $statusCode = 0
        $responseHeaders = $null
        $response = Invoke-RestMethod -Method Get -Uri $Uri -Headers $headers -TimeoutSec $TimeoutSeconds -MaximumRedirection 5 -SkipHttpErrorCheck -StatusCodeVariable statusCode -ResponseHeadersVariable responseHeaders
        if ($statusCode -ge 200 -and $statusCode -lt 300) { return $response }

        $errorText = Get-PowerBIErrorText -Response $response
        if ($statusCode -eq 401 -and -not $reauthenticated) {
            WriteLog -Message 'Power BI returned HTTP 401. Refreshing the Power BI authenticated session once.' -Level 'WARNING'
            Connect-PowerBIActivitySession -Force
            $reauthenticated = $true
            continue
        }
        if ($statusCode -eq 429) {
            $delay = Get-PowerBIRetryAfterSeconds -Headers $responseHeaders
            WriteLog -Message ("Power BI throttled the request (HTTP 429). Retrying in {0} second(s); Attempt={1}/{2}." -f $delay, $attempt, $MaxRetryCount) -Level 'WARNING'
            Start-Sleep -Seconds $delay
            continue
        }
        if ($statusCode -in @(408, 409) -or $statusCode -ge 500) {
            $delay = [int][Math]::Min(60, [Math]::Pow(2, [Math]::Min($attempt, 6)))
            WriteLog -Message ("Transient Power BI API failure HTTP {0}. Retrying in {1} second(s); Attempt={2}/{3}." -f $statusCode, $delay, $attempt, $MaxRetryCount) -Level 'WARNING'
            Start-Sleep -Seconds $delay
            continue
        }
        throw "Power BI Activity Events request failed with HTTP $statusCode. Response=$errorText"
    }
    throw "Power BI Activity Events request failed after $MaxRetryCount attempts. Uri=$Uri"
}

function Invoke-PowerBIActivityCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$DailyWindows,
        [int]$MaximumItems = 0
    )

    $details = New-Object 'System.Collections.Generic.List[object]'
    $pageCount = 0
    foreach ($window in $DailyWindows) {
        $uri = New-PowerBIActivityInitialUri -StartUtc $window.StartUtc -EndUtc $window.EndUtc
        $dayItems = 0
        WriteLog -Message ("Collecting Power BI and Fabric activity for UTC day {0}. Start={1}; End={2}" -f $window.Date.ToString('yyyy-MM-dd'), $window.StartUtc.ToString('o'), $window.EndUtc.ToString('o')) -Level 'INFO'

        while (-not [string]::IsNullOrWhiteSpace($uri)) {
            $response = Invoke-PowerBIActivityApiRequest -Uri $uri -MaxRetryCount ([int](Get-PowerBIActivityConfigValue -Name 'MaxRetryCount' -DefaultValue 8)) -TimeoutSeconds ([int](Get-PowerBIActivityConfigValue -Name 'RequestTimeoutSeconds' -DefaultValue 120))
            $pageCount++
            foreach ($activityEventItem in @($response.activityEventEntities)) {
                $details.Add((ConvertTo-PowerBIActivityDetail -ActivityEvent $activityEventItem)) | Out-Null
                $dayItems++
                if ($MaximumItems -gt 0 -and $details.Count -ge $MaximumItems) { break }
            }
            if ($MaximumItems -gt 0 -and $details.Count -ge $MaximumItems) {
                WriteLog -Message ("MaxItems={0} reached. Collection stopped without updating canonical CSV names." -f $MaximumItems) -Level 'WARNING'
                break
            }
            $uri = Get-PowerBIActivityNextUri -Response $response
        }

        WriteLog -Message ("UTC day {0} completed. Events={1}; TotalEvents={2}; Pages={3}" -f $window.Date.ToString('yyyy-MM-dd'), $dayItems, $details.Count, $pageCount) -Level 'INFO'
        if ($MaximumItems -gt 0 -and $details.Count -ge $MaximumItems) { break }
    }

    return [pscustomobject]@{
        Details = @($details.ToArray() | Sort-Object CreationTime -Descending)
        PageCount = $pageCount
    }
}

function Assert-PowerBIActivityValidation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$DateRange,
        [Parameter(Mandatory)][object[]]$DailyWindows
    )

    if ($DailyWindows.Count -ne $DateRange.DayCount) { throw 'Daily window count validation failed.' }
    foreach ($window in $DailyWindows) {
        if ($window.StartUtc.Kind -ne [DateTimeKind]::Utc -or $window.EndUtc.Kind -ne [DateTimeKind]::Utc) { throw 'A daily window is not UTC.' }
        if ($window.StartUtc.Date -ne $window.EndUtc.Date) { throw 'A daily window crosses a UTC date boundary.' }
        $null = New-PowerBIActivityInitialUri -StartUtc $window.StartUtc -EndUtc $window.EndUtc
    }

    $continuationUri = "$($script:PowerBIActivityApiUri)?continuationToken=sample-uri-token"
    $nextFromUri = Get-PowerBIActivityNextUri -Response ([pscustomobject]@{ continuationUri = $continuationUri; continuationToken = 'ignored' })
    if ($nextFromUri -ne $continuationUri) { throw 'continuationUri validation failed.' }
    $nextFromToken = Get-PowerBIActivityNextUri -Response ([pscustomobject]@{ continuationToken = '%2BRID%3Asample%23RT%3A1' })
    if ($nextFromToken -notmatch 'continuationToken=%2BRID%3Asample%23RT%3A1$') { throw 'continuationToken validation failed.' }
    if ($null -ne (Get-PowerBIActivityNextUri -Response ([pscustomobject]@{}))) { throw 'Pagination termination validation failed.' }

    $now = [datetime]::UtcNow
    $twoHundredRequests = @(1..200 | ForEach-Object { $now.AddMinutes(-10) })
    if ((Get-PowerBIRateLimitDelaySeconds -RequestTimesUtc $twoHundredRequests -UtcNow $now) -le 0) { throw 'The 200 requests-per-hour limiter validation failed.' }
    $oneHundredNinetyNineRequests = @(1..199 | ForEach-Object { $now.AddMinutes(-10) })
    if ((Get-PowerBIRateLimitDelaySeconds -RequestTimesUtc $oneHundredNinetyNineRequests -UtcNow $now) -ne 0) { throw 'The request limiter blocked before 200 requests.' }
    if ((Get-PowerBIRateLimitDelaySeconds -RequestTimesUtc @() -UtcNow $now) -ne 0) { throw 'The request limiter rejected an empty initial request ledger.' }

    if (($detailColumns | Select-Object -Unique).Count -ne $detailColumns.Count) { throw 'Detailed CSV schema contains duplicate columns.' }
    if (($userActivityColumns | Select-Object -Unique).Count -ne $userActivityColumns.Count) { throw 'User activity CSV schema contains duplicate columns.' }

    $sampleDate = $DateRange.ToDate.ToString('yyyy-MM-dd')
    $samples = @(
        [pscustomobject]@{ CreationTime = "${sampleDate}T08:00:00Z"; Activity = 'ViewReport'; Operation = 'ViewReport'; Workload = 'PowerBI'; UserType = 0; UserId = 'user@example.com'; WorkspaceId = 'w1'; ReportId = 'r1'; ReportName = 'Report 1'; DatasetId = 'd1'; IsSuccess = $true },
        [pscustomobject]@{ CreationTime = "${sampleDate}T09:00:00Z"; Activity = 'ViewDashboard'; Operation = 'ViewDashboard'; Workload = 'PowerBI'; UserType = 10; UserId = 'guest_example.com#EXT#@tenant.onmicrosoft.com'; WorkspaceId = 'w1'; IsSuccess = $true },
        [pscustomobject]@{ CreationTime = "${sampleDate}T10:00:00Z"; Activity = 'RefreshDataset'; Operation = 'RefreshDataset'; Workload = 'PowerBI'; UserType = 5; UserKey = 'sample-application'; DatasetId = 'd2'; IsSuccess = $true },
        [pscustomobject]@{ CreationTime = "${sampleDate}T11:00:00Z"; Activity = 'ExportData'; Operation = 'ExportData'; Workload = 'Fabric'; UserType = 6; UserKey = 'sample-service-principal'; WorkspaceId = 'w2'; IsSuccess = $true }
    )
    $sampleDetails = @($samples | ForEach-Object { ConvertTo-PowerBIActivityDetail -ActivityEvent $_ })
    $sampleTypes = @($sampleDetails.PrincipalType | Sort-Object -Unique)
    foreach ($expectedType in @('User', 'Guest', 'Application', 'ServicePrincipal')) {
        if ($sampleTypes -notcontains $expectedType) { throw "Principal classification validation failed for $expectedType." }
    }
    $sampleSummary = @(ConvertTo-PowerBIUserActivitySummary -Details $sampleDetails)
    if ($sampleSummary.Count -ne 4) { throw 'User activity aggregation validation failed.' }
    if (($sampleSummary | Measure-Object -Property TotalActivityCount -Sum).Sum -ne 4) { throw 'User activity total count validation failed.' }

    $emptyDetail = [pscustomobject][ordered]@{}
    foreach ($column in $detailColumns) { Add-Member -InputObject $emptyDetail -MemberType NoteProperty -Name $column -Value '' }
    $emptyUser = [pscustomobject][ordered]@{}
    foreach ($column in $userActivityColumns) { Add-Member -InputObject $emptyUser -MemberType NoteProperty -Name $column -Value '' }
    if ((@($emptyDetail.PSObject.Properties.Name) -join '|') -ne ($detailColumns -join '|')) { throw 'Detailed empty-schema validation failed.' }
    if ((@($emptyUser.PSObject.Properties.Name) -join '|') -ne ($userActivityColumns -join '|')) { throw 'User activity empty-schema validation failed.' }
}

$dateRange = Resolve-PowerBIActivityDateRange -RequestedLookbackDays $LookbackDays -RequestedFromDate $FromDate -RequestedToDate $ToDate
$dailyWindows = @(New-PowerBIActivityDailyWindows -RangeFromDate $dateRange.FromDate -RangeToDate $dateRange.ToDate)

$dataAllRoot = [string](Get-PowerBIActivityConfigValue -Name 'DataAllRootPath' -DefaultValue '')
$logAllRoot = [string](Get-PowerBIActivityConfigValue -Name 'LogAllRootPath' -DefaultValue '')
if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) {
    $LatestCsvFolderPath = [string](Get-PowerBIActivityConfigValue -Name 'LatestCsvFolderPath' -DefaultValue '')
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = [string](Get-PowerBIActivityConfigValue -Name 'OutputPath' -DefaultValue (Join-Path -Path $dataAllRoot -ChildPath 'M365\PowerBI\Activity'))
}

$logFolder = Join-Path -Path $logAllRoot -ChildPath 'SmartM365-PowerBIFabricActivity-Inventory'
# Keep the file variable distinct from global:LogPath; PowerShell variable names are case-insensitive.
$logFilePath = Join-Path -Path $logFolder -ChildPath ("SmartM365-PowerBIFabricActivity-Inventory_{0}.log" -f $runId)
$transcriptPath = Join-Path -Path $logFolder -ChildPath ("SmartM365-PowerBIFabricActivity-Inventory_{0}_Transcript.log" -f $runId)

Set-SmartM365CoreContext -RunId $runId -RunOutputRoot $OutputPath -LatestOutputRoot $LatestCsvFolderPath
$global:BasePath = $OutputPath

Initialize-SmartM365DefaultCsvValidationRules
Add-SmartM365CsvValidationRule -Rules $global:SmartM365CsvValidationRules -BaseFileName 'M365_PowerBI_Fabric_ActivityEvents' -CriticalFields @('CreationTime', 'Activity', 'PrincipalType') -RequiredColumns $detailColumns -AllowEmptyDataset -Name 'Power BI and Fabric activity events'
Add-SmartM365CsvValidationRule -Rules $global:SmartM365CsvValidationRules -BaseFileName 'M365_PowerBI_Fabric_UserActivity' -CriticalFields @('PrincipalType', 'LastActivityDate') -RequiredColumns $userActivityColumns -AllowEmptyDataset -Name 'Power BI and Fabric user activity'

if ($ValidateOnly) {
    try {
        Assert-PowerBIActivityValidation -DateRange $dateRange -DailyWindows $dailyWindows
        $powerBIModule = Get-Module -ListAvailable -Name 'MicrosoftPowerBIMgmt.Profile' | Sort-Object Version -Descending | Select-Object -First 1
        $moduleStatus = if ($powerBIModule) { "$($powerBIModule.Version)" } else { 'not installed; required only for a live collection' }
        WriteLog -Message ("ValidateOnly passed. Tenant={0}; UTC days={1}; FromDate={2}; ToDate={3}; Pagination=continuationUri+continuationToken; RateLimit={4}/hour; DetailedColumns={5}; UserActivityColumns={6}; MicrosoftPowerBIMgmt.Profile={7}" -f $Tenant, $dateRange.DayCount, $dateRange.FromDate.ToString('yyyy-MM-dd'), $dateRange.ToDate.ToString('yyyy-MM-dd'), $script:PowerBIRequestLimitPerHour, $detailColumns.Count, $userActivityColumns.Count, $moduleStatus) -Level 'SUCCESS'
        Complete-SmartM365ExecutionContext -Status Success
        return
    }
    catch {
        try { Complete-SmartM365ExecutionContext -Status Failed -ErrorRecord $_ -FailureStage 'ValidateOnly' } catch {}
        throw
    }
}

$global:LogTextFile = $logFilePath
$global:logTextFile = $logFilePath
$global:LogPath = $logFolder
$global:logTranscriptFile = $transcriptPath
Set-SmartM365CoreContext -RunId $runId -RunOutputRoot $OutputPath -LatestOutputRoot $LatestCsvFolderPath -LogPath $logFilePath

function Stop-PowerBIActivityTranscript {
    [CmdletBinding()]
    param()
    try {
        Stop-Transcript | Out-Null
        Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile
    }
    catch {}
}

try {
    foreach ($folder in @($OutputPath, $LatestCsvFolderPath, $logFolder)) {
        if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
    }
    Start-Transcript -Path $transcriptPath -Append | Out-Null
    WriteLog -Message ("Starting {0}. Tenant={1}; FromDate={2}; ToDate={3}; UtcDays={4}; InteractiveAuth={5}; MaxItems={6}" -f $TaskName, $Tenant, $dateRange.FromDate.ToString('yyyy-MM-dd'), $dateRange.ToDate.ToString('yyyy-MM-dd'), $dateRange.DayCount, [bool]$InteractiveAuth, $MaxItems) -Level 'INFO'

    Invoke-SmartM365Preflight -ScriptName $TaskName -RequiredModules @('MicrosoftPowerBIMgmt.Profile') -RequiredCommands @('Connect-PowerBIServiceAccount', 'Get-PowerBIAccessToken', 'Disconnect-PowerBIServiceAccount') -OutputPaths @($OutputPath, $LatestCsvFolderPath, $logFolder) | Out-Null
    $requiredPowerBIModuleVersion = [version]'1.2.1111'
    $availablePowerBIModule = Get-Module -ListAvailable -Name 'MicrosoftPowerBIMgmt.Profile' | Sort-Object Version -Descending | Select-Object -First 1
    if ($null -eq $availablePowerBIModule) {
        throw 'MicrosoftPowerBIMgmt.Profile was not found after the successful preflight.'
    }
    $availablePowerBIModuleVersion = [version]$availablePowerBIModule.Version
    if ($availablePowerBIModuleVersion.CompareTo($requiredPowerBIModuleVersion) -lt 0) {
        throw "MicrosoftPowerBIMgmt.Profile $requiredPowerBIModuleVersion or later is required. Detected version: $availablePowerBIModuleVersion."
    }
    WriteLog -Message ("Power BI module version validated. Required={0}; Detected={1}; Path={2}" -f $requiredPowerBIModuleVersion, $availablePowerBIModuleVersion, $availablePowerBIModule.Path) -Level 'SUCCESS'

    Connect-PowerBIActivitySession
    $collection = Invoke-PowerBIActivityCollection -DailyWindows $dailyWindows -MaximumItems $MaxItems
    $details = @($collection.Details)
    $userActivity = @(ConvertTo-PowerBIUserActivitySummary -Details $details)

    $detailResult = Export-SmartM365Csv -BaseFileName 'M365_PowerBI_Fabric_ActivityEvents' -OutputPath $OutputPath -GlobalPath $LatestCsvFolderPath -Data $details -Columns $detailColumns
    $userResult = Export-SmartM365Csv -BaseFileName 'M365_PowerBI_Fabric_UserActivity' -OutputPath $OutputPath -GlobalPath $LatestCsvFolderPath -Data $userActivity -Columns $userActivityColumns

    $summary = "UTC range=$($dateRange.FromDate.ToString('yyyy-MM-dd'))..$($dateRange.ToDate.ToString('yyyy-MM-dd')); Events=$($details.Count); Principals=$($userActivity.Count); Pages=$($collection.PageCount); DetailedCsv=$($detailResult.PublishedPath); UserActivityCsv=$($userResult.PublishedPath)"
    WriteLog -Message ("Power BI and Fabric activity inventory completed. {0}" -f $summary) -Level 'SUCCESS'
    Send-SmartM365TeamsNotification -Level SUCCESS -Channel Infos -Title 'SmartM365 Power BI and Fabric activity inventory completed' -Message $summary -ResultSummary $summary -Facts @{
        Tenant = $Tenant
        UtcRange = "$($dateRange.FromDate.ToString('yyyy-MM-dd'))..$($dateRange.ToDate.ToString('yyyy-MM-dd'))"
        Events = $details.Count
        Principals = $userActivity.Count
        Pages = $collection.PageCount
        RunId = $runId
        OutputPath = $LatestCsvFolderPath
    } | Out-Null

    Stop-PowerBIActivityTranscript
    Complete-SmartM365ExecutionContext -Status Auto
}
catch {
    $failure = $_
    $message = $failure.Exception.Message
    WriteLog -Message ("Power BI and Fabric activity inventory failed: {0}" -f $message) -Level 'ERROR'
    try {
        Send-SmartM365TeamsNotification -Level ERROR -Channel Alerts -Title 'SmartM365 Power BI and Fabric activity inventory failed' -Message $message -Facts @{
            Tenant = $Tenant
            UtcRange = "$($dateRange.FromDate.ToString('yyyy-MM-dd'))..$($dateRange.ToDate.ToString('yyyy-MM-dd'))"
            RunId = $runId
            LogPath = $logFilePath
            TranscriptPath = $transcriptPath
            OutputPath = $OutputPath
        } | Out-Null
    }
    catch {
        WriteLog -Message ("Teams alert notification failed: {0}" -f $_.Exception.Message) -Level 'WARNING'
    }
    Stop-PowerBIActivityTranscript
    try { Complete-SmartM365ExecutionContext -Status Failed -ErrorRecord $failure -FailureStage 'PowerBIFabricActivityInventory' } catch {}
    throw
}
finally {
    if ($script:PowerBIConnected -and (Get-Command -Name Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue)) {
        try { Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue | Out-Null } catch {}
    }
    try { RemoveOldFiles -Path $logFolder -Filter '*.log' -KeepCount $global:RetentionMaxLogs -LogFile $global:LogTextFile } catch {}
    try { Stop-PowerBIActivityTranscript } catch {}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCugqB/lIdYWZ0Y
# QBovWYzVwaAXcv6pdWYFmb3/t8mWFqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDPTp8dBrQcTdUNhY7a1WwfRqNIyjlrpnpnhcVeoKetHMA0GCSqG
# SIb3DQEBAQUABIIBgHIbzMKfXFxbEVs9+R3vA/RNes5KnkX4ZqK1pck2wyAQJZdv
# z2H83CJumcvTM+Zj5CQB/VZItR/PqFkO3FArqMZ6Kpc/+dctKMdU/l8QVSpc1Shq
# 0hdnp7tGmHhK83Xli+Ybv8FRyuC265w46gudgxF/PIXRH41P8CGYTKCg2DCNYI8U
# ifZWSdL2asYYJfot3pxo140R7ohjmQMqftrWlIwB2lQNOg6CNkicGsGQ+njEqGGc
# g1Cmlz/2dAh1/MvIIC2M4T3MNB0vqMXsJ1a1fj9poktuytS5UzzM5auuGYjkKqkH
# XOjcwluog+CL/3zWaBHoMsKhyOZcQ0BPChc+XAAWjGqJOTjL78KqZZKi77giOVYn
# E3w7ipfWpdt3Lm8v6vP1r+QBfRk+F2U58KtCukdfgEM8Sgo89YdkD4G0WLH2UgQq
# j1oyq1uQinOo0SpHQvbtR3REGyoNy5qlMKYgV1uO7YQsBTryYbMiyD1p9lPWvYQu
# ahGao/7pOSbyK3KS7aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTgyMTMx
# NTBaMC8GCSqGSIb3DQEJBDEiBCCzw8euepBCBOUW1n83ol1ht0MnAaVJYM6MN1PP
# 1lZaZjANBgkqhkiG9w0BAQEFAASCAgCeuMDAFoLRQfM75L7RmR6GrNzcrYT9fH6I
# XUjWLEx1euypHKvHHxUlUXe97zv56i/j8/EwPGV86mnMuZPqedspsWAg5DZx1+6g
# IWSdU1X+l/K8eYw4iwsNWStScWC+YTssvo3BUW2LSBOhar9bcaru/VQFQkXb+aDU
# yx+XMvtodIlKJzWdgkUCvfEEO1GdlklSNJfA/hD6nAuGhpodn887B6Ak84PerXsv
# wVy0LXEpZKrsytunL+hgTHsn6uxSEeY5BuQOkKUPofb2wZoWGlo3IV0YTraCRrfN
# S/9QWdHbXLvy6akTg8WBK+PG0bYmUHLsAQf4cYbpA0FmcA5RmiWf5wXzKEKhmuJR
# +ulSspklAfgoHA0tOtSCWIk9rIrwic3B6wIZVWg9gU6obxznyt7/GBZM5Tl3O79w
# 6NaBDlvTbjt/NiYuinZA55TRkQvgkkY8XUrNvacTtAPd8nZHBLmaeKCRqHSYtdk9
# oZrFlLn8zlrs9h2gO+tBazXgfci0br4dbQpRrzF4i9qXQQV8GfpkSnW3KR+AhC3b
# BmT+ekw/gLHvz5wosmSkvg5UQqR+kag+QBTU6i/Fwd/1tODpKNFoCDGpOdzlqw8d
# YCHQbDYiOOkVn4f83dRo3DiGUGhSCN7RQrVfRFxvZP8vYqjqwFYZhOc+jJq5Q/n2
# auE12z7k4g==
# SIG # End signature block
