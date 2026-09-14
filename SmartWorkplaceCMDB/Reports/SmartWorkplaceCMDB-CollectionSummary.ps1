<#
.SYNOPSIS
Creates and sends the SmartWorkplaceCMDB full-collection delta summary.

.DESCRIPTION
Builds a tenant-private aggregate snapshot from curated CMDB CSVs, compares it
with the previous full snapshot and the latest snapshots at or before 7 and 30
days, saves an HTML copy, and sends it through Microsoft Graph or SMTP.

.VERSION
1.3.5
#>
[CmdletBinding()]
param(
    [Alias('ProfileKey')][string]$Tenant = 'default',
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
    [string]$RunId = ([guid]::NewGuid().ToString('N')),
    [string]$RunStatus = 'Completed',
    [string]$OperationalError,
    [string]$FailedStep,
    [string]$FailureLogPath,
    [string]$FailureTranscriptPath,
    [datetimeoffset]$SnapshotDateTime = [datetimeoffset]::UtcNow,
    [switch]$CaptureBaselineOnly,
    [switch]$PreviewOnly,
    [switch]$ValidateOnly,
    [switch]$SendMailTestOnly,
    [switch]$ReuseLatestSnapshot,
    [switch]$NoConfigWrite
)

$ScriptVersion = '1.3.5'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Get-SmartWorkplaceCMDBSummarySetting {
    param($Configuration, [Parameter(Mandatory)][string]$Name, $DefaultValue)
    if ($null -eq $Configuration) { return $DefaultValue }
    if ($Configuration -is [Collections.IDictionary] -and $Configuration.Contains($Name)) {
        return $Configuration[$Name]
    }
    $property = $Configuration.PSObject.Properties[$Name]
    if ($null -ne $property) { return $property.Value }
    return $DefaultValue
}

function ConvertTo-SmartWorkplaceCMDBSummaryHtml {
    param([AllowNull()]$Value)
    return [Net.WebUtility]::HtmlEncode([string]$Value)
}

function Get-SmartWorkplaceCMDBSummaryMailLogo {
    param($NotificationConfiguration)

    $productRoot = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        ''
    }
    else {
        Split-Path -Parent $PSScriptRoot
    }
    $configuredPath = [string](Get-SmartWorkplaceCMDBSummarySetting `
        $NotificationConfiguration 'MailClientLogoPath' '')
    if ($configuredPath -in @('__USE_GLOBAL__','USE_GLOBAL')) { $configuredPath = '' }
    $configuredPath = [Environment]::ExpandEnvironmentVariables($configuredPath.Trim())
    if (-not [string]::IsNullOrWhiteSpace($configuredPath) -and
        -not [IO.Path]::IsPathRooted($configuredPath)) {
        $configuredPath = Join-Path $productRoot $configuredPath
    }

    $maxLogoKB = 200
    $configuredMaximum = Get-SmartWorkplaceCMDBSummarySetting `
        $NotificationConfiguration 'MailClientLogoMaxKB' 200
    [void][int]::TryParse([string]$configuredMaximum, [ref]$maxLogoKB)
    if ($maxLogoKB -lt 1) { $maxLogoKB = 200 }

    $clientName = [string](Get-SmartWorkplaceCMDBSummarySetting `
        $NotificationConfiguration 'MailClientName' 'Client')
    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($configuredPath)) {
        $candidates += [pscustomobject]@{
            Path = $configuredPath
            Alt = $clientName
            IsDefault = $false
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($productRoot)) {
        $candidates += [pscustomobject]@{
            Path = Join-Path $productRoot 'Assets\WorkplaceCloudHub-mail.png'
            Alt = 'WorkplaceCloudHub'
            IsDefault = $true
        }
    }

    foreach ($candidate in $candidates) {
        if (-not (Test-Path -LiteralPath $candidate.Path -PathType Leaf)) { continue }
        $item = Get-Item -LiteralPath $candidate.Path -ErrorAction Stop
        if ($item.Length -gt ([int64]$maxLogoKB * 1KB)) { continue }
        $extension = [IO.Path]::GetExtension($item.FullName).ToLowerInvariant()
        $mediaType = switch ($extension) {
            '.png'  { 'image/png' }
            '.jpg'  { 'image/jpeg' }
            '.jpeg' { 'image/jpeg' }
            '.gif'  { 'image/gif' }
            default { '' }
        }
        if ([string]::IsNullOrWhiteSpace($mediaType)) { continue }
        return [pscustomobject]@{
            Path = $item.FullName
            Alt = $candidate.Alt
            IsDefault = $candidate.IsDefault
            MediaType = $mediaType
        }
    }
    return $null
}

function Add-SmartWorkplaceCMDBSummaryMailBranding {
    param(
        [AllowNull()][string]$BodyHtml,
        $NotificationConfiguration
    )

    if ([string]::IsNullOrWhiteSpace($BodyHtml)) { $BodyHtml = '' }
    if ($BodyHtml -match 'SmartWorkplaceCMDBMailBranding:v1') { return $BodyHtml }
    $logo = Get-SmartWorkplaceCMDBSummaryMailLogo $NotificationConfiguration
    if ($null -eq $logo) {
        return $BodyHtml -replace '(?is)(<body\b[^>]*>)', '$1<!-- SmartWorkplaceCMDBMailBranding:v1 -->'
    }

    $dataUri = 'data:{0};base64,{1}' -f $logo.MediaType,
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($logo.Path))
    $safeDataUri = ConvertTo-SmartWorkplaceCMDBSummaryHtml $dataUri
    $safeAlt = ConvertTo-SmartWorkplaceCMDBSummaryHtml $logo.Alt
    $imageHtml = '<img src="{0}" alt="{1}" width="104" style="display:block;margin:0 auto;width:104px;max-width:104px;height:auto;border:0;outline:none;text-decoration:none;" />' -f `
        $safeDataUri,$safeAlt
    if ($logo.IsDefault) {
        $imageHtml = '<a href="https://workplacecloudhub.com/" style="display:inline-block;text-decoration:none;">{0}</a>' -f $imageHtml
    }
    $brandingHtml = @"
<!-- SmartWorkplaceCMDBMailBranding:v1 -->
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#eef3f8;border-collapse:collapse;">
  <tr><td align="center" style="padding:12px 24px 10px 24px;text-align:center;">$imageHtml</td></tr>
</table>
"@
    if ($BodyHtml -match '(?is)<body\b[^>]*>') {
        return [regex]::Replace($BodyHtml, '(?is)(<body\b[^>]*>)', ('$1' + "`n" + $brandingHtml), 1)
    }
    return ($brandingHtml + "`n" + $BodyHtml)
}

function ConvertTo-SmartWorkplaceCMDBSummaryInt64 {
    param([AllowNull()]$Value)
    $number = [int64]0
    if ([int64]::TryParse([string]$Value, [Globalization.NumberStyles]::Integer,
            [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }
    return [int64]0
}

function Get-SmartWorkplaceCMDBSummaryFingerprint {
    param([Parameter(Mandatory)][string[]]$Path)
    $source = @($Path | Sort-Object | ForEach-Object {
            '{0}|{1}' -f [IO.Path]::GetFileName($_),
            (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash
        }) -join "`n"
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString(
                $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($source))) -replace '-', '')
    }
    finally { $sha.Dispose() }
}

function Get-SmartWorkplaceCMDBLicenseFamily {
    param([string]$SkuPartNumber, $NotificationConfiguration)
    $defaults = [ordered]@{
        F1 = '^(M365_F1|Microsoft_365_F1)(?:_.*)?$'
        F3 = '^(SPE_F1|SPE_F3|M365_F3|Microsoft_365_F3)(?:_.*)?$'
        E3 = '^(SPE_E3|M365_E3|Microsoft_365_E3)(?:_.*)?$'
        E5 = '^(SPE_E5|M365_E5|Microsoft_365_E5)(?:_.*)?$'
        Copilot = '^(Microsoft_365_Copilot|M365_COPILOT)(?:_.*)?$'
    }
    $configured = Get-SmartWorkplaceCMDBSummarySetting `
        $NotificationConfiguration 'LicenseFamilyPatterns' $null
    foreach ($family in $defaults.Keys) {
        $pattern = [string](Get-SmartWorkplaceCMDBSummarySetting `
                $configured $family $defaults[$family])
        if (-not [string]::IsNullOrWhiteSpace($pattern) -and
            $SkuPartNumber -match $pattern) {
            return $family
        }
    }
    return ''
}

function New-SmartWorkplaceCMDBSummarySnapshot {
    param(
        [Parameter(Mandatory)]$Paths,
        [Parameter(Mandatory)]$NotificationConfiguration,
        [Parameter(Mandatory)][string]$SnapshotRunId,
        [Parameter(Mandatory)][datetimeoffset]$DateTime,
        [Parameter(Mandatory)][string]$Status
    )
    $cmdbRoot = Join-Path $Paths.LatestOutputRootPath 'CMDB'
    $files = [ordered]@{
        Devices = Join-Path $cmdbRoot 'CMDB_Devices.csv'
        Users = Join-Path $cmdbRoot 'CMDB_Users.csv'
        Mailboxes = Join-Path $cmdbRoot 'CMDB_Mailboxes.csv'
        Licenses = Join-Path $cmdbRoot 'CMDB_Licenses.csv'
        Relationships = Join-Path $cmdbRoot 'CMDB_Relationships.csv'
        SharePointSites = Join-Path $Paths.LatestOutputRootPath 'PowerBI\DimSharePointSite.csv'
        Teams = Join-Path $Paths.LatestOutputRootPath 'PowerBI\DimTeam.csv'
        TeamMembers = Join-Path $Paths.LatestOutputRootPath 'PowerBI\FactTeamMember.csv'
    }
    foreach ($path in $files.Values) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "A required curated summary input is missing: '$path'."
        }
    }

    $licenses = @(Import-Csv -LiteralPath $files.Licenses)
    $relationships = @(Import-Csv -LiteralPath $files.Relationships)
    $sharePointSites = @(Import-Csv -LiteralPath $files.SharePointSites)
    $teams = @(Import-Csv -LiteralPath $files.Teams)
    $teamMembers = @(Import-Csv -LiteralPath $files.TeamMembers)
    $familyByLicenseId = @{}
    $capacity = @{F1=0L;F3=0L;E3=0L;E5=0L;Copilot=0L}
    foreach ($license in $licenses) {
        $family = Get-SmartWorkplaceCMDBLicenseFamily `
            -SkuPartNumber ([string]$license.SkuPartNumber) `
            -NotificationConfiguration $NotificationConfiguration
        if (-not [string]::IsNullOrWhiteSpace($family)) {
            $familyByLicenseId[[string]$license.CmdbLicenseId] = $family
            $capacity[$family] += ConvertTo-SmartWorkplaceCMDBSummaryInt64 $license.EnabledUnits
        }
    }

    $assignedUsers = @{
        F1 = New-Object 'System.Collections.Generic.HashSet[string]'
        F3 = New-Object 'System.Collections.Generic.HashSet[string]'
        E3 = New-Object 'System.Collections.Generic.HashSet[string]'
        E5 = New-Object 'System.Collections.Generic.HashSet[string]'
        Copilot = New-Object 'System.Collections.Generic.HashSet[string]'
    }
    foreach ($relationship in $relationships) {
        if ([string]$relationship.RelationshipType -ne 'AssignedLicense' -or
            [string]$relationship.FromEntityType -ne 'User' -or
            [string]$relationship.ToEntityType -ne 'License') { continue }
        $licenseId = [string]$relationship.ToEntityId
        if ($familyByLicenseId.ContainsKey($licenseId)) {
            [void]$assignedUsers[$familyByLicenseId[$licenseId]].Add(
                [string]$relationship.FromEntityId)
        }
    }

    $completeTeams = @($teams | Where-Object MembershipCoverageStatus -eq 'Complete')
    $sharePointStorageUsedBytes = ($sharePointSites | Measure-Object StorageUsedBytes -Sum).Sum
    $sharePointStorageAllocatedBytes = ($sharePointSites | Measure-Object StorageAllocatedBytes -Sum).Sum
    return [pscustomobject][ordered]@{
        Version = 2
        RunId = $SnapshotRunId
        RunStatus = $Status
        SnapshotDateTime = $DateTime.ToUniversalTime().ToString('o')
        TenantKey = $Paths.TenantKey
        DatasetFingerprint = Get-SmartWorkplaceCMDBSummaryFingerprint @($files.Values)
        Devices = @(Import-Csv -LiteralPath $files.Devices).Count
        Users = @(Import-Csv -LiteralPath $files.Users).Count
        Mailboxes = @(Import-Csv -LiteralPath $files.Mailboxes).Count
        SharePointSites = $sharePointSites.Count
        SharePointActiveSites = @($sharePointSites | Where-Object ActivityState -eq 'Active (90d)').Count
        SharePointInactiveSites = @($sharePointSites | Where-Object ActivityState -eq 'Inactive (>90d)').Count
        SharePointOwnerMissingSites = @($sharePointSites | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.OwnerPrincipalName) }).Count
        SharePointStorageUsedGB = [Math]::Round(($sharePointStorageUsedBytes / 1GB), 1)
        SharePointStorageAllocatedGB = [Math]::Round(($sharePointStorageAllocatedBytes / 1GB), 1)
        SharePointStorageUtilizationPercent = if ($sharePointStorageAllocatedBytes -gt 0) { [Math]::Round((100 * $sharePointStorageUsedBytes / $sharePointStorageAllocatedBytes), 1) } else { $null }
        SharePointAverageStorageGB = if ($sharePointSites.Count -gt 0) { [Math]::Round((($sharePointSites | Measure-Object StorageUsedBytes -Average).Average / 1GB), 1) } else { 0 }
        Teams = $teams.Count
        TeamsActive = @($teams | Where-Object ActivityState -eq 'Active (90d)').Count
        TeamsInactive = @($teams | Where-Object ActivityState -eq 'Inactive (>90d)').Count
        TeamsArchived = @($teams | Where-Object IsArchived -eq 'True').Count
        TeamsWithoutOwner = @($teams | Where-Object { (ConvertTo-SmartWorkplaceCMDBSummaryInt64 $_.OwnerCount) -eq 0 }).Count
        TeamsWithGuests = @($teams | Where-Object { (ConvertTo-SmartWorkplaceCMDBSummaryInt64 $_.GuestCount) -gt 0 }).Count
        TeamsMembershipPartial = @($teams | Where-Object MembershipCoverageStatus -eq 'Partial').Count
        TeamMembers = $teamMembers.Count
        TeamGuests = @($teamMembers | Where-Object UserType -eq 'Guest').Count
        TeamAverageMembers = if ($completeTeams.Count -gt 0) { [Math]::Round(((($completeTeams | ForEach-Object { ConvertTo-SmartWorkplaceCMDBSummaryInt64 $_.MemberCount }) | Measure-Object -Sum).Sum / $completeTeams.Count), 1) } else { 0 }
        M365F1Assigned = $assignedUsers.F1.Count
        M365F1Capacity = $capacity.F1
        M365F3Assigned = $assignedUsers.F3.Count
        M365F3Capacity = $capacity.F3
        M365E3Assigned = $assignedUsers.E3.Count
        M365E3Capacity = $capacity.E3
        M365E5Assigned = $assignedUsers.E5.Count
        M365E5Capacity = $capacity.E5
        M365CopilotAssigned = $assignedUsers.Copilot.Count
        M365CopilotCapacity = $capacity.Copilot
    }
}

function Get-SmartWorkplaceCMDBSummaryHistory {
    param([Parameter(Mandatory)][string]$HistoryRootPath)
    if (-not (Test-Path -LiteralPath $HistoryRootPath)) { return @() }
    return @(Get-ChildItem -LiteralPath $HistoryRootPath -Filter '*.csv' -File -Recurse |
        ForEach-Object {
            try { Import-Csv -LiteralPath $_.FullName | Select-Object -First 1 }
            catch { Write-Warning "Ignored unreadable summary snapshot '$($_.FullName)'." }
        } | Where-Object { $null -ne $_ } | Sort-Object {
            [datetimeoffset]::Parse([string]$_.SnapshotDateTime)
        })
}

function Save-SmartWorkplaceCMDBSummarySnapshot {
    param([Parameter(Mandatory)]$Snapshot, [Parameter(Mandatory)]$Paths)
    $date = [datetimeoffset]::Parse([string]$Snapshot.SnapshotDateTime)
    $historyFolder = Join-Path $Paths.DataAllRootPath (
        'CollectionSummary\{0}\{1}' -f $date.ToString('yyyy'), $date.ToString('MM'))
    $historyPath = Join-Path $historyFolder (
        'SmartWorkplaceCMDB_CollectionSummary_{0}_{1}.csv' -f
        $date.ToString('yyyyMMdd-HHmmssfff'), $Snapshot.RunId)
    $latestPath = Join-Path $Paths.LatestOutputRootPath `
        'Summary\SmartWorkplaceCMDB_CollectionSummary.csv'
    foreach ($path in @($historyPath, $latestPath)) {
        New-Item -ItemType Directory -Path (Split-Path $path -Parent) -Force | Out-Null
        @($Snapshot) | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
    }
    return [pscustomobject]@{HistoryPath=$historyPath;LatestPath=$latestPath}
}

function Format-SmartWorkplaceCMDBSummaryNumber {
    param([AllowNull()]$Value)
    return (ConvertTo-SmartWorkplaceCMDBSummaryInt64 $Value).ToString('N0',
        [Globalization.CultureInfo]::GetCultureInfo('en-US'))
}

function Format-SmartWorkplaceCMDBSummaryDelta {
    param([AllowNull()]$Current, [AllowNull()]$Reference)
    if ($null -eq $Reference) { return 'n/a' }
    $delta = (ConvertTo-SmartWorkplaceCMDBSummaryInt64 $Current) -
        (ConvertTo-SmartWorkplaceCMDBSummaryInt64 $Reference)
    if ($delta -gt 0) { return '+' + (Format-SmartWorkplaceCMDBSummaryNumber $delta) }
    return Format-SmartWorkplaceCMDBSummaryNumber $delta
}

function Get-SmartWorkplaceCMDBSummaryProperty {
    param([AllowNull()]$Snapshot,[Parameter(Mandatory)][string]$Name)
    if ($null -eq $Snapshot) { return $null }
    $property = $Snapshot.PSObject.Properties[$Name]
    if ($null -eq $property -or [string]::IsNullOrWhiteSpace([string]$property.Value)) { return $null }
    return $property.Value
}

function Format-SmartWorkplaceCMDBSummaryMetric {
    param([AllowNull()]$Value,[string]$Format='integer')
    if ($null -eq $Value) { return 'n/a' }
    if ($Format -in @('decimal','percent')) {
        $number = [double]0
        if (-not [double]::TryParse([string]$Value,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$number)) { return 'n/a' }
        $text=$number.ToString('N1',[Globalization.CultureInfo]::GetCultureInfo('en-US'))
        if($Format-eq'percent'){return $text+'%'}
        return $text
    }
    return Format-SmartWorkplaceCMDBSummaryNumber $Value
}

function Format-SmartWorkplaceCMDBSummaryMetricDelta {
    param([AllowNull()]$Current,[AllowNull()]$Reference,[string]$Format='integer')
    if ($null -eq $Reference) { return 'n/a' }
    if ($Format -in @('decimal','percent')) {
        $currentNumber=[double]0;$referenceNumber=[double]0
        if (-not [double]::TryParse([string]$Current,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$currentNumber) -or -not [double]::TryParse([string]$Reference,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$referenceNumber)) { return 'n/a' }
        $delta=$currentNumber-$referenceNumber;$text=[Math]::Round($delta,1).ToString('N1',[Globalization.CultureInfo]::GetCultureInfo('en-US'));if($Format-eq'percent'){$text+=' pp'};if($delta -gt 0){return '+'+$text};return $text
    }
    return Format-SmartWorkplaceCMDBSummaryDelta $Current $Reference
}

function New-SmartWorkplaceCMDBSummaryHtml {
    param($Current, $Previous, $Day7, $Day30, $NotificationConfiguration)
    $metrics = @(
        @{Section='Core population';Label='Devices';Property='Devices';Capacity='';Format='integer'},
        @{Section='Core population';Label='Users';Property='Users';Capacity='';Format='integer'},
        @{Section='Core population';Label='Mailboxes';Property='Mailboxes';Capacity='';Format='integer'},
        @{Section='SharePoint';Label='SharePoint sites';Property='SharePointSites';Capacity='';Format='integer'},
        @{Section='SharePoint';Label='Active sites (90d)';Property='SharePointActiveSites';Capacity='';Format='integer'},
        @{Section='SharePoint';Label='Inactive sites (>90d)';Property='SharePointInactiveSites';Capacity='';Format='integer'},
        @{Section='SharePoint';Label='Sites without reported owner';Property='SharePointOwnerMissingSites';Capacity='';Format='integer'},
        @{Section='SharePoint';Label='Storage used (GB)';Property='SharePointStorageUsedGB';Capacity='';Format='decimal'},
        @{Section='SharePoint';Label='Storage allocated (GB)';Property='SharePointStorageAllocatedGB';Capacity='';Format='decimal'},
        @{Section='SharePoint';Label='Storage utilization';Property='SharePointStorageUtilizationPercent';Capacity='';Format='percent'},
        @{Section='SharePoint';Label='Average storage per site (GB)';Property='SharePointAverageStorageGB';Capacity='';Format='decimal'},
        @{Section='Teams';Label='Teams';Property='Teams';Capacity='';Format='integer'},
        @{Section='Teams';Label='Active Teams (90d)';Property='TeamsActive';Capacity='';Format='integer'},
        @{Section='Teams';Label='Inactive Teams (>90d)';Property='TeamsInactive';Capacity='';Format='integer'},
        @{Section='Teams';Label='Archived Teams';Property='TeamsArchived';Capacity='';Format='integer'},
        @{Section='Teams';Label='Teams without owner';Property='TeamsWithoutOwner';Capacity='';Format='integer'},
        @{Section='Teams';Label='Teams with guests';Property='TeamsWithGuests';Capacity='';Format='integer'},
        @{Section='Teams';Label='Teams with partial membership coverage';Property='TeamsMembershipPartial';Capacity='';Format='integer'},
        @{Section='Teams';Label='Unique Team memberships';Property='TeamMembers';Capacity='';Format='integer'},
        @{Section='Teams';Label='Guest memberships';Property='TeamGuests';Capacity='';Format='integer'},
        @{Section='Teams';Label='Average members per Team';Property='TeamAverageMembers';Capacity='';Format='decimal'},
        @{Section='Microsoft 365 licensing';Label='Microsoft 365 F1 assignments';Property='M365F1Assigned';Capacity='M365F1Capacity';Format='integer'},
        @{Section='Microsoft 365 licensing';Label='Microsoft 365 F3 assignments';Property='M365F3Assigned';Capacity='M365F3Capacity';Format='integer'},
        @{Section='Microsoft 365 licensing';Label='Microsoft 365 E3 assignments';Property='M365E3Assigned';Capacity='M365E3Capacity';Format='integer'},
        @{Section='Microsoft 365 licensing';Label='Microsoft 365 E5 assignments';Property='M365E5Assigned';Capacity='M365E5Capacity';Format='integer'},
        @{Section='Microsoft 365 licensing';Label='Microsoft 365 Copilot assignments';Property='M365CopilotAssigned';Capacity='M365CopilotCapacity';Format='integer'}
    )
    $section = ''
    $rows = foreach ($metric in $metrics) {
        if ($metric.Section -ne $section) { $section=$metric.Section;'<tr class="section"><td colspan="5">{0}</td></tr>' -f (ConvertTo-SmartWorkplaceCMDBSummaryHtml $section) }
        $currentValue = $Current.($metric.Property)
        $currentText = Format-SmartWorkplaceCMDBSummaryMetric $currentValue $metric.Format
        if (-not [string]::IsNullOrWhiteSpace($metric.Capacity)) {
            $capacity = ConvertTo-SmartWorkplaceCMDBSummaryInt64 $Current.($metric.Capacity)
            $ratio = if ($capacity -gt 0) {
                '{0:P1}' -f ((ConvertTo-SmartWorkplaceCMDBSummaryInt64 $currentValue) / $capacity)
            } else { 'n/a' }
            $currentText = '{0} / {1} ({2})' -f $currentText,
                (Format-SmartWorkplaceCMDBSummaryNumber $capacity), $ratio
        }
        '<tr><td>{0}</td><td class="number">{1}</td><td class="number">{2}</td><td class="number">{3}</td><td class="number">{4}</td></tr>' -f
            (ConvertTo-SmartWorkplaceCMDBSummaryHtml $metric.Label),
            (ConvertTo-SmartWorkplaceCMDBSummaryHtml $currentText),
            (ConvertTo-SmartWorkplaceCMDBSummaryHtml (Format-SmartWorkplaceCMDBSummaryMetricDelta $currentValue (Get-SmartWorkplaceCMDBSummaryProperty $Previous $metric.Property) $metric.Format)),
            (ConvertTo-SmartWorkplaceCMDBSummaryHtml (Format-SmartWorkplaceCMDBSummaryMetricDelta $currentValue (Get-SmartWorkplaceCMDBSummaryProperty $Day7 $metric.Property) $metric.Format)),
            (ConvertTo-SmartWorkplaceCMDBSummaryHtml (Format-SmartWorkplaceCMDBSummaryMetricDelta $currentValue (Get-SmartWorkplaceCMDBSummaryProperty $Day30 $metric.Property) $metric.Format))
    }
    $client = [string](Get-SmartWorkplaceCMDBSummarySetting `
        $NotificationConfiguration 'MailClientName' 'Smart Workplace')
    $previousLabel = if ($Previous) { [datetimeoffset]::Parse([string]$Previous.SnapshotDateTime).ToString('yyyy-MM-dd HH:mm UTC') } else { 'not available' }
    $currentLabel = [datetimeoffset]::Parse([string]$Current.SnapshotDateTime).ToString('yyyy-MM-dd HH:mm UTC')
    $day7Label = if ($Day7) { [datetimeoffset]::Parse([string]$Day7.SnapshotDateTime).ToString('yyyy-MM-dd HH:mm UTC') } else { 'not available' }
    $day30Label = if ($Day30) { [datetimeoffset]::Parse([string]$Day30.SnapshotDateTime).ToString('yyyy-MM-dd HH:mm UTC') } else { 'not available' }
    return @"
<!doctype html><html><head><meta charset="utf-8"><style>
body{font-family:Segoe UI,Arial,sans-serif;background:#eef3f8;color:#172033;margin:0;padding:24px}.card{max-width:1040px;margin:auto;background:#fff;border:1px solid #d9e5f0;border-radius:14px;overflow:hidden;box-shadow:0 8px 24px rgba(15,23,42,.08)}.header{background:linear-gradient(125deg,#0f766e,#2563eb);color:#fff;padding:26px 30px}.header h1{font-size:24px;margin:0}.header p{margin:8px 0 0;color:#e5f4ff}.pill{display:inline-block;background:#107c10;color:#fff;border-radius:999px;padding:4px 10px;font-size:11px;font-weight:700}.content{padding:24px 28px}.kpis{width:100%;border-spacing:10px;border-collapse:separate;margin:0 -10px 14px}.kpis td{width:25%;background:#f8fafc;border:1px solid #dde7f0;border-radius:9px;padding:14px}.kpiLabel{font-size:11px;color:#64748b;text-transform:uppercase}.kpiValue{font-size:23px;font-weight:700;color:#0f4c81;margin-top:4px}table.data{width:100%;border-collapse:collapse;margin-top:18px}table.data th{background:#eaf3fb;color:#23415d;text-align:left;font-size:12px;padding:10px;border-bottom:1px solid #c7d9e8}table.data td{padding:9px 10px;border-bottom:1px solid #e4edf5;font-size:13px}.number{text-align:right!important;white-space:nowrap}.section td{background:#f5f8fb;color:#0f4c81;font-weight:700;text-transform:uppercase;font-size:11px!important;letter-spacing:.04em}.meta{font-size:12px;color:#52677b;line-height:1.55}.notice{margin-top:18px;background:#fff7e6;border-left:4px solid #ff8c00;padding:11px 13px;color:#6b4c00;font-size:12px}.footer{padding:16px 28px;background:#f7fafc;color:#66788a;font-size:11px}</style></head><body>
<div class="card"><div class="header"><h1>Smart Workplace CMDB — full collection summary</h1><p>$(ConvertTo-SmartWorkplaceCMDBSummaryHtml $client) &nbsp; <span class="pill">$(ConvertTo-SmartWorkplaceCMDBSummaryHtml $Current.RunStatus)</span></p></div><div class="content">
<table class="kpis"><tr><td><div class="kpiLabel">Devices</div><div class="kpiValue">$(Format-SmartWorkplaceCMDBSummaryNumber $Current.Devices)</div></td><td><div class="kpiLabel">Users</div><div class="kpiValue">$(Format-SmartWorkplaceCMDBSummaryNumber $Current.Users)</div></td><td><div class="kpiLabel">SharePoint sites</div><div class="kpiValue">$(Format-SmartWorkplaceCMDBSummaryNumber $Current.SharePointSites)</div></td><td><div class="kpiLabel">Teams</div><div class="kpiValue">$(Format-SmartWorkplaceCMDBSummaryNumber $Current.Teams)</div></td></tr></table>
<div class="meta">Current snapshot: $(ConvertTo-SmartWorkplaceCMDBSummaryHtml $currentLabel)<br>Previous collection: $(ConvertTo-SmartWorkplaceCMDBSummaryHtml $previousLabel)<br>J-7 baseline: $(ConvertTo-SmartWorkplaceCMDBSummaryHtml $day7Label)<br>J-30 baseline: $(ConvertTo-SmartWorkplaceCMDBSummaryHtml $day30Label)</div>
<table class="data"><thead><tr><th>Metric</th><th class="number">Current<br>$(ConvertTo-SmartWorkplaceCMDBSummaryHtml $currentLabel)</th><th class="number">Since previous<br>$(ConvertTo-SmartWorkplaceCMDBSummaryHtml $previousLabel)</th><th class="number">Since J-7<br>$(ConvertTo-SmartWorkplaceCMDBSummaryHtml $day7Label)</th><th class="number">Since J-30<br>$(ConvertTo-SmartWorkplaceCMDBSummaryHtml $day30Label)</th></tr></thead><tbody>$($rows -join "`n")</tbody></table>
<div class="notice"><b>Membership coverage:</b> SharePoint member counts are not present in the Graph site-usage report and are therefore not estimated. The Teams average includes only Teams with complete exact membership enumeration; partial Teams are counted separately. Active means activity reported within 90 days; missing activity remains unknown.</div>
</div><div class="footer">© 2026 WorkplaceCloudHub — https://workplacecloudhub.com/ · Aggregate notification only; no user, device, mailbox address, or tenant row is included.</div></div></body></html>
"@
}

function Get-SmartWorkplaceCMDBSummaryRetryDelay {
    param([Parameter(Mandatory)]$ErrorRecord, [int]$Attempt, [int]$MaximumSeconds = 120)
    $value = $null
    try {
        if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.Headers) {
            $headers = $ErrorRecord.Exception.Response.Headers
            try { $value = @($headers.GetValues('Retry-After') | Select-Object -First 1)[0] }
            catch { try { $value = $headers['Retry-After'] } catch {} }
        }
    }
    catch {}
    if ($null -eq $value) {
        try { $value = $ErrorRecord.Exception.Data['Retry-After'] }
        catch {}
    }
    $seconds = 0
    if ($null -ne $value -and [int]::TryParse([string]$value, [ref]$seconds) -and $seconds -gt 0) {
        return [Math]::Min($seconds, $MaximumSeconds)
    }
    $retryDate = [datetimeoffset]::MinValue
    if ($null -ne $value -and [datetimeoffset]::TryParse([string]$value, [ref]$retryDate)) {
        $dateDelay = [int][Math]::Ceiling(($retryDate.UtcDateTime - [datetime]::UtcNow).TotalSeconds)
        if ($dateDelay -gt 0) { return [Math]::Min($dateDelay, $MaximumSeconds) }
    }
    return [Math]::Min($MaximumSeconds, (5 * [Math]::Pow(2, [Math]::Max(0, $Attempt - 1))))
}

function Test-SmartWorkplaceCMDBSummaryTransientError {
    param([Parameter(Mandatory)]$ErrorRecord)
    $statusCode = $null
    try {
        if ($ErrorRecord.Exception.Response) {
            $statusCode = [int]$ErrorRecord.Exception.Response.StatusCode
        }
    }
    catch {}
    if ($statusCode -in @(408, 429, 500, 502, 503, 504)) { return $true }
    return ([string]$ErrorRecord.Exception.Message -match '(?i)throttl|too many requests|temporar|timeout|timed out|connection.*closed')
}

function Invoke-SmartWorkplaceCMDBSummaryGraphMailRequest {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$Body,
        [int]$MaxAttempts = 4
    )
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $Body `
                -ContentType 'application/json' -ErrorAction Stop | Out-Null
            return
        }
        catch {
            if (-not (Test-SmartWorkplaceCMDBSummaryTransientError -ErrorRecord $_) -or
                $attempt -ge $MaxAttempts) {
                $statusCode = 0
                try {
                    if ($_.Exception.Response) {
                        $statusCode = [int]$_.Exception.Response.StatusCode
                    }
                }
                catch {}
                $details = New-Object System.Collections.Generic.List[string]
                if (-not [string]::IsNullOrWhiteSpace([string]$_.Exception.Message)) {
                    $details.Add([string]$_.Exception.Message)
                }
                if ($_.ErrorDetails -and
                    -not [string]::IsNullOrWhiteSpace([string]$_.ErrorDetails.Message)) {
                    $details.Add([string]$_.ErrorDetails.Message)
                }
                try {
                    $responseContent = [string]$_.Exception.Response.Content
                    if (-not [string]::IsNullOrWhiteSpace($responseContent)) {
                        $details.Add($responseContent)
                    }
                }
                catch {}
                $detailText = @($details | Select-Object -Unique) -join ' | '
                throw "Graph mail request failed. Status=$statusCode; $detailText"
            }
            $delay = Get-SmartWorkplaceCMDBSummaryRetryDelay -ErrorRecord $_ -Attempt $attempt
            Write-Warning ("Graph mail transient failure; attempt {0}/{1}; retrying in {2}s." -f
                $attempt, $MaxAttempts, $delay)
            Start-Sleep -Seconds $delay
        }
    }
}

function Invoke-SmartWorkplaceCMDBSummaryGraphMailWithTimeout {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$CertificateThumbprint,
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$Body,
        [ValidateRange(1, 900)][int]$TimeoutSeconds = 120,
        [ValidateRange(1, 10)][int]$MaxAttempts = 4,
        [scriptblock]$WorkerScriptBlock
    )

    if ($null -eq $WorkerScriptBlock) {
        $WorkerScriptBlock = {
            param($WorkerTenantId, $WorkerClientId, $WorkerThumbprint,
                $WorkerUri, $WorkerBody, $WorkerMaxAttempts)

            $ErrorActionPreference = 'Stop'
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            $connected = $false
            try {
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                Connect-MgGraph -TenantId $WorkerTenantId `
                    -ClientId $WorkerClientId `
                    -CertificateThumbprint $WorkerThumbprint `
                    -ContextScope Process -NoWelcome -ErrorAction Stop | Out-Null
                $connected = $true
                for ($attempt = 1; $attempt -le $WorkerMaxAttempts; $attempt++) {
                    try {
                        Invoke-MgGraphRequest -Method POST -Uri $WorkerUri `
                            -Body $WorkerBody -ContentType 'application/json' `
                            -ErrorAction Stop | Out-Null
                        return
                    }
                    catch {
                        $statusCode = 0
                        try {
                            if ($_.Exception.Response) {
                                $statusCode = [int]$_.Exception.Response.StatusCode
                            }
                        }
                        catch {}
                        $transient = $statusCode -in @(408, 429, 500, 502, 503, 504) -or
                            [string]$_.Exception.Message -match '(?i)throttl|too many requests|temporar|timeout|timed out|connection.*closed'
                        if (-not $transient -or $attempt -ge $WorkerMaxAttempts) {
                            $details = New-Object System.Collections.Generic.List[string]
                            if (-not [string]::IsNullOrWhiteSpace([string]$_.Exception.Message)) {
                                $details.Add([string]$_.Exception.Message)
                            }
                            if ($_.ErrorDetails -and
                                -not [string]::IsNullOrWhiteSpace([string]$_.ErrorDetails.Message)) {
                                $details.Add([string]$_.ErrorDetails.Message)
                            }
                            throw "Graph mail request failed. Status=$statusCode; $(@($details | Select-Object -Unique) -join ' | ')"
                        }
                        $retryAfter = $null
                        try {
                            if ($_.Exception.Response -and
                                $_.Exception.Response.Headers) {
                                try {
                                    $retryAfter = @(
                                        $_.Exception.Response.Headers.GetValues(
                                            'Retry-After') |
                                            Select-Object -First 1
                                    )[0]
                                }
                                catch {
                                    try {
                                        $retryAfter =
                                            $_.Exception.Response.Headers['Retry-After']
                                    }
                                    catch {}
                                }
                            }
                        }
                        catch {}
                        $delaySeconds = 0
                        if ($null -eq $retryAfter -or
                            -not [int]::TryParse(
                                [string]$retryAfter,
                                [ref]$delaySeconds) -or
                            $delaySeconds -lt 1) {
                            $delaySeconds = 5 * [Math]::Pow(
                                2,
                                [Math]::Max(0, $attempt - 1)
                            )
                        }
                        $delaySeconds = [Math]::Min(120, $delaySeconds)
                        Start-Sleep -Seconds $delaySeconds
                    }
                }
            }
            finally {
                if ($connected) {
                    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                }
            }
        }
    }

    Import-Module Microsoft.PowerShell.ThreadJob -ErrorAction Stop
    $job = Start-ThreadJob -ScriptBlock $WorkerScriptBlock -ArgumentList @(
        $TenantId,
        $ClientId,
        $CertificateThumbprint,
        $Uri,
        $Body,
        $MaxAttempts
    )
    try {
        $completed = Wait-Job -Job $job -Timeout $TimeoutSeconds
        if ($null -eq $completed) {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
            throw "Graph mail operation exceeded the configured timeout of $TimeoutSeconds seconds."
        }
        Receive-Job -Job $job -ErrorAction Stop | Out-Null
        if ($job.State -ne 'Completed') {
            throw "Graph mail worker ended with state '$($job.State)'."
        }
    }
    finally {
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    }
}

function New-SmartWorkplaceCMDBSummaryGraphMailBody {
    param(
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$BodyHtml,
        [Parameter(Mandatory)][object[]]$ToRecipients,
        [object[]]$CcRecipients = @()
    )

    if ($ToRecipients.Count -eq 0) {
        throw 'At least one Graph mail recipient is required.'
    }
    $message = @{
        subject = $Subject
        body = @{contentType='HTML';content=$BodyHtml}
        toRecipients = [object[]]$ToRecipients
    }
    if ($CcRecipients.Count -gt 0) {
        $message.ccRecipients = [object[]]$CcRecipients
    }
    return @{message=$message;saveToSentItems=$false} |
        ConvertTo-Json -Depth 12
}

function Send-SmartWorkplaceCMDBSummaryGraphMail {
    param($NotificationConfiguration, $GraphConfiguration, [string]$ResolvedTenantId,
        [string]$Subject, [string]$BodyHtml)
    $from = [string](Get-SmartWorkplaceCMDBSummarySetting $NotificationConfiguration 'From' '')
    $to = @(([string](Get-SmartWorkplaceCMDBSummarySetting $NotificationConfiguration 'To' '')) -split '[;,]' | Where-Object {$_.Trim()})
    $cc = @(([string](Get-SmartWorkplaceCMDBSummarySetting $NotificationConfiguration 'Cc' '')) -split '[;,]' | Where-Object {$_.Trim()})
    if ([string]::IsNullOrWhiteSpace($from) -or $to.Count -eq 0) { throw 'Notifications.From and Notifications.To are required.' }
    $clientId = [string](Get-SmartWorkplaceCMDBSummarySetting $GraphConfiguration 'ClientId' '')
    $thumbprint = [string](Get-SmartWorkplaceCMDBSummarySetting $GraphConfiguration 'CertificateThumbprint' '')
    if ([string]::IsNullOrWhiteSpace($ResolvedTenantId) -or
        [string]::IsNullOrWhiteSpace($clientId) -or
        [string]::IsNullOrWhiteSpace($thumbprint)) {
        throw 'MicrosoftGraph.TenantId, ClientId, and CertificateThumbprint are required for Graph mail.'
    }
    $toRecipients = @($to | ForEach-Object {
            @{emailAddress=@{address=$_.Trim()}}
        })
    $ccRecipients = @($cc | ForEach-Object {
            @{emailAddress=@{address=$_.Trim()}}
        })
    $body = New-SmartWorkplaceCMDBSummaryGraphMailBody `
        -Subject $Subject `
        -BodyHtml $BodyHtml `
        -ToRecipients $toRecipients `
        -CcRecipients $ccRecipients
    $uri = 'https://graph.microsoft.com/v1.0/users/{0}/sendMail' -f [uri]::EscapeDataString($from)
    $timeoutSeconds = 0
    $configuredTimeout = Get-SmartWorkplaceCMDBSummarySetting `
        $NotificationConfiguration 'MailTimeoutSeconds' 120
    if (-not [int]::TryParse([string]$configuredTimeout, [ref]$timeoutSeconds) -or
        $timeoutSeconds -lt 15 -or $timeoutSeconds -gt 900) {
        throw 'Notifications.MailTimeoutSeconds must be an integer from 15 through 900.'
    }
    Invoke-SmartWorkplaceCMDBSummaryGraphMailWithTimeout `
        -TenantId $ResolvedTenantId `
        -ClientId $clientId `
        -CertificateThumbprint $thumbprint `
        -Uri $uri `
        -Body $body `
        -TimeoutSeconds $timeoutSeconds
}

function Send-SmartWorkplaceCMDBSummarySmtpMail {
    param($NotificationConfiguration, [string]$Subject, [string]$BodyHtml)
    $server = [string](Get-SmartWorkplaceCMDBSummarySetting $NotificationConfiguration 'SmtpServer' '')
    $port = [int](Get-SmartWorkplaceCMDBSummarySetting $NotificationConfiguration 'SmtpPort' 25)
    $from = [string](Get-SmartWorkplaceCMDBSummarySetting $NotificationConfiguration 'From' '')
    $to = @(([string](Get-SmartWorkplaceCMDBSummarySetting $NotificationConfiguration 'To' '')) -split '[;,]' | Where-Object {$_.Trim()})
    $cc = @(([string](Get-SmartWorkplaceCMDBSummarySetting $NotificationConfiguration 'Cc' '')) -split '[;,]' | Where-Object {$_.Trim()})
    if ([string]::IsNullOrWhiteSpace($server) -or [string]::IsNullOrWhiteSpace($from) -or $to.Count -eq 0) {
        throw 'Notifications.SmtpServer, From, and To are required for SMTP.'
    }
    $mail = New-Object Net.Mail.MailMessage
    $mail.From = New-Object Net.Mail.MailAddress($from)
    foreach ($address in $to) { [void]$mail.To.Add($address.Trim()) }
    foreach ($address in $cc) { [void]$mail.CC.Add($address.Trim()) }
    $mail.Subject = $Subject
    $mail.Body = $BodyHtml
    $mail.IsBodyHtml = $true
    $client = New-Object Net.Mail.SmtpClient($server, $port)
    $client.EnableSsl = [bool](Get-SmartWorkplaceCMDBSummarySetting $NotificationConfiguration 'SmtpUseSsl' $false)
    $client.UseDefaultCredentials = [bool](Get-SmartWorkplaceCMDBSummarySetting $NotificationConfiguration 'SmtpUseIntegratedAuth' $true)
    try { $client.Send($mail) } finally { $mail.Dispose(); $client.Dispose() }
}

$projectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$coreModulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
Import-Module $coreModulePath -Force
$bound = @{}
foreach ($key in $PSBoundParameters.Keys) { $bound[$key] = $PSBoundParameters[$key] }
$context = Resolve-SmartWorkplaceCMDBContext -BoundParameters $bound `
    -GlobalConfigPath $GlobalConfigPath -TenantConfigPath $TenantConfigPath `
    -NoConfigWrite:($NoConfigWrite -or $ValidateOnly -or $PreviewOnly -or $SendMailTestOnly)
$paths = Resolve-SmartWorkplaceCMDBCollectionPaths -Paths $context.Paths `
    -ExplicitDataRoot:([bool]$DataRootPath) `
    -NoWrite:($ValidateOnly -or $PreviewOnly -or $SendMailTestOnly)
$notifications = Get-SmartWorkplaceCMDBSummarySetting $context.Configuration 'Notifications' $null
$graph = Get-SmartWorkplaceCMDBSummarySetting $context.Configuration 'MicrosoftGraph' $null
$historyRoot = Join-Path $paths.DataAllRootPath 'CollectionSummary'

if ($SendMailTestOnly) {
    $client = [string](Get-SmartWorkplaceCMDBSummarySetting `
        $notifications 'MailClientName' $paths.TenantKey)
    $subjectPrefix = [string](Get-SmartWorkplaceCMDBSummarySetting `
        $notifications 'Subject' 'Smart Workplace CMDB')
    $subject = '{0} - mail transport test - {1}' -f `
        $subjectPrefix,$SnapshotDateTime.ToString('yyyy-MM-dd HH:mm')
    $html = @"
<!doctype html><html><head><meta charset="utf-8"></head><body style="font-family:Segoe UI,Arial,sans-serif">
<h1>Smart Workplace CMDB - mail transport test</h1>
<p>Tenant: $(ConvertTo-SmartWorkplaceCMDBSummaryHtml $client)</p>
<p>Run: $(ConvertTo-SmartWorkplaceCMDBSummaryHtml $RunId)</p>
<p>Date: $(ConvertTo-SmartWorkplaceCMDBSummaryHtml $SnapshotDateTime.ToString('o'))</p>
<p>Script version: $(ConvertTo-SmartWorkplaceCMDBSummaryHtml $ScriptVersion)</p>
</body></html>
"@
    $html = Add-SmartWorkplaceCMDBSummaryMailBranding $html $notifications
    if ($ValidateOnly) {
        [pscustomobject]@{
            Status='Validated';ScriptVersion=$ScriptVersion;Subject=$subject
            HtmlPath='';BodyHtml=$html
        }
        return
    }
    if (-not [bool](Get-SmartWorkplaceCMDBSummarySetting `
            $notifications 'Enabled' $false)) {
        throw 'Collection summary email notification is not enabled in Notifications.Enabled.'
    }
    $mailMode = [string](Get-SmartWorkplaceCMDBSummarySetting `
        $notifications 'SendMailMode' 'Graph')
    switch ($mailMode.ToUpperInvariant()) {
        'GRAPH' {
            Send-SmartWorkplaceCMDBSummaryGraphMail `
                $notifications $graph $paths.TenantId $subject $html
        }
        'SMTP' {
            Send-SmartWorkplaceCMDBSummarySmtpMail $notifications $subject $html
        }
        'BOTH' {
            try {
                Send-SmartWorkplaceCMDBSummaryGraphMail `
                    $notifications $graph $paths.TenantId $subject $html
            }
            catch {
                Send-SmartWorkplaceCMDBSummarySmtpMail `
                    $notifications $subject $html
            }
        }
        default {
            throw "Unsupported Notifications.SendMailMode '$mailMode'. Use Graph, SMTP, or Both."
        }
    }
    [pscustomobject]@{
        Status='TestSent';ScriptVersion=$ScriptVersion;Subject=$subject
        HtmlPath='';BodyHtml=$html
    }
    return
}

if (-not [string]::IsNullOrWhiteSpace($OperationalError)) {
    $client = [string](Get-SmartWorkplaceCMDBSummarySetting $notifications 'MailClientName' $paths.TenantKey)
    $subjectPrefix = [string](Get-SmartWorkplaceCMDBSummarySetting $notifications 'Subject' 'Smart Workplace CMDB')
    $subject = '{0} - collection failed - {1}' -f $subjectPrefix,$SnapshotDateTime.ToString('yyyy-MM-dd HH:mm')
    $html = @"
<!doctype html><html><head><meta charset="utf-8"><style>
body{font-family:Segoe UI,Arial,sans-serif;background:#f4f7fb;color:#172033;padding:24px}.card{max-width:900px;margin:auto;background:white;border:1px solid #d8e2ef;border-radius:12px;overflow:hidden}.header{background:#b42318;color:white;padding:20px}.content{padding:20px}.label{font-weight:600;color:#475467}.value{margin:4px 0 16px;white-space:pre-wrap}
</style></head><body><div class="card"><div class="header"><h1>Smart Workplace CMDB - collection failed</h1><p>$(ConvertTo-SmartWorkplaceCMDBSummaryHtml $client)</p></div><div class="content">
<div class="label">Run</div><div class="value">$(ConvertTo-SmartWorkplaceCMDBSummaryHtml $RunId)</div>
<div class="label">Date</div><div class="value">$(ConvertTo-SmartWorkplaceCMDBSummaryHtml $SnapshotDateTime.ToString('o'))</div>
<div class="label">Failed step</div><div class="value">$(ConvertTo-SmartWorkplaceCMDBSummaryHtml $FailedStep)</div>
<div class="label">Error</div><div class="value">$(ConvertTo-SmartWorkplaceCMDBSummaryHtml $OperationalError)</div>
<div class="label">Log</div><div class="value">$(ConvertTo-SmartWorkplaceCMDBSummaryHtml $FailureLogPath)</div>
<div class="label">Transcript</div><div class="value">$(ConvertTo-SmartWorkplaceCMDBSummaryHtml $FailureTranscriptPath)</div>
</div></div></body></html>
"@
    $html = Add-SmartWorkplaceCMDBSummaryMailBranding $html $notifications
    if ($ValidateOnly) {
        [pscustomobject]@{Status='Validated';Subject=$subject;BodyHtml=$html;HtmlPath=''}
        return
    }
    $htmlFolder = Join-Path $paths.LogRootPath ('OperationalAlerts\{0}\{1}' -f $SnapshotDateTime.ToString('yyyy'),$SnapshotDateTime.ToString('MM'))
    New-Item -ItemType Directory -Path $htmlFolder -Force | Out-Null
    $htmlPath = Join-Path $htmlFolder ('SmartWorkplaceCMDB_Failure_{0}_{1}.html' -f $SnapshotDateTime.ToString('yyyyMMdd-HHmmssfff'),$RunId)
    [IO.File]::WriteAllText($htmlPath, $html, [Text.UTF8Encoding]::new($false))
    $status = 'Previewed'
    if (-not $PreviewOnly) {
        if (-not [bool](Get-SmartWorkplaceCMDBSummarySetting $notifications 'Enabled' $false)) {
            throw 'Collection failure email notification is not enabled in Notifications.Enabled.'
        }
        $mailMode = [string](Get-SmartWorkplaceCMDBSummarySetting $notifications 'SendMailMode' 'Graph')
        switch ($mailMode.ToUpperInvariant()) {
            'GRAPH' { Send-SmartWorkplaceCMDBSummaryGraphMail $notifications $graph $paths.TenantId $subject $html }
            'SMTP' { Send-SmartWorkplaceCMDBSummarySmtpMail $notifications $subject $html }
            'BOTH' {
                try { Send-SmartWorkplaceCMDBSummaryGraphMail $notifications $graph $paths.TenantId $subject $html }
                catch { Send-SmartWorkplaceCMDBSummarySmtpMail $notifications $subject $html }
            }
            default { throw "Unsupported Notifications.SendMailMode '$mailMode'. Use Graph, SMTP, or Both." }
        }
        $status = 'AlertSent'
    }
    [pscustomobject]@{Status=$status;ScriptVersion=$ScriptVersion;Subject=$subject;HtmlPath=$htmlPath;BodyHtml=$html}
    return
}

if ($CaptureBaselineOnly) {
    $saved = $null
    $buildManifestPath = Join-Path $paths.LatestOutputRootPath 'CMDB\CMDB_BuildManifest.csv'
    if (Test-Path -LiteralPath $buildManifestPath -PathType Leaf) {
        $SnapshotDateTime = [datetimeoffset](Get-Item -LiteralPath $buildManifestPath).LastWriteTimeUtc
    }
    $baseline = New-SmartWorkplaceCMDBSummarySnapshot $paths $notifications `
        ('baseline-' + $RunId) $SnapshotDateTime 'PreviousCompleted'
    $existing = @(Get-SmartWorkplaceCMDBSummaryHistory $historyRoot | Where-Object {
            $_.DatasetFingerprint -eq $baseline.DatasetFingerprint })
    if (-not $ValidateOnly -and $existing.Count -eq 0) {
        $saved = Save-SmartWorkplaceCMDBSummarySnapshot $baseline $paths
    }
    [pscustomobject]@{Status=$(if($ValidateOnly){'Validated'}elseif($existing.Count){'AlreadyCaptured'}else{'Captured'});Snapshot=$baseline;Paths=$saved}
    return
}

$current = New-SmartWorkplaceCMDBSummarySnapshot $paths $notifications $RunId $SnapshotDateTime $RunStatus
$allHistory = @(Get-SmartWorkplaceCMDBSummaryHistory $historyRoot)
$snapshotReused = $false
if ($ReuseLatestSnapshot) {
    $matchingSnapshot = @($allHistory | Where-Object {
            $_.DatasetFingerprint -eq $current.DatasetFingerprint -and
            $_.RunStatus -notin @('Failed', 'PreviousCompleted')
        } | Select-Object -Last 1)[0]
    if ($null -ne $matchingSnapshot) {
        $current = $matchingSnapshot
        $SnapshotDateTime = [datetimeoffset]::Parse(
            [string]$current.SnapshotDateTime)
        $snapshotReused = $true
    }
}
$history = @($allHistory | Where-Object {
        [datetimeoffset]::Parse([string]$_.SnapshotDateTime) -lt $SnapshotDateTime })
$previous = @($history | Select-Object -Last 1)[0]
$cutoff7 = $SnapshotDateTime.AddDays(-7)
$cutoff30 = $SnapshotDateTime.AddDays(-30)
$day7 = @($history | Where-Object {[datetimeoffset]::Parse([string]$_.SnapshotDateTime) -le $cutoff7} | Select-Object -Last 1)[0]
$day30 = @($history | Where-Object {[datetimeoffset]::Parse([string]$_.SnapshotDateTime) -le $cutoff30} | Select-Object -Last 1)[0]
$html = New-SmartWorkplaceCMDBSummaryHtml $current $previous $day7 $day30 $notifications
$html = Add-SmartWorkplaceCMDBSummaryMailBranding $html $notifications
$subjectPrefix = [string](Get-SmartWorkplaceCMDBSummarySetting $notifications 'Subject' 'Smart Workplace CMDB')
$subject = '{0} - full collection summary - {1}' -f $subjectPrefix,$SnapshotDateTime.ToString('yyyy-MM-dd')

if ($ValidateOnly) {
    [pscustomobject]@{Status='Validated';Snapshot=$current;BodyHtml=$html;Previous=$previous;Day7=$day7;Day30=$day30;SnapshotReused=$snapshotReused}
    return
}

$saved = Save-SmartWorkplaceCMDBSummarySnapshot $current $paths
$htmlFolder = Join-Path $paths.LogRootPath ('CollectionSummary\{0}\{1}' -f $SnapshotDateTime.ToString('yyyy'),$SnapshotDateTime.ToString('MM'))
New-Item -ItemType Directory -Path $htmlFolder -Force | Out-Null
$htmlPath = Join-Path $htmlFolder ('SmartWorkplaceCMDB_CollectionSummary_{0}.html' -f $SnapshotDateTime.ToString('yyyyMMdd-HHmmssfff'))
[IO.File]::WriteAllText($htmlPath, $html, [Text.UTF8Encoding]::new($false))

$status = 'Previewed'
if (-not $PreviewOnly) {
    if (-not [bool](Get-SmartWorkplaceCMDBSummarySetting $notifications 'Enabled' $false)) {
        throw 'Full-collection email notification is not enabled in Notifications.Enabled.'
    }
    $mode = [string](Get-SmartWorkplaceCMDBSummarySetting $notifications 'SendMailMode' 'Graph')
    switch ($mode.ToUpperInvariant()) {
        'GRAPH' {
            Send-SmartWorkplaceCMDBSummaryGraphMail `
                -NotificationConfiguration $notifications `
                -GraphConfiguration $graph `
                -ResolvedTenantId $paths.TenantId `
                -Subject $subject `
                -BodyHtml $html
        }
        'SMTP' {
            Send-SmartWorkplaceCMDBSummarySmtpMail `
                -NotificationConfiguration $notifications -Subject $subject -BodyHtml $html
        }
        'BOTH' {
            try {
                Send-SmartWorkplaceCMDBSummaryGraphMail `
                    -NotificationConfiguration $notifications `
                    -GraphConfiguration $graph `
                    -ResolvedTenantId $paths.TenantId `
                    -Subject $subject `
                    -BodyHtml $html
            }
            catch {
                Send-SmartWorkplaceCMDBSummarySmtpMail `
                    -NotificationConfiguration $notifications -Subject $subject -BodyHtml $html
            }
        }
        default { throw "Unsupported Notifications.SendMailMode '$mode'. Use Graph, SMTP, or Both." }
    }
    $status = 'Sent'
}

[pscustomobject]@{
    Status=$status;ScriptVersion=$ScriptVersion;Subject=$subject;Snapshot=$current
    Previous=$previous;Day7=$day7;Day30=$day30;HistoryPath=$saved.HistoryPath
    LatestPath=$saved.LatestPath;HtmlPath=$htmlPath;BodyHtml=$html
    SnapshotReused=$snapshotReused
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD389azhRoF6HHj
# m8sTfKuF8/JNS5XQ16hjqnNCX1+FnKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILl0FYcDmIV2ReZdw152O//LAREaRWjZvzK9LpptB56yMA0GCSqG
# SIb3DQEBAQUABIIBgCMkb9BSzTta4cpOyDs9GUN/WmfMgYFNPoJIkdRK3yzONdsT
# cBr0n/0lvDvYXn47fm/LKNcon1c3O0jj36rXiKmYfoe6e0q3Rwx0Dgtfc3yrB7uK
# jPHYyPVAFM9NbsGX5drjivSzpz9Zw7cnLZA0yDM011l45Q9dJsOeV3wjpKDk+8b8
# UgDZ9luaEOhYmyTDbRDC5yHLeQn7j7ZfYDH5vpyOlH69h8pCGXQtRdIKnnjgtm5q
# 7M+TvTRAzFgVMiK25Zdbkrx/LkOk8/0CiwwbUuVcLQkjt1FsgUTDFRTjWJpl94xA
# d92Sxvq7Q/K73VinKX78qSBiu4lMtlON6g2LwEsta+/5gedpNpiaXVd8UeWNgX7l
# LfQP7S3KNSwzCdJtsf9n7k0PriwXn9wY0O4dG5oLTeOC1jVccOiPdRVxUsDTqMPk
# 2pFDGXU4ZwjB1pulvUWPPfbmgziecYjiCgPoUR802rOa+bxww6qXEI3AazpKVefS
# u3JRj8m3UVtUxxEHuaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTQxODM1
# MTBaMC8GCSqGSIb3DQEJBDEiBCAEB76ERgvK/KsyROvStqPWHEaEp3EKa3sw9CfQ
# Yx8gTTANBgkqhkiG9w0BAQEFAASCAgBWQsw0017psbNxqqFZfNIc0SGjw/KeIgeS
# awXTlk4BkY7T8fpXju7E9DokWE+1ffZv8x4Hhj2iyxeg4HiewENF7RDCgvoySSfe
# geJk+rLtkRQ0IeCfyfBp4f7QJm8+vUzLJ46N7A1MeQyfLicxu11e2NmXX1x/onG+
# uv1tKOYq1VIUhz9Of2wYVugS4G04jMVsCeSG5yL/N4PDu8tPJF/xaJcRkDunKki5
# mE06LX9xoX/4J8uX8Ye3iEPB9NiG9cJdi/GqlDfn5m2Hc4i7pxM2qBUE3CO9Lrsv
# 2ZeehA98jJlFgLgLLJhWca9+lHVlT6Y0y3HfYp+uLXk+18GPLhodTRJg83xBOfvB
# /w+byVQRd4DVBmChN4HvObroMu5Dm9RRouWb+5ZENtjvAEtDhHJmqGAz1yNe35YG
# X4N/Uto0jFiSy02RoCtgPkgIC8u/CtKPt92yq1zmH1g+XKk06FHORJNL3dNKRLLf
# QSzpPFosud6LWP4/KSmsEia5zzKM1tglYTF1/xaACWpXpssYsMtHsCPPKiHa9WDx
# FQVfOob39P46BLDvAAfdUYmR5wsI1va1hBz4KHO6xobIMGLIQyK4OxS2GnxGeX/H
# 6j779gTD1C0FyqSLkfcv5g0MfsBOFTqyo7D4ZcsvrQJwCs47gLV+pmUo+RgoN/zH
# HxKMgQgkTA==
# SIG # End signature block
