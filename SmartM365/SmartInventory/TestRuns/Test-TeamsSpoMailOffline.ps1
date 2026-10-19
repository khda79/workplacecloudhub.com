#Requires -Version 7.0
<#
.SYNOPSIS
    Synthetic Teams and SharePoint workbook/mail regressions; no dependency repair.
.VERSION
    1.0.4
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Import-Module ImportExcel -ErrorAction Stop

function Import-ScriptFunction {
    param([string]$Path,[string[]]$Names)
    $tokens = $null
    $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)
    if (@($errors).Count -gt 0) { throw "Parser errors in $Path" }
    foreach ($name in $Names) {
        $definition = $ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name},$true)
        if (-not $definition) { throw "Missing function: $name" }
        Invoke-Expression ("function global:{0} {1}" -f $name,$definition.Body.Extent.Text)
    }
}
function Assert-Equal {
    param($Actual,$Expected,[string]$Message)
    if ($Actual -ne $Expected) { throw "$Message. Expected '$Expected'; got '$Actual'." }
}

$inventoryRoot = Split-Path $PSScriptRoot -Parent
$smartM365Root = Split-Path $inventoryRoot -Parent
Import-Module -Name (Join-Path $smartM365Root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -Force -ErrorAction Stop
# Workbook regression checks may load the installed module, but never repair/install it.
function global:Initialize-SmartM365ImportExcel {
    param([switch]$InstallIfMissing)
    SmartM365.Core\Initialize-SmartM365ImportExcel -CheckOnly | Out-Null
}
$teamsPath = Join-Path $inventoryRoot 'M365Inventory/Teams/SmartM365-Teams-Inventory.ps1'
$spoPath = Join-Path $inventoryRoot 'M365Inventory/SharePoint/SmartM365-SPO-Inventory.ps1'
Import-ScriptFunction -Path $teamsPath -Names @('IsoUtc','Get-TeamsGraphPropertyValue','Get-TeamsChannelSiteId','Get-TeamsDriveSiteId','Add-TeamsChannelSiteReferences','Get-TeamsBatchSeed','Get-TeamsCsvColumnNames','Import-TeamsMailCsvSet','Invoke-TeamsDailySummaryMail','New-TeamsTimestampedWorkbook','Add-TeamsWorkbookOverview','Get-TeamsMailStatus','ConvertTo-HtmlReport')
Import-ScriptFunction -Path $spoPath -Names @('ConvertTo-SpoHtml','Get-SpoCsvColumnNames','ConvertTo-SpoSiteUrlKey','Measure-SpoTeamsSiteLinks','Get-SpoTeamsSiteLinksFromCsv','Get-SpoLicenseCapacityEstimate','Add-SpoLicenseEstimateToSummary','Import-SpoMailCsvSet','Invoke-SpoDailySummaryMail','New-SpoTimestampedWorkbook','Add-SpoWorkbookOverview','Get-SpoMailStatus','New-SpoHtmlSummary')
function WriteLog { param($Message,$Level) }
function Write-SpoLog { param($Message,$Level) }

$script:batchLengths = [System.Collections.Generic.List[int]]::new()
$script:batchMode = 'Seed'
function global:Invoke-TeamsGraphBatch {
    param([object[]]$Requests)
    if ($script:batchMode -eq 'Seed') {
        [void]$script:batchLengths.Add($Requests.Count)
        return @()
    }
    foreach ($request in $Requests) {
        if ($script:batchMode -eq 'Unavailable') { [pscustomobject]@{id=$request.id;status=404;body=$null} }
        elseif ($script:batchMode -eq 'DriveMetadata') {
            if ($request.url -like '/drives/*') {
                [pscustomobject]@{id=$request.id;status=200;body=[pscustomobject]@{sharepointIds=[pscustomobject]@{siteId='11111111-1111-1111-1111-111111111111';webId='22222222-2222-2222-2222-222222222222';siteUrl='https://tenant.example/sites/channel'}}}
            }
            else { [pscustomobject]@{id=$request.id;status=200;body=[pscustomobject]@{parentReference=[pscustomobject]@{driveId='drive-channel'}}} }
        }
        else { [pscustomobject]@{id=$request.id;status=200;body=[pscustomobject]@{parentReference=[pscustomobject]@{siteId='tenant.example,site-channel,web-channel'}}} }
    }
}
$seed = Get-TeamsBatchSeed -Teams @([pscustomobject]@{id='a'},[pscustomobject]@{id='b'},[pscustomobject]@{id='c'},[pscustomobject]@{id='d'})
Assert-Equal ($script:batchLengths -join ',') '18,6' 'Teams prefetch stays within Graph 20-request batch limit'
Assert-Equal $seed['a'].PSObject.Properties['Site'].Name 'Site' 'Teams prefetch includes the site identity'
$script:batchMode = 'Channel'
$channelLinkRows = @(
    [pscustomobject]@{TeamId='a';ChannelId='private-1';MembershipType='private'},
    [pscustomobject]@{TeamId='a';ChannelId='standard-1';MembershipType='standard'}
)
Add-TeamsChannelSiteReferences -ChannelRows $channelLinkRows
Assert-Equal $channelLinkRows[0].SharePointSiteId 'tenant.example,site-channel,web-channel' 'Private channel site identity'
Assert-Equal $channelLinkRows[0].SharePointSiteLinkStatus 'Collected' 'Private channel link status'
Assert-Equal $channelLinkRows[1].SharePointSiteLinkStatus 'NotApplicable' 'Standard channel reuses the team site'
$script:batchMode = 'DriveMetadata'
Add-TeamsChannelSiteReferences -ChannelRows $channelLinkRows
Assert-Equal $channelLinkRows[0].SharePointSiteId 'tenant.example,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222' 'Channel site ID is recovered from drive SharePoint IDs'
Assert-Equal $channelLinkRows[0].SharePointSiteLinkStatus 'Collected' 'Drive metadata channel link status'
$script:batchMode = 'Unavailable'
Add-TeamsChannelSiteReferences -ChannelRows $channelLinkRows
Assert-Equal $channelLinkRows[0].SharePointSiteLinkStatus 'Unavailable' 'Failed channel link cannot reuse a prior identity'

$teamsSummary = @{TotalTeams=769;OwnerlessTeams=60;HighQuotaTeams=2;InactiveTeams=100;ActiveTeams=600;ArchivedTeams=69;PublicTeams=20;TeamsWithGuests=50;CriticalCount=62;WarningCount=384}
Assert-Equal (Get-TeamsMailStatus -Summary $teamsSummary -PartialInventory $false) 'Warning' 'Localized Teams findings'
$teamsSummary.OwnerlessTeams = 77
Assert-Equal (Get-TeamsMailStatus -Summary $teamsSummary -PartialInventory $false) 'Critical' 'Systemic Teams findings'
Assert-Equal (Get-TeamsMailStatus -Summary $teamsSummary -PartialInventory $true) 'Warning' 'Partial Teams inventory'
$teamsSummary.OwnerlessTeams = 60

$spoSummary = [ordered]@{SitesProcessed=1915;OwnerlessSites=5;HighQuotaSites=3;InactiveSites=789;SharePointSites=1915;OneDriveSites=0;ListsProcessed=200;StorageUsedGB=500;TenantStorageCapacityGB='';TenantStorageUtilizationPercent='';CriticalAlerts=8;WarningAlerts=789;RunId='synthetic-run';Duration='00:01:00';InventoryMode='GraphOnly'}
Assert-Equal (Get-SpoMailStatus -Summary $spoSummary -PartialInventory $false) 'Warning' 'Localized SPO findings'
$spoSummary.CriticalAlerts = 0
$spoSummary.WarningAlerts = 0
$spoSummary.TenantStorageUtilizationPercent = 82
Assert-Equal (Get-SpoMailStatus -Summary $spoSummary -PartialInventory $false) 'Warning' 'Measured tenant capacity warning'
$spoSummary.CriticalAlerts = 8
$spoSummary.WarningAlerts = 789
$spoSummary.TenantStorageUtilizationPercent = 95
Assert-Equal (Get-SpoMailStatus -Summary $spoSummary -PartialInventory $false) 'Critical' 'Measured tenant capacity'
$spoSummary.TenantStorageUtilizationPercent = ''

$teamsAlerts = @(
    [pscustomobject]@{TeamId='one';TeamDisplayName='Example & One';Status='Critical';Check='Owners';NumericValue=0;TextValue='NoOwner';Threshold='Owners >= 2'},
    [pscustomobject]@{TeamId='one';TeamDisplayName='Example & One';Status='Critical';Check='StorageQuotaPercent';NumericValue=94;TextValue='';Threshold='<= 90 percent'},
    [pscustomobject]@{TeamId='two';TeamDisplayName='Example Two';Status='Warning';Check='Inactivity';NumericValue=245;TextValue='';Threshold='<= 180 days'}
)
$teamsRows = @(
    [pscustomobject]@{TeamId='one';ActivityEvidenceStatus='Observed';ActivityReportRefreshDate='2026-10-10';StorageCollectionStatus='Collected'},
    [pscustomobject]@{TeamId='two';ActivityEvidenceStatus='NoReportRow';ActivityReportRefreshDate='';StorageCollectionStatus='Unavailable'}
)
$teamsHtml = ConvertTo-HtmlReport -AlertRows $teamsAlerts -TeamRows $teamsRows -Summary $teamsSummary -Worst 'Warning' -Started (Get-Date).AddMinutes(-1) -Ended (Get-Date)
if ($teamsHtml -notmatch 'Teams to review' -or $teamsHtml -notmatch '2 distinct teams require review' -or $teamsHtml -notmatch 'Action required') { throw 'Teams action summary is missing.' }
if ($teamsHtml -notmatch 'Example &amp; One' -or $teamsHtml -match 'Example & One' -or $teamsHtml -match '<a\s') { throw 'Teams mail escaping or default link behavior is invalid.' }
if ([regex]::Matches($teamsHtml,'Example &amp; One').Count -ne 1) { throw 'Teams priority table must list each team only once.' }
if ($teamsHtml -notmatch 'no activity report row' -or $teamsHtml -notmatch 'Teams without measured storage: 1') { throw 'Teams evidence coverage is missing.' }
if ($teamsHtml.IndexOf('<h2>Summary</h2>') -gt $teamsHtml.IndexOf('Action required') -or $teamsHtml.IndexOf('Run details') -lt $teamsHtml.IndexOf('Action required')) { throw 'Teams mail section order is invalid.' }
if ($teamsHtml -notmatch 'Priority' -or $teamsHtml -notmatch 'Storage above 90%') { throw 'Teams priority table or source threshold is missing.' }
if ($teamsHtml -notmatch 'colspan="2"' -or $teamsHtml -notmatch 'background:#eef2ff') { throw 'Teams KPI grid does not use the Exchange-style wide card.' }
$teamsHtmlWithLinks = ConvertTo-HtmlReport -AlertRows @() -Summary $teamsSummary -Worst 'OK' -Started (Get-Date).AddMinutes(-1) -Ended (Get-Date) -FileLinksHtml '<a href="https://example.invalid/export">Export</a>'
if ($teamsHtmlWithLinks -notmatch 'https://example.invalid/export') { throw 'Teams mail links cannot be enabled.' }
$teamsSnapshotHtml = ConvertTo-HtmlReport -AlertRows @() -Summary $teamsSummary -Worst 'OK' -Started (Get-Date) -Ended (Get-Date) -SourceNote 'Snapshot & no new collection'
if ($teamsSnapshotHtml -notmatch 'Snapshot &amp; no new collection' -or $teamsSnapshotHtml -notmatch 'Machine: Not recorded in CSV') { throw 'Teams snapshot provenance is missing or unescaped.' }
$spoAlerts = @(
    [pscustomobject]@{Severity='Critical';Category='Ownership';SiteUrl='https://example.invalid/a';ObjectName='';Value='';Threshold='Valid owner required'},
    [pscustomobject]@{Severity='Critical';Category='Storage';SiteUrl='https://example.invalid/a';ObjectName='';Value='96';Threshold='90%'},
    [pscustomobject]@{Severity='Warning';Category='ListViewThreshold';SiteUrl='https://example.invalid/b';ObjectName='List & One';Value='8200';Threshold='5000'}
)
$spoSites = @(
    [pscustomobject]@{SiteUrl='https://example.invalid/a';IsInactive=$true;LastActivityUtc='2026-01-01';GraphEnrichmentStatus='Collected';ReportRefreshDate='2026-10-10';Threshold='90%; 180 days'},
    [pscustomobject]@{SiteUrl='https://example.invalid/b';IsInactive=$true;LastActivityUtc='';GraphEnrichmentStatus='Unavailable';ReportRefreshDate='2026-10-10';Threshold='90%; 180 days'}
)
$spoSummary.IncludeOneDrive=$false
$spoSummary.DeepSharingScan=$false
$spoSummary.TenantCapacitySource='Get-PnPTenantUnauthorized'
$spoSummary.EstimatedTenantCapacityGB=91024
$spoSummary.EstimatedTenantUtilizationPercent=0.55
$spoSummary.CapacityEstimateEligibleUnits=9000
$spoSummary.CapacityEstimateAsOfUtc='2026-10-13T02:00:00Z'
$statusFromEstimate = [ordered]@{SitesProcessed=100;OwnerlessSites=0;HighQuotaSites=0;CriticalAlerts=0;WarningAlerts=0;TenantStorageUtilizationPercent='';EstimatedTenantCapacityGB=91024;EstimatedTenantUtilizationPercent=99}
Assert-Equal (Get-SpoMailStatus -Summary $statusFromEstimate -PartialInventory $false) 'OK' 'Estimated capacity cannot change mail status'
$spoHtml = New-SpoHtmlSummary -WorstStatus 'Warning' -Alerts $spoAlerts -SiteRows $spoSites -Summary $spoSummary
if ($spoHtml -notmatch 'Sites to review' -or $spoHtml -notmatch 'Action required' -or $spoHtml -notmatch 'Usage-report sites') { throw 'SPO mail summary is missing.' }
if ($spoHtml -notmatch 'List &amp; One' -or $spoHtml -match 'List & One' -or $spoHtml -match '<a\s') { throw 'SPO mail escaping or default link behavior is invalid.' }
if ([regex]::Matches($spoHtml,'https://example.invalid/a').Count -ne 1 -or $spoHtml -notmatch 'Inactive with date</div>\s*<div[^>]*>1') { throw 'SPO priority or dated-inactivity count is invalid.' }
if ($spoHtml -notmatch 'Sites without an activity date: 1' -or $spoHtml -notmatch 'Sites without Graph enrichment: 1') { throw 'SPO evidence coverage is missing.' }
if ($spoHtml -notmatch 'Tenant utilization \(estimated\)' -or $spoHtml -notmatch '0.55%' -or $spoHtml -notmatch '91(?:&#160;|&#8239;|,|\u00A0|\u202F)024 GB estimated capacity' -or $spoHtml -notmatch 'License estimate: 1 TB \+ 10 GB x 9000' -or $spoHtml -match 'Not measured') { throw 'SPO must show estimated utilization, the capacity basis, and avoid an unmeasured card.' }
if ($spoHtml -notmatch 'colspan="2"' -or $spoHtml -notmatch 'background:#eef2ff') { throw 'SPO KPI grid does not use the Exchange-style wide card.' }
if ($spoHtml.IndexOf('<h2>Summary</h2>') -gt $spoHtml.IndexOf('Action required') -or $spoHtml -notmatch 'Priority' -or $spoHtml -notmatch 'Site storage above 90%' -or $spoHtml -notmatch 'No activity recorded for over 180 days') { throw 'SPO KPI grid must precede the action summary, which shows source thresholds.' }
$spoSummary.TenantStorageCapacityGB = 1000
$spoSummary.TenantStorageUtilizationPercent = 50
$spoHtmlMeasured = New-SpoHtmlSummary -WorstStatus 'Warning' -Alerts @() -SiteRows $spoSites -Summary $spoSummary
if ($spoHtmlMeasured -notmatch 'Capacity: 1000 GB' -or $spoHtmlMeasured -notmatch 'Tenant utilization</div>\s*<div[^>]*>50%') { throw 'SPO measured tenant capacity is missing from mail.' }
$spoSummary.TenantStorageCapacityGB = ''
$spoSummary.TenantStorageUtilizationPercent = ''
if ($spoHtml.IndexOf('Run details') -lt $spoHtml.IndexOf('Action required')) { throw 'SPO run details must follow the summary.' }
$spoHtmlWithLinks = New-SpoHtmlSummary -WorstStatus 'OK' -Alerts @() -Summary $spoSummary -FileLinksHtml '<a href="https://example.invalid/export">Export</a>'
if ($spoHtmlWithLinks -notmatch 'https://example.invalid/export') { throw 'SPO mail links cannot be enabled.' }
$spoSummary.InventoryMode='CSVSnapshot'
$spoSnapshotHtml = New-SpoHtmlSummary -WorstStatus 'OK' -Alerts @() -Summary $spoSummary -SourceNote 'Snapshot & no new collection'
if ($spoSnapshotHtml -notmatch 'Snapshot &amp; no new collection' -or $spoSnapshotHtml -notmatch 'Not verifiable from CSV') { throw 'SPO snapshot provenance is missing or unescaped.' }
$spoSummary.InventoryMode='GraphOnly'

$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ("SmartM365-MailOffline-{0}" -f [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporaryRoot -Force | Out-Null
try {
    $licenseDate = [datetimeoffset]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ',[Globalization.CultureInfo]::InvariantCulture)
    $licensePath = Join-Path $temporaryRoot 'M365_Licenses_Tenant.csv'
    @(
        [pscustomobject]@{TenantKey='prod';TenantId='tenant-synthetic';TenantSkuPartNumber='SPE_E3';TenantPrepaidEnabled='100';CapabilityStatus='Enabled';AppliesTo='User';CollectedAtUtc=$licenseDate},
        [pscustomobject]@{TenantKey='prod';TenantId='tenant-synthetic';TenantSkuPartNumber='SPE_F1';TenantPrepaidEnabled='200';CapabilityStatus='Enabled';AppliesTo='User';CollectedAtUtc=$licenseDate}
    ) | Export-Csv -LiteralPath $licensePath -NoTypeInformation
    $estimate = Get-SpoLicenseCapacityEstimate -Folder $temporaryRoot -ExpectedTenantKey 'prod' -ExpectedTenantId 'tenant-synthetic' -ReferenceDateUtc $licenseDate
    Assert-Equal $estimate.CapacityGB 2024 'Eligible E3 seats add 10 GB each; F1 seats do not'
    Assert-Equal (Get-SpoLicenseCapacityEstimate -Folder $temporaryRoot -ExpectedTenantKey 'other' -ExpectedTenantId 'tenant-synthetic' -ReferenceDateUtc $licenseDate).Available $false 'License estimate tenant isolation'
    Assert-Equal (Get-SpoLicenseCapacityEstimate -Folder $temporaryRoot -ExpectedTenantKey 'prod' -ExpectedTenantId 'tenant-synthetic' -ReferenceDateUtc '2026-01-01T00:00:00Z').Available $false 'License estimate freshness'
    $estimateSummary = [ordered]@{TenantStorageCapacityGB='';TenantStorageUtilizationPercent='';StorageUsedGB=506}
    Add-SpoLicenseEstimateToSummary -Summary $estimateSummary -Folder $temporaryRoot -ExpectedTenantKey 'prod' -ExpectedTenantId 'tenant-synthetic' -ReferenceDateUtc $licenseDate
    Assert-Equal $estimateSummary.EstimatedTenantCapacityGB 2024 'License estimate added to mail summary'
    Assert-Equal $estimateSummary.EstimatedTenantUtilizationPercent 25 'Estimated utilization uses inventoried storage and estimated capacity'
    Add-SpoLicenseEstimateToSummary -Summary $estimateSummary -Folder $temporaryRoot -ExpectedTenantKey 'prod' -ExpectedTenantId 'tenant-synthetic' -ReferenceDateUtc $licenseDate -PartialInventory:$true
    Assert-Equal $estimateSummary.EstimatedTenantUtilizationPercent '' 'Partial inventory cannot claim tenant utilization'

    $teamsDate = [datetimeoffset]::UtcNow.AddMinutes(-10).ToString('yyyy-MM-ddTHH:mm:ssZ',[Globalization.CultureInfo]::InvariantCulture)
    $spoDate = [datetimeoffset]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ',[Globalization.CultureInfo]::InvariantCulture)
    $siteLinkRows = @(
        [pscustomobject]@{SiteId='tenant.example,site-one,web-one';SiteUrl='https://tenant.example/sites/one';IsOneDrive='False';SiteIdentitySource='GraphSite';GraphEnrichmentStatus='Collected'},
        [pscustomobject]@{SiteId='tenant.example,site-two,web-two';SiteUrl='https://tenant.example/sites/two';IsOneDrive='False';SiteIdentitySource='GraphSite';GraphEnrichmentStatus='Collected'},
        [pscustomobject]@{SiteId='tenant.example,site-channel,web-channel';SiteUrl='https://tenant.example/sites/channel';IsOneDrive='False';SiteIdentitySource='GraphSite';GraphEnrichmentStatus='Collected'},
        [pscustomobject]@{SiteId='tenant.example,site-other,web-other';SiteUrl='https://tenant.example/sites/other';IsOneDrive='False';SiteIdentitySource='GraphSite';GraphEnrichmentStatus='Collected'},
        [pscustomobject]@{SiteId='tenant.example,site-drive,web-drive';SiteUrl='https://tenant.example/personal/user';IsOneDrive='True';SiteIdentitySource='GraphSite';GraphEnrichmentStatus='Collected'}
    )
    $teamLinkRows = @(
        [pscustomobject]@{TenantKey='prod';TenantId='tenant-synthetic';RunId='teams-run';RunDateUtc=$teamsDate;TeamId='team-one';TeamDisplayName='Team One';SharePointSiteId='tenant.example,site-one,web-one';SharePointSiteUrl='https://tenant.example/sites/one';SharePointSiteLinkStatus='Collected';IsPartialInventory='False'},
        [pscustomobject]@{TenantKey='prod';TenantId='tenant-synthetic';RunId='teams-run';RunDateUtc=$teamsDate;TeamId='team-two';TeamDisplayName='Team Two';SharePointSiteId='tenant.example,site-two,web-two';SharePointSiteUrl='https://tenant.example/sites/two';SharePointSiteLinkStatus='Collected';IsPartialInventory='False'}
    )
    $channelSiteRows = @(
        [pscustomobject]@{TenantKey='prod';TenantId='tenant-synthetic';RunId='teams-run';RunDateUtc=$teamsDate;TeamId='team-one';ChannelId='channel-private';ChannelDisplayName='Private';MembershipType='private';SharePointSiteId='tenant.example,site-channel,web-channel';SharePointSiteLinkStatus='Collected'},
        [pscustomobject]@{TenantKey='prod';TenantId='tenant-synthetic';RunId='teams-run';RunDateUtc=$teamsDate;TeamId='team-one';ChannelId='channel-standard';ChannelDisplayName='General';MembershipType='standard';SharePointSiteId='';SharePointSiteLinkStatus='NotApplicable'}
    )
    $teamLinkPath = Join-Path $temporaryRoot 'M365_Teams_Teams.csv'
    $channelLinkPath = Join-Path $temporaryRoot 'M365_Teams_Channels.csv'
    $teamLinkRows | Export-Csv -LiteralPath $teamLinkPath -NoTypeInformation
    $channelSiteRows | Export-Csv -LiteralPath $channelLinkPath -NoTypeInformation
    $linkSummary = Get-SpoTeamsSiteLinksFromCsv -Folder $temporaryRoot -ExpectedTenantKey 'prod' -ExpectedTenantId 'tenant-synthetic' -ReferenceDateUtc $spoDate -SiteRows $siteLinkRows
    Assert-Equal $linkSummary.Available $true 'Same-tenant Teams site mapping qualifies'
    Assert-Equal $linkSummary.LinkedSites 3 'Teams site links are deduplicated by SPO site'
    Assert-Equal $linkSummary.TotalSites 4 'OneDrive is excluded from Teams site ratio denominator'
    Assert-Equal $linkSummary.Percent 75 'Teams site percentage'
    Assert-Equal $linkSummary.ParentSites 2 'Team main sites counted once'
    Assert-Equal $linkSummary.ChannelSites 1 'Private channel site counted separately'
    Assert-Equal (Get-SpoTeamsSiteLinksFromCsv -Folder $temporaryRoot -ExpectedTenantKey 'prod' -ExpectedTenantId 'other-tenant' -ReferenceDateUtc $spoDate -SiteRows $siteLinkRows).Available $false 'Teams site mapping tenant isolation'
    Assert-Equal (Get-SpoTeamsSiteLinksFromCsv -Folder $temporaryRoot -ExpectedTenantKey 'prod' -ExpectedTenantId 'tenant-synthetic' -ReferenceDateUtc ([datetimeoffset]::UtcNow.AddDays(2).ToString('o')) -SiteRows $siteLinkRows).Available $false 'Teams site mapping freshness'
    Assert-Equal (Get-SpoTeamsSiteLinksFromCsv -Folder $temporaryRoot -ExpectedTenantKey 'prod' -ExpectedTenantId 'tenant-synthetic' -ReferenceDateUtc $spoDate -SiteRows $siteLinkRows -PartialInventory:$true).Available $false 'Limited SPO inventory cannot claim Teams site ratio'
    $siteLinkRows[0].GraphEnrichmentStatus = 'Unavailable'
    Assert-Equal (Get-SpoTeamsSiteLinksFromCsv -Folder $temporaryRoot -ExpectedTenantKey 'prod' -ExpectedTenantId 'tenant-synthetic' -ReferenceDateUtc $spoDate -SiteRows $siteLinkRows).Available $false 'SPO site without verified Graph identity cannot claim Teams site ratio'
    $siteLinkRows[0].GraphEnrichmentStatus = 'Collected'
    $spoSummary.TeamsSiteLinks = $linkSummary
    $spoLinkedHtml = New-SpoHtmlSummary -WorstStatus 'Warning' -Alerts @() -SiteRows @() -Summary $spoSummary
    if ($spoLinkedHtml -notmatch 'Teams-linked sites:</b> 3 / 4 \(75%\)' -or $spoLinkedHtml -notmatch 'private/shared channel sites: 1' -or $spoLinkedHtml.IndexOf('Teams-linked sites:') -gt $spoLinkedHtml.IndexOf('Action required')) { throw 'Teams site ratio is missing from the SPO Summary.' }
    $channelSiteRows[0].SharePointSiteLinkStatus = 'Unavailable'
    $channelSiteRows | Export-Csv -LiteralPath $channelLinkPath -NoTypeInformation
    $unavailableLinks = Get-SpoTeamsSiteLinksFromCsv -Folder $temporaryRoot -ExpectedTenantKey 'prod' -ExpectedTenantId 'tenant-synthetic' -ReferenceDateUtc $spoDate -SiteRows $siteLinkRows
    Assert-Equal $unavailableLinks.Available $false 'Unresolved channel site prevents a false percentage'
    $spoSummary.TeamsSiteLinks = $unavailableLinks
    $spoUnlinkedHtml = New-SpoHtmlSummary -WorstStatus 'Warning' -Alerts @() -SiteRows @() -Summary $spoSummary
    if ($spoUnlinkedHtml -notmatch 'Teams-linked sites:</b> unavailable' -or $spoUnlinkedHtml -match '3 / 4 \(75%\)') { throw 'Unqualified Teams site evidence must not show a percentage.' }
    $spoSummary.TeamsSiteLinks = $linkSummary

    $teamsFiles = @()
    foreach ($name in @('Teams','Members','Channels','Guests')) {
        $path = Join-Path $temporaryRoot "$name.csv"
        [pscustomobject]@{Name='Synthetic';Status='OK'} | Export-Csv -LiteralPath $path -NoTypeInformation
        $teamsFiles += [pscustomobject]@{Path=$path;WorksheetName=$name;TableName="TestTeams$name"}
    }
    $teamsBook = Join-Path $temporaryRoot 'Teams.xlsx'
    New-TeamsTimestampedWorkbook -CsvFiles $teamsFiles -Path $teamsBook | Out-Null
    Add-TeamsWorkbookOverview -Path $teamsBook -Summary $teamsSummary -Alerts @([pscustomobject]@{Status='Critical';TeamDisplayName='Synthetic';Check='Owners';NumericValue='0';TextValue='';Threshold='1';Details='No owner'}) -Status 'Warning'
    $book = Open-ExcelPackage -Path $teamsBook
    try { Assert-Equal $book.Workbook.Worksheets.Count 6 'Teams worksheet count'; Assert-Equal $book.Workbook.Worksheets['Findings'].Cells[2,1].Text 'Critical' 'Teams findings worksheet' }
    finally { Close-ExcelPackage -ExcelPackage $book -NoSave }
    $teamsEmptyBook = Join-Path $temporaryRoot 'Teams-empty.xlsx'
    New-TeamsTimestampedWorkbook -CsvFiles $teamsFiles -Path $teamsEmptyBook | Out-Null
    Add-TeamsWorkbookOverview -Path $teamsEmptyBook -Summary $teamsSummary -Alerts @() -Status 'OK'
    $book = Open-ExcelPackage -Path $teamsEmptyBook
    try { Assert-Equal $book.Workbook.Worksheets['Findings'].Cells[1,1].Text 'Status' 'Empty Teams findings worksheet header' }
    finally { Close-ExcelPackage -ExcelPackage $book -NoSave }

    $spoFiles = @()
    foreach ($name in @('Sites','Lists','Permissions','External sharing','Tenant capacity')) {
        $path = Join-Path $temporaryRoot (($name -replace ' ','_') + '.csv')
        [pscustomobject]@{Name='Synthetic';Status='OK'} | Export-Csv -LiteralPath $path -NoTypeInformation
        $spoFiles += [pscustomobject]@{Path=$path;WorksheetName=$name;TableName=('TestSpo' + ($name -replace ' ',''))}
    }
    $spoBook = Join-Path $temporaryRoot 'SPO.xlsx'
    New-SpoTimestampedWorkbook -CsvFiles $spoFiles -Path $spoBook | Out-Null
    Add-SpoWorkbookOverview -Path $spoBook -Summary $spoSummary -Alerts @([pscustomobject]@{Severity='Critical';Category='Ownership';SiteUrl='https://example.invalid';ObjectName='';Metric='Owner';Value='';Threshold='Valid owner';Details='No owner'}) -Status 'Warning'
    $book = Open-ExcelPackage -Path $spoBook
    try { Assert-Equal $book.Workbook.Worksheets.Count 8 'SPO worksheet count with qualified Teams links'; Assert-Equal $book.Workbook.Worksheets['Findings'].Cells[2,1].Text 'Critical' 'SPO findings worksheet'; Assert-Equal $book.Workbook.Worksheets['Summary'].Cells[11,2].Text '91024' 'Estimated SPO tenant capacity in workbook'; Assert-Equal $book.Workbook.Worksheets['Summary'].Cells[12,1].Text 'Tenant storage utilization estimated percent' 'SPO utilization provenance in workbook'; Assert-Equal ([double]$book.Workbook.Worksheets['Summary'].Cells[12,2].Value) 0.55 'Estimated SPO utilization in workbook'; Assert-Equal $book.Workbook.Worksheets['Teams site links'].Cells[2,1].Text 'https://tenant.example/sites/channel' 'Teams site mapping worksheet' }
    finally { Close-ExcelPackage -ExcelPackage $book -NoSave }
    $spoEmptyBook = Join-Path $temporaryRoot 'SPO-empty.xlsx'
    New-SpoTimestampedWorkbook -CsvFiles $spoFiles -Path $spoEmptyBook | Out-Null
    Add-SpoWorkbookOverview -Path $spoEmptyBook -Summary $spoSummary -Alerts @() -Status 'OK'
    $book = Open-ExcelPackage -Path $spoEmptyBook
    try { Assert-Equal $book.Workbook.Worksheets['Findings'].Cells[1,1].Text 'Severity' 'Empty SPO findings worksheet header' }
    finally { Close-ExcelPackage -ExcelPackage $book -NoSave }

    $sourceDate = '2026-10-10T10:00:00Z'
    $teamsLatest = Join-Path $temporaryRoot 'TeamsLatest'
    $spoLatest = Join-Path $temporaryRoot 'SpoLatest'
    New-Item -ItemType Directory -Path $teamsLatest,$spoLatest | Out-Null
    foreach($name in @('Teams','Members','Channels','Guests')){
        [pscustomobject]@{TenantKey='prod';TenantId='tenant-synthetic';RunId='run-teams';RunDateUtc=$sourceDate;TeamId='team-one';TeamDisplayName='Example';OwnerCount=1;GuestCount=0;InactiveDays='';StorageQuotaPercent='';IsArchived='False';Visibility='Private';Status='OK';Details='';Threshold='Owners >= 2; inactive <= 180 days; storage <= 90 percent';ActivityEvidenceStatus='Observed';StorageCollectionStatus='Collected'} |
            Export-Csv -LiteralPath (Join-Path $teamsLatest "M365_Teams_$name.csv") -NoTypeInformation
    }
    $teamsSnapshot=Import-TeamsMailCsvSet -Folder $teamsLatest -ExpectedTenantKey 'prod'
    Assert-Equal $teamsSnapshot.RunId 'run-teams' 'Teams source run identity'
    Assert-Equal $teamsSnapshot.Files.Count 4 'Teams mail CSV count'
    $emptyTeamsPath=Join-Path $teamsLatest 'M365_Teams_Guests.csv'
    Set-Content -LiteralPath $emptyTeamsPath -Value '"TenantKey","TenantId","RunId","RunDateUtc"' -Encoding utf8
    Assert-Equal (Import-TeamsMailCsvSet -Folder $teamsLatest -ExpectedTenantKey 'prod').Rows.Guests.Count 0 'Empty Teams supporting CSV'
    (Get-Item -LiteralPath $emptyTeamsPath).LastWriteTimeUtc=(Get-Date).ToUniversalTime().AddHours(-2)
    $rejected=$false
    try { Import-TeamsMailCsvSet -Folder $teamsLatest -ExpectedTenantKey 'prod' | Out-Null } catch { $rejected=$true }
    if(-not $rejected){throw 'Teams MailOnly accepted an old empty CSV.'}
    [pscustomobject]@{TenantKey='prod';TenantId='tenant-synthetic';RunId='different-run';RunDateUtc=$sourceDate} |
        Export-Csv -LiteralPath (Join-Path $teamsLatest 'M365_Teams_Members.csv') -NoTypeInformation
    $rejected=$false
    try { Import-TeamsMailCsvSet -Folder $teamsLatest -ExpectedTenantKey 'prod' | Out-Null } catch { $rejected=$true }
    if(-not $rejected){throw 'Teams MailOnly accepted mixed run IDs.'}

    foreach($name in @('Sites','Lists','Permissions','ExternalSharing','Tenant')){
        [pscustomobject]@{TenantKey='prod';TenantId='tenant-synthetic';RunId='run-spo';RunDateUtc=$sourceDate;SiteUrl='https://example.invalid/site';IsOneDrive='False';IsInactive='False';IsOrphaned='False';StorageQuotaPercent='';Status='OK';Details='';Threshold='90%; 180 days';GraphEnrichmentStatus='Collected';LastActivityUtc='2026-10-09';ListTitle='Example list';ItemCount=1;IsPartialInventory='False';StorageCapacityGB='';StorageUtilizationPercent='';CapacitySource='NotMeasured'} |
            Export-Csv -LiteralPath (Join-Path $spoLatest "M365_SPO_$name.csv") -NoTypeInformation
    }
    $spoSnapshot=Import-SpoMailCsvSet -Folder $spoLatest -ExpectedTenantKey 'prod'
    Assert-Equal $spoSnapshot.RunId 'run-spo' 'SPO source run identity'
    Assert-Equal $spoSnapshot.Files.Count 5 'SPO mail CSV count'
    $emptySpoPath=Join-Path $spoLatest 'M365_SPO_Permissions.csv'
    Set-Content -LiteralPath $emptySpoPath -Value '"TenantKey","TenantId","RunId","RunDateUtc"' -Encoding utf8
    Assert-Equal (Import-SpoMailCsvSet -Folder $spoLatest -ExpectedTenantKey 'prod').Rows.Permissions.Count 0 'Empty SPO supporting CSV'
    $rejected=$false
    try { Import-SpoMailCsvSet -Folder $spoLatest -ExpectedTenantKey 'other' | Out-Null } catch { $rejected=$true }
    if(-not $rejected){throw 'SPO MailOnly accepted a different tenant.'}

    foreach($case in @(@{Name='Teams';Invoke={param($p,$a,$f) Invoke-TeamsDailySummaryMail -MarkerPath $p -SendAction $a -ForceSend:$f}},@{Name='SPO';Invoke={param($p,$a,$f) Invoke-SpoDailySummaryMail -MarkerPath $p -SendAction $a -ForceSend:$f}})){
        $marker=Join-Path $temporaryRoot "$($case.Name)-LastSent.txt"
        $today=(Get-Date).ToString('yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)
        [IO.File]::WriteAllText($marker,$today)
        $script:mailCount=0
        $action={ $script:mailCount++ }
        Assert-Equal (& $case.Invoke $marker $action $false) $false "$($case.Name) daily mail lock"
        Assert-Equal $script:mailCount 0 "$($case.Name) skipped send"
        Assert-Equal (& $case.Invoke $marker $action $true) $true "$($case.Name) forced send"
        Assert-Equal $script:mailCount 1 "$($case.Name) forced send count"
        Assert-Equal (Get-Content -LiteralPath $marker -Raw) $today "$($case.Name) marker after forced send"
        $failed=$false
        try { & $case.Invoke $marker { throw 'Synthetic send failure' } $true | Out-Null } catch { $failed=$true }
        if(-not $failed){throw "$($case.Name) send failure was swallowed."}
        Assert-Equal (Get-Content -LiteralPath $marker -Raw) $today "$($case.Name) marker after failed forced send"
    }
}
finally {
    $resolvedTemporaryRoot = [IO.Path]::GetFullPath($temporaryRoot)
    $resolvedTempParent = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\','/') + [IO.Path]::DirectorySeparatorChar
    if ($resolvedTemporaryRoot.StartsWith($resolvedTempParent,[StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolvedTemporaryRoot) -like 'SmartM365-MailOffline-*') {
        Remove-Item -LiteralPath $resolvedTemporaryRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
Write-Host 'Teams/SPO mail and workbook offline checks passed.'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDMrQ1GnN+6VHZU
# e2bIT4wXJ7T0Nz1fCbuAeZfxoYzaw6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIADELAPTPBTf9Tk6NLyRZvgRMvE3FVv6bC9XsdGMk5RBMA0GCSqG
# SIb3DQEBAQUABIIBgDFzxuYBo49BIsc1DvA+rZm2EPF797Z25VLvreIrZCMM/+vL
# WtFXGf96bzOw/nUMywg/zfZSpO2/U2Xv/4Uvd92NQT+S0TfL9EX/MzeyZ54U49hy
# PB36/rrKrblMZrMPYZ77dBPsYqtpjDoHJzP6xbymm4mjjoA+W8ix9WrSUpaq9KPF
# mniDAU5Plww3v3pVdAPX5NZEUwldBhx1t+/MBUkKy3VVijBPezscKMIcNdwMUwNg
# vo6xG4hWSenm7qI//NmfTWWQAAj/zzJAwoBg25R2tRFTajyChozCR1GzkGZmzUL3
# pKWrM27XWsmwHikvBfjr7YhjPYgUNAR85UUXRjEz+zRKortMVsGuiTCLuZHWlFFS
# 2TcnZwNo52p9viUNKV+ywRJgUNv7Qm3VI3/Df7yuJTOGvSrJCflemiTPVTx1obq/
# aLC4xz9eKvH05OdSHm1FktxT4+Nea7srTPzpUWuhmRniefF6asb4cA9fcUBAH1jt
# qL9lDcxWFpagEfipUKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTUwODA0
# MDZaMC8GCSqGSIb3DQEJBDEiBCC1e4XzXZmykrjn/+JbKNyf3mQupeLJF2escyr4
# 7Ho39zANBgkqhkiG9w0BAQEFAASCAgAgAVhfw/s0qIVRIeUcSMoGLH7GOJTrmHun
# jFqWg5nFbjVGhTr+b0pjUX5i5kJAuGrL3VUgf0tKtCtzhXLxJNuKE2xGgf1WB73f
# VdSe5zYg/Sk30XLpNNE2A8LcWKteN/LIhZoVbIDaW/W+rCDHNmHSWFt53x5fwCyS
# N3sv4KIcMpMzPdArb2aBQV3wT3oUMe+svK23KVfM7tCQiaQy3bn2eYn4KRTAmryV
# MvyuRVzZ/wYjwXev1+rDIgF+PZeSoGe4PQEu0/kZI1nkzukTSQYeFV/XDjP1fS50
# w109Ec7ID51Ah+CiJMX0cIFEdf2aO40WD29BYcth64ioVA+Qo+wB9xXbqzTsFAhZ
# orLajLnpU0CMHhV0dqLa9/qYS+FBgjgS7KoBJ05v7BePHpfNOGbl06DFoDh8vaRh
# O/ynN8Xc1fk3YNlYW98B5gW9vj/FH2Fj3H3gTSfOCGXwzN1mtT+oel8SVHqPx4Vq
# 1COvkXOMADqd7SWzb85xPG3Rx3rHr6dEEzVkCT3sjK2Wh0vnVeZ+85v1px+Hwb2S
# fsSKeR/K/kuRQDQI6XD0pXRh74LPk+ypyQk1RMq8Hll+H+Kb4aBca4C6s7qDKsoN
# 8SdCLP4FEa81IGLRcuJG66RjHzDy6K7NyqosSFMjoXDznz71elORsoIzPoYeWTgL
# Fr9azHITSg==
# SIG # End signature block
