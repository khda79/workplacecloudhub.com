#Requires -Version 7.0
<#
.SYNOPSIS
Offline checks for mail reconstruction from the latest previously published exports.
.VERSION
1.2.0
#>
[CmdletBinding()]
param([switch]$SkipWorkbookChecks)

$ErrorActionPreference='Stop'
$smartRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$core=Join-Path $smartRoot 'Modules/SmartM365.Core'
$gui=Join-Path $smartRoot 'SmartInventory/Orchestrator'
Import-Module (Join-Path $core 'SmartM365.Core.psd1')
. (Join-Path $smartRoot 'SmartInventory/Common/SmartM365.ReportInputs.ps1')
. (Join-Path $smartRoot 'SmartInventory/Common/SmartM365.ReportRenderers.ps1')
Import-Module (Join-Path $gui 'SmartM365.Orchestrator.RebuiltReport.psm1')
. (Join-Path $core 'SmartM365-MailAdminRouting.ps1')
. (Join-Path $core 'SmartM365-MailMaintenance.ps1')

function Assert-Rebuild($Condition,[string]$Message) { if (-not $Condition) { throw $Message } }

# All 13 scripts listed in the cadence editor have a reader or a native mail-only branch.
$inventoryRoot=Join-Path $smartRoot 'SmartInventory'
$templates = @(Get-ChildItem -Path $inventoryRoot -Recurse -Filter '*.local.json.template' |
    Where-Object { (Get-Content -LiteralPath $_.FullName -Raw) -match '"ReportMailCadence"' })
$covered = @($templates | ForEach-Object {
    Get-SmartM365RebuiltReportMode -Producer ($_.Name -replace '\.local\.json\.template$')
})
Assert-Rebuild ($covered.Count -eq 13 -and @($covered | Where-Object { $_ -eq 'Exports' }).Count -eq 9) 'Expected nine export adapters and four native entry points.'

$base=[IO.Path]::GetTempPath()
$root=[IO.Path]::GetFullPath((Join-Path $base ('Wch-RebuildReport-' + [guid]::NewGuid().ToString('N'))))
$latest=Join-Path $root 'DATA-LAST'
$scratch=Join-Path $root 'scratch'
[void][IO.Directory]::CreateDirectory($latest)
[void][IO.Directory]::CreateDirectory($scratch)
$cfg=[pscustomobject]@{Smart365Administrators='admin@example.invalid;admin2@example.invalid'}
$script:mailRestricted=$false
function Get-SmartM365EffectiveModuleGlobalConfig { $cfg }
function Resolve-SmartM365ConfigValue { param($Value) $Value }
function Resolve-SmartM365MaintenanceMailRecipientRoute {
    param($To,$Cc,$MailPurpose)
    [pscustomobject]@{To=$To;Cc=$Cc;Restricted=$script:mailRestricted}
}
try {
    $global:SmartM365TenantKey='demo'
    function Export-FixtureCsv {
        param([string]$Name,[object[]]$Rows)
        $Rows | Export-Csv -LiteralPath (Join-Path $latest $Name) -NoTypeInformation
    }
    function Save-FixtureInputs {
        param([string]$Producer,[string[]]$Files,[hashtable]$InputData,[string]$BatchId='', [string]$BatchManifestHash='')
        Save-SmartM365ReportInputs -Producer $Producer -LatestFolder $latest -TenantKey demo -Files $Files -InputData $InputData -BatchId $BatchId -BatchManifestHash $BatchManifestHash
    }
    function Build-FixtureReport {
        param([string]$Producer)
        New-SmartM365RebuiltReport -Producer $Producer -LatestFolder $latest -TenantKey demo -ScratchFolder $scratch
    }

    $source = Join-Path $latest 'Exchange_Hybrid_FreeBusy_Summary.csv'
    $findingName = 'Exchange_Hybrid_FreeBusy_Findings.csv'
    $stamp = '2026-11-18T19:00:00+01:00'
    $summaryRow = [pscustomobject]@{TenantKey='demo';Side='ExchangeOnline';Auth='OAuth';RunStatus='Completed';SideStatus='Completed'
        RunMessage='';MissingRoles='';GenerationTime=$stamp;FreeBusyCheckerVersion='test';ErrorCount=0;WarningCount=1;NotAuthorizedCount=0}
    $findingRow = [pscustomobject]@{TenantKey='demo';Side='ExchangeOnline';Auth='OAuth';Status='Warning';GenerationTime=$stamp
        RowNumber=1;Category='Configuration';Item='Synthetic old setting';Value='old value'}
    Export-FixtureCsv -Name 'Exchange_Hybrid_FreeBusy_Summary.csv' -Rows $summaryRow
    Export-FixtureCsv -Name $findingName -Rows $findingRow
    $producer = 'SmartM365-Exchange-Hybrid-FreeBusy-Inventory'
    $first = Build-FixtureReport $producer
    $expected = ConvertTo-FreeBusyReportHtml -Summary @($summaryRow) -Findings @($findingRow) -RunStatus 'Completed' -CheckerVersion 'test' -ReportRunId '20261118-190000'
    Assert-Rebuild ($first.BodyHtml -ceq (Add-SmartM365FreeBusyTemplateMarker $expected) -and $first.Attachments[0] -eq $source) 'Free/Busy did not use the normal report renderer and file reference.'
    $findingRow.Item = 'Synthetic current setting'
    Export-FixtureCsv -Name $findingName -Rows $findingRow
    $second = Build-FixtureReport $producer
    Assert-Rebuild ($second.BodyHtml -match 'Synthetic current setting' -and $second.BodyHtml -notmatch 'Synthetic old setting') 'Free/Busy did not use the latest export.'
    $summaryRow.TenantKey = 'other'
    Export-FixtureCsv -Name 'Exchange_Hybrid_FreeBusy_Summary.csv' -Rows $summaryRow
    $failed=$false
    try { $null=Build-FixtureReport $producer } catch {$failed=$_.Exception.Message -like '*different or empty tenant*'}
    Assert-Rebuild $failed 'Cross-tenant data was accepted.'

    $producer = 'SmartM365-ActiveDirectory-HealthCheck'
    $healthName='AD_HealthCheck.csv'
    $healthRow=[pscustomobject]@{TenantKey='demo';Forest='demo.invalid';Domain='demo.invalid';DC='dc1';Category='Connectivity';Check='Ping';Status='Warning';NumericValue=0;Threshold='Reachable';Details='Synthetic check'}
    Export-FixtureCsv -Name $healthName -Rows $healthRow
    $healthInputs=@{Status='Warning';Started='2026-11-18T19:00:00Z';Ended='2026-11-18T19:03:00Z';Csv='health-run.csv'
        ReportRunId='health-run-1';ReportHost='AD-DEMO';Facts=@([pscustomobject]@{Domain='demo.invalid';DCCount=1;GlobalCatalogCount=1})
        ForestInfo=@{Name='demo.invalid';RootDomain='demo.invalid';ForestMode='Windows2016Forest';Domains=@('demo.invalid')};RemoteDcAdminChecksEnabled=$false}
    Save-FixtureInputs -Producer $producer -Files $healthName -InputData ($healthInputs + @{RunDateUtc='2026-11-18T19:00:00Z'})
    $health=Build-FixtureReport $producer
    $healthBody=ConvertTo-ReportHtml @healthInputs -r @($healthRow)
    Assert-Rebuild ($health.BodyHtml -ceq $healthBody -and $health.Subject -match '\[WARNING\]') 'AD HealthCheck differs from the normal report.'
    $healthRow.Details='changed after input publication'
    Export-FixtureCsv -Name $healthName -Rows $healthRow
    $failed=$false
    try { $null=Build-FixtureReport $producer } catch {$failed=$_.Exception.Message -like '*do not match the latest export*'}
    Assert-Rebuild $failed 'Mismatched report inputs and current CSV were accepted.'

    $producer='SmartM365-SharePoint-OnPrem-Content-Inventory'
    $spNames=@('SharePoint_OnPrem_SiteCollections.csv','SharePoint_OnPrem_Webs.csv','SharePoint_OnPrem_CollectionCoverage.csv','SharePoint_OnPrem_SiteAdministrators.csv')
    foreach ($name in $spNames) { Export-FixtureCsv -Name $name -Rows $healthRow }
    $spInputs=@{Status='Partial';ContentStatus='Qualified';Infrastructure=[pscustomobject]@{Status='Stale';Detail='Synthetic infrastructure';CollectedAtUtc='';Rows=$null}
        Stats=@{WebApplications=1;ContentDatabases=1;SiteCollections=2;Webs=3;CoverageIssues=0;ReadOnlyCollections=0;UnverifiedLocks=0;CoverageSamples=@();ReadOnlySamples=@();SiteAdministrators=1;CollectionCoverageRows=2};Failures=@();UploadFailures=0;TenantName='demo';Duration=[timespan]::FromSeconds(35);HostName='SP-DEMO';ReportGeneratedAt='2026-11-18 19:00:00 +00:00'}
    Save-FixtureInputs -Producer $producer -Files $spNames -InputData $spInputs
    Assert-Rebuild ((Build-FixtureReport $producer).BodyHtml -ceq (New-ContentMailHtml @spInputs)) 'SharePoint on-premises differs from the normal report.'

    $generated = '2026-11-18 19:00:00 +00:00'
    $producer='SmartM365-ActiveDirectory-Inventory'
    $adNames=@('AD_Users_AllDomains.csv','AD_Computers_AllDomains.csv')
    foreach ($name in $adNames) { Export-FixtureCsv -Name $name -Rows $healthRow }
    # Test the worker entry point with the real renderer and a simulated transport.
    $runtime = Join-Path $root 'worker-runtime'
    $workerDir = Join-Path $runtime 'SmartInventory/Orchestrator'
    $fixtureConfig = Join-Path $runtime 'Config'
    $fixtureCore = Join-Path $runtime 'Modules/SmartM365.Core'
    $fixtureCommon = Join-Path $runtime 'SmartInventory/Common'
    foreach ($directory in @($workerDir,$fixtureConfig,$fixtureCore,$fixtureCommon)) { $null = [IO.Directory]::CreateDirectory($directory) }
    foreach ($name in @('SmartM365-SendPreparedReportMail.ps1','SmartM365.Orchestrator.RebuiltReport.psm1')) {
        Copy-Item -LiteralPath (Join-Path $gui $name) -Destination $workerDir
    }
    foreach ($name in @('SmartM365.ReportRenderers.ps1','SmartM365.ReportInputs.ps1')) {
        Copy-Item -LiteralPath (Join-Path $inventoryRoot "Common/$name") -Destination $fixtureCommon
    }
    @'
function Initialize-SmartM365TenantContext {
    param($Tenant,$StartPath,$ScriptName)
    if ($Tenant -ne 'demo') { throw 'Unexpected fixture tenant.' }
    [pscustomobject]@{TenantKey='demo';LatestCsvFolderPath=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../DATA-LAST'))}
}
function Get-SmartM365ConfigurationValue {
    param($Name,$LocalConfig,$EffectiveConfig,$DefaultValue,$EffectiveFallback)
    $p = $EffectiveConfig.PSObject.Properties[$Name]
    if ($p) { $p.Value } else { $DefaultValue }
}
'@ | Set-Content -LiteralPath (Join-Path $fixtureConfig 'SmartM365-TenantContext.ps1')
    @'
@{RootModule='SmartM365.Core.psm1';ModuleVersion='1.0.132';FunctionsToExport='*'}
'@ | Set-Content -LiteralPath (Join-Path $fixtureCore 'SmartM365.Core.psd1')
    @'
function InitializeScriptEnvironment {
    param($OutputPath,$LogFileName,$CallerScriptPath)
    if (-not $OutputPath -or -not [IO.Directory]::Exists($OutputPath)) { throw 'The mail-only worker must provide an existing output folder.' }
    if (-not ([IO.FileSystemAclExtensions]::GetAccessControl([IO.DirectoryInfo]::new($OutputPath))).AreAccessRulesProtected) { throw 'The work folder is not private.' }
    Write-Host 'FIXTURE: environment initialized'
}
function WriteLog {
    param($Message,$Level)
    Write-Host "FIXTURE LOG: $Message"
}
function Complete-SmartM365ExecutionContext { param($Status) }
function ConvertTo-SmartM365EmailHtmlText { param($Value) [Net.WebUtility]::HtmlEncode([string]$Value) }
function New-SmartM365EmailBody { '<p>Fictional report body</p>' }
function Set-SmartM365ManualReportAudience { param($Audience) $script:Audience=$Audience }
function SendEmailHtmlReport {
    param($Subject,$BodyHtml,$To,$Cc,$From,$MailPurpose,$DeliveryClass,$Attachments,$AllowAttachments,$SuppressAttachmentLinks,$SmtpServer,$SmtpPort)
    if ($script:Audience -ne 'Administrators' -or $MailPurpose -ne 'Report' -or $DeliveryClass -ne 'Operational' -or $BodyHtml -notmatch 'Fictional report body') { throw 'Unexpected simulated delivery.' }
    Write-Host 'FIXTURE: delivery simulated'
}
'@ | Set-Content -LiteralPath (Join-Path $fixtureCore 'SmartM365.Core.psm1')
    @'
function Get-SmartM365OrchestratorMailOnlySource {
    param($InventoryRoot,$Producer)
    if ($Producer -ne 'SmartM365-ActiveDirectory-Inventory') { throw 'Unexpected fixture report producer.' }
    [pscustomobject]@{Config=[pscustomobject]@{}}
}
'@ | Set-Content -LiteralPath (Join-Path $workerDir 'SmartM365.Orchestrator.ReportCadence.psm1')
    $workerPath = Join-Path $workerDir 'SmartM365-SendPreparedReportMail.ps1'
    function Assert-FixtureWorker {
        param([int]$ExpectedCode,[string]$ExpectedOutput,[switch]$ExpectedDelivery)
        $output = & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $workerPath -Tenant demo -Producer 'SmartM365-ActiveDirectory-Inventory' -Audience Administrators 2>&1 | Out-String
        $actualCode = $LASTEXITCODE
        Assert-Rebuild ($actualCode -eq $ExpectedCode -and $output -match 'FIXTURE: environment initialized' -and $output -like "*$ExpectedOutput*") "Unexpected mail worker outcome ($actualCode instead of $ExpectedCode): $output"
        Assert-Rebuild (($output -match 'FIXTURE: delivery simulated') -eq [bool]$ExpectedDelivery) 'The worker unexpectedly reached or skipped the simulated transport.'
    }
    Assert-FixtureWorker -ExpectedCode 21 -ExpectedOutput 'missing the original report inputs'

    $snapshot=[pscustomobject]@{
        SnapshotDate='2026-11-18';GeneratedAt=$generated;DomainCount=1;TotalUsers=12;TotalComputers=4;TotalGroups=3;TotalOUs=2;TotalContacts=1
        DistinctDuplicateUPNs=0;AffectedDuplicateUPNAccounts=0;DistinctDuplicateSMTPAddresses=0;AffectedDuplicateSMTPEntries=0
    }
    $domainCounts=@{}
    foreach ($name in 'Users','Computers','Groups','OUs','Contacts','DuplicateUPN','DuplicateSMTP') { $domainCounts[$name]=@{'demo.invalid'=0} }
    $adInputs=@{SummarySnapshot=$snapshot;PreviousSnapshot=$null;WeekSnapshot=$null;MonthSnapshot=$null;LogicalCountBaselineAt=$generated
        SourceFolder='DATA-ALL';SummaryCsvPath='AD_Inventory_DailySummary.csv';LatestSummaryPath='';SummarySharePointUploads=@()
        ShowMailLinks=$false;DomainCounts=$domainCounts;DomainNames=@('demo.invalid');Tenant='demo';ReportHost='AD-DEMO';ReportGeneratedAt=$generated}
    Save-FixtureInputs -Producer $producer -Files $adNames -InputData $adInputs
    $normalAd=New-SmartM365AdInventoryMailReport @adInputs
    $rebuiltAd=Build-FixtureReport $producer
    Assert-Rebuild ($rebuiltAd.Subject -ceq $normalAd.Subject -and $rebuiltAd.BodyHtml -ceq $normalAd.BodyHtml) 'AD inventory differs from the normal report.'
    Assert-FixtureWorker -ExpectedCode 0 -ExpectedOutput 'Rebuilt report mail accepted' -ExpectedDelivery
    $adSourcePath = Join-Path $latest $adNames[0]
    $adCsv = [IO.File]::ReadAllBytes($adSourcePath)
    try {
        [IO.File]::AppendAllText($adSourcePath,[Environment]::NewLine)
        Assert-FixtureWorker -ExpectedCode 22 -ExpectedOutput 'do not match the latest export'
    } finally {
        [IO.File]::WriteAllBytes($adSourcePath,$adCsv)
    }


    $producer='SmartM365-Exchange-Local-Mailboxes-Inventory'
    $mailboxName='Exchange_OnPrem_Mailboxes_AllDomains.csv'
    Export-FixtureCsv -Name $mailboxName -Rows $healthRow
    $mailboxRow=[pscustomobject]@{DomainName='demo.invalid';TotalMailboxCount=12;TotalLocalMailboxCount=5;TotalRemoteMailboxCount=7;EnabledAccounts=9;DisabledAccounts=3;TotalLocalMailboxSizeGB=17.5}
    $mailboxInputs=@{ReportRows=@($mailboxRow);DailyStatsCsv='daily.csv';SummaryCsv='summary.csv';Title='Exchange mailbox report'
        TenantName='demo';ReportHost='EX-DEMO';ReportGeneratedAt=$generated;FileRowCounts=@{'daily.csv'=1;'summary.csv'=1}}
    $mailboxSnapshot=[pscustomobject]@{SnapshotDate='2026-11-18';GeneratedAt=$generated;DomainCount=1;TotalMailboxCount=12;TotalLocalMailboxCount=5;TotalRemoteMailboxCount=7;EnabledAccounts=9;DisabledAccounts=3;TotalLocalMailboxSizeGB=17.5;LocalSizeCoverageComplete=$true}
    Save-FixtureInputs -Producer $producer -Files $mailboxName -InputData @{Mail=$mailboxInputs;CurrentSnapshot=$mailboxSnapshot;History=@()}
    $normalMailbox=New-SmartM365ExchangeLocalMailboxReportEmailBody @mailboxInputs -DailyDiffSection (New-SmartM365ExchangeMailboxDailyDiffSection -Current $mailboxSnapshot)
    Assert-Rebuild ((Build-FixtureReport $producer).BodyHtml -ceq $normalMailbox) 'Exchange mailbox inventory differs from the normal report.'

    $producer='SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory'
    $infraNames=@('Exchange_OnPrem_Servers_Inventory_Summary.csv','Exchange_OnPrem_Infrastructure_PerServerSummary.csv','Exchange_OnPrem_MigrationReadiness_Config.csv')
    $infraSummary=[pscustomobject]@{TenantKey='demo';RunId='infra-1';ExecutionDate='2026-11-18';ComputeCollectionErrors=0;DiskDriveCollectionErrors=0;LogicalDiskCollectionErrors=0}
    Export-FixtureCsv -Name $infraNames[0] -Rows $infraSummary
    [IO.File]::WriteAllText((Join-Path $latest $infraNames[1]),"TenantKey,ExchangeServerName`n")
    [IO.File]::WriteAllText((Join-Path $latest $infraNames[2]),"TenantKey,Category,CollectionStatus,Importance`n")
    $infraInputs=@{Files=@();Tenant='demo';WarningCount=0;ErrorCount=0;CertificateBudgetSeconds=120;ReferenceDate=[datetime]'2026-11-18T19:00:00';ReportHost='EX-DEMO'}
    Save-FixtureInputs -Producer $producer -Files $infraNames -InputData ($infraInputs + @{Subject='Infrastructure normal'})
    Set-Alias -Name New-CoreSmartM365EmailBody -Value New-SmartM365EmailBody
    Set-Alias -Name New-CoreSmartM365EmailKpiGridHtml -Value New-SmartM365EmailKpiGridHtml
    Set-Alias -Name New-CoreSmartM365EmailTableHtml -Value New-SmartM365EmailTableHtml
    Assert-Rebuild ((Build-FixtureReport $producer).BodyHtml -ceq (New-ServersAndStorageMailBodyHtml @infraInputs -Summary $infraSummary)) 'Exchange infrastructure differs from the normal report.'

    $producer='SmartM365-WinUpdate_Status_From_Intune'
    $intuneName='Intune_WindowsUpdate_Status.csv'
    Export-FixtureCsv -Name $intuneName -Rows $healthRow
    $intuneInputs=@{RunId='intune-1';OrgDomain='demo.invalid';policyCount=0;count=0;durationSeconds=12;allPolicySummary=@();actionRows=@();actionLimit=10
        coverageAvailable=$false;readinessMode='None';EnableReadinessEnrichment=$false;reportStatus='OK'
        reportHealth=[pscustomobject]@{ActionRequiredRatePct=0;Priority0Threshold=0;WarningRateThresholdPct=0;CriticalRateThresholdPct=0}
        statusAgeSourceColumn='NotAvailable';statusAgeCoveragePct=0;statusAgePopulatedRows=0;computedRowCount=0;CsvLastFinal='Intune_WindowsUpdate_Status.csv'}
    Save-FixtureInputs -Producer $producer -Files $intuneName -InputData $intuneInputs
    $normalIntune=New-SmartM365WindowsUpdateMailReport @intuneInputs
    $rebuiltIntune=Build-FixtureReport $producer
    Assert-Rebuild ($rebuiltIntune.Subject -ceq $normalIntune.Subject -and $rebuiltIntune.BodyHtml -ceq $normalIntune.BodyHtml) 'Windows Update differs from the normal report.'

    if (-not $SkipWorkbookChecks) {
        $producer='SmartM365-EXO-QuarantineMessages-Report'
        $csvName='Exchange_EXO_QuarantineMessages.csv'
        $end=[datetime]::SpecifyKind([datetime]'2026-11-18T19:00:00', [DateTimeKind]::Utc)
        $columns=@(Get-SmartM365QuarantineReportColumns)
        $row=[ordered]@{}
        foreach ($column in $columns) { $row[$column]='' }
        $row.TenantKey='demo'; $row.RunId='quarantine-run'; $row.CollectedAtUtc=$end.ToString('o')
        $row.ReceivedTimeUtc=$end.AddDays(-2).ToString('o'); $row.AgeDays=2.0; $row.QuarantineType='Spam'
        $row.Subject="'=1+1"
        $row=[pscustomobject]$row
        Export-FixtureCsv -Name $csvName -Rows $row
        Save-FixtureInputs -Producer $producer -Files $csvName -InputData @{EndUtc=$end; LookbackDays=7; WorkbookGeneratedUtc=$end.AddMinutes(1); LocalDate='2026-11-18'}
        $mail=Build-FixtureReport $producer
        $book=$mail.Attachments[0]
        $normal=Join-Path $root ([IO.Path]::GetFileName($book))
        Initialize-SmartM365ImportExcel -CheckOnly | Out-Null
        Export-SmartM365QuarantineWorkbook -Rows @($row) -DetailColumns $columns -Path $normal -StartReceivedDateUtc $end.AddDays(-7) -EndReceivedDateUtc $end -EffectiveLookbackDays 7 -GeneratedUtc $end.AddMinutes(1) | Out-Null
        Assert-Rebuild ($mail.BodyHtml -ceq (Format-SmartM365QuarantineMailBody -Rows @($row) -StartReceivedDateUtc $end.AddDays(-7) -EndReceivedDateUtc $end -WorkbookPath $normal)) 'The quarantine mail did not use the native body.'
        foreach ($sheet in 'Summary','By_Reason','Quarantine_Messages') {
            Assert-Rebuild (((Import-Excel -Path $book -WorksheetName $sheet) | ConvertTo-Json -Compress) -ceq ((Import-Excel -Path $normal -WorksheetName $sheet) | ConvertTo-Json -Compress)) "Quarantine worksheet differs from the normal workbook: $sheet"
        }
        $adSourceNames = @(Get-SmartM365AdDiagnosticWorkbookSource -Folder $latest | Select-Object -ExpandProperty Name)
        foreach ($name in $adSourceNames) {
            [IO.File]::WriteAllText((Join-Path $latest $name),"TenantKey,DomainName`n")
        }
        $identityReference = [pscustomobject]@{ Name='AD_Users_IdentityAndMailRoutingIssues.xlsx';Path='DATA-ALL/AD_Users_IdentityAndMailRoutingIssues.xlsx';SizeKB=10;WebUrl='';SharePointPath='' }
        $identityData = @{SharePointUploads=@();PathRows=@();WorkbookFile=$identityReference;PreviewLimit=5;UserDomains=@('demo.invalid')}
        Save-FixtureInputs -Producer 'SmartM365-ActiveDirectory-Inventory' -Files ($adNames + $adSourceNames) -InputData ($adInputs + @{Identity=$identityData})
        $adIdentity = Build-FixtureReport 'SmartM365-ActiveDirectory-Inventory'
        Assert-Rebuild ($adIdentity.Subject -match 'and identity diagnostics' -and $adIdentity.BodyHtml -match 'Remote routing issues' -and $adIdentity.BodyHtml -match 'Technical files' -and $adIdentity.Attachments.Count -eq 1) 'The AD combined report did not rebuild its native diagnostic attachment.'
        $pkg = Open-ExcelPackage -Path $adIdentity.Attachments[0]
        try { Assert-Rebuild ($pkg.Workbook.Worksheets.Count -eq 4) 'AD diagnostic workbook lost worksheets.' }
        finally { $pkg.Dispose() }
    }

    $prepared=Join-Path $root 'DATA-POWERBI'
    $batch=Join-Path $prepared 'batches/batch-1'
    [void][IO.Directory]::CreateDirectory($batch)
    $manifestPath=Join-Path $batch 'batch.json.txt'
    [IO.File]::WriteAllText($manifestPath, '{"BatchId":"batch-1","TenantKey":"demo","CreatedUtc":"2026-11-18T20:00:00Z","Files":[{"File":"example.csv"}]}')
    $pointerPath=Join-Path $prepared 'current.json.txt'
    [IO.File]::WriteAllText($pointerPath, (@{SchemaVersion=1;TenantKey='demo';BatchId='batch-1';ManifestSHA256=(Get-FileHash -LiteralPath $manifestPath).Hash}|ConvertTo-Json))
    $prepInputs=@{Details=[ordered]@{'Batch ID'='batch-1';'Prepared CSV files'=1};Warnings=@('Synthetic preparation warning')}
    Save-SmartM365ReportInputs -Producer 'SmartM365-WorkplaceEvidence-Prepare' -LatestFolder $latest -TenantKey demo -BatchId 'batch-1' -BatchManifestHash (Get-FileHash -LiteralPath $manifestPath).Hash -InputData $prepInputs
    $prep=Build-FixtureReport 'SmartM365-WorkplaceEvidence-Prepare'
    Assert-Rebuild ($prep.BodyHtml -ceq (New-SmartM365PreparedCompletionMailBody @prepInputs) -and -not $prep.PSObject.Properties['Attachments']) 'Prepare mail differs from the normal completion message.'
    [IO.File]::AppendAllText($manifestPath,' ')
    $failed=$false
    try { $null=New-SmartM365RebuiltReport -Producer 'SmartM365-WorkplaceEvidence-Prepare' -LatestFolder $latest -TenantKey demo -ScratchFolder $scratch } catch {$failed=$_.Exception.Message -like '*does not match its pointer*'}
    Assert-Rebuild $failed 'A divergent prepared manifest was accepted.'

    $script:SmartM365ReportReplayAudience='Administrators'
    $route=Resolve-SmartM365UnifiedMailRoute -To 'report@example.invalid' -Cc 'copy@example.invalid' -Bcc 'hidden@example.invalid'
    Assert-Rebuild ($route.To -eq $cfg.Smart365Administrators -and -not $route.Cc -and -not $route.Bcc) 'The administrator-only route included another address.'
    $script:SmartM365ReportReplayAudience='Configured'
    $route=Resolve-SmartM365UnifiedMailRoute -To 'report@example.invalid' -Cc 'copy@example.invalid'
    Assert-Rebuild ($route.To -like 'report@example.invalid;*' -and $route.Cc -eq 'copy@example.invalid') 'Configured routing did not use current recipients.'
    $script:mailRestricted=$true
    $failed=$false
    try { $null=Resolve-SmartM365UnifiedMailRoute -To 'report@example.invalid' } catch { $failed=$_.Exception.Message -like '*restricted by the mail policy*' }
    Assert-Rebuild $failed 'The configured route ignored the current restriction.'
    $workbookResult = if ($SkipWorkbookChecks) { 'workbook checks skipped' } else { 'AD and quarantine workbook parity' }
    "PASS: 13 routes, native mail parity, $workbookResult, latest exports, tenant isolation, and both recipient choices. No live operations."
}
finally {
    if ([IO.Path]::GetFileName($root) -notmatch '^Wch-RebuildReport-[a-f0-9]{32}$' -or (Split-Path $root -Parent).TrimEnd('\') -ne $base.TrimEnd('\')) { throw 'Unexpected fixture root.' }
    Remove-Item -LiteralPath $root -Recurse -Force
}

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCdtM927Uf1/d5m
# zP+MK3oGZYlpAq/EetIxHvbjeD0UN6CCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgL2WoEZADg0OI
# zPgQvch1VanSgU1rgIJubXjzXtq8/I4wDQYJKoZIhvcNAQEBBQAEggGAm2nh1feB
# flL7VYx/ugX1mrjauXL1DD7Q0igZhEgZhW+1nZ/mrZBwNhJa2p3ztqse7zoDqZAl
# JG4lCdJdj2cei7FU6loMuNVB0XsypI/3ewxcIvEITbopV4G2vgztEWhVM13/tk0z
# f7S7Bx8vFwMYY0TWx5jmYQ2v5Udkl4wTIlIyJvoiSpudMCNW5WODmm+nGSAKsb55
# Mm5JYHK6C72n+SskLuVMBdZqIpCvwiTgEhTbb2I2zM4dJly06s7vQOTxkDbeKTKB
# kZNUW+qw2FiwVMx948ibm4MLKSFldM5+VFFB3r17MEc4L2oIMGIqRIEz/eO9fEJP
# 76TWvSB+mW71iy0C02mvtI5C08iHUxfZRJnz14k7hKm0QWxV777jnP66ZRklt3Qv
# QGD67L12Mst5nicUwGPGPgq4eCbyu6pvA7ac4wI00+mdKo6X42hldUQKonSXlaY1
# aAWT3hbpiI7UzJdfshvxUd6gbbikaHfHap8Oh9msjdz3gScOTEDUK2nxoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MTAxNTE3WjA/BgkqhkiG9w0BCQQxMgQw2xKwOEO9Xc3mVuI2zCpmqcVc
# KuKnwlb/Qk2mZp3InRtNqvAPzyrO2Uh3Zvz1hF9XMA0GCSqGSIb3DQEBAQUABIIC
# AHOJAjDoPuO3d5e0GkAgcdwkenBA16nwksORzDmgdpqNadCxNAJyEkfFfFlg/rPA
# vxmH5gcyhF7Yl6FbtonzGMrb2XLKlcOBYTvic/sGqk7Rjfmi8qDnhtNV99C4Pz2z
# k4hRrh/4qhF8QttDJOiK+uGo7boX88646iQWELa29DO3VuXkX4LFyWAtTgxnp8Z3
# DxTRi9e3uT5viZtsHqY2XF3khM6Agog3XzBLB111Z08TYWVotF0bM7X8VthhFnXb
# 8O6cxZ2kbamOEzvATXkaMgyaCiumzFZw0W6U1KIyUwE9n7PlIRU6hh7/7ljKx2h0
# kqgMMaAGt/G6z7XMszkS46mL/j3957Un9j1ya35ElUwVw021FBdRCOmv81EOgNy+
# N7ZuL0VVr7drDFy+Rq+aD70y0kbca+Iy6S1zMYfJE+3NDp/0WoyJY0wQdv1Dtl1L
# 7AxBCcFG65vnbrmJVMtowKQLwFKSa8Y+rxI6+pyzQm+Fmr9U/uU5X49RjzlTqLL4
# 3mSt/zEbmF0bp5W2coIGvy4xiKjIuADEIr0QTkve3GfqaAGdKf20AZ2hJQ38hfiu
# Zjf2pKc3/MfHfU21irmY8yB3NX6z4ixLR+GUPPARDxHC9SvIbx3HGUwUYVBmjtbJ
# hvWROUUbIBxKHjK4WUqz/6A4tj+dxmvHZu6qbdjbhjvx
# SIG # End signature block
