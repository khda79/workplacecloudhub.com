<#
.SYNOPSIS
Synthetic database email tests; extracts functions without running the collector.
.VERSION
1.0.0
#>
[CmdletBinding()]
param([string]$SourceRoot)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $SourceRoot) { $SourceRoot = Split-Path $PSScriptRoot -Parent }
$collectorPath = Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/ServersAndStorage/SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1'
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($collectorPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw "Collector parse failed: $parseErrors" }
$functionNames = @('Format-HtmlValue', 'ConvertTo-ServersAndStorageDatabaseSectionHtml', 'New-HtmlExecutiveSummary',
    'Import-ServersAndStorageExistingCsv', 'Invoke-ServersAndStorageHtmlReportRegeneration', 'Send-ServersAndStorageHtmlReport')
$definitions = foreach ($name in $functionNames) {
    $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    if ($null -eq $node) { throw "Missing function: $name" }
    $node.Extent.Text
}
# Never import the runtime Core module or load tenant configuration. All external
# operations reachable from regeneration/mail are replaced inside this module.
$module = New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")))
& $module {
    Set-StrictMode -Version Latest
    $script:CapturedMail = $null
    function script:Import-SmartM365CoreModule { }
    function script:Write-Log { param($Message) }
    function script:Invoke-ServersAndStorageSharePointUpload { param($LocalFilePath) }
    function script:Get-SmartM365EffectiveConfigValue { param($Name, $DefaultValue) return $DefaultValue }
    function script:CoreSendEmailHtmlReport {
        param($From, $To, $Subject, $BodyHtml, $Attachments)
        $script:CapturedMail = [pscustomobject]@{ Body=$BodyHtml; Attachments=$Attachments }
    }
    function script:Get-MailboxDatabase { throw 'A synthetic report must not collect Exchange data.' }
}
$script:Checks = 0
function Assert-DatabaseMail {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
    $script:Checks++
}
function New-ReadinessRow {
    param([string]$Name, [string]$Setting, [AllowNull()][object]$Value, [string]$Status = 'OK', [string]$Category = 'MailboxDatabase')
    [pscustomobject]@{Category=$Category; ObjectName=$Name; Setting=$Setting; Value=$Value; CollectionStatus=$Status; Importance='Review'; MigrationFocus='Synthetic fixture'; ErrorMessage=''}
}
function Get-DatabaseSection {
    param([AllowNull()][object[]]$Rows)
    & $module { param($r,$s) ConvertTo-ServersAndStorageDatabaseSectionHtml -ReadinessInventory $r -Summary $s } $Rows $summary
}
function Assert-Counters {
    param([string]$Html, [string]$Total, [string]$Mounted, [string]$NotMounted, [string]$Unknown)
    foreach ($pair in @(@('Total databases observed',$Total), @('Mounted',$Mounted), @('Not mounted',$NotMounted), @('Unknown status',$Unknown))) {
        Assert-DatabaseMail ($Html -cmatch ([regex]::Escape($pair[0]) + ': <strong>' + [regex]::Escape($pair[1]) + '</strong>')) "Wrong counter: $($pair[0]) = $($pair[1])"
    }
}

$summary = [pscustomobject]@{ExecutionDate='2020-01-02 03:04:05';RunId='SYNTHETIC-OLD-RUN';ExchangeServersCount=1;TotalLogicalProcessorCount=4;TotalMemoryGB=16;TotalDiskDriveCount=1;TotalDiskDriveSizeTB=1;ExchangeSchemaRangeUpper='N/A';ExchangeOrgObjectVersion='N/A';ExcludedServers='EX-EXCLUDED'}
$server = [pscustomobject]@{ExchangeServerName='EX-01';ServerRole='Mailbox';LogicalProcessorCount=4;MemoryGB=16;DiskDriveCount=1;DiskDriveTotalSizeGB=1024;ComputeCollectionStatus='OK';DiskDriveCollectionStatus='OK';LogicalDiskCollectionStatus='OK';LowSpaceLogicalDiskCount=0}
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-DatabaseMail-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)
try {
    $rows = @(
        New-ReadinessRow 'DB-A' 'Mounted' $true
        New-ReadinessRow 'DB-A' 'Server' 'EX-01'
        New-ReadinessRow 'DB-A' 'DatabaseSize' '10 GB (10,737,418,240 bytes)'
        New-ReadinessRow 'DB-A' 'AvailableNewMailboxSpace' '0 B (0 bytes)'
        New-ReadinessRow 'DB-B' 'Mounted' $false
        New-ReadinessRow 'DB-B' 'Server' 'EX-EXCLUDED'
        New-ReadinessRow 'DB-C' 'Mounted' ''
        New-ReadinessRow 'DB-D' 'Mounted' 'not-a-boolean'
        New-ReadinessRow 'DB-D' 'Recovery' 'True'
        New-ReadinessRow 'DAG-A' 'Servers' 'EX-01; EX-02' -Category 'DatabaseAvailabilityGroup'
        New-ReadinessRow 'DAG-B' 'Servers' 'EX-03; EX-04' -Category 'DatabaseAvailabilityGroup'
    )
    $html = Get-DatabaseSection $rows
    Assert-Counters $html 4 1 1 2
    Assert-DatabaseMail ($html -match '<td>DB-B</td><td>EX-EXCLUDED</td><td[^>]*>Not mounted</td>') 'Boolean false or excluded server association lost.'
    Assert-DatabaseMail ($html -match '<td>10 GB \(10,737,418,240 bytes\)</td><td>0 B \(0 bytes\)</td>') 'Raw Exchange sizes or measured zero lost.'
    Assert-DatabaseMail ($html -match 'DatabaseSize: missing\.' -and $html -match 'Mounted: unrecognized value\.') 'Missing values and invalid boolean not disclosed.'
    Assert-DatabaseMail ($html -notmatch '<td>DAG-A</td>|<td>DAG-B</td>') 'DAGs counted as databases.'
    foreach ($notice in @('not free disk space', 'detailed DAG copy states and per-copy servers are not collected', 'Mailboxes per database: N/A', 'Zero observed mailboxes would not prove an empty database', 'SYNTHETIC-OLD-RUN', '2020-01-02 03:04:05', 'Exchange_OnPrem_MigrationReadiness_Config.csv')) {
        Assert-DatabaseMail ($html.Contains($notice)) "Missing provenance or limitation: $notice"
    }

    $csvPath = Join-Path $testRoot 'Exchange_OnPrem_MigrationReadiness_Config.csv'
    $rows | Export-Csv -LiteralPath $csvPath -Delimiter ';' -NoTypeInformation -Encoding UTF8
    $roundTrip = @(Import-Csv -LiteralPath $csvPath -Delimiter ';')
    Assert-DatabaseMail ((Get-DatabaseSection $roundTrip) -eq $html) 'Boolean values changed after CSV serialization.'
    Assert-Counters (Get-DatabaseSection @(New-ReadinessRow 'DB-ONLY' 'Mounted' 'fAlSe')) 1 0 1 0
    Assert-Counters (Get-DatabaseSection @(New-ReadinessRow 'DB-ONLY' 'Server' 'EX-01')) 1 0 0 1

    foreach ($empty in @(@(), @($null), @(New-ReadinessRow 'DAG' 'Servers' 'EX-01' -Category 'DatabaseAvailabilityGroup'))) {
        $absent = Get-DatabaseSection $empty
        Assert-Counters $absent 'N/A' 'N/A' 'N/A' 'N/A'
        Assert-DatabaseMail ($absent.Contains('this does not prove that there are no databases')) 'Absent evidence represented as zero.'
    }
    $failed = @(New-ReadinessRow 'MailboxDatabase' 'Collection' 'Failed' -Status 'ERROR' -Category 'CollectorError')
    $failedHtml = Get-DatabaseSection $failed
    Assert-Counters $failedHtml 'N/A' 'N/A' 'N/A' 'N/A'
    Assert-DatabaseMail ($failedHtml.Contains('coverage is incomplete')) 'Failed collection not disclosed.'
    $partial = Get-DatabaseSection ($rows + $failed)
    Assert-Counters $partial 4 1 1 2
    Assert-DatabaseMail ($partial.Contains('observed counts are not estate totals')) 'Partial counts presented as estate totals.'

    $duplicates = Get-DatabaseSection @(
        New-ReadinessRow 'DB-A' 'Mounted' 'True'
        New-ReadinessRow 'db-a' 'Mounted' 'False'
        New-ReadinessRow 'DB-A' 'Server' 'EX-01'
        New-ReadinessRow 'DB-A' 'Server' 'EX-02'
    )
    Assert-Counters $duplicates 1 0 0 1
    Assert-DatabaseMail ($duplicates.Contains('Mounted: ambiguous duplicate records.') -and $duplicates.Contains('Server: ambiguous duplicate records.')) 'Ambiguous records silently picked.'
    $invalid = Get-DatabaseSection @(
        New-ReadinessRow 'DB-A' 'Mounted' 'True' -Status 'ERROR'
        New-ReadinessRow 'DB-A' 'DatabaseSize' '20 GB' -Status 'WARNING'
        New-ReadinessRow '' 'Mounted' 'True'
    )
    Assert-Counters $invalid 1 0 0 1
    Assert-DatabaseMail ($invalid.Contains('1 record(s) have no database name') -and $invalid.Contains('DatabaseSize: collection not OK.')) 'Invalid data not disclosed.'
    Assert-DatabaseMail (-not $invalid.Contains('<td>20 GB</td>')) 'Failed measurement rendered as valid.'
    $escaped = Get-DatabaseSection @(New-ReadinessRow '<script>DB</script>' 'Server' 'EX<&>')
    Assert-DatabaseMail ($escaped.Contains('&lt;script&gt;DB&lt;/script&gt;') -and $escaped.Contains('EX&lt;&amp;&gt;') -and -not $escaped.Contains('<script>')) 'Database names/server HTML not escaped.'

    # Exercise the actual regeneration and mail body paths on a historical fixture.
    $summary | Export-Csv -LiteralPath (Join-Path $testRoot 'Exchange_OnPrem_Servers_Inventory_Summary.csv') -Delimiter ';' -NoTypeInformation -Encoding UTF8
    $server | Export-Csv -LiteralPath (Join-Path $testRoot 'Exchange_OnPrem_Infrastructure_PerServerSummary.csv') -Delimiter ';' -NoTypeInformation -Encoding UTF8
    & $module { param($root) Invoke-ServersAndStorageHtmlReportRegeneration -RunFolder $root -SendMail } $testRoot
    $report = Get-Content -LiteralPath (Join-Path $testRoot 'Exchange_OnPrem_InfrastructureAndReadiness_Report.html') -Raw
    $mail = & $module { $script:CapturedMail }
    Assert-DatabaseMail ($report.Contains($html)) 'Database summary missing from the regenerated full report.'
    Assert-DatabaseMail ($mail.Body.Contains($html) -and $mail.Attachments.Count -eq 3) 'Email did not retain database section and CSV attachments.'
    Assert-DatabaseMail ($report.Contains('Excluded</td>') -and $report.Contains('Per-server infrastructure summary') -and $report.Contains('Exchange readiness configuration')) 'Existing report sections regressed.'
    Assert-Counters $report 4 1 1 2
    Write-Host "PASS: $script:Checks synthetic Exchange database mail checks (PowerShell $($PSVersionTable.PSVersion)). No collector or real mail executed."
}
finally {
    Remove-Module $module -Force
    $resolved = [IO.Path]::GetFullPath($testRoot)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'SmartM365-DatabaseMail-*') { throw 'Unsafe test cleanup path.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
