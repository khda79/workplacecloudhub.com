<#
.SYNOPSIS
    Approval-gated item-scoped batches for source 401 cases after a reviewed pilot.
.VERSION
    1.0.4
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ProjectRoot,
    [Parameter(Mandatory)][string]$AnalysisDirectory,
    [Parameter(Mandatory)][string]$WitnessDirectory,
    [Parameter(Mandatory)][string]$PilotDirectory,
    [Parameter(Mandatory)][string]$QualificationDirectory,
    [Parameter(Mandatory)][string]$PathCorrectionDirectory,
    [Parameter(Mandatory)][string]$SessionId,
    [ValidateRange(1,100)][int]$BatchSize = 50,
    [ValidateRange(0,100)][int]$MaxErrorsPerBatch = 0,
    [ValidateRange(0,1440)][int]$MaintenanceMarginMinutes = 60,
    [ValidateRange(0,1000000)][int]$ExpectedOriginalItemCount = 0,
    [ValidateRange(0,1000000)][int]$ExpectedRemainingItemCount = 0,
    [ValidateNotNullOrEmpty()][string]$FarmTimeZoneId = 'W. Europe Standard Time',
    [ValidatePattern('^[0-9A-Fa-f]{64}$')][string]$ExpectedAnalysisHash = '',
    [ValidatePattern('^[0-9A-Fa-f]{64}$')][string]$ExpectedWitnessHash = '',
    [ValidatePattern('^[0-9A-Fa-f]{64}$')][string]$ExpectedPilotManifestHash = '',
    [ValidatePattern('^[0-9A-Fa-f]{64}$')][string]$ExpectedQualificationHash = '',
    [ValidatePattern('^[0-9A-Fa-f]{64}$')][string]$ExpectedPathCorrectionHash = '',
    [ValidatePattern('^[0-9A-Fa-f]{64}$')][string]$ExpectedPlanHash = '',
    [switch]$DryRun,
    [switch]$Run,
    [switch]$ConfirmBatch
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$version = '1.0.4'
. (Join-Path $PSScriptRoot '..\Launchers\SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
$script:ConsoleLifecycleContext = Start-SmartM365MigrationConsoleLifecycle -ScriptPath $PSCommandPath
$script:ConsoleLifecycleFailure = $null
$script:ConsoleLifecycleStatus = 'SUCCESS'
try {
if ($DryRun -and $Run) { throw 'Choose either -DryRun or -Run.' }
if ($Run) { throw 'Real item-scoped batch copy remains disabled pending review of the placement-proof DryRun and separate approval.' }
if ($ConfirmBatch -and -not $Run) { throw '-ConfirmBatch applies only with -Run.' }
if ($Run -and -not $ConfirmBatch) { throw 'Real copies require both -Run and -ConfirmBatch.' }
if ($Run -and (-not $ExpectedAnalysisHash -or -not $ExpectedWitnessHash -or -not $ExpectedPilotManifestHash -or
    -not $ExpectedQualificationHash -or -not $ExpectedPathCorrectionHash -or
    -not $ExpectedPlanHash -or -not $ExpectedOriginalItemCount -or -not $ExpectedRemainingItemCount)) {
    throw 'Real copies require reviewed analysis, witness, pilot, qualification, correction, and plan SHA256 values plus both item counts.'
}
if ($PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5) { throw 'Use Windows PowerShell 5.1 for ShareGate.' }
. (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-FarmMaintenance.ps1')
. (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-TransientEvidence.ps1')
. (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-DestinationPath.ps1')
. (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-ShareGateReportReader.ps1')
. (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-PlacementEvidence.ps1')
$reportAliases = Get-SmartM365ShareGateReportAliases -ConfigRoot (Join-Path $PSScriptRoot '..\..\Config')
$farmZone = [TimeZoneInfo]::FindSystemTimeZoneById($FarmTimeZoneId)
$null = Assert-SmartM365OutsideFarmMaintenance -FarmTimeZone $farmZone -Phase 'Transient batch preparation'
$evidence = Get-SmartM365TransientEvidence -ProjectRoot $ProjectRoot -AnalysisDirectory $AnalysisDirectory -PilotDirectory $PilotDirectory -SessionId $SessionId
$witness = Get-SmartM365DiagnosticsChild -Parent $evidence.DiagnosticsRoot -Path $WitnessDirectory
$witnessCsv = Join-Path $witness 'Witness-Results.csv'
if (-not (Test-Path -LiteralPath $witnessCsv -PathType Leaf)) { throw 'Witness-Results.csv is missing.' }
$witnessHash = (Get-FileHash -LiteralPath $witnessCsv -Algorithm SHA256).Hash
$pilotLog = Get-Content -LiteralPath $evidence.PilotLogPath -Raw -ErrorAction Stop
if ($pilotLog -notmatch ('(?i)WitnessSHA256=' + [regex]::Escape($witnessHash) + '(?:;|\s)')) { throw 'Pilot log does not match the selected witness.' }
if ($ExpectedAnalysisHash -and $ExpectedAnalysisHash.ToUpperInvariant() -ne $evidence.AnalysisSHA256) { throw 'Analysis hash differs from the reviewed value.' }
if ($ExpectedWitnessHash -and $ExpectedWitnessHash.ToUpperInvariant() -ne $witnessHash) { throw 'Witness hash differs from the reviewed value.' }
if ($ExpectedPilotManifestHash -and $ExpectedPilotManifestHash.ToUpperInvariant() -ne $evidence.PilotManifestSHA256) { throw 'Pilot manifest hash differs from the reviewed value.' }
if ($ExpectedOriginalItemCount -and $ExpectedOriginalItemCount -ne @($evidence.AccessItems).Count) { throw 'Original access item count differs from the reviewed value.' }

$copied = @($evidence.PilotItems | Where-Object { $_.Result -eq 'Success' -and $_.ImportStatus -match '(?i)^Finished$' })
$skipped = @($evidence.PilotItems | Where-Object { $_.Result -eq 'Skipped' -and $_.SourceItemId -eq 1 -and $_.SourcePath -eq 'Home.aspx' })
if ($copied.Count -ne 4 -or $skipped.Count -ne 1 -or @($evidence.PilotItems | Where-Object { $_.Result -notin @('Success','Skipped') }).Count) {
    throw 'Pilot evidence must show four imported Success items and one separately handled Home.aspx Skipped item.'
}
$placement = Get-SmartM365ShareGatePlacementEvidence -DiagnosticsRoot $evidence.DiagnosticsRoot `
    -QualificationDirectory $QualificationDirectory -PathCorrectionDirectory $PathCorrectionDirectory `
    -SessionId $SessionId -AnalysisSHA256 $evidence.AnalysisSHA256 `
    -AccessItems $evidence.AccessItems -PilotSuccessItems $copied -Aliases $reportAliases
if ($ExpectedQualificationHash -and $ExpectedQualificationHash.ToUpperInvariant() -ne $placement.QualificationSHA256) {
    throw 'Qualification evidence hash differs from the reviewed value.'
}
if ($ExpectedPathCorrectionHash -and $ExpectedPathCorrectionHash.ToUpperInvariant() -ne $placement.PathCorrectionSHA256) {
    throw 'Path correction evidence hash differs from the reviewed value.'
}
$excluded = @{}
foreach ($item in $evidence.PilotItems) { $excluded[$item.ItemKey] = $true }
$otherIdOne = @($evidence.AccessItems | Where-Object { $_.SourceItemId -eq 1 -and -not $excluded.ContainsKey($_.ItemKey) })
$separatePages = @($otherIdOne | Where-Object {
    $_.SourceList -match '(?i)^(Pages du site|Site Pages|SitePages)$' -and
    ($_.ItemName -match '(?i)^Home\.aspx$' -or $_.'Raw: Source path' -match '(?i)(^|/)Home\.aspx$')
})
foreach ($page in $separatePages) { $excluded[$page.ItemKey] = $true }
$remaining = @($evidence.AccessItems | Where-Object { -not $excluded.ContainsKey($_.ItemKey) })
foreach ($item in $remaining) {
    $item | Add-Member -NotePropertyName DestinationRoute -NotePropertyValue (Resolve-SmartM365ShareGateDestinationPath -Row $item)
}
if ($remaining.Count -ne @($evidence.AccessItems).Count - 5 - $separatePages.Count) {
    throw 'The pilot and separately handled page item keys were not uniquely excluded.'
}
if ($ExpectedRemainingItemCount -and $ExpectedRemainingItemCount -ne $remaining.Count) { throw 'Remaining item count differs from the reviewed value.' }
$outOfBatch = @($evidence.OutOfBatchRows | Sort-Object RowId)

function Get-TransientEndpointKey {
    param($Row)
    return ($Row.SourceUrl.TrimEnd('/').ToLowerInvariant() + '|' + $Row.SourceListId.ToLowerInvariant() + '|' +
        $Row.SourceList.ToLowerInvariant() + '|' + $Row.DestinationUrl.TrimEnd('/').ToLowerInvariant() + '|' +
        $Row.DestinationList.ToLowerInvariant() + '|' + $Row.DestinationRoute.DestinationFolder.ToLowerInvariant())
}
$groups = @($remaining | Group-Object { Get-TransientEndpointKey -Row $_ } | Sort-Object Name)
$batches = [System.Collections.Generic.List[object]]::new()
$planRows = [System.Collections.Generic.List[object]]::new()
foreach ($group in $groups) {
    $items = @($group.Group | Sort-Object @{Expression={ [int]$_.SourceItemId }}, ItemKey)
    $duplicateIds = @($items | Group-Object SourceItemId | Where-Object Count -GT 1)
    if ($duplicateIds.Count) { throw "Duplicate source IDs occur in one source list: $($group.Name)" }
    for ($start = 0; $start -lt $items.Count; $start += $BatchSize) {
        $end = [math]::Min($start + $BatchSize - 1, $items.Count - 1)
        $slice = @($items[$start..$end])
        $number = $batches.Count + 1
        $first = $slice[0]
        $ids = @($slice | ForEach-Object { [int]$_.SourceItemId })
        $batch = [pscustomobject]@{
            Number=$number; SourceUrl=$first.SourceUrl; SourceList=$first.SourceList;
            DestinationUrl=$first.DestinationUrl; DestinationList=$first.DestinationList;
            DestinationFolder=$first.DestinationRoute.DestinationFolder;
            Items=$slice; Ids=$ids
        }
        $batches.Add($batch)
        foreach ($item in $slice) {
            $planRows.Add([pscustomobject]@{
                BatchNumber=$number; ItemKey=$item.ItemKey; SourceUrl=$item.SourceUrl; SourceList=$item.SourceList;
                SourceItemId=[int]$item.SourceItemId; DestinationUrl=$item.DestinationUrl; DestinationList=$item.DestinationList;
                SourceFilePath=$item.DestinationRoute.SourceFilePath;
                DestinationFilePath=$item.DestinationRoute.DestinationFilePath;
                DestinationFolder=$item.DestinationRoute.DestinationFolder
            })
        }
    }
}
if ($planRows.Count -ne $remaining.Count -or -not $batches.Count) { throw 'The batch plan is incomplete.' }
$canonicalRows = [System.Collections.Generic.List[string]]::new()
$canonicalRows.Add('schema=4;batchSize=' + $BatchSize + ';qualification=' + $placement.QualificationSHA256 +
    ';correction=' + $placement.PathCorrectionSHA256)
foreach ($row in $planRows) {
    $canonicalRows.Add('COPY' + "`t" + (@($row.BatchNumber,$row.ItemKey,$row.SourceUrl,$row.SourceList,$row.SourceItemId,$row.DestinationUrl,$row.DestinationList,$row.SourceFilePath,$row.DestinationFilePath,$row.DestinationFolder) -join "`t"))
}
foreach ($page in @($separatePages | Sort-Object ItemKey)) {
    $canonicalRows.Add('SEPARATE-PAGE' + "`t" + (@($page.ItemKey,$page.ItemName,$page.ObjectType,$page.SourceUrl,$page.'Raw: Source path',$page.DestinationUrl) -join "`t"))
}
foreach ($row in $outOfBatch) {
    $canonicalRows.Add('NO-ID' + "`t" + (@($row.RowId,$row.ObjectType,$row.ItemName,$row.SourceUrl,$row.SourcePath,$row.DestinationUrl) -join "`t"))
}
$canonical = ($canonicalRows -join "`n") + "`n"
$hasher = [Security.Cryptography.SHA256]::Create()
try { $planHash = ([BitConverter]::ToString($hasher.ComputeHash([Text.UTF8Encoding]::new($false).GetBytes($canonical)))).Replace('-', '') }
finally { $hasher.Dispose() }
if ($ExpectedPlanHash -and $ExpectedPlanHash.ToUpperInvariant() -ne $planHash) { throw 'Batch plan hash differs from the reviewed value.' }

$pilotStarts = @{}
$pilotDurations = [System.Collections.Generic.List[double]]::new()
foreach ($line in (Get-Content -LiteralPath $evidence.PilotLogPath -ErrorAction Stop)) {
    if ($line -match '^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}) Starting real item copy (\d+)/5:') {
        $pilotStarts[[int]$Matches[2]] = [datetime]::ParseExact($Matches[1], 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
    }
    elseif ($line -match '^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}) Finished item (\d+)/5:') {
        $index = [int]$Matches[2]
        if ($index -ge 1 -and $index -le 5 -and $evidence.PilotItems[$index - 1].Result -eq 'Success' -and $pilotStarts.ContainsKey($index)) {
            $finished = [datetime]::ParseExact($Matches[1], 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
            $seconds = ($finished - $pilotStarts[$index]).TotalSeconds
            if ($seconds -gt 0) { $pilotDurations.Add($seconds) }
        }
    }
}
if ($pilotDurations.Count -ne 4) { throw 'Cannot estimate duration: four completed pilot Success timings are required.' }
$correctionStarts = @{}
$correctionDurations = [System.Collections.Generic.List[double]]::new()
$correctionLogPath = Join-Path $placement.PathCorrectionDirectory 'PathCorrection.log'
foreach ($line in (Get-Content -LiteralPath $correctionLogPath -ErrorAction Stop)) {
    if ($line -match '^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}) Starting item \d/3; ID=(\d+);') {
        $correctionStarts[$Matches[2]] = [datetime]::ParseExact($Matches[1], 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
    }
    elseif ($line -match '^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}) Completed ID=(\d+); Status=Success;') {
        $id = $Matches[2]
        if ($correctionStarts.ContainsKey($id)) {
            $finished = [datetime]::ParseExact($Matches[1], 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
            $seconds = ($finished - $correctionStarts[$id]).TotalSeconds
            if ($seconds -gt 0) { $correctionDurations.Add($seconds) }
        }
    }
}
if ($correctionDurations.Count -ne 3) { throw 'Cannot estimate duration: three completed path-correction item timings are required.' }
$secondsPerItem = [math]::Max(($pilotDurations | Measure-Object -Maximum).Maximum,
    ($correctionDurations | Measure-Object -Maximum).Maximum)
$estimatedSeconds = [math]::Ceiling($remaining.Count * $secondsPerItem)

function Get-TransientScheduleAssessment {
    param([int]$ItemCount, [datetime]$UtcNow = [datetime]::UtcNow)
    $utc = if ($UtcNow.Kind -eq [DateTimeKind]::Utc) { $UtcNow } else { $UtcNow.ToUniversalTime() }
    $farmNow = [TimeZoneInfo]::ConvertTimeFromUtc($utc, $farmZone)
    $nextWindowFarm = $farmNow.Date.AddHours(23).AddMinutes(45)
    if ($nextWindowFarm -le $farmNow) { $nextWindowFarm = $nextWindowFarm.AddDays(1) }
    $nextWindowUtc = [TimeZoneInfo]::ConvertTimeToUtc($nextWindowFarm, $farmZone)
    $projectedEndUtc = $utc.AddSeconds([math]::Ceiling($ItemCount * $secondsPerItem))
    $projectedEndWithMarginUtc = $projectedEndUtc.AddMinutes($MaintenanceMarginMinutes)
    return [pscustomobject]@{
        FarmNow=$farmNow; NextWindowFarm=$nextWindowFarm;
        ProjectedEndFarm=[TimeZoneInfo]::ConvertTimeFromUtc($projectedEndUtc, $farmZone);
        EndWithMarginFarm=[TimeZoneInfo]::ConvertTimeFromUtc($projectedEndWithMarginUtc, $farmZone);
        IsSafe=($projectedEndWithMarginUtc -lt $nextWindowUtc)
    }
}

function Assert-TransientSchedule {
    param([int]$ItemCount, [string]$Phase)
    $schedule = Get-TransientScheduleAssessment -ItemCount $ItemCount
    if (-not $schedule.IsSafe) {
        throw ('Refusing {0}: estimated completion {1:yyyy-MM-dd HH:mm:ss} plus {2} min margin reaches maintenance start {3:yyyy-MM-dd HH:mm:ss} ({4}).' -f
            $Phase,$schedule.ProjectedEndFarm,$MaintenanceMarginMinutes,$schedule.NextWindowFarm,$farmZone.Id)
    }
    return $schedule
}

function Write-TransientLog {
    param([string]$Message)
    $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    if ($script:logPath) { Add-Content -LiteralPath $script:logPath -Value $line -Encoding UTF8 }
    Write-Output $line
}
$script:logPath = ''
Write-TransientLog ('Mode={0}; script=v{1}; session={2}; source 401 items={3}; pilot Success excluded=4; pilot Home.aspx Skipped excluded=1; other Home.aspx separate={4}; remaining={5}; batches={6}; batchSize={7}; maxErrorsPerBatch={8}; copySettings=IncrementalUpdate; maintenance=23:45-00:15 farm time; writes={9}' -f
    $(if ($Run) { 'Run' } else { 'DryRun' }),$version,$SessionId,@($evidence.AccessItems).Count,$separatePages.Count,$remaining.Count,$batches.Count,$BatchSize,$MaxErrorsPerBatch,$(if ($Run) { 'possible after confirmation' } else { 'none' }))
Write-TransientLog ('AnalysisSHA256={0}; WitnessSHA256={1}; PilotManifestSHA256={2}; PlanSHA256={3}' -f
    $evidence.AnalysisSHA256,$witnessHash,$evidence.PilotManifestSHA256,$planHash)
Write-TransientLog ('Placement proof: qualified=1; corrected=3; ShareGate={0}; QualificationSHA256={1}; PathCorrectionSHA256={2}; correction folder={3}' -f
    $placement.ShareGateVersion,$placement.QualificationSHA256,$placement.PathCorrectionSHA256,$placement.PathCorrectionDirectory)
foreach ($item in $otherIdOne) {
    $disposition = if ($excluded.ContainsKey($item.ItemKey)) { 'Hors lot - page a examiner a part' } else { 'Included in item batch' }
    Write-TransientLog ('Other Source ID 1: title={0}; type={1}; source site URL={2}; source path={3}; target site URL={4}; list={5}; disposition={6}' -f
        $item.ItemName,$item.ObjectType,$item.SourceUrl,$item.'Raw: Source path',$item.DestinationUrl,$item.SourceList,$disposition)
}
Write-TransientLog ('Hors lot - à traiter à part: {0} lines without Source ID (Site={1}; File={2}).' -f
    $outOfBatch.Count,@($outOfBatch | Where-Object ObjectType -EQ 'Site').Count,@($outOfBatch | Where-Object ObjectType -EQ 'File').Count)
foreach ($row in $outOfBatch) {
    Write-TransientLog ('Hors lot - à traiter à part: RowId={0}; type={1}; title={2}; source site URL={3}; source path={4}' -f
        $row.RowId,$row.ObjectType,$row.ItemName,$row.SourceUrl,$row.SourcePath)
}
$schedule = Get-TransientScheduleAssessment -ItemCount $remaining.Count
Write-TransientLog ('Observed item seconds: original pilot={0}; path correction={1}; conservative estimate={2:N2} s/item; estimated batch duration={3}; projected farm end={4:yyyy-MM-dd HH:mm:ss}; with margin {5} min={6:yyyy-MM-dd HH:mm:ss}; next maintenance={7:yyyy-MM-dd HH:mm:ss}; start allowed={8}' -f
    ($pilotDurations -join ','),($correctionDurations -join ','),$secondsPerItem,([timespan]::FromSeconds($estimatedSeconds)),
    $schedule.ProjectedEndFarm,$MaintenanceMarginMinutes,$schedule.EndWithMarginFarm,$schedule.NextWindowFarm,$schedule.IsSafe)
foreach ($batch in $batches) {
    Write-TransientLog ('Batch {0}/{1}: {2} | {3} -> {4} | {5}; destinationFolder={6}; items={7}; SourceItemIds={8}' -f
        $batch.Number,$batches.Count,$batch.SourceUrl,$batch.SourceList,$batch.DestinationUrl,$batch.DestinationList,$batch.DestinationFolder,
        $batch.Ids.Count,($batch.Ids -join ','))
}
$phrase = 'COPY ' + $remaining.Count + ' ITEMS ' + $SessionId + ' v' + $version
Write-TransientLog ('Confirmation phrase for this script version: ' + $phrase)
if (-not $Run) { return }
$null = Assert-TransientSchedule -ItemCount $remaining.Count -Phase 'transient batch run before confirmation'
$entered = Read-Host ('Type exactly "{0}" to authorize the listed item-scoped batches' -f $phrase)
if ($entered -cne $phrase) { throw 'Exact confirmation phrase was not entered. No ShareGate copy was started.' }
$null = Assert-SmartM365OutsideFarmMaintenance -FarmTimeZone $farmZone -Phase 'Transient real copy startup'
$null = Assert-TransientSchedule -ItemCount $remaining.Count -Phase 'transient real copy startup'

$module = @(Get-Module -ListAvailable -Name ShareGate | Sort-Object Version -Descending | Select-Object -First 1)
if (-not $module.Count) { throw 'ShareGate module is not discoverable in Windows PowerShell 5.1.' }
Import-Module -Name $module[0].Path -ErrorAction Stop
if ([string]$module[0].Version -ne $placement.ShareGateVersion) {
    throw 'Installed ShareGate version differs from the qualified path-correction version.'
}
foreach ($name in @('Connect-Site','Get-List','Get-Folder','Copy-Content','New-CopySettings','Export-Report')) {
    if (-not (Get-Command -Name $name -Module ShareGate -ErrorAction SilentlyContinue)) { throw "Required ShareGate cmdlet is missing: $name" }
}
$copyCommand = Get-Command -Name Copy-Content -Module ShareGate
$requiredParameters = @('SourceList','DestinationList','SourceItemId','DestinationFolder','CopySettings','TaskName')
if (-not @($copyCommand.ParameterSets | Where-Object {
    $names = @($_.Parameters | ForEach-Object Name)
    @($requiredParameters | Where-Object { $_ -notin $names }).Count -eq 0
}).Count) { throw 'Installed Copy-Content lacks the item-scoped parameter set.' }
$settings = ShareGate\New-CopySettings -OnContentItemExists IncrementalUpdate -ErrorAction Stop
if (-not $settings) { throw 'New-CopySettings returned no settings.' }

$lockPath = Join-Path $evidence.DiagnosticsRoot ('Transient-' + $SessionId + '.lock')
try { $runLock = [IO.File]::Open($lockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
catch { throw "Another transient batch run may be active for session $SessionId. Exclusive lock: $lockPath. $($_.Exception.Message)" }
$runLock.SetLength(0)
$lockText = '{0} Actor={1}\{2}; Machine={3}; Script=v{4}' -f (Get-Date -Format 'o'),$env:USERDOMAIN,$env:USERNAME,$env:COMPUTERNAME,$version
$lockBytes = [Text.UTF8Encoding]::new($false).GetBytes($lockText)
$runLock.Write($lockBytes, 0, $lockBytes.Length)
$runLock.Flush()

$runId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'),[guid]::NewGuid().ToString('N')
$output = Join-Path $evidence.DiagnosticsRoot ('Transient-' + $runId)
$reportDir = Join-Path $output 'Reports'
$objectDir = Join-Path $output 'CopyResults'
New-Item -ItemType Directory -Path $reportDir,$objectDir -Force | Out-Null
$script:logPath = Join-Path $output 'Transient.log'
$planPath = Join-Path $output 'Transient-Plan.csv'
$resultPath = Join-Path $output 'Transient-Results.csv'
$outOfBatchPath = Join-Path $output 'Transient-HorsLot.csv'
$summaryPath = Join-Path $output 'Transient-Summary.json.txt'
$planColumns = @('BatchNumber','ItemKey','SourceUrl','SourceList','SourceItemId','DestinationUrl','DestinationList','SourceFilePath','DestinationFilePath','DestinationFolder')
$resultColumns = @('BatchNumber','ItemKey','SourceUrl','SourceList','SourceItemId','DestinationUrl','DestinationList','DestinationFilePath','DestinationFolder','ExportedDestinationPath','Result','CopySessionId','ReportPath','Error')
$outOfBatchColumns = @('Disposition','Reason','RowId','Timestamp','ObjectType','ItemName','SourceUrl','SourceList','SourcePath','DestinationUrl','DestinationList','DestinationPath')
Export-SmartM365TransientCsv -Path $planPath -Rows $planRows.ToArray() -Columns $planColumns
Export-SmartM365TransientCsv -Path $outOfBatchPath -Rows $outOfBatch -Columns $outOfBatchColumns
$results = [System.Collections.Generic.List[object]]::new()
$byKey = @{}
foreach ($item in $planRows) {
    $result = [pscustomobject]@{
        BatchNumber=$item.BatchNumber; ItemKey=$item.ItemKey; SourceUrl=$item.SourceUrl; SourceList=$item.SourceList;
        SourceItemId=$item.SourceItemId; DestinationUrl=$item.DestinationUrl; DestinationList=$item.DestinationList;
        DestinationFilePath=$item.DestinationFilePath; DestinationFolder=$item.DestinationFolder;
        ExportedDestinationPath=''; Result='NotAttempted'; CopySessionId=''; ReportPath=''; Error=''
    }
    $results.Add($result)
    $byKey[$item.ItemKey] = $result
}
$siteCache = @{}
$listCache = @{}
$completedBatches = 0
$runStatus = 'Running'

function Save-TransientState {
    Export-SmartM365TransientCsv -Path $resultPath -Rows $results.ToArray() -Columns $resultColumns
    $summary = [ordered]@{
        SchemaVersion=1; ScriptVersion=$version; SessionId=$SessionId; RunStatus=$runStatus;
        GeneratedAtUtc=[datetime]::UtcNow.ToString('o'); Actor=($env:USERDOMAIN + '\' + $env:USERNAME); Machine=$env:COMPUTERNAME;
        PlannedItems=$results.Count; BatchCount=$batches.Count; CompletedBatches=$completedBatches;
        Success=@($results | Where-Object Result -EQ 'Success').Count;
        Skipped=@($results | Where-Object Result -EQ 'Skipped').Count;
        Error=@($results | Where-Object Result -EQ 'Error').Count;
        Warning=@($results | Where-Object Result -EQ 'Warning').Count;
        Mixed=@($results | Where-Object Result -EQ 'Mixed').Count;
        Unreported=@($results | Where-Object Result -EQ 'Unreported').Count;
        NotAttempted=@($results | Where-Object Result -EQ 'NotAttempted').Count;
        AnalysisSHA256=$evidence.AnalysisSHA256; WitnessSHA256=$witnessHash;
        PilotManifestSHA256=$evidence.PilotManifestSHA256;
        QualificationSHA256=$placement.QualificationSHA256; PathCorrectionSHA256=$placement.PathCorrectionSHA256;
        PlacementShareGateVersion=$placement.ShareGateVersion; PlanSHA256=$planHash;
        PilotSuccessSecondsPerItem=$secondsPerItem; EstimatedDurationSeconds=$estimatedSeconds;
        MaintenanceMarginMinutes=$MaintenanceMarginMinutes;
        SeparatePageItems=$separatePages.Count;
        OutOfBatchLabel='Hors lot - à traiter à part'; OutOfBatchLines=$outOfBatch.Count;
        OutOfBatchSiteLines=@($outOfBatch | Where-Object ObjectType -EQ 'Site').Count;
        OutOfBatchFileLines=@($outOfBatch | Where-Object ObjectType -EQ 'File').Count;
        OutputDirectory=$output; PlanPath=$planPath; ResultsPath=$resultPath;
        OutOfBatchPath=$outOfBatchPath; LogPath=$script:logPath
    }
    Export-SmartM365TransientJson -Path $summaryPath -Value $summary
}

function Get-ExactTransientList {
    param([ValidateSet('Source','Destination')][string]$Side,[string]$SiteUrl,[string]$ListName)
    $siteKey = $Side + '|' + $SiteUrl.TrimEnd('/').ToLowerInvariant()
    if (-not $siteCache.ContainsKey($siteKey)) {
        $null = Assert-SmartM365OutsideFarmMaintenance -FarmTimeZone $farmZone -Phase ("$Side site connection")
        $site = if ($Side -eq 'Source') { ShareGate\Connect-Site -Url $SiteUrl -ErrorAction Stop }
                else { ShareGate\Connect-Site -Url $SiteUrl -Browser -ErrorAction Stop }
        if (-not $site) { throw "Connect-Site returned no $Side site: $SiteUrl" }
        $siteCache[$siteKey] = $site
    }
    $listKey = $siteKey + '|' + $ListName.ToLowerInvariant()
    if (-not $listCache.ContainsKey($listKey)) {
        $found = @(ShareGate\Get-List -Site $siteCache[$siteKey] -Name $ListName -ErrorAction Stop)
        $exact = @($found | Where-Object {
            ($_.PSObject.Properties['Title'] -and $_.Title -eq $ListName) -or
            ($_.PSObject.Properties['Name'] -and $_.Name -eq $ListName)
        })
        if ($exact.Count -ne 1) { throw "Get-List returned $($exact.Count) exact matches for $Side '$ListName'." }
        $listCache[$listKey] = $exact[0]
    }
    return $listCache[$listKey]
}

function Get-TransientCopySessionId {
    param($CopyResult)
    foreach ($name in @('SessionId','SessionID','CopySessionId','Id')) {
        $property = $CopyResult.PSObject.Properties[$name]
        if ($property -and [string]$property.Value -match '^\d{6}-\d+$') { return [string]$property.Value }
    }
    return ''
}

Save-TransientState
Write-TransientLog ('Actor={0}\{1}; Machine={2}; ShareGate={3}; SourceAuth=CurrentWindowsIdentity; DestinationAuth=Browser; CopySettings=IncrementalUpdate; plan={4}; results={5}; horsLot={6}' -f
    $env:USERDOMAIN,$env:USERNAME,$env:COMPUTERNAME,$module[0].Version,$planPath,$resultPath,$outOfBatchPath)
$currentBatch = $null
$currentCopyStarted = $false
$currentCopySession = ''
try {
    foreach ($batch in $batches) {
        $currentBatch = $batch
        $currentCopyStarted = $false
        $currentCopySession = ''
        $null = Assert-TransientSchedule -ItemCount @($results | Where-Object Result -EQ 'NotAttempted').Count -Phase ("batch $($batch.Number)/$($batches.Count) preparation")
        $null = Assert-SmartM365OutsideFarmMaintenance -FarmTimeZone $farmZone -Phase ("batch $($batch.Number)/$($batches.Count) preparation")
        if ((Get-SmartM365PlacementEvidenceHash -Paths $placement.QualificationFiles) -ne $placement.QualificationSHA256 -or
            (Get-SmartM365PlacementEvidenceHash -Paths $placement.CorrectionFiles) -ne $placement.PathCorrectionSHA256 -or
            (Get-SmartM365PilotManifest -PilotDirectory $evidence.PilotDirectory).SHA256 -ne $evidence.PilotManifestSHA256 -or
            (Get-FileHash -LiteralPath $evidence.ClassifiedPath -Algorithm SHA256).Hash -ne $evidence.AnalysisSHA256) {
            throw 'Reviewed placement, pilot, or analysis evidence changed before the next batch.'
        }
        $sourceList = Get-ExactTransientList -Side Source -SiteUrl $batch.SourceUrl -ListName $batch.SourceList
        $destinationList = Get-ExactTransientList -Side Destination -SiteUrl $batch.DestinationUrl -ListName $batch.DestinationList
        Assert-SmartM365ShareGateDestinationFolder -DestinationList $destinationList -DestinationFolder $batch.DestinationFolder
        $null = Assert-SmartM365OutsideFarmMaintenance -FarmTimeZone $farmZone -Phase ("batch $($batch.Number)/$($batches.Count) copy")
        $taskName = 'SmartM365 Transient 401 ' + $SessionId + ' ' + $runId + ' batch ' + ('{0:D3}' -f $batch.Number)
        $reportPath = Join-Path $reportDir ('Batch-{0:D3}.csv' -f $batch.Number)
        $temporaryReport = Join-Path $reportDir ('.Batch-{0:D3}-{1}.csv' -f $batch.Number,[guid]::NewGuid().ToString('N'))
        $objectPath = Join-Path $objectDir ('CopyResult-{0:D3}.txt' -f $batch.Number)
        Write-TransientLog ('Starting batch {0}/{1}; destinationFolder={2}; items={3}; task={4}' -f
            $batch.Number,$batches.Count,$batch.DestinationFolder,$batch.Ids.Count,$taskName)
        $currentCopyStarted = $true
        $copyParameters = @{ SourceList=$sourceList; DestinationList=$destinationList; SourceItemId=$batch.Ids;
            CopySettings=$settings; TaskName=$taskName; ErrorAction='Stop' }
        if ($batch.DestinationFolder) { $copyParameters.DestinationFolder = $batch.DestinationFolder }
        $copyResult = ShareGate\Copy-Content @copyParameters
        if (-not $copyResult -or @($copyResult).Count -ne 1) { throw 'Copy-Content did not return exactly one CopyResult.' }
        $copySession = Get-TransientCopySessionId -CopyResult $copyResult
        $currentCopySession = $copySession
        @('Type: ' + $copyResult.GetType().FullName,'Session ID: ' + $(if ($copySession) { $copySession } else { '(not exposed)' }),
          'Successes: ' + $copyResult.Successes,'Warnings: ' + $copyResult.Warnings,'Errors: ' + $copyResult.Errors) |
            Set-Content -LiteralPath $objectPath -Encoding UTF8
        ShareGate\Export-Report -CopyResult $copyResult -Path $temporaryReport -ErrorAction Stop | Out-Null
        if (-not (Test-Path -LiteralPath $temporaryReport -PathType Leaf)) { throw 'Export-Report did not create the batch CSV.' }
        Move-Item -LiteralPath $temporaryReport -Destination $reportPath -Force
        $reportRows = @(Import-Csv -LiteralPath $reportPath -Encoding UTF8)
        if (-not $reportRows.Count -or
            -not (Test-SmartM365ShareGateReportField -Row $reportRows[0] -Field 'SourceItemId' -Aliases $reportAliases) -or
            -not (Test-SmartM365ShareGateReportField -Row $reportRows[0] -Field 'DestinationPath' -Aliases $reportAliases)) {
            throw 'Batch export has no rows or lacks source-ID/destination-path columns.'
        }
        $importStates = @($reportRows | ForEach-Object { [string]$_.'Microsoft 365 Import: Status' } | Where-Object { $_ } | Sort-Object -Unique)
        $importFinished = ($importStates.Count -eq 1 -and $importStates[0] -eq 'Finished')
        $plannedIds = @{}
        foreach ($id in $batch.Ids) { $plannedIds[[string]$id] = $true }
        $unexpected = @($reportRows | Where-Object {
            $id = 0
            [int]::TryParse((Get-SmartM365ShareGateReportValue -Row $_ -Field 'SourceItemId' -Aliases $reportAliases),[ref]$id) -and
                $id -gt 0 -and -not $plannedIds.ContainsKey([string]$id)
        })
        foreach ($item in $batch.Items) {
            $result = $byKey[$item.ItemKey]
            $itemRows = @($reportRows | Where-Object {
                (Get-SmartM365ShareGateReportValue -Row $_ -Field 'SourceItemId' -Aliases $reportAliases) -eq [string]$item.SourceItemId
            })
            $statuses = @($itemRows | ForEach-Object {
                Get-SmartM365ShareGateReportValue -Row $_ -Field 'Status' -Aliases $reportAliases
            } | Where-Object { $_ } | Sort-Object -Unique)
            $paths = @($itemRows | ForEach-Object {
                Get-SmartM365ShareGateReportValue -Row $_ -Field 'DestinationPath' -Aliases $reportAliases
            } | Where-Object { $_ } | Sort-Object -Unique)
            $reportedProblems = @($itemRows | ForEach-Object {
                (Get-SmartM365ShareGateReportValue -Row $_ -Field 'Errors' -Aliases $reportAliases),
                (Get-SmartM365ShareGateReportValue -Row $_ -Field 'Warnings' -Aliases $reportAliases)
            } | Where-Object { $_ })
            $state = if (-not $itemRows.Count -or -not $statuses.Count) { 'Unreported' }
                     elseif ($statuses.Count -gt 1) { 'Mixed' }
                     elseif ($statuses[0] -match '(?i)^(success|skipped|error|warning)$') { $statuses[0] }
                     else { 'Unreported' }
            if ($state -eq 'Success' -and ($paths.Count -ne 1 -or
                (ConvertTo-SmartM365ShareGateRelativePath -Path $paths[0] -Side ExportedDestination) -ne $item.DestinationRoute.DestinationFilePath -or
                -not $importFinished)) {
                $state = 'Unreported'
                $result.Error = 'Success was not proven at the planned destination path with a finished Microsoft 365 import.'
            }
            elseif ($state -eq 'Success' -and $reportedProblems.Count) {
                $state = 'Error'
                $result.Error = $reportedProblems -join ' | '
            }
            $result.Result = $state
            $result.ExportedDestinationPath = $paths -join '; '
            $result.CopySessionId = $copySession
            $result.ReportPath = $reportPath
            if ($state -eq 'Unreported' -and -not $result.Error) { $result.Error = 'No recognized item result in ShareGate export.' }
            elseif ($state -eq 'Mixed') { $result.Error = 'Mixed version statuses; inspect ShareGate export.' }
            elseif ($state -eq 'Error' -and -not $result.Error) { $result.Error = $reportedProblems -join ' | ' }
            elseif ($state -eq 'Skipped') { $result.Error = 'Skipped item requires separate placement review.' }
        }
        $completedBatches++
        $itemErrors = @($batch.Items | Where-Object { $byKey[$_.ItemKey].Result -eq 'Error' }).Count
        $reportErrors = @($reportRows | Where-Object {
            (Get-SmartM365ShareGateReportValue -Row $_ -Field 'Status' -Aliases $reportAliases) -match '(?i)^(error|failed|failure)$'
        }).Count
        $copyErrors = if ($copyResult.PSObject.Properties['Errors']) { [int]$copyResult.Errors } else { 0 }
        $errorCount = [math]::Max($itemErrors,[math]::Max($reportErrors,$copyErrors))
        $unknownCount = @($batch.Items | Where-Object { $byKey[$_.ItemKey].Result -in @('Unreported','Mixed') }).Count
        if ($unexpected.Count) { $runStatus = 'StoppedUnexpectedSourceId' }
        elseif ($unknownCount) { $runStatus = 'StoppedUnreportedItems' }
        elseif ($errorCount -gt $MaxErrorsPerBatch) { $runStatus = 'StoppedErrorThreshold' }
        elseif (@($batch.Items | Where-Object { $byKey[$_.ItemKey].Result -eq 'Skipped' }).Count) { $runStatus = 'StoppedSkippedItems' }
        Save-TransientState
        Write-TransientLog ('Finished batch {0}/{1}; session={2}; export={3}; Success={4}; Skipped={5}; Error={6}; Unknown={7}; errorsForThreshold={8}/{9}; status={10}' -f
            $batch.Number,$batches.Count,$copySession,$reportPath,
            @($batch.Items | Where-Object { $byKey[$_.ItemKey].Result -eq 'Success' }).Count,
            @($batch.Items | Where-Object { $byKey[$_.ItemKey].Result -eq 'Skipped' }).Count,
            $itemErrors,$unknownCount,$errorCount,$MaxErrorsPerBatch,$runStatus)
        if ($runStatus -ne 'Running') { throw "Batch $($batch.Number) requires review: $runStatus. No further copy was started." }
        $currentBatch = $null
    }
    $runStatus = if (@($results | Where-Object Result -EQ 'Skipped').Count) { 'CompletedWithSkipped' } else { 'Completed' }
}
catch {
    if ($runStatus -eq 'Running') { $runStatus = 'Failed' }
    if ($currentBatch -and $currentCopyStarted) {
        foreach ($item in $currentBatch.Items) {
            $result = $byKey[$item.ItemKey]
            if ($result.Result -eq 'NotAttempted') {
                $result.Result = 'Unreported'
                $result.CopySessionId = $currentCopySession
                $result.Error = 'Copy may have started; inspect the ShareGate task and log. ' + $_.Exception.Message
            }
        }
    }
    Write-TransientLog ('Batch run stopped: ' + $_.Exception.Message)
    throw
}
finally {
    try {
        Save-TransientState
        Write-TransientLog ('Global result: status={0}; planned={1}; Success={2}; Skipped={3}; Error={4}; Warning={5}; Mixed={6}; Unreported={7}; NotAttempted={8}; Hors lot - à traiter à part={9} (Site={10}; File={11}); separate pages={12}; Results={13}; HorsLot={14}' -f
            $runStatus,$results.Count,
            @($results | Where-Object Result -EQ 'Success').Count,
            @($results | Where-Object Result -EQ 'Skipped').Count,
            @($results | Where-Object Result -EQ 'Error').Count,
            @($results | Where-Object Result -EQ 'Warning').Count,
            @($results | Where-Object Result -EQ 'Mixed').Count,
            @($results | Where-Object Result -EQ 'Unreported').Count,
            @($results | Where-Object Result -EQ 'NotAttempted').Count,$outOfBatch.Count,
            @($outOfBatch | Where-Object ObjectType -EQ 'Site').Count,
            @($outOfBatch | Where-Object ObjectType -EQ 'File').Count,
            $separatePages.Count,$resultPath,$outOfBatchPath)
    }
    finally { $runLock.Dispose() }
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC9OLW8Rm53gq3h
# mdU/7TUX3L1u20c69naFSd5DJZjFeqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKkgvqhKAqSuCHMF1zc4JnU12Kx0gCqTvb21Bm/hUEEMMA0GCSqG
# SIb3DQEBAQUABIIBgHv85v67luJpB7jMxigA+GXkQx8anwljLlY3HOyjmLzB371U
# RG2NMYYrhD98IJ5/WvIs3EZ6wZFH+G4TlvAw0wuqxQEkUvBeJcOduyvKZm5R4pHj
# wuEWFzV3yMaPPpRF9KB8AMmBN0t1nvoW3RRuMhx0SfVdM0102WLreOt4U3SdrC0U
# bHs6aWyXo7vK2pK6rh/Z5YG9/FsobkNDsMwpHiwFyEZPjzrKajfLeG3U56TOlPV3
# HJfWPklNOtqvTZW7BnElwa786scchC25k8za9QnIloDFXYALj3qiPKi3drNcYNEW
# BQ8OzXWJijc51TO8pODUaQNBF+etJqtdg2DRO4XnBjWqKhVwn8Zm0Op1ARgks8/q
# zQqTOc3MLSLYs8KyRKnurInycGX6esShqFWKvyVKW3ytjKeploU0nDhCe3dDguxQ
# onmfp2FVffBYm2xec1ScDEgJoyDCjozuJyE30U3w+cV3JaRsxK00yCHp0KCT2Byk
# zl8DesCW1LomDhCerKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxODE3
# NDZaMC8GCSqGSIb3DQEJBDEiBCA9WBzd2K62rw+oMs3KUqbrb2OeMtEdSTj0w1DD
# oqFQKzANBgkqhkiG9w0BAQEFAASCAgBUqfpyU9PJlu/gXDu/+x+w9MLV/cuPU1sI
# y5acaTyDOiq0wXNc9VZ4HdPdAwjXfulxqDuQFcsED7fm27YBZRd/EoS/U0cm9DoN
# /Zgz9RTSoQpcTP7qjRuHoBqJ7nwSfqAC1YaIFaru+GxF3ss/Mmq7O4ZiKeVbfS0k
# UsN2JUa6NqDylfUlkxEO6vqF+EKm2zgWI8TzrBGuaRvydpUx6BwCNngVbN5CqKA4
# uk96WTnmMPkYUFGG+S8hDfQJ406YCCvz+uBn+1x7Oj/ej2RamTdjZ17LccyoEz5R
# lF7AU2dQP6JWIUCEjxTHSmNqhocUUvJf1lX3GhHoavwZAMOSu+xXGQvHZT8yZWlV
# XkSv6R04t3En2Se6RuAT8OQ0rxpdqGDtybkwwEGiyKYD6z0/Vlwke7LEvAuE6XMj
# Y9n+zT4G4bWAAKpsQ+0I2wsRBIg70U/nuwvXmhlRB4LXAV2G2KtDVq8ao+eyMu8j
# mE0weJNcxK6xxJzajjAEyNR5pkYkA3cbsMt+a0B4JvagcL7t/gP7yowKRWrNoZG5
# jLKco2RUjegkAslzqKqs7HD4o76CZs9Yd5Js04qTuMD8iYS185x02Cl2ZJut9F1j
# FhwUvUjKNCi2nsSb0X0Dhu9KdtB2IegNW+IV9lTsTzgZXv97DLoph32pfm92Wm/a
# Jp4cz7M/Mw==
# SIG # End signature block
