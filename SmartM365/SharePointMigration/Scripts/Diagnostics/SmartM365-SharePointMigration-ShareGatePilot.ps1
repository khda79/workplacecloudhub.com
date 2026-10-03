<#
.SYNOPSIS
    Approval-gated, five-item ShareGate remediation pilot for source 401 cases.
.VERSION
    1.0.3
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ProjectRoot,
    [Parameter(Mandatory)][string]$AnalysisDirectory,
    [Parameter(Mandatory)][string]$WitnessDirectory,
    [Parameter(Mandatory)][string]$SessionId,
    [ValidateNotNullOrEmpty()][string]$FarmTimeZoneId = 'W. Europe Standard Time',
    [ValidatePattern('^[0-9A-Fa-f]{64}$')][string]$ExpectedAnalysisHash = '',
    [ValidatePattern('^[0-9A-Fa-f]{64}$')][string]$ExpectedWitnessHash = '',
    [switch]$DryRun,
    [switch]$Run,
    [switch]$ConfirmPilot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Launchers\SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
$script:ConsoleLifecycleContext = Start-SmartM365MigrationConsoleLifecycle -ScriptPath $PSCommandPath
$script:ConsoleLifecycleFailure = $null
$script:ConsoleLifecycleStatus = 'SUCCESS'
try {
if ($DryRun -and $Run) { throw 'Choose either -DryRun or -Run.' }
if ($Run) { throw 'Real item-scoped copy is disabled: validation found that Copy-Content -SourceItemId created files at the destination library root while the originals already existed in their subfolders. Review paths and use a corrected remediation plan before any further copy.' }
if ($ConfirmPilot -and -not $Run) { throw '-ConfirmPilot applies only with -Run.' }
if ($Run -and -not $ConfirmPilot) { throw 'A real copy requires both -Run and -ConfirmPilot.' }
if ($Run -and (-not $ExpectedAnalysisHash -or -not $ExpectedWitnessHash)) { throw 'A real copy requires both reviewed analysis and witness SHA256 hashes.' }
if ($PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5) { throw 'Use Windows PowerShell 5.1 (powershell.exe) for ShareGate.' }
. (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-FarmMaintenance.ps1')
. (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-ShareGateReportReader.ps1')
$reportAliases = Get-SmartM365ShareGateReportAliases -ConfigRoot (Join-Path $PSScriptRoot '..\..\Config')
$farmTimeZone = [TimeZoneInfo]::FindSystemTimeZoneById($FarmTimeZoneId)
$null = Assert-SmartM365OutsideFarmMaintenance -FarmTimeZone $farmTimeZone -Phase 'pilot preparation'

$project = (Resolve-Path -LiteralPath $ProjectRoot -ErrorAction Stop).ProviderPath
$analysis = (Resolve-Path -LiteralPath $AnalysisDirectory -ErrorAction Stop).ProviderPath
$witness = (Resolve-Path -LiteralPath $WitnessDirectory -ErrorAction Stop).ProviderPath
$diagnosticsRoot = Join-Path $project 'ShareGate\Diagnostics'
foreach ($inputDirectory in @($analysis, $witness)) {
    if (-not $inputDirectory.StartsWith($diagnosticsRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'AnalysisDirectory and WitnessDirectory must be under this project ShareGate\Diagnostics folder.'
    }
}
if ($analysis -eq $witness) { throw 'AnalysisDirectory and WitnessDirectory must differ.' }
$classifiedPath = Join-Path $analysis 'ClassifiedRows.csv'
$witnessResultsPath = Join-Path $witness 'Witness-Results.csv'
$witnessLogPath = Join-Path $witness 'Witness.log'
foreach ($inputPath in @($classifiedPath, $witnessResultsPath, $witnessLogPath)) {
    if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) { throw "Required input is missing: $inputPath" }
}
$analysisHash = (Get-FileHash -LiteralPath $classifiedPath -Algorithm SHA256).Hash
if ($ExpectedAnalysisHash -and $analysisHash -ne $ExpectedAnalysisHash.ToUpperInvariant()) { throw 'ClassifiedRows.csv hash differs from the reviewed analysis.' }
$witnessLog = Get-Content -LiteralPath $witnessLogPath -Raw -ErrorAction Stop
if ($witnessLog -notmatch ('(?i)AnalysisSHA256=' + [regex]::Escape($analysisHash)) -or $witnessLog -notmatch ('(?i)Session=' + [regex]::Escape($SessionId) + '(?:;|\s)')) {
    throw 'Witness log does not match the analysis hash and session.'
}
$witnessRows = @(Import-Csv -LiteralPath $witnessResultsPath -Encoding UTF8)
if ($witnessRows.Count -ne 5 -or @($witnessRows | Where-Object SessionId -NE $SessionId).Count) { throw 'Witness results must contain exactly five rows for this session.' }
$selectionHash = (Get-FileHash -LiteralPath $witnessResultsPath -Algorithm SHA256).Hash
if ($ExpectedWitnessHash -and $selectionHash -ne $ExpectedWitnessHash.ToUpperInvariant()) { throw 'Witness-Results.csv hash differs from the reviewed witness.' }

function Assert-WitnessRow {
    param($Row, [string]$Role, [string]$ExpectedStatus, [bool]$MustContainRows)
    if ($null -eq $Row -or $Row.Role -ne $Role -or $Row.Status -ne $ExpectedStatus -or $Row.Error) { throw "Witness role $Role did not pass its control." }
    $expectedCount = 0
    if (-not [int]::TryParse([string]$Row.ReportRows, [ref]$expectedCount)) { throw "Witness role $Role has an invalid report count." }
    if ($MustContainRows -ne ($expectedCount -gt 0)) { throw "Witness role $Role has an unexpected report count." }
    $report = (Resolve-Path -LiteralPath $Row.ReportPath -ErrorAction Stop).ProviderPath
    if (-not $report.StartsWith($witness + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'A witness report is outside WitnessDirectory.' }
    $actualCount = @(Import-Csv -LiteralPath $report -Encoding UTF8).Count
    if ($actualCount -ne $expectedCount) { throw "Witness role $Role report count changed." }
}
$roles = @('WitnessShortcut','WitnessModernLink','AccessDominantFirst','AccessDominantLast','AccessOtherList')
foreach ($role in $roles) {
    $matches = @($witnessRows | Where-Object Role -EQ $role)
    if ($matches.Count -ne 1) { throw "Witness role $role must occur exactly once." }
    Assert-WitnessRow -Row $matches[0] -Role $role -ExpectedStatus $(if ($role -like 'Witness*') { 'Pre-check warning' } else { 'Undetermined - empty report' }) -MustContainRows ($role -like 'Witness*')
}

$classified = @(Import-Csv -LiteralPath $classifiedPath -Encoding UTF8 | Where-Object SessionId -EQ $SessionId)
$accessByKey = @{}
foreach ($row in $classified) {
    if ($row.RuleId -ne 'SG-ACCESS-SOURCE' -or $row.State -ne 'To fix' -or $row.AccessSide -ne 'Source' -or -not $row.ItemKey -or
        -not $row.SourceUrl -or -not $row.SourceList -or -not $row.DestinationUrl -or -not $row.DestinationList) { continue }
    $id = 0
    if (-not [int]::TryParse([string]$row.SourceItemId, [ref]$id) -or $id -le 0) { continue }
    if (-not $accessByKey.ContainsKey($row.ItemKey) -or ($accessByKey[$row.ItemKey].ObjectType -ne 'File' -and $row.ObjectType -eq 'File')) { $accessByKey[$row.ItemKey] = $row }
}
$access = @($accessByKey.Values)
if ($access.Count -lt 5) { throw 'Fewer than five eligible source 401 items remain.' }

function Get-EndpointKey {
    param($Row)
    return (($Row.SourceUrl.TrimEnd('/') + '|' + $Row.SourceList + '|' + $Row.DestinationUrl.TrimEnd('/') + '|' + $Row.DestinationList).ToLowerInvariant())
}
function Get-ReviewedAccessRow {
    param([string]$Role)
    $witnessRow = @($witnessRows | Where-Object Role -EQ $Role)[0]
    if (-not $accessByKey.ContainsKey($witnessRow.ItemKey)) { throw "Witness access item $Role is absent from the reviewed analysis." }
    $row = $accessByKey[$witnessRow.ItemKey]
    foreach ($name in @('SourceUrl','SourceList','SourceItemId','DestinationUrl','DestinationList')) {
        if ([string]$row.$name -ne [string]$witnessRow.$name) { throw "Witness access item $Role differs from the reviewed analysis." }
    }
    return $row
}
$first = Get-ReviewedAccessRow -Role 'AccessDominantFirst'
$last = Get-ReviewedAccessRow -Role 'AccessDominantLast'
$other = Get-ReviewedAccessRow -Role 'AccessOtherList'
$dominantKey = Get-EndpointKey -Row $first
if ((Get-EndpointKey -Row $last) -ne $dominantKey -or (Get-EndpointKey -Row $other) -eq $dominantKey) { throw 'Witness access items do not match the expected two-list pilot scope.' }
$dominant = @($access | Where-Object { (Get-EndpointKey -Row $_) -eq $dominantKey } | Sort-Object @{Expression={ [int]$_.SourceItemId }}, ItemKey)
if ($dominant.Count -lt 4) { throw 'The dominant list has too few distinct 401 items for this pilot.' }
$selectedKeys = @($first.ItemKey, $last.ItemKey, $other.ItemKey)
function Get-NearestAdditional {
    param([int]$TargetIndex)
    $candidates = @(for ($index = 0; $index -lt $dominant.Count; $index++) {
        if ($dominant[$index].ItemKey -notin $selectedKeys) {
            [pscustomobject]@{ Distance=[math]::Abs($index - $TargetIndex); Index=$index; Row=$dominant[$index] }
        }
    })
    $ranked = @($candidates | Sort-Object Distance, Index | Select-Object -First 1)
    if ($ranked.Count -ne 1) { throw 'Could not select an additional source 401 item.' }
    return $ranked[0].Row
}
$middleOne = Get-NearestAdditional -TargetIndex ([math]::Floor(($dominant.Count - 1) / 3))
$selectedKeys += $middleOne.ItemKey
$middleTwo = Get-NearestAdditional -TargetIndex ([math]::Floor(2 * ($dominant.Count - 1) / 3))
$selection = @(
    [pscustomobject]@{ Role='AccessDominantFirst'; Row=$first },
    [pscustomobject]@{ Role='AccessDominantMiddle1'; Row=$middleOne },
    [pscustomobject]@{ Role='AccessDominantMiddle2'; Row=$middleTwo },
    [pscustomobject]@{ Role='AccessDominantLast'; Row=$last },
    [pscustomobject]@{ Role='AccessOtherList'; Row=$other }
)
if (@($selection | ForEach-Object { $_.Row.ItemKey } | Sort-Object -Unique).Count -ne 5) { throw 'Pilot item keys are not unique.' }
Write-Output ('{0} Mode={1}; Session={2}; AnalysisSHA256={3}; WitnessSHA256={4}; selected=5; writes={5}; FarmTimeZone={6}; Maintenance=23:45-00:15' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $(if ($Run) { 'Run' } else { 'DryRun' }), $SessionId, $analysisHash, $selectionHash, $(if ($Run) { 'possible after confirmation' } else { 'none' }), $farmTimeZone.Id)
foreach ($entry in $selection) {
    $row = $entry.Row
    Write-Output ('{0} {1}: ID={2}; {3} | {4} -> {5} | {6}; original={7}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $entry.Role, $row.SourceItemId, $row.SourceUrl, $row.SourceList, $row.DestinationUrl, $row.DestinationList, $row.RuleId)
}
if (-not $Run) { return }

$phrase = 'COPY 5 ITEMS ' + $SessionId
$entered = Read-Host ('Type exactly "{0}" to authorize these five real copies' -f $phrase)
if ($entered -cne $phrase) { throw 'Pilot confirmation was not entered exactly. No ShareGate copy was started.' }
$null = Assert-SmartM365OutsideFarmMaintenance -FarmTimeZone $farmTimeZone -Phase 'real pilot startup'

$runId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N')
$output = Join-Path $diagnosticsRoot ('Pilot-' + $runId)
$reports = Join-Path $output 'Reports'
$objects = Join-Path $output 'CopyResults'
New-Item -ItemType Directory -Path $reports,$objects -Force | Out-Null
$logPath = Join-Path $output 'Pilot.log'
$resultPath = Join-Path $output 'Pilot-Results.csv'
$resultColumns = @('Role','SessionId','ItemKey','SourceUrl','SourceList','SourceItemId','DestinationUrl','DestinationList','TaskName','CopySessionId','Status','ShareGateResult','ReportRows','ReportPath','CopyResultPath','DestinationPath','DestinationItemId','DestinationItemUrl','DestinationItemUrlEvidence','Error')
$results = [System.Collections.Generic.List[object]]::new()
foreach ($entry in $selection) {
    $row = $entry.Row
    $results.Add([pscustomobject]@{ Role=$entry.Role; SessionId=$SessionId; ItemKey=$row.ItemKey; SourceUrl=$row.SourceUrl; SourceList=$row.SourceList; SourceItemId=$row.SourceItemId; DestinationUrl=$row.DestinationUrl; DestinationList=$row.DestinationList; TaskName=''; CopySessionId=''; Status='Not attempted'; ShareGateResult='Not attempted'; ReportRows=0; ReportPath=''; CopyResultPath=''; DestinationPath=''; DestinationItemId=''; DestinationItemUrl=''; DestinationItemUrlEvidence='Unavailable'; Error='' })
}
$siteCache = @{}
$listCache = @{}

function Write-PilotLog {
    param([string]$Message)
    $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
    Write-Output $line
}
function Export-AtomicCsv {
    param([string]$Path, [object[]]$Rows, [string[]]$Columns)
    $temporary = Join-Path (Split-Path -Parent $Path) ('.' + [guid]::NewGuid().ToString('N') + '.tmp')
    try {
        if ($Rows.Count) { $Rows | Select-Object -Property $Columns | Export-Csv -LiteralPath $temporary -NoTypeInformation -Encoding UTF8 }
        else { Set-Content -LiteralPath $temporary -Value (($Columns | ForEach-Object { '"' + $_.Replace('"','""') + '"' }) -join ',') -Encoding UTF8 }
        Move-Item -LiteralPath $temporary -Destination $Path -Force
    }
    finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force } }
}
function Get-ExactList {
    param([string]$Side, [string]$Url, [string]$Name)
    $siteKey = $Side + '|' + $Url.TrimEnd('/').ToLowerInvariant()
    if (-not $siteCache.ContainsKey($siteKey)) {
        if ($Side -eq 'Source') { $siteCache[$siteKey] = ShareGate\Connect-Site -Url $Url -ErrorAction Stop }
        else { $siteCache[$siteKey] = ShareGate\Connect-Site -Url $Url -Browser -ErrorAction Stop }
        if (-not $siteCache[$siteKey]) { throw "Connect-Site returned no site for $Side $Url." }
    }
    $listKey = $siteKey + '|' + $Name.ToLowerInvariant()
    if (-not $listCache.ContainsKey($listKey)) {
        $found = @(ShareGate\Get-List -Site $siteCache[$siteKey] -Name $Name -ErrorAction Stop)
        $exact = @($found | Where-Object {
            ($_.PSObject.Properties['Title'] -and $_.Title -eq $Name) -or
            ($_.PSObject.Properties['Name'] -and $_.Name -eq $Name)
        })
        if ($exact.Count -ne 1) { throw "Get-List returned $($exact.Count) exact matches for $Side '$Name' at $Url." }
        $listCache[$listKey] = $exact[0]
    }
    return $listCache[$listKey]
}
function Get-CopySessionId {
    param($CopyResult)
    foreach ($name in @('SessionId','SessionID','CopySessionId','Id')) {
        $property = $CopyResult.PSObject.Properties[$name]
        if ($property -and [string]$property.Value -match '^\d{6}-\d+$') { return [string]$property.Value }
    }
    return ''
}
function Get-ShareGateReportReview {
    param([object[]]$Rows, [int]$SourceItemId)
    if (-not $Rows.Count) {
        return [pscustomobject]@{ Result='No report rows'; HasError=$false; DestinationPath=''; DestinationItemId='' }
    }
    $itemRows = $Rows
    if (Test-SmartM365ShareGateReportField -Row $Rows[0] -Field 'SourceItemId' -Aliases $reportAliases) {
        $matched = @($Rows | Where-Object { (Get-SmartM365ShareGateReportValue -Row $_ -Field 'SourceItemId' -Aliases $reportAliases) -eq [string]$SourceItemId })
        if ($matched.Count) { $itemRows = $matched }
    }
    $values = @($itemRows | ForEach-Object { Get-SmartM365ShareGateReportValue -Row $_ -Field 'Status' -Aliases $reportAliases } | Where-Object { $_ })
    $groups = @($values | Group-Object | Sort-Object Name)
    $result = if ($groups.Count) { ($groups | ForEach-Object { '{0}={1}' -f $_.Name, $_.Count }) -join '; ' } else { 'Result column unavailable' }
    $hasError = @($values | Where-Object { $_ -match '(?i)^(error|failed|failure|erreur|échec)$' }).Count -gt 0
    if (-not $values.Count) {
        $hasError = @($itemRows | Where-Object { Get-SmartM365ShareGateReportValue -Row $_ -Field 'Errors' -Aliases $reportAliases }).Count -gt 0
    }
    $paths = @($itemRows | ForEach-Object { Get-SmartM365ShareGateReportValue -Row $_ -Field 'DestinationPath' -Aliases $reportAliases } | Where-Object { $_ } | Select-Object -Unique)
    $ids = @($itemRows | ForEach-Object { Get-SmartM365ShareGateReportValue -Row $_ -Field 'DestinationItemId' -Aliases $reportAliases } | Where-Object { $_ } | Select-Object -Unique)
    return [pscustomobject]@{
        Result=$result
        HasError=$hasError
        DestinationPath=$(if ($paths.Count -eq 1) { $paths[0] } else { '' })
        DestinationItemId=$(if ($ids.Count -eq 1) { $ids[0] } else { '' })
    }
}
function Test-ExpectedDestinationUrl {
    param([string]$Candidate, [string]$SiteUrl)
    $uri = $null
    if (-not [uri]::TryCreate($Candidate, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https') { return $false }
    $site = [uri]$SiteUrl
    return ($uri.Host -eq $site.Host -and $uri.AbsolutePath.StartsWith($site.AbsolutePath.TrimEnd('/') + '/', [StringComparison]::OrdinalIgnoreCase))
}
function Resolve-DestinationItemUrl {
    param($DestinationList, [string]$SiteUrl, [string]$DestinationPath, [string]$DestinationItemId, [string]$PathEvidence)
    $relativePath = $DestinationPath.Trim().TrimStart('/')
    if ($relativePath -and (Get-Command -Name Get-File -Module ShareGate -ErrorAction SilentlyContinue)) {
        try {
            $files = @(ShareGate\Get-File -List $DestinationList -Path $relativePath -ErrorAction Stop)
            if ($files.Count -eq 1 -and $files[0].PSObject.Properties['Address']) {
                $address = [string]$files[0].Address
                if (Test-ExpectedDestinationUrl -Candidate $address -SiteUrl $SiteUrl) {
                    return [pscustomobject]@{ Url=$address; Evidence='Verified by ShareGate Get-File' }
                }
            }
        }
        catch { Write-PilotLog ('Destination Get-File lookup failed: ' + $_.Exception.Message) | Out-Null }
    }
    $numericId = 0
    if ([int]::TryParse($DestinationItemId, [ref]$numericId) -and $numericId -gt 0 -and
        (Get-Command -Name Get-ListItem -Module ShareGate -ErrorAction SilentlyContinue)) {
        try {
            $items = @(ShareGate\Get-ListItem -List $DestinationList -Id $numericId -ErrorAction Stop)
            if ($items.Count -eq 1 -and $items[0].PSObject.Properties['Address']) {
                $address = [string]$items[0].Address
                if (Test-ExpectedDestinationUrl -Candidate $address -SiteUrl $SiteUrl) {
                    return [pscustomobject]@{ Url=$address; Evidence='Verified by ShareGate Get-ListItem' }
                }
            }
        }
        catch { Write-PilotLog ('Destination Get-ListItem lookup failed: ' + $_.Exception.Message) | Out-Null }
    }
    if ($relativePath -and $DestinationList.PSObject.Properties['RootFolder']) {
        $root = [string]$DestinationList.RootFolder
        if ($root) {
            try {
                $site = [uri]$SiteUrl
                if ($root -match '^https://') { $base = $root.TrimEnd('/') }
                elseif ($root.StartsWith('/')) { $base = $site.GetLeftPart([UriPartial]::Authority) + $root.TrimEnd('/') }
                else { $base = $SiteUrl.TrimEnd('/') + '/' + $root.Trim('/') }
                $encoded = (($relativePath -split '/' | ForEach-Object { [uri]::EscapeDataString([uri]::UnescapeDataString($_)) }) -join '/')
                $candidate = ([uri]($base + '/' + $encoded)).AbsoluteUri
                if (Test-ExpectedDestinationUrl -Candidate $candidate -SiteUrl $SiteUrl) {
                    return [pscustomobject]@{ Url=$candidate; Evidence=('Inferred from ' + $PathEvidence + ' and destination RootFolder; verify in SPO') }
                }
            }
            catch { Write-PilotLog ('Destination URL inference failed: ' + $_.Exception.Message) | Out-Null }
        }
    }
    return [pscustomobject]@{ Url=''; Evidence='Item URL unavailable; use destination site and report' }
}
function Write-PilotSummary {
    foreach ($item in $results) {
        Write-PilotLog ('SUMMARY ID={0}; ShareGate={1}; Status={2}; Export={3}; SPO item={4}; URL evidence={5}; Destination site={6}' -f
            $item.SourceItemId, $item.ShareGateResult, $item.Status,
            $(if ($item.ReportPath) { $item.ReportPath } else { '(none)' }),
            $(if ($item.DestinationItemUrl) { $item.DestinationItemUrl } else { '(unavailable)' }),
            $item.DestinationItemUrlEvidence, $item.DestinationUrl)
    }
}

try {
    Write-PilotLog ('Actor={0}\{1}; Machine={2}; RealCopy=True; Session={3}; AnalysisSHA256={4}; WitnessSHA256={5}; FarmTimeZone={6}; Maintenance=23:45-00:15' -f $env:USERDOMAIN, $env:USERNAME, $env:COMPUTERNAME, $SessionId, $analysisHash, $selectionHash, $farmTimeZone.Id)
    $plan = @($selection | ForEach-Object { [pscustomobject]@{ Role=$_.Role; ItemKey=$_.Row.ItemKey; SourceUrl=$_.Row.SourceUrl; SourceList=$_.Row.SourceList; SourceItemId=$_.Row.SourceItemId; DestinationUrl=$_.Row.DestinationUrl; DestinationList=$_.Row.DestinationList } })
    Export-AtomicCsv -Path (Join-Path $output 'Pilot-Plan.csv') -Rows $plan -Columns @('Role','ItemKey','SourceUrl','SourceList','SourceItemId','DestinationUrl','DestinationList')
    Export-AtomicCsv -Path $resultPath -Rows $results.ToArray() -Columns $resultColumns
    $module = @(Get-Module -ListAvailable -Name ShareGate | Sort-Object Version -Descending | Select-Object -First 1)
    if (-not $module.Count) { throw 'ShareGate module is not discoverable in Windows PowerShell 5.1.' }
    Import-Module -Name $module[0].Path -ErrorAction Stop
    Write-PilotLog ('ShareGate module version={0}; path={1}' -f $module[0].Version, $module[0].Path)
    foreach ($name in @('Connect-Site','Get-List','Copy-Content','New-CopySettings','Export-Report')) {
        if (-not (Get-Command -Name $name -Module ShareGate -ErrorAction SilentlyContinue)) { throw "Required ShareGate cmdlet is missing: $name" }
    }
    $copy = Get-Command Copy-Content -Module ShareGate
    $requiredCopyParameters = @('SourceItemId','SourceList','DestinationList','CopySettings','TaskName')
    if (-not @($copy.ParameterSets | Where-Object {
        $names = @($_.Parameters | ForEach-Object Name)
        @($requiredCopyParameters | Where-Object { $_ -notin $names }).Count -eq 0
    }).Count) { throw 'Installed Copy-Content lacks the required item-scoped parameter set.' }
    $settings = ShareGate\New-CopySettings -OnContentItemExists IncrementalUpdate -ErrorAction Stop
    if (-not $settings) { throw 'New-CopySettings returned no CopySettings object.' }
    for ($index = 0; $index -lt $selection.Count; $index++) {
        $entry = $selection[$index]
        $row = $entry.Row
        $taskName = 'SmartM365 401 pilot ' + $SessionId + ' ' + $runId + ' ' + $entry.Role + ' ID=' + $row.SourceItemId
        $reportPath = Join-Path $reports ('Pilot-{0:D2}.csv' -f ($index + 1))
        $objectPath = Join-Path $objects ('CopyResult-{0:D2}.txt' -f ($index + 1))
        $copySession = ''
        $reportRows = 0
        $status = 'Failed'
        $shareGateResult = 'Copy not completed'
        $destinationPath = ''
        $destinationItemId = ''
        $itemUrl = ''
        $urlEvidence = 'Unavailable'
        $errorText = ''
        try {
            $null = Assert-SmartM365OutsideFarmMaintenance -FarmTimeZone $farmTimeZone -Phase ('item {0}/5 connection' -f ($index + 1))
            $sourceList = Get-ExactList -Side 'Source' -Url $row.SourceUrl -Name $row.SourceList
            $destinationList = Get-ExactList -Side 'Destination' -Url $row.DestinationUrl -Name $row.DestinationList
            $null = Assert-SmartM365OutsideFarmMaintenance -FarmTimeZone $farmTimeZone -Phase ('item {0}/5 copy' -f ($index + 1))
            Write-PilotLog ('Starting real item copy {0}/5: {1}; ID={2}; TaskName={3}' -f ($index + 1), $entry.Role, $row.SourceItemId, $taskName)
            $copyResult = ShareGate\Copy-Content -SourceList $sourceList -DestinationList $destinationList -SourceItemId @([int]$row.SourceItemId) -CopySettings $settings -TaskName $taskName -ErrorAction Stop
            if (-not $copyResult -or @($copyResult).Count -ne 1) { throw 'Copy-Content did not return exactly one CopyResult; inspect the ShareGate session before continuing.' }
            $copySession = Get-CopySessionId -CopyResult $copyResult
            @('Type: ' + $copyResult.GetType().FullName, 'Session ID: ' + $(if ($copySession) { $copySession } else { '(not exposed)' }), '', ($copyResult | Format-List * -Force | Out-String -Width 4096)) |
                Set-Content -LiteralPath $objectPath -Encoding UTF8
            ShareGate\Export-Report -CopyResult $copyResult -Path $reportPath -ErrorAction Stop | Out-Null
            if (-not (Test-Path -LiteralPath $reportPath -PathType Leaf)) { throw 'Export-Report did not create a CSV.' }
            $exportedRows = @(Import-Csv -LiteralPath $reportPath -Encoding UTF8)
            $reportRows = $exportedRows.Count
            $review = Get-ShareGateReportReview -Rows $exportedRows -SourceItemId ([int]$row.SourceItemId)
            $shareGateResult = $review.Result
            $destinationPath = $review.DestinationPath
            $destinationItemId = $review.DestinationItemId
            $pathEvidence = 'current ShareGate export'
            if (-not $destinationPath -and $row.PSObject.Properties['Raw: Destination path']) {
                $destinationPath = [string]$row.'Raw: Destination path'
                $pathEvidence = 'prior ShareGate report'
            }
            $url = Resolve-DestinationItemUrl -DestinationList $destinationList -SiteUrl $row.DestinationUrl -DestinationPath $destinationPath -DestinationItemId $destinationItemId -PathEvidence $pathEvidence
            $itemUrl = $url.Url
            $urlEvidence = $url.Evidence
            $status = if ($review.HasError) { 'ShareGate error - stopped' } else { 'Completed - review report' }
            if ($review.HasError) { $errorText = 'The ShareGate export contains an Error result; pilot stopped before the next item.' }
        }
        catch {
            $errorText = $_.Exception.Message
            $status = 'Failed - stopped'
        }
        $result = $results[$index]
        $result.TaskName = $taskName
        $result.CopySessionId = $copySession
        $result.Status = $status
        $result.ShareGateResult = $shareGateResult
        $result.ReportRows = $reportRows
        $result.ReportPath = if (Test-Path -LiteralPath $reportPath -PathType Leaf) { $reportPath } else { '' }
        $result.CopyResultPath = if (Test-Path -LiteralPath $objectPath -PathType Leaf) { $objectPath } else { '' }
        $result.DestinationPath = $destinationPath
        $result.DestinationItemId = $destinationItemId
        $result.DestinationItemUrl = $itemUrl
        $result.DestinationItemUrlEvidence = $urlEvidence
        $result.Error = $errorText
        Export-AtomicCsv -Path $resultPath -Rows $results.ToArray() -Columns $resultColumns
        Write-PilotLog ('Finished item {0}/5: ID={1}; ShareGate={2}; status={3}; copySession={4}; reportRows={5}; export={6}; SPO item={7}; URL evidence={8}; error={9}' -f
            ($index + 1), $row.SourceItemId, $shareGateResult, $status, $copySession, $reportRows,
            $(if ($result.ReportPath) { $result.ReportPath } else { '(none)' }),
            $(if ($itemUrl) { $itemUrl } else { '(unavailable)' }), $urlEvidence, $errorText)
        if ($errorText) { throw "Pilot stopped after item $($index + 1): $errorText" }
    }
    Write-PilotLog ('Completed: 5 item-scoped real Copy-Content calls; results={0}' -f $resultPath)
}
catch {
    Write-PilotLog ('Pilot failed: ' + $_.Exception.Message)
    throw
}
finally {
    if (Test-Path -LiteralPath $logPath -PathType Leaf) { Write-PilotSummary }
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAULLCfEa6y+yb8
# rgz7Bm3hSsr66Qdqk5u3uqGcLQfR/aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIK9iXIwCb6HUYQGJutEBNG9NZDfe87WnQ1jIr70KjgaRMA0GCSqG
# SIb3DQEBAQUABIIBgD7U1tPhE4kkZ0rCnmUGIcywTNqM9nHBSh7N55Fa+INDTWeG
# pFP0c3q6YX1Ll1VcSkB1V0I7QgQXi5DI7eimQx8/zd6VHKoOtrNxKg4ufpkrl3ra
# h7WqQ4CxeT3cH38PcOpSM3G19J5OtBzRxaRks/Sp8OuFSdLrMrNm32zwyawamRVB
# pyoWuW1zHOHO+yLuzAObBnFhQRDZAwAMGIjwwLl8WoE30FRZY16PHdn2DH+fr2Wy
# 97ZeKVRrQgo2CRrTYNus4dCwFJCR1ugOcJV1fGYkkfxjqcS0VkBWRlI0pxEw07t3
# 7gM11cfAYEmN4V0Wh6WH9dV2/7Cfi7usc5hebtVJS4jKXT4B1e4lEb05GeGR1o5b
# 02AwDOdlknAu6z/5U8N7GCR+dpfdGYTTAOuOLHK82TMZNBKRVsTUUMwdO1SjsBsy
# dLwdnL+upHNljxMHmkEzfcCQgXddSYI24sw9bRE9eH2azfw+2RcmVPwV2k11/Dl+
# vOFk04QYWK4zYRKupqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxNTI3
# NDNaMC8GCSqGSIb3DQEJBDEiBCDQa1C38S5nnf3grSbDAO5kcIc+BiUo67e4Xai6
# wzekSDANBgkqhkiG9w0BAQEFAASCAgAMe+o6GVS9nu05lsXjwVcJhnDqyfUaTSAd
# d83tbhVHMNtg7sSb94Oe0hYysXVvVk2mBeziPshNpx3jDwhcJFqMiie8fXMjHKrF
# nIIDsGSkuLcCofCMM62hnJckxPytBMVbkIoe3aYbhaty8QlmYFnO5OmMT4De3Zy2
# HdsPG7eLryJGS0v4CiWAuHSHDJNepA3ohELmJ+tZuoaI6kM0nwAhe++SvDrB3JY0
# zh/iVB1qaACrsfdBd3u46AYynu0khZ2lE28fmY57oPZYITEjYcukEKXuICQAc6Rf
# wSM2oB/3HUZSvoJidXH3JDLeyJX+aSHxl808+FBoWypSs3Iail1AC+HvxU856tpg
# KoSL8CdXhe5InLhfUd4vOh1gaAKIcIH8SMAKqpVlSEuaxBA/AGgTiAPW2+Jsq2wo
# TNVLZpTxqqwlNNrv31BIERjDwAieNJfG0i7WZgLE7KPjZESepNBmM85El5I4YMmh
# jmU//8OZw9kyngMXm6aRFrqbqkC801NGqdk6YdtH6KTPtF/qO6Rm3ZolNdKchRNB
# bXVWETkmuLPksDQfdPcFMDSfQr7q9WXWdGrRvTX3Owtse9F75jVmfHuHMu6iaHUW
# J7F/17+w749YlAnKGRwbFwOXlFUk9O2iF8iJU9RuZeGJOA2C/w+Fo+H/F0Z5Wj2r
# PuzCpg/mSg==
# SIG # End signature block
