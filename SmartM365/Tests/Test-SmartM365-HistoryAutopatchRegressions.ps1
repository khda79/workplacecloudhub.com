<#
.SYNOPSIS
Offline regression tests for weekly history, Autopatch and mailbox finalization.
.VERSION
1.0.4
#>
[CmdletBinding()]
param([ValidateSet('History','Autopatch','Mailbox','All')][string]$Case='All')
$ErrorActionPreference='Stop'
# Collector functions call exported Core helpers (Graph error classification, export download), as in production.
Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.99' -Force -Global -WarningAction SilentlyContinue
# Endpoint Analytics and Autopatch import Core with -Prefix Core: their functions call the prefixed names.
Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.99' -Prefix Core -Global -WarningAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Force
$module=Get-Module SmartM365.JsonTransport
$originalPolicy=& $module { (Get-Command Get-SmartM365JsonTransportPolicy).ScriptBlock }
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-HistoryRegression-' + [guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory $root
$script:passed=0
function Check($Value,$Message) {if(-not $Value){throw $Message};$script:passed++}
function Reject($Action,$Message) {$failed=$false;try{& $Action | Out-Null}catch{$failed=$true};Check $failed $Message}
function Definition($Path,$Name) {
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $Path),[ref]$tokens,[ref]$errors)
    if($errors.Count){throw ($errors|Out-String)}
    $node=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name},$true)
    if(-not $node){throw "Missing function $Name"}
    $node.Extent.Text
}
try {
    & $module {function script:Get-SmartM365JsonTransportPolicy { @{Mode='JsonText';QualifiedUncRoots=@();QualifiedLocalDataRoots=@()} }}
    if($Case -in 'History','All') {
        $mailboxLabels=@('Exchange on-prem mailboxes','Exchange on-prem remote mailboxes','Exchange on-prem mailbox daily stats')
        foreach($label in @('SmartM365 inventory') + $mailboxLabels) {
            $history=Join-Path $root (([guid]::NewGuid().ToString('N')) + '/Tenants/fixture/DATA-ALL/Exchange/OnPrem/Mailboxes/WeeklyHistory')
            $week=Join-Path $history '2026-W29';$null=New-Item -ItemType Directory $week -Force
            $legacy=Join-Path $week 'manifest.json'
            $doc=[ordered]@{Week='2026-W29';HistoryLabel=$label;HistoryRootPath=$history;Files=@('Exchange_OnPrem_Mailboxes_AllDomains.csv')}
            [IO.File]::WriteAllText($legacy,($doc|ConvertTo-Json),[Text.UTF8Encoding]::new($true))
            [IO.File]::WriteAllText((Join-Path $week $doc.Files[0]),"Id`r`n42`r`n")
            $hash=(Get-FileHash $legacy).Hash
            foreach($request in $mailboxLabels) {
                $null=Resolve-SmartM365WeeklyManifestPaths -HistoryRootPath $history -HistoryLabel $request
                Check ((Get-FileHash "$legacy.txt").Hash -eq $hash) 'Compatible label rewrote history.'
            }
            Check (-not (Test-Path $legacy)) 'Legacy name retained after verified transition.'
            $doc.HistoryLabel='Foreign collector'
            [IO.File]::WriteAllText("$legacy.txt",($doc|ConvertTo-Json))
            Reject {Resolve-SmartM365WeeklyManifestPaths -HistoryRootPath $history -HistoryLabel $request} 'Foreign label accepted.'
        }
        foreach($scenario in @('generic','relocated','foreignTenant','foreignRoot','incomplete','foreignLabel','invalid','identical','different','resume')) {
            $history=Join-Path $root "$scenario/Tenants/fixture/DATA-ALL/M365/Licensing/Licenses/WeeklyHistory"
            $week=Join-Path $history '2026-W28';$null=New-Item -ItemType Directory $week -Force
            $legacy=Join-Path $week 'manifest.json'
            $doc=[ordered]@{Week='2026-W28';HistoryLabel='SmartM365 inventory';HistoryRootPath=$history;Files=@('M365_Licenses_Users.csv')}
            if($scenario -eq 'relocated'){$doc.HistoryRootPath='C:\old\Tenants\fixture\DATA-ALL\M365\Licensing\Licenses\WeeklyHistory'}
            if($scenario -eq 'foreignTenant'){$doc.HistoryRootPath=$history.Replace('\fixture\','\another\').Replace('/fixture/','/another/')}
            if($scenario -eq 'foreignRoot'){$doc.HistoryRootPath=$history.Replace('Licenses','Different')}
            if($scenario -eq 'foreignLabel'){$doc.HistoryLabel='Foreign'}
            [IO.File]::WriteAllText($legacy,($doc|ConvertTo-Json),[Text.UTF8Encoding]::new($true))
            [IO.File]::WriteAllText((Join-Path $week $doc.Files[0]),"Id`r`n42`r`n")
            if($scenario -eq 'incomplete'){[IO.File]::Move((Join-Path $week $doc.Files[0]),(Join-Path $week 'unrelated.csv'))}
            $hash=(Get-FileHash $legacy).Hash
            if($scenario -in 'identical','resume'){Copy-Item $legacy "$legacy.txt"}
            if($scenario -eq 'invalid'){[IO.File]::WriteAllText("$legacy.txt",'{')}
            if($scenario -eq 'different'){[IO.File]::WriteAllText("$legacy.txt",(($doc|ConvertTo-Json) + "`n"))}
            if($scenario -eq 'resume') {
                [IO.File]::WriteAllText("$legacy.migration.log",((@{Owner='WeeklyHistory:SmartM365 inventory';Phase='Prepared';SHA256=$hash}|ConvertTo-Json -Compress) + "`n"))
            }
            if($scenario -in 'generic','relocated','identical','resume') {
                $null=Resolve-SmartM365WeeklyManifestPaths -HistoryRootPath $history -HistoryLabel 'M365 licenses inventory'
                Check ((Get-FileHash "$legacy.txt").Hash -eq $hash -and -not (Test-Path $legacy)) "Migration failed: $scenario"
            } else {
                Reject {Resolve-SmartM365WeeklyManifestPaths -HistoryRootPath $history -HistoryLabel 'M365 licenses inventory'} "Unsafe $scenario accepted"
                Check ((Get-FileHash $legacy).Hash -eq $hash) "Unsafe $scenario destroyed old bytes"
            }
        }
    }
    if($Case -in 'Autopatch','All') {
        Set-StrictMode -Version Latest
        $file='../SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1'
        foreach($name in @('Test-AutopatchAlertMessage','Convert-FeatureRowsToAlertDetails','Convert-QualityRowsToAlertDetails','Convert-QualityErrorRowsToAlertDetails','Group-AlertSummary')) { . ([scriptblock]::Create((Definition $file $name))) }
        foreach($name in @('Convert-FeatureRowsToAlertDetails','Convert-QualityRowsToAlertDetails','Convert-QualityErrorRowsToAlertDetails')) {
            Check (@(& $name -Rows @() -PolicyMap @{}).Count -eq 0) 'Empty completed report failed.'
            Reject {& $name -Rows $null -PolicyMap @{}} 'Unnormalized null accepted.'
            $row=[pscustomobject]@{LatestAlertMessage='0';AlertMessage='0';DeviceId='device';DeviceName='fixture';PolicyId='policy';EventDateTimeUTC='2026-01-01';LastWUScanTimeUTC='2026-01-01';AggregateState='Error';CurrentDeviceUpdateStatus='Error';ExpediteQUReleaseDate='2026-01-01'}
            Check (@(& $name -Rows @($row) -PolicyMap @{}).Count -eq 0) 'No-alert sentinel emitted a row.'
            $row.LatestAlertMessage='UnknownAlert';$row.AlertMessage='UnknownAlert'
            $alerts=@(& $name -Rows @($row,$row) -PolicyMap @{policy='Fixture'})
            Check ($alerts.Count -eq 2 -and $alerts[0].DeviceId -eq 'device') 'Alert rows lost.'
            $summary=@(Group-AlertSummary -Details $alerts)
            Check ($summary.Count -eq 1 -and $summary[0].Impact -eq 2) 'Summary lost alert count.'
        }
        Check (@(Group-AlertSummary -Details @()).Count -eq 0) 'Empty alert summary failed.'

        # Exercise the actual feature/quality collection and publication phase with completed synthetic reports.
        $source=[IO.File]::ReadAllText((Join-Path $PSScriptRoot $file))
        $start=$source.IndexOf('    $detailRows = New-Object')
        $end=$source.IndexOf('    Remove-CoreSmartM365TimestampedFilesOlderThan', $start)
        Check ($start -ge 0 -and $end -gt $start) 'Autopatch collection phase not found.'
        $phase=[scriptblock]::Create($source.Substring($start,$end-$start))
        function Write-Log {param($Message,$Level)}
        function Get-FeatureUpdatePolicyMap { @{empty='Empty';alert='Alert'} }
        function Get-QualityUpdatePolicyMap { @{} }
        function Import-ExportedCsv {
            param($ReportName,$Select,$Filter)
            $script:requested += $ReportName
            if($script:reportCase -eq 'failed'){throw 'Synthetic Graph report failure'}
            if($ReportName -eq 'FeatureUpdatePolicyStatusSummary'){return @([pscustomobject]@{PolicyId='empty'},[pscustomobject]@{PolicyId='alert'})}
            if($ReportName -eq 'FeatureUpdateDeviceState' -and $Filter -match "'alert'" -and $script:reportCase -eq 'alerts') {return $script:alertFixture}
            # Completed header-only CSV: PowerShell emits no pipeline objects.
        }
        function Publish-CoreSmartM365Csv {
            param($Data,$TimestampedPath,$LatestPath,$Columns)
            $script:published += [pscustomobject]@{Path=$TimestampedPath;Rows=@($Data).Count;Columns=$Columns}
        }
        $script:alertFixture=$row
        $IncludeFeatureUpdates=$true;$IncludeQualityUpdates=$true;$MaxItems=0
        $PolicyCsvPath='policy.csv';$PolicyLatestCsvPath='latest-policy.csv';$PolicyColumns=@('PolicyId')
        $DetailCsvPath='details.csv';$DetailLatestCsvPath='latest-details.csv';$DetailColumns=@('DeviceId','AlertName')
        $SummaryCsvPath='summary.csv';$SummaryLatestCsvPath='latest-summary.csv';$SummaryColumns=@('AlertName','Impact')
        foreach($script:reportCase in @('empty','alerts','failed')) {
            $script:published=@();$script:requested=@()
            if($script:reportCase -eq 'failed') {
                Reject {& $phase} 'Graph failure converted to an empty successful report.'
                Check ($script:published.Count -eq 0) 'Failed report published partial output.'
            } else {
                & $phase
                Check ($script:requested.Count -eq 4) 'Empty feature report prevented subsequent policy/quality requests.'
                Check ($script:published.Count -eq 3) 'Missing policy, detail or summary publication.'
                $expected=if($script:reportCase -eq 'alerts'){1}else{0}
                Check (($script:published|Where-Object Path -eq 'details.csv').Rows -eq $expected) 'Detail publication count changed.'
                Check (($script:published|Where-Object Path -eq 'summary.csv').Rows -eq $expected) 'Summary publication count changed.'
                Check (@($script:published|Where-Object {$_.Columns.Count -eq 0}).Count -eq 0) 'Empty publication lost CSV schemas.'
            }
        }
        # Actual CSV importer, synthetic ZIP files and a mocked download; no Graph request.
        . ([scriptblock]::Create((Definition $file 'Import-ExportedCsv')))
        function Start-ExportJob {param($ReportName,$Select,$Filter) @{id=[guid]::NewGuid().ToString('N')} }
        function Wait-ExportJob {param($JobId) @{url='https://example.test/synthetic'} }
        function Invoke-WebRequest {param($Uri,$OutFile) Copy-Item -LiteralPath $script:fixtureZip -Destination $OutFile }
        function Ensure-Folder {param($Path)$null=New-Item -ItemType Directory -Path $Path -Force}
        # Isolate ZIP plumbing from the locally installed Archive module; test the collector importer.
        function Expand-Archive {param($Path,$DestinationPath,[switch]$Force) [IO.Compression.ZipFile]::ExtractToDirectory($Path,$DestinationPath,$true)}
        foreach($payload in @('empty','row','noHeader','missingCsv')) {
            $folder=Join-Path $root "zip-$payload";$null=New-Item -ItemType Directory $folder
            $csv=Join-Path $folder 'report.csv'
            $content=switch($payload){'empty'{"DeviceId,PolicyId`r`n"};'row'{"DeviceId,PolicyId`r`ndevice,policy`r`n"};default{''}}
            if($payload -eq 'missingCsv'){$csv=Join-Path $folder 'missing.txt'}
            [IO.File]::WriteAllText($csv,$content)
            $script:fixtureZip=Join-Path $root "$payload.zip"
            [IO.Compression.ZipFile]::CreateFromDirectory($folder,$script:fixtureZip)
            if($payload -in 'noHeader','missingCsv') {
                Reject {Import-ExportedCsv -ReportName 'Synthetic'} 'Malformed or missing CSV was accepted.'
            } else {
                $rows=@(Import-ExportedCsv -ReportName 'Synthetic')
                $expected=if($payload -eq 'row'){1}else{0}
                Check ($rows.Count -eq $expected) 'CSV importer changed row count.'
            }
        }

    }
    if($Case -in 'Mailbox','All') {
        $path=Join-Path $PSScriptRoot '../SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
        $ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$null,[ref]$null)
        $failure=$ast.Find({param($n)$n -is [Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -eq '-not $InventoryCompletedSuccessfully -and -not (Test-SmartM365InterruptedRun)'},$true)
        Check ($null -ne $failure) 'Mailbox final failure branch missing.'
        foreach($scenario in @('record','fallback','success')) {
            $fixture=Join-Path $root ("Mailbox-$scenario.ps1")
            $lines=@(
                'param([string]$Scenario)'
                '$StartTime=Get-Date'
                '$InventoryCompletedSuccessfully=$Scenario -eq "success"'
                '$InventoryFailureRecord=if($Scenario -eq "record"){try{throw "Synthetic inventory failure"}catch{$_}}else{$null}'
                'function Stop-SmartM365TranscriptSafely {}'
                'function Test-SmartM365InterruptedRun {$false}'
                'function Complete-SmartM365ExecutionContext {param($Status,$FailureStage,$ErrorRecord)}'
                ('if ('+$failure.Clauses[0].Item1.Extent.Text+') '+$failure.Clauses[0].Item2.Extent.Text)
            )
            [IO.File]::WriteAllText($fixture,($lines -join "`r`n"))
            foreach($engine in @((Join-Path $PSHOME 'pwsh.exe'), "$env:SystemRoot/System32/WindowsPowerShell/v1.0/powershell.exe")) {
                if(-not (Test-Path $engine)){continue}
                $result=& $engine -NoProfile -ExecutionPolicy Bypass -File $fixture $scenario 2>&1
                $expected=if($scenario -eq 'success'){0}else{1}
                Check ($LASTEXITCODE -eq $expected) "Mailbox $scenario returned false status with $engine."
                if($scenario -eq 'record'){Check (($result|Out-String) -match 'Synthetic inventory failure') 'Original failure message lost.'}
            }
        }
    }
    [pscustomobject]@{Passed=$script:passed;Case=$Case;FixtureRoot=$root;Evidence='Synthetic local fixtures only; no collection or remote access'}
} finally {& $module {param($p)Set-Item Function:script:Get-SmartM365JsonTransportPolicy -Value $p} $originalPolicy}
