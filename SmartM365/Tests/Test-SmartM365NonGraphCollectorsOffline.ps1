<#
.SYNOPSIS
Synthetic regression tests for non-Graph SmartInventory CSV publication paths.
.VERSION
1.0.10
#>
[CmdletBinding()]
param(
    [string]$SourceRoot,
    [string]$ResultPath
)
$ErrorActionPreference = 'Stop'
if (-not $SourceRoot) { $SourceRoot = Split-Path $PSScriptRoot -Parent }
$results = New-Object 'System.Collections.Generic.List[object]'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-NonGraph-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)

function Assert-Offline {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}
function Test-OfflineCase {
    param([string]$Name, [scriptblock]$Body)
    try { & $Body; $results.Add([pscustomobject]@{ Name=$Name; Passed=$true; Error='' }) }
    catch { $results.Add([pscustomobject]@{ Name=$Name; Passed=$false; Error=$_.Exception.Message }) }
}
function Get-FunctionText {
    param([string]$Path, [string[]]$Names)
    $tokens=$null; $errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)
    if($errors.Count){throw "Source parse failed: $Path"}
    foreach($name in $Names){
        $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
        if($null -eq $node){throw "Function not found: $name"}
        $node.Extent.Text
    }
}
function New-IssueExporterModule {
    param([string]$CollectorPath)
    $corePath=Join-Path $SourceRoot 'Modules/SmartM365.Core/SmartM365.Core.psm1'
    $coreNames=@(
        'Get-SmartM365CoreContextValue','Get-SmartM365CsvValidationBaseName','Get-SmartM365CsvValidationRule',
        'Assert-SmartM365CsvDataCompleteness','Add-SmartM365TenantKeyToCsvData',
        'Write-SmartM365CsvAtomically','Copy-SmartM365FileAtomically'
    )
    $definitions=@(Get-FunctionText -Path $corePath -Names $coreNames)
    $definitions+=@(Get-FunctionText -Path $CollectorPath -Names @('Log','Warn','FlattenPath','CopyCsv','Remove-DetailedArchiveFilesOlderThan','ExportIssues'))
    $module=New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")))
    & $module {
        $script:SmartM365CoreTenantKey='synthetic-a'
        $script:SmartM365CoreOrganizationKey='synthetic-org'
        $script:SmartM365CoreEnvironmentKey='test'
        $script:SmartM365CoreTenantId='00000000-0000-0000-0000-000000000001'
        $script:DetailedArchiveRetentionDays=7
    }
    $module
}
function New-Windows11ReadinessCsvModule {
    param([string]$CollectorPath)
    $definitions=@(Get-FunctionText -Path $CollectorPath -Names @('Log','Warn','WarnIfInputCsvStale','Csv','CsvAny','CsvAnyProjected'))
    New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")))
}
function Invoke-Windows11ReadinessCsvLoad {
    param(
        [System.Management.Automation.PSModuleInfo]$Module,
        [string]$Folder,
        [datetime]$RunStartedAt,
        [double]$ThresholdHours,
        [string[]]$Names,
        [bool]$Required = $false
    )
    & $Module {
        param($folder,$runStartedAt,$thresholdHours,$names,$required)
        $script:DataLastFolder=$folder
        $script:RunStartedAt=$runStartedAt
        $script:InputCsvFreshnessWarningHours=$thresholdHours
        $script:WarningCount=0
        $stream=@(CsvAnyProjected -Names $names -Columns @('Id','Value') -Req:$required 3>&1)
        $warningRecords=@($stream | Where-Object {$_ -is [Management.Automation.WarningRecord]})
        $rows=@($stream | Where-Object {$_ -isnot [Management.Automation.WarningRecord]})
        [pscustomobject]@{
            Rows=$rows
            WarningCount=$script:WarningCount
            WarningMessages=@($warningRecords | ForEach-Object {$_.Message})
        }
    } $Folder $RunStartedAt $ThresholdHours $Names $Required
}

try {
    $collectors=@(
        [pscustomobject]@{
            Name='Exchange hybrid identity issues'
            RelativePath='SmartInventory/ExchangeInventory/Migration/SmartM365-Exchange-HybridIdentity-Issues-Inventory.ps1'
            Detail='Exchange_HybridIdentity_Issues.csv'
            Summary='Exchange_HybridIdentity_Issues_Summary.csv'
            Data=@([pscustomobject]@{IssueNumber=101;ObjectGUID='synthetic-object-1';Potential_Issue='Synthetic mismatch';IssueCategory='Identity'})
        },
        [pscustomobject]@{
            Name='Windows 11 readiness issues'
            RelativePath='SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-Intune-Windows11-Readiness-Issues-Inventory.ps1'
            Detail='Intune_Windows11_Readiness_Issues.csv'
            Summary='Intune_Windows11_Readiness_Issues_Summary.csv'
            Data=@([pscustomobject]@{IssueCode='SYN-001';Area='Readiness';ObjectGUID_Norm='synthetic-device-1';Potential_Issue='Synthetic blocker';IssueCategory='Hardware';PriorityScore=10;IsBlocking=$true;RecommendedAction='Test only';ImpactMigration='Test only'})
        }
    )

    foreach($collector in $collectors){
        $path=Join-Path $SourceRoot $collector.RelativePath
        Test-OfflineCase "$($collector.Name) uses atomic final publication" {
            $copyText=@(Get-FunctionText -Path $path -Names @('CopyCsv'))[0]
            $exportText=@(Get-FunctionText -Path $path -Names @('ExportIssues'))[0]
            Assert-Offline ($copyText -match 'Copy-SmartM365FileAtomically' -and $copyText -notmatch '\bCopy-Item\b') 'CopyCsv still writes a final path directly.'
            Assert-Offline ($exportText -match 'Write-SmartM365CsvAtomically' -and $exportText -notmatch '\bExport-Csv\b') 'ExportIssues still writes a final path directly.'
        }
        Test-OfflineCase "$($collector.Name) DATA-ALL DATA-LAST schema and byte parity" {
            $output=Join-Path $testRoot (($collector.Name -replace '[^A-Za-z0-9]','-')+'-all')
            $latest=Join-Path $testRoot (($collector.Name -replace '[^A-Za-z0-9]','-')+'-last')
            [void][IO.Directory]::CreateDirectory($output);[void][IO.Directory]::CreateDirectory($latest)
            $module=New-IssueExporterModule -CollectorPath $path
            try {
                & $module {param($o,$l,$d) $script:OutputFolder=$o;$script:LatestFolder=$l;$script:RunStamp='20260101-000000';ExportIssues $d} $output $latest $collector.Data | Out-Null
                $allDetail=Join-Path $output $collector.Detail; $lastDetail=Join-Path $latest $collector.Detail
                $allSummary=Join-Path $output $collector.Summary; $lastSummary=Join-Path $latest $collector.Summary
                Assert-Offline ((Get-FileHash $allDetail).Hash -eq (Get-FileHash $lastDetail).Hash) 'Detail DATA-ALL/DATA-LAST bytes differ.'
                Assert-Offline ((Get-FileHash $allSummary).Hash -eq (Get-FileHash $lastSummary).Hash) 'Summary DATA-ALL/DATA-LAST bytes differ.'
                $header=(Import-Csv -LiteralPath $lastDetail | Select-Object -First 1).PSObject.Properties.Name
                Assert-Offline (($header | Select-Object -First 4) -join ',' -eq 'TenantKey,OrganizationKey,EnvironmentKey,TenantId') 'Identity-first schema changed.'
            } finally {Remove-Module $module}
        }
        Test-OfflineCase "$($collector.Name) locked DATA-LAST preserves last valid export" {
            $output=Join-Path $testRoot (($collector.Name -replace '[^A-Za-z0-9]','-')+'-locked-all')
            $latest=Join-Path $testRoot (($collector.Name -replace '[^A-Za-z0-9]','-')+'-locked-last')
            [void][IO.Directory]::CreateDirectory($output);[void][IO.Directory]::CreateDirectory($latest)
            $lastDetail=Join-Path $latest $collector.Detail
            [IO.File]::WriteAllText($lastDetail,'LAST VALID SYNTHETIC EXPORT')
            $lock=[IO.File]::Open($lastDetail,'Open','Read','Read')
            $module=New-IssueExporterModule -CollectorPath $path
            try {
                $caught=$false
                try {& $module {param($o,$l,$d) $script:OutputFolder=$o;$script:LatestFolder=$l;$script:RunStamp='20260101-000000';ExportIssues $d} $output $latest $collector.Data | Out-Null}
                catch {$caught=$true}
                Assert-Offline $caught 'Locked DATA-LAST publication was acknowledged.'
            } finally {$lock.Dispose();Remove-Module $module}
            Assert-Offline ([IO.File]::ReadAllText($lastDetail) -eq 'LAST VALID SYNTHETIC EXPORT') 'Locked DATA-LAST was modified.'
        }
    }

    $readinessPath=Join-Path $SourceRoot 'SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-Intune-Windows11-Readiness-Issues-Inventory.ps1'
    $freshnessRoot=Join-Path $testRoot 'windows11-readiness-input-freshness'
    [void][IO.Directory]::CreateDirectory($freshnessRoot)
    $runStartedAt=[datetime]::SpecifyKind([datetime]'2026-09-23T12:00:00',[DateTimeKind]::Utc)

    Test-OfflineCase 'Windows 11 readiness input younger than threshold does not warn' {
        $folder=Join-Path $freshnessRoot 'younger'
        [void][IO.Directory]::CreateDirectory($folder)
        $path=Join-Path $folder 'younger.csv'
        [IO.File]::WriteAllText($path,"Id,Value`r`n1,young`r`n")
        [IO.File]::SetLastWriteTimeUtc($path,$runStartedAt.AddHours(-23))
        $module=New-Windows11ReadinessCsvModule -CollectorPath $readinessPath
        try {
            $observed=Invoke-Windows11ReadinessCsvLoad -Module $module -Folder $folder -RunStartedAt $runStartedAt -ThresholdHours 24 -Names @('younger.csv') -Required $true
            Assert-Offline ($observed.WarningCount -eq 0 -and $observed.WarningMessages.Count -eq 0) 'A CSV younger than 24 hours emitted a freshness warning.'
            Assert-Offline ($observed.Rows.Count -eq 1 -and $observed.Rows[0].Value -ceq 'young') 'The fresh CSV was not imported.'
        } finally {Remove-Module $module -Force}
    }
    Test-OfflineCase 'Windows 11 readiness input older than threshold warns and remains imported' {
        $folder=Join-Path $freshnessRoot 'older'
        [void][IO.Directory]::CreateDirectory($folder)
        $path=Join-Path $folder 'older.csv'
        [IO.File]::WriteAllText($path,"Id,Value`r`n2,old`r`n")
        $lastWriteTimeUtc=$runStartedAt.AddHours(-25)
        [IO.File]::SetLastWriteTimeUtc($path,$lastWriteTimeUtc)
        $module=New-Windows11ReadinessCsvModule -CollectorPath $readinessPath
        try {
            $observed=Invoke-Windows11ReadinessCsvLoad -Module $module -Folder $folder -RunStartedAt $runStartedAt -ThresholdHours 24 -Names @('older.csv') -Required $true
            Assert-Offline ($observed.WarningCount -eq 1 -and $observed.WarningMessages.Count -eq 1) 'A CSV older than 24 hours did not emit exactly one warning.'
            Assert-Offline ($observed.Rows.Count -eq 1 -and $observed.Rows[0].Value -ceq 'old') 'The stale CSV import did not continue.'
            $message=$observed.WarningMessages[0]
            Assert-Offline ($message -match [regex]::Escape('older.csv') -and $message -match 'age=25\.00 hour' -and $message -match [regex]::Escape($lastWriteTimeUtc.ToString('o')) -and $message -match 'threshold=24\.00 hour' -and $message -match 'Processing continues\.') 'The stale CSV warning does not contain the required operational details.'
        } finally {Remove-Module $module -Force}
    }
    Test-OfflineCase 'Windows 11 readiness input exactly at threshold does not warn' {
        $folder=Join-Path $freshnessRoot 'exact'
        [void][IO.Directory]::CreateDirectory($folder)
        $path=Join-Path $folder 'exact.csv'
        [IO.File]::WriteAllText($path,"Id,Value`r`n3,exact`r`n")
        [IO.File]::SetLastWriteTimeUtc($path,$runStartedAt.AddHours(-24))
        $module=New-Windows11ReadinessCsvModule -CollectorPath $readinessPath
        try {
            $observed=Invoke-Windows11ReadinessCsvLoad -Module $module -Folder $folder -RunStartedAt $runStartedAt -ThresholdHours 24 -Names @('exact.csv') -Required $true
            Assert-Offline ($observed.WarningCount -eq 0 -and $observed.WarningMessages.Count -eq 0) 'A CSV exactly 24 hours old emitted a freshness warning.'
            Assert-Offline ($observed.Rows.Count -eq 1) 'The threshold-boundary CSV was not imported.'
        } finally {Remove-Module $module -Force}
    }
    Test-OfflineCase 'Windows 11 readiness required input remains blocking when absent' {
        $folder=Join-Path $freshnessRoot 'required-absent'
        [void][IO.Directory]::CreateDirectory($folder)
        $module=New-Windows11ReadinessCsvModule -CollectorPath $readinessPath
        try {
            $caught=$null
            try {Invoke-Windows11ReadinessCsvLoad -Module $module -Folder $folder -RunStartedAt $runStartedAt -ThresholdHours 24 -Names @('required.csv') -Required $true | Out-Null}
            catch {$caught=$_}
            Assert-Offline ($null -ne $caught -and $caught.Exception.Message -match 'Required CSV not found') 'An absent required CSV is no longer blocking.'
        } finally {Remove-Module $module -Force}
    }
    Test-OfflineCase 'Windows 11 readiness optional input remains non-blocking when absent' {
        $folder=Join-Path $freshnessRoot 'optional-absent'
        [void][IO.Directory]::CreateDirectory($folder)
        $module=New-Windows11ReadinessCsvModule -CollectorPath $readinessPath
        try {
            $observed=Invoke-Windows11ReadinessCsvLoad -Module $module -Folder $folder -RunStartedAt $runStartedAt -ThresholdHours 24 -Names @('optional.csv')
            Assert-Offline ($observed.Rows.Count -eq 0 -and $observed.WarningCount -eq 0) 'An absent optional CSV did not retain its empty, non-warning behavior.'
        } finally {Remove-Module $module -Force}
    }
    Test-OfflineCase 'Windows 11 readiness checks only the selected alternative input' {
        $folder=Join-Path $freshnessRoot 'alternatives'
        [void][IO.Directory]::CreateDirectory($folder)
        $preferred=Join-Path $folder 'preferred.csv'
        $fallback=Join-Path $folder 'fallback.csv'
        [IO.File]::WriteAllText($preferred,"Id,Value`r`n4,preferred`r`n")
        [IO.File]::WriteAllText($fallback,"Id,Value`r`n5,fallback`r`n")
        [IO.File]::SetLastWriteTimeUtc($preferred,$runStartedAt.AddHours(-1))
        [IO.File]::SetLastWriteTimeUtc($fallback,$runStartedAt.AddHours(-48))
        $module=New-Windows11ReadinessCsvModule -CollectorPath $readinessPath
        try {
            $preferredObserved=Invoke-Windows11ReadinessCsvLoad -Module $module -Folder $folder -RunStartedAt $runStartedAt -ThresholdHours 24 -Names @('preferred.csv','fallback.csv') -Required $true
            Assert-Offline ($preferredObserved.WarningCount -eq 0 -and $preferredObserved.Rows[0].Value -ceq 'preferred') 'The unselected stale fallback CSV was checked or imported.'
            Remove-Item -LiteralPath $preferred -Force
            $fallbackObserved=Invoke-Windows11ReadinessCsvLoad -Module $module -Folder $folder -RunStartedAt $runStartedAt -ThresholdHours 24 -Names @('preferred.csv','fallback.csv') -Required $true
            Assert-Offline ($fallbackObserved.WarningCount -eq 1 -and $fallbackObserved.WarningMessages[0] -match [regex]::Escape('fallback.csv') -and $fallbackObserved.Rows[0].Value -ceq 'fallback') 'The selected fallback CSV did not receive the freshness check.'
        } finally {Remove-Module $module -Force}
    }
    Test-OfflineCase 'Windows 11 readiness freshness check preserves imported schema and content' {
        $folder=Join-Path $freshnessRoot 'content-parity'
        [void][IO.Directory]::CreateDirectory($folder)
        $path=Join-Path $folder 'parity.csv'
        [IO.File]::WriteAllText($path,"Id,Value`r`n6,unchanged`r`n7,also-unchanged`r`n")
        $module=New-Windows11ReadinessCsvModule -CollectorPath $readinessPath
        try {
            [IO.File]::SetLastWriteTimeUtc($path,$runStartedAt.AddHours(-1))
            $fresh=Invoke-Windows11ReadinessCsvLoad -Module $module -Folder $folder -RunStartedAt $runStartedAt -ThresholdHours 24 -Names @('parity.csv') -Required $true
            [IO.File]::SetLastWriteTimeUtc($path,$runStartedAt.AddHours(-48))
            $stale=Invoke-Windows11ReadinessCsvLoad -Module $module -Folder $folder -RunStartedAt $runStartedAt -ThresholdHours 24 -Names @('parity.csv') -Required $true
            Assert-Offline (($fresh.Rows[0].PSObject.Properties.Name -join ',') -ceq 'Id,Value') 'The freshness path changed the projected input schema.'
            Assert-Offline (($fresh.Rows | ConvertTo-Json -Depth 5 -Compress) -ceq ($stale.Rows | ConvertTo-Json -Depth 5 -Compress)) 'Freshness changed imported row content.'
            Assert-Offline ($fresh.WarningCount -eq 0 -and $stale.WarningCount -eq 1) 'The parity fixture did not exercise both fresh and stale paths.'
        } finally {Remove-Module $module -Force}
    }

    Test-OfflineCase 'AD HealthCheck uses atomic current and append paths' {
        $path=Join-Path $SourceRoot 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-HealthCheck.ps1'
        $source=Get-Content -LiteralPath $path -Raw
        Assert-Offline ($source -match 'Add-SmartM365CsvRowsAtomically\s+-Data\s+\$all') 'AD history append is not atomic.'
        Assert-Offline ($source -match 'Write-SmartM365CsvAtomically\s+-Data\s+\$all\s+-Path\s+\$latestCsv') 'AD latest export is not atomic.'
        Assert-Offline ($source -notmatch 'Export-Csv\s+\$latestCsv') 'AD latest direct Export-Csv remains.'
        Assert-Offline ($source -match '-Encoding\s+utf8BOM') 'AD HealthCheck encoding contract changed.'
        Assert-Offline ($source.Contains('Health findings: Critical={0}; Warning={1}; NotMeasured={2}')) 'AD HealthCheck does not distinguish business findings and unmeasured checks in its log.'
    }
    Test-OfflineCase 'AD HealthCheck keeps unmeasured checks distinct from healthy checks' {
        $path=Join-Path $SourceRoot 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-HealthCheck.ps1'
        $definitions=@(Get-FunctionText -Path $path -Names @('Num','Add-Row','Worst','Rank'))
        $module=New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")))
        try {
            $observed=& $module {
                $script:Rows=[System.Collections.ArrayList]::new()
                $script:RunId='synthetic'
                $script:RunDateUtc='2026-09-22'
                Add-Row 'forest.test' 'domain.test' 'dc.test' 'Disk' 'FreePercent' 'NotMeasured' '' 'NotMeasured' '>=10' 'No T0 access' 1
                $unmeasuredWorst=Worst @($script:Rows.ToArray())
                Add-Row 'forest.test' 'domain.test' 'dc.test' 'DNS' 'Resolve' 'Warning' 0 'Failed' '>=1' 'Synthetic warning' 1
                [pscustomobject]@{Rows=@($script:Rows.ToArray());UnmeasuredWorst=$unmeasuredWorst;MixedWorst=(Worst @($script:Rows.ToArray()));RankNotMeasured=(Rank 'NotMeasured')}
            }
            Assert-Offline ($observed.Rows[0].Status -eq 'NotMeasured' -and $observed.Rows[0].TextValue -eq 'NotMeasured') 'An unmeasured check was marked healthy.'
            Assert-Offline ($observed.UnmeasuredWorst -eq 'NotMeasured' -and $observed.MixedWorst -eq 'Warning') 'Health ranking promoted unmeasured status or hid a real warning.'
            Assert-Offline ($observed.RankNotMeasured -gt ( & $module {Rank 'Warning'})) 'Unmeasured health ranking outranks a business warning.'
        }
        finally {Remove-Module $module -Force}
    }
    Test-OfflineCase 'AD HealthCheck distinguishes DFSR measurements from diagnostic failures' {
        $path=Join-Path $SourceRoot 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-HealthCheck.ps1'
        $definition=@(Get-FunctionText -Path $path -Names @('Get-DfsrBacklog'))[0]
        $module=New-Module -ScriptBlock ([scriptblock]::Create($definition))
        try {
            $observed=& $module {
                function script:dfsrdiag.exe { 'Backlog File Count : 42'; $global:LASTEXITCODE=0 }
                $measured=Get-DfsrBacklog 'src.test' 'dst.test'
                function script:dfsrdiag.exe { 'No Backlog'; $global:LASTEXITCODE=0 }
                $zero=Get-DfsrBacklog 'src.test' 'dst.test'
                function script:dfsrdiag.exe { 'Backlog File Count : 42'; $global:LASTEXITCODE=5 }
                $failed=Get-DfsrBacklog 'src.test' 'dst.test'
                function script:dfsrdiag.exe { 'synthetic unrecognized response'; $global:LASTEXITCODE=0 }
                $unrecognized=Get-DfsrBacklog 'src.test' 'dst.test'
                function script:Get-Command { param($Name,$ErrorAction) if($Name -eq 'dfsrdiag.exe'){return $null} }
                $missing=Get-DfsrBacklog 'src.test' 'dst.test'
                [pscustomobject]@{Measured=$measured;Zero=$zero;Failed=$failed;Unrecognized=$unrecognized;Missing=$missing}
            }
            Assert-Offline ($observed.Measured.Count -eq 42 -and $observed.Zero.Count -eq 0) 'Valid DFSR backlog responses were not measured.'
            Assert-Offline ($null -eq $observed.Failed.Count -and $observed.Failed.Reason -match 'DfsrdiagExitCode=5') 'A nonzero dfsrdiag exit code was treated as a measurement.'
            Assert-Offline ($null -eq $observed.Unrecognized.Count -and $observed.Unrecognized.Reason -match 'DfsrdiagOutputUnrecognized') 'Unrecognized DFSR output has no diagnosis.'
            Assert-Offline ($null -eq $observed.Missing.Count -and $observed.Missing.Reason -match 'DfsrdiagUnavailable') 'A missing dfsrdiag executable has no diagnosis.'
        }
        finally {Remove-Module $module -Force}
    }
    Test-OfflineCase 'Exchange infrastructure uses paired atomic publication' {
        $path=Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/ServersAndStorage/SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1'
        $text=@(Get-FunctionText -Path $path -Names @('Export-ServersAndStorageCsv'))[0]
        Assert-Offline ($text -match 'Publish-CoreSmartM365Csv[\s\S]+-LatestPath\s+\$latestPath') 'LatestPath is not part of the paired Core publication.'
        Assert-Offline ($text -notmatch '\bCopy-Item\b') 'Infrastructure latest path still uses direct Copy-Item.'
        Assert-Offline ($text -match "-Delimiter\s+';'" ) 'Semicolon delimiter contract changed.'
    }
    Test-OfflineCase 'Exchange executive summary renders ERROR status in red' {
        $path=Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/ServersAndStorage/SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1'
        $definitions=@(Get-FunctionText -Path $path -Names @('Format-HtmlValue','New-HtmlExecutiveSummary'))
        $module=New-Module -ScriptBlock ([scriptblock]::Create($definitions -join "`n"))
        $htmlPath=Join-Path $testRoot 'exchange-executive-summary.html'
        $summary=[pscustomobject]@{ExecutionDate='2026-09-21';RunId='synthetic';ExchangeServersCount=1;TotalLogicalProcessorCount=0;TotalMemoryGB=0;TotalDiskDriveCount=0;TotalDiskDriveSizeTB=0;ExchangeSchemaRangeUpper='N/A';ExchangeOrgObjectVersion='N/A'}
        $server=[pscustomobject]@{ExchangeServerName='SYNTHETIC-EXCHANGE';ServerRole='Mailbox';LogicalProcessorCount=$null;MemoryGB=$null;DiskDriveCount=0;DiskDriveTotalSizeGB=0;ComputeCollectionStatus='ERROR';DiskDriveCollectionStatus='OK';LogicalDiskCollectionStatus='OK';LowSpaceLogicalDiskCount=0}
        try {
            & $module {param($s,$r,$p) New-HtmlExecutiveSummary -Summary $s -PerServerSummary @($r) -ReadinessInventory @() -Path $p} $summary $server $htmlPath
            $html=[IO.File]::ReadAllText($htmlPath)
            Assert-Offline ($html -match "class='status status-error'\s+style='font-weight:700;color:#b91c1c;'[^>]*>ERROR</td>") 'Per-server ERROR status is not rendered with an inline red color.'
        }
        finally {Remove-Module $module -Force}
    }
    Test-OfflineCase 'AD full inventory blocks incomplete sequential domains' {
        $path=Join-Path $SourceRoot 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-Inventory.ps1'
        $source=Get-Content -LiteralPath $path -Raw
        foreach($label in @('OU','Computer','User','Group','Contact')){
            $failurePattern=[regex]::Escape("$label inventory failed for domain")+'[\s\S]{0,250}\bthrow\b'
            Assert-Offline ($source -match $failurePattern) "$label domain failure is still non-blocking."
        }
        $copyText=@(Get-FunctionText -Path $path -Names @('Copy-SmartM365AdFileWithRetry'))[0]
        Assert-Offline ($copyText -match 'Copy-SmartM365FileAtomically' -and $copyText -notmatch '\bCopy-Item\b') 'AD publication retry does not use atomic copy.'
        Assert-Offline ($source -match 'Add-SmartM365CsvRowsAtomically[\s\S]+AD daily summary snapshot') 'AD daily history append is not atomic.'
    }
    Test-OfflineCase 'AD enrichment generic list publication is atomic and PowerShell 7 safe' {
        $syntheticRows=New-Object System.Collections.Generic.List[object]
        [void]$syntheticRows.Add([pscustomobject]@{Id='synthetic'})
        $syntheticArray=$syntheticRows.ToArray()
        Assert-Offline ($syntheticArray.Count-eq1 -and $syntheticArray[0].Id-eq'synthetic') 'Generic object list ToArray conversion failed.'
        foreach($relative in @(
            'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-Enrichment.ps1',
            'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-UsersEnrichment.ps1'
        )){
            $source=Get-Content -LiteralPath (Join-Path $SourceRoot $relative) -Raw
            Assert-Offline ($source -match 'Write-SmartM365CsvAtomically\s+-Data\s+\$enrichedRows\.ToArray\(\)') "$relative does not publish a PowerShell 7-safe object array atomically."
            Assert-Offline ($source -notmatch '-Data\s+@\(\$enrichedRows\)') "$relative still uses the failing generic List[object] array-subexpression conversion."
            Assert-Offline ($source -notmatch '\$enrichedRows[\s\S]{0,80}\bExport-Csv\b') "$relative still exports directly."
        }
    }
    Test-OfflineCase 'Exchange local mailbox current copies are atomic and blocking' {
        $path=Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
        $publishText=@(Get-FunctionText -Path $path -Names @('Publish-SmartM365ExchangeLocalMailboxCsv'))[0]
        Assert-Offline ($publishText -match 'Copy-SmartM365FileAtomically') 'Mailbox DATA-LAST copy is not atomic.'
        Assert-Offline ($publishText -match "Failed to publish latest CSV copy[\s\S]+-Level 'ERROR'[\s\S]+throw") 'Mailbox DATA-LAST failure is still acknowledged as a warning.'
        $source=Get-Content -LiteralPath $path -Raw
        Assert-Offline ($source -notmatch 'Copy-Item\s+-LiteralPath\s+\$(dailyCsv|summaryCsv)') 'Mailbox daily or summary latest copy remains direct.'
    }
    Test-OfflineCase 'Calendar inventory stops and caches unavailable backends' {
        $path=Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/CalendarPermissions/SmartM365-Exchange-MailboxCalendarPermissions-Inventory.ps1'
        $definitions=@(Get-FunctionText -Path $path -Names @('Get-SmartM365CalendarFailureCategory','Get-SmartM365CalendarBackendName','New-SmartM365CalendarLookupException','Test-SmartM365CalendarBackendPreflight','Get-CalendarFoldersSafe'))
        $module=New-Module -ScriptBlock ([scriptblock]::Create($definitions -join "`n"))
        try {
            $observed=&$module {
                function script:Invoke-Quiet {param([scriptblock]$Script)&$Script}
                function script:Get-MailboxFolderStatistics {
                    param($Identity,$FolderScope,$ErrorAction)
                    $script:StatisticsCallCount++
                    throw "Cannot open mailbox on server 'SYNTHETIC-EX01' through cn=Servers/cn=SYNTHETIC-EX01/cn=Microsoft System Attendant because the Information Store is unavailable."
                }
                function script:Get-ExceptionMetadata {
                    param($ErrorRecord)
                    $exception=$ErrorRecord.Exception
                    [pscustomobject]@{
                        Category=[string]$exception.Data['SmartM365Category']
                        Attempts=[int]$exception.Data['SmartM365Attempts']
                        Backend=[string]$exception.Data['SmartM365Backend']
                    }
                }
                $script:UnavailableCalendarBackends=@{}
                $script:CalendarBackendPreflightResults=@{}
                $script:BackendPreflightTimeoutSeconds=0
                $script:StatisticsCallCount=0
                $mailbox=[pscustomobject]@{Identity='synthetic-one';UserPrincipalName='synthetic-one@example.test';PrimarySmtpAddress='synthetic-one@example.test';Guid=[guid]::Empty;ServerName='SYNTHETIC-EX01';Database='SYNTHETIC-DB01'}
                $first=$null
                try {Get-CalendarFoldersSafe -Mbx $mailbox -PrimaryOnly:$true|Out-Null}catch{$first=Get-ExceptionMetadata $_}
                $callsAfterFirst=$script:StatisticsCallCount
                $mailbox.Identity='synthetic-two';$mailbox.UserPrincipalName='synthetic-two@example.test';$mailbox.PrimarySmtpAddress='synthetic-two@example.test';$mailbox.ServerName='SYNTHETIC-EX02'
                $second=$null
                try {Get-CalendarFoldersSafe -Mbx $mailbox -PrimaryOnly:$true|Out-Null}catch{$second=Get-ExceptionMetadata $_}
                [pscustomobject]@{First=$first;Second=$second;CallsAfterFirst=$callsAfterFirst;TotalCalls=$script:StatisticsCallCount}
            }
            Assert-Offline ($observed.CallsAfterFirst -eq 1 -and $observed.TotalCalls -eq 1) 'Unavailable backend caused redundant identity or mailbox retries.'
            Assert-Offline ($observed.First.Category -ceq 'BackendUnavailable' -and $observed.First.Attempts -eq 1) 'First infrastructure failure metadata is incorrect.'
            Assert-Offline ($observed.Second.Category -ceq 'BackendPreviouslyUnavailable' -and $observed.Second.Attempts -eq 0) 'Cached backend failure did not skip the next mailbox.'
            Assert-Offline ($observed.First.Backend -ceq 'SYNTHETIC-EX01' -and $observed.Second.Backend -ceq 'SYNTHETIC-EX02') 'Backend name was not preserved in failure metadata.'
        } finally {Remove-Module $module -Force}
    }
    Test-OfflineCase 'Calendar inventory preflight bounds and caches unavailable databases' {
        $path=Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/CalendarPermissions/SmartM365-Exchange-MailboxCalendarPermissions-Inventory.ps1'
        $definitions=@(Get-FunctionText -Path $path -Names @('Get-SmartM365CalendarFailureCategory','Get-SmartM365CalendarBackendName','New-SmartM365CalendarLookupException','Test-SmartM365CalendarBackendPreflight','Get-CalendarFoldersSafe'))
        $module=New-Module -ScriptBlock ([scriptblock]::Create($definitions -join "`n"))
        try {
            $observed=&$module {
                function script:Test-HasCommand {param([string]$Name)return $true}
                function script:Invoke-Quiet {param([scriptblock]$Script)&$Script}
                function script:WriteLog {param($Message,$Level)}
                function script:Test-MAPIConnectivity {
                    param($Database,$Server,$PerConnectionTimeout,$AllConnectionsTimeout,$ErrorAction)
                    $script:PreflightCallCount++
                    $script:ObservedPerConnectionTimeout=$PerConnectionTimeout
                    $script:ObservedAllConnectionsTimeout=$AllConnectionsTimeout
                    [pscustomobject]@{Result='Failure';Error='Synthetic MAPI information store unavailable.'}
                }
                function script:Get-MailboxFolderStatistics {
                    param($Identity,$FolderScope,$ErrorAction)
                    $script:StatisticsCallCount++
                }
                function script:Get-ExceptionMetadata {
                    param($ErrorRecord)
                    $exception=$ErrorRecord.Exception
                    [pscustomobject]@{
                        Category=[string]$exception.Data['SmartM365Category']
                        Attempts=[int]$exception.Data['SmartM365Attempts']
                    }
                }
                $script:UnavailableCalendarBackends=@{}
                $script:CalendarBackendPreflightResults=@{}
                $script:BackendPreflightTimeoutSeconds=17
                $script:PreflightCallCount=0
                $script:StatisticsCallCount=0
                $mailbox=[pscustomobject]@{Identity='synthetic-one';UserPrincipalName='synthetic-one@example.test';PrimarySmtpAddress='synthetic-one@example.test';Guid=[guid]::Empty;ServerName='SYNTHETIC-EX01';Database='SYNTHETIC-DB01'}
                $first=$null
                try {Get-CalendarFoldersSafe -Mbx $mailbox -PrimaryOnly:$true|Out-Null}catch{$first=Get-ExceptionMetadata $_}
                $mailbox.Identity='synthetic-two';$mailbox.UserPrincipalName='synthetic-two@example.test';$mailbox.PrimarySmtpAddress='synthetic-two@example.test'
                $second=$null
                try {Get-CalendarFoldersSafe -Mbx $mailbox -PrimaryOnly:$true|Out-Null}catch{$second=Get-ExceptionMetadata $_}
                [pscustomobject]@{
                    First=$first
                    Second=$second
                    PreflightCalls=$script:PreflightCallCount
                    StatisticsCalls=$script:StatisticsCallCount
                    PerConnectionTimeout=$script:ObservedPerConnectionTimeout
                    AllConnectionsTimeout=$script:ObservedAllConnectionsTimeout
                }
            }
            Assert-Offline ($observed.PreflightCalls -eq 1 -and $observed.StatisticsCalls -eq 0) 'Unavailable database was not stopped and cached by the preflight.'
            Assert-Offline ($observed.PerConnectionTimeout -eq 17 -and $observed.AllConnectionsTimeout -eq 17) 'Configured MAPI preflight timeout was not applied.'
            Assert-Offline ($observed.First.Category -ceq 'BackendUnavailable' -and $observed.First.Attempts -eq 0) 'Preflight failure metadata is incorrect.'
            Assert-Offline ($observed.Second.Category -ceq 'BackendPreviouslyUnavailable' -and $observed.Second.Attempts -eq 0) 'Cached preflight failure did not skip the next mailbox.'
        } finally {Remove-Module $module -Force}
    }
    Test-OfflineCase 'Calendar weekly history is published once as a validated group' {
        $path=Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/CalendarPermissions/SmartM365-Exchange-MailboxCalendarPermissions-Inventory.ps1'
        $source=Get-Content -LiteralPath $path -Raw
        $tokens=$null;$parseErrors=$null
        $ast=[Management.Automation.Language.Parser]::ParseInput($source,[ref]$tokens,[ref]$parseErrors)
        if($parseErrors.Count){throw 'Calendar source parse failed.'}
        $exportCommands=@($ast.FindAll({param($node)$node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'ExportAndCopyCsv'},$true))
        Assert-Offline ($exportCommands.Count -eq 2 -and @($exportCommands|Where-Object{$_.Extent.Text -notmatch '-SkipWeeklyHistory'}).Count -eq 0) 'One or more calendar exports still publish WeeklyHistory independently.'
        $publishCommands=@($ast.FindAll({param($node)$node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Publish-SmartM365CalendarWeeklyHistory'},$true))
        Assert-Offline ($publishCommands.Count -eq 1) 'Calendar WeeklyHistory group is not published exactly once.'

        $definition=@(Get-FunctionText -Path $path -Names @('Publish-SmartM365CalendarWeeklyHistory'))[0]
        $module=New-Module -ScriptBlock ([scriptblock]::Create($definition))
        $mainPath=Join-Path $testRoot 'calendar-main.csv'
        $errorPath=Join-Path $testRoot 'calendar-errors.csv'
        Set-Content -LiteralPath $mainPath -Value 'Mailbox' -Encoding UTF8
        Set-Content -LiteralPath $errorPath -Value 'Mailbox' -Encoding UTF8
        try {
            $observed=&$module {
                param($main,$errors,$root)
                function script:Test-SmartM365MaxItemsMode {return $false}
                function script:ConvertTo-SmartM365ConfigBoolean {param($Value,[bool]$DefaultValue)if($Value -is [bool]){return $Value};return $DefaultValue}
                function script:Get-ScriptLocalConfigValue {
                    param($Config,[string]$Name,$DefaultValue)
                    $property=$Config.PSObject.Properties[$Name]
                    if($property){return $property.Value}
                    return $DefaultValue
                }
                function script:WriteLog {param($Message,$Level)}
                function script:Add-SmartM365WeeklyHistory {
                    param([string[]]$SourceCsvPaths,[string]$HistoryRootPath,[int]$RetentionWeeks,[string]$HistoryLabel,[switch]$UploadChangedFilesOnly)
                    $script:PublishCount++
                    $script:Published=[pscustomobject]@{Sources=@($SourceCsvPaths);Root=$HistoryRootPath;Retention=$RetentionWeeks;ChangedOnly=[bool]$UploadChangedFilesOnly}
                }
                $script:PublishCount=0
                $config=[pscustomobject]@{EnableWeeklyHistory=$true;WeeklyHistoryFolderPath=(Join-Path $root 'weekly');WeeklyHistoryRetentionWeeks=8}
                Publish-SmartM365CalendarWeeklyHistory -SourceCsvPaths @($main,$errors,$main) -Config $config -FallbackRootPath $root
                [pscustomobject]@{Count=$script:PublishCount;Published=$script:Published}
            } $mainPath $errorPath $testRoot
            Assert-Offline ($observed.Count -eq 1 -and $observed.Published.Sources.Count -eq 2) 'Calendar WeeklyHistory helper did not group and de-duplicate both CSVs.'
            Assert-Offline ($observed.Published.Retention -eq 8 -and $observed.Published.ChangedOnly) 'Calendar WeeklyHistory configuration or changed-file upload mode was not preserved.'
        } finally {Remove-Module $module -Force}
    }
    Test-OfflineCase 'Exchange local mailbox issues are structured and published' {
        $path=Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
        $definition=@(Get-FunctionText -Path $path -Names @('Add-SmartM365LocalMailboxIssue'))[0]
        $module=New-Module -ScriptBlock ([scriptblock]::Create($definition))
        try {
            $issue=&$module {
                $script:LocalMailboxIssues=New-Object 'System.Collections.Generic.List[object]'
                $script:LocalMailboxIssueSequence=0
                Add-SmartM365LocalMailboxIssue -Category 'SyntheticFailure' -Operation 'SyntheticOperation' -MailboxIdentity 'synthetic-mailbox' -Message 'Synthetic message'
            }
            Assert-Offline ($issue.IssueIndex -eq 1 -and $issue.Category -ceq 'SyntheticFailure' -and $issue.Operation -ceq 'SyntheticOperation' -and $issue.Message -ceq 'Synthetic message') 'Structured mailbox issue fields changed.'
        } finally {Remove-Module $module -Force}
        $source=Get-Content -LiteralPath $path -Raw
        Assert-Offline ($source -match 'function\s+Export-SmartM365LocalMailboxIssues[\s\S]+Export-SmartM365Csv') 'Mailbox issue CSV does not use the shared publisher.'
        Assert-Offline ($source -match "-Operation 'Get-MailboxStatistics'" -and $source -match "-Operation 'Get-MobileDevice'" -and $source -match "-Operation 'Get-RemoteMailbox'") 'One or more required mailbox issue sources are not recorded.'
        Assert-Offline ($source -match 'Export-SmartM365LocalMailboxIssues') 'Mailbox issue export is not invoked.'
    }
    Test-OfflineCase 'Mailbox inventory summary sends at most once per local day' {
        $path=Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
        $definition=@(Get-FunctionText -Path $path -Names @('Invoke-SmartM365MailboxDailySummaryMail'))[0]
        $module=New-Module -ScriptBlock ([scriptblock]::Create($definition))
        $marker=Join-Path $testRoot 'mailbox-summary.sent'
        try {
            &$module { function script:WriteLog { param($Message,$Level) }; $script:SendCount=0 }
            $first=&$module {param($p)Invoke-SmartM365MailboxDailySummaryMail -MarkerPath $p -SendAction {$script:SendCount++}} $marker
            $second=&$module {param($p)Invoke-SmartM365MailboxDailySummaryMail -MarkerPath $p -SendAction {$script:SendCount++}} $marker
            $count=&$module {$script:SendCount}
            Assert-Offline ($first -and -not $second -and $count -eq 1) 'Mailbox summary was sent more than once on the same day.'
            Assert-Offline ((Get-Content -LiteralPath $marker -Raw).Trim() -eq (Get-Date).ToString('yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)) 'Mailbox summary marker does not store the local send date.'
        } finally {Remove-Module $module -Force}
    }
    Test-OfflineCase 'Exchange warnings and alert mail subjects retain their operational contracts' {
        $exchange=Get-Content -LiteralPath (Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/ServersAndStorage/SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1') -Raw
        $mailbox=Get-Content -LiteralPath (Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1') -Raw
        $winUpdate=Get-Content -LiteralPath (Join-Path $SourceRoot 'SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-WinUpdate_Status_From_Intune.ps1') -Raw
        $adHealth=Get-Content -LiteralPath (Join-Path $SourceRoot 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-HealthCheck.ps1') -Raw
        $exchangeCompletionLine = '$completionStatus = if ($script:ServersAndStorageWarningCount -gt 0) { ''CompletedWithWarnings'' } else { ''Success'' }'
        Assert-Offline ([regex]::Matches($exchange,[regex]::Escape($exchangeCompletionLine)).Count -eq 2) 'Exchange completion does not propagate collector warnings.'
        Assert-Offline ($exchange -match '\$reportedWarningCount\s*=\s+\$exchangeReadinessWarningRows\.Count\s*\+\s*\$lowSpaceRows\.Count') 'Exchange readiness and capacity warnings are not included in completion reporting.'
        $warningExitPattern = [regex]::Escape('$script:CompletionStatus = $completionStatus') + '[\s\S]{0,180}' + [regex]::Escape('exit 3')
        Assert-Offline ([regex]::Matches($exchange,$warningExitPattern).Count -eq 2) 'Exchange warnings do not return the orchestrator warning exit code.'
        Assert-Offline ($mailbox -match 'Invoke-SmartM365MailboxDailySummaryMail\s+-MarkerPath') 'Mailbox summary email is not protected by the daily guard.'
        Assert-Offline (($winUpdate | Select-String -Pattern 'SMART365 - \[\$reportStatus\] WinUpdate Feature Update' -AllMatches).Matches.Count -eq 2) 'WinUpdate alert subject lacks the SMART365 prefix.'
        Assert-Offline (($adHealth | Select-String -Pattern 'SMART365 - \[\$\(\$worst\.ToUpperInvariant\(\)\)\] Active Directory Health Check' -AllMatches).Matches.Count -eq 1) 'AD Health alert subject lacks the SMART365 prefix.'
        Assert-Offline ($adHealth -match 'SMART365 - \[CRITICAL\] Active Directory Health Check failed') 'AD Health failure subject lacks the SMART365 prefix.'
    }
    Test-OfflineCase 'Weekly CSV histories and manifests use atomic primitives' {
        $teams=Get-Content -LiteralPath (Join-Path $SourceRoot 'SmartInventory/M365Inventory/Teams/SmartM365-Teams-Inventory.ps1') -Raw
        $spo=Get-Content -LiteralPath (Join-Path $SourceRoot 'SmartInventory/M365Inventory/SharePoint/SmartM365-SPO-Inventory.ps1') -Raw
        $licenses=Get-Content -LiteralPath (Join-Path $SourceRoot 'SmartInventory/M365Inventory/Licensing/SmartM365-Licences-Inventory.ps1') -Raw
        Assert-Offline ($teams -match 'Add-SmartM365CsvRowsAtomically[\s\S]{0,200}-Encoding\s+utf8BOM') 'Teams append history is not atomic or lost its BOM contract.'
        Assert-Offline ($spo -match 'Add-SmartM365CsvRowsAtomically[\s\S]{0,200}-Encoding\s+utf8BOM') 'SharePoint append history is not atomic or lost its BOM contract.'
        Assert-Offline ($licenses -match 'Write-SmartM365(?:TextAtomically|JsonBytesAtomically)\s+-Path\s+\$manifestPath') 'Licensing weekly manifest is not atomic.'
    }
    Test-OfflineCase 'Entra empty canonical current copy is atomic' {
        $source=Get-Content -LiteralPath (Join-Path $SourceRoot 'SmartInventory/M365Inventory/Devices/SmartM365-EntraDevices-Inventory.ps1') -Raw
        Assert-Offline ($source -match 'Copy-SmartM365FileAtomically\s+-SourcePath\s+\$emptyPendingExport\.TimestampedPath\s+-DestinationPath\s+\$currentPendingPath') 'Entra header-only current file is not promoted atomically.'
    }
}
finally {
    $resolved=[IO.Path]::GetFullPath($testRoot)
    $expectedParent=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if(-not $resolved.StartsWith($expectedParent,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'SmartInventory-NonGraph-*'){throw 'Unsafe cleanup root.'}
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
if($ResultPath){$results|ConvertTo-Json -Depth 5|Set-Content -LiteralPath $ResultPath -Encoding UTF8}
$results|Format-Table Name,Passed,Error -AutoSize
$failed=@($results|Where-Object{-not $_.Passed}).Count
Write-Output ("Cases={0}; Passed={1}; Failed={2}; PowerShell={3}" -f $results.Count,($results.Count-$failed),$failed,$PSVersionTable.PSVersion)
if($failed){exit 1}
