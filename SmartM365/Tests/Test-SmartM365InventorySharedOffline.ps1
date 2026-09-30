<#
.SYNOPSIS
Synthetic regression tests for shared inventory identity and atomic persistence.
.VERSION
1.1.9
#>
[CmdletBinding()]
param(
    [string]$SourceRoot,
    [string]$ResultPath
)
$ErrorActionPreference = 'Stop'
if (-not $SourceRoot) { $SourceRoot = Split-Path $PSScriptRoot -Parent }
$results = New-Object 'System.Collections.Generic.List[object]'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-Offline-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)

function Assert-Offline {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}
function Test-OfflineCase {
    param([string]$Name, [scriptblock]$Body)
    try { & $Body; $results.Add([pscustomobject]@{ Name = $Name; Passed = $true; Error = '' }) }
    catch { $results.Add([pscustomobject]@{ Name = $Name; Passed = $false; Error = $_.Exception.Message }) }
}
function Import-OfflineFunctions {
    param([string]$Path, [string[]]$Names)
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw "Source parse failed: $Path" }
    $definitions = foreach ($name in $Names) {
        $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        if ($null -eq $node) { throw "Function not found: $name" }
        $node.Extent.Text
    }
    # Only the named function definitions run. No module initializers, tenant loader,
    # collector entry point, credentials, task scheduler, mail or network calls.
    New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")))
}
# Weekly history and orchestrator distributed state go through the JSON transport (.json.txt
# names, owned atomic writes). Load it with its deployment policy, as the production modules do.
Import-Module (Join-Path $SourceRoot 'Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Force -Global

try {
    $tenantContextModule = Import-OfflineFunctions (Join-Path $SourceRoot 'Config/SmartM365-TenantContext.ps1') @('Initialize-SmartM365TenantContext')
    try {
        & $tenantContextModule {
            function script:Write-SmartM365StartupBanner {}
            function script:Get-SmartM365EffectiveGlobalConfig {
                param($StartPath, $ProfileKey)
                return $script:SyntheticEffectiveConfig
            }
        }
        Test-OfflineCase 'Tenant context accepts PSCustomObject without MailClientName' {
            & $tenantContextModule {
                $script:SyntheticEffectiveConfig = [pscustomobject]@{
                    ProfileKey = 'prod'; OrganizationKey = 'contoso'; EnvironmentKey = 'prod'
                    TenantKey = 'contoso-prod'; TenantId = '00000000-0000-0000-0000-000000000001'
                }
                Initialize-SmartM365TenantContext -Tenant prod | Out-Null
            }
            Assert-Offline ($global:SmartM365MailTenantName -ceq 'CONTOSO') 'PSCustomObject tenant context did not use the organization fallback.'
        }
        Test-OfflineCase 'Tenant context accepts dictionary MailClientName override' {
            & $tenantContextModule {
                $script:SyntheticEffectiveConfig = @{
                    ProfileKey = 'prod'; OrganizationKey = 'contoso'; EnvironmentKey = 'prod'
                    TenantKey = 'contoso-prod'; TenantId = '00000000-0000-0000-0000-000000000001'
                    MailClientName = 'CONTOSO France'
                }
                Initialize-SmartM365TenantContext -Tenant prod | Out-Null
            }
            Assert-Offline ($global:SmartM365MailTenantName -ceq 'CONTOSO France') 'Dictionary tenant context ignored MailClientName.'
        }
    }
    finally {
        Remove-Module $tenantContextModule -Force
    }

    $completionModule = Import-OfflineFunctions (Join-Path $SourceRoot 'Modules/SmartM365.Core/SmartM365.Core.psm1') @('Complete-SmartM365ExecutionContext')
    try {
        Test-OfflineCase 'Successful execution omits stale failure stage' {
            $observed = & $completionModule {
                function WriteLog {
                    param([string]$Message,[string]$Level)
                    $script:SummaryLines.Add($Message) | Out-Null
                    if ($Level -eq 'ERROR') { $global:SmartM365ErrorCount++ }
                }
                function Write-SmartM365CompletionBanner { param($Status,$ScriptName,$StartedAt,$EndedAt,$WarningCount,$ErrorCount,$GeneratedCsvFiles,$LogPath) }
                $script:SummaryLines = [System.Collections.Generic.List[string]]::new()
                $global:SmartM365ExecutionSummaryWritten = $false
                $global:SmartM365WarningCount = 0
                $global:SmartM365ErrorCount = 0
                $global:LogTextFile = ''
                $global:logTranscriptFile = ''
                $global:SmartM365SharePointUploadedFiles = $null
                $global:SmartM365MailHtmlFiles = $null
                Complete-SmartM365ExecutionContext -Status Success -FailureStage 'PreviousPhase'
                [pscustomobject]@{ Lines=@($script:SummaryLines); Errors=$global:SmartM365ErrorCount }
            }
            Assert-Offline (-not (@($observed.Lines) -match 'FailureStage:').Count) 'Successful summary retained FailureStage.'
            Assert-Offline ($observed.Errors -eq 0) 'Successful completion generated a false error.'
        }
        Test-OfflineCase 'Failed execution logs and counts an otherwise silent error' {
            $observed = & $completionModule {
                function WriteLog {
                    param([string]$Message,[string]$Level)
                    $script:SummaryLines.Add($Message) | Out-Null
                    if ($Level -eq 'ERROR') { $global:SmartM365ErrorCount++ }
                }
                function Write-SmartM365CompletionBanner { param($Status,$ScriptName,$StartedAt,$EndedAt,$WarningCount,$ErrorCount,$GeneratedCsvFiles,$LogPath) }
                $script:SummaryLines = [System.Collections.Generic.List[string]]::new()
                $global:SmartM365ExecutionSummaryWritten = $false
                $global:SmartM365WarningCount = 0
                $global:SmartM365ErrorCount = 0
                Complete-SmartM365ExecutionContext -Status Failed -FailureStage 'SyntheticPhase' -ErrorRecord ([pscustomobject]@{Exception=[Exception]::new('Synthetic failure')})
                [pscustomobject]@{ Lines=@($script:SummaryLines); Errors=$global:SmartM365ErrorCount }
            }
            Assert-Offline ($observed.Errors -eq 1) 'Failed completion did not increment Errors exactly once.'
            Assert-Offline ((@($observed.Lines) -match 'FailureStage: SyntheticPhase').Count -eq 1) 'Failed summary lost FailureStage.'
            Assert-Offline ((@($observed.Lines) -match 'Execution failed during SyntheticPhase').Count -eq 1) 'Failed completion did not write the missing ERROR entry.'
            Assert-Offline ((@($observed.Lines) -match 'Synthetic failure').Count -ge 1) 'The failure diagnostic was omitted from the log.'
        }
    }
    finally { Remove-Module $completionModule -Force }

    $weeklyModule = Import-OfflineFunctions (Join-Path $SourceRoot 'Modules/SmartM365.Core/SmartM365.Core.psm1') @('Save-SmartM365WeeklyInventoryHistory')
    try {
        & $weeklyModule {
            function script:Get-SmartM365IsoWeekName { '2026-W39' }
            function script:Get-SmartM365WeeklyHistoryFileName { param($Path) [IO.Path]::GetFileName($Path) }
            function script:Copy-SmartM365FileAtomically { param($SourcePath,$DestinationPath) Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force }
            function script:Write-SmartM365TextAtomically { param($Path,$Content,$Encoding) Set-Content -LiteralPath $Path -Value $Content -Encoding utf8 }
            function script:Invoke-SmartM365SharePointCsvUpload { param($LocalFilePath) $script:Uploads++; $script:UploadedPaths += $LocalFilePath }
            function script:WriteLog { param($Message,$Level) if ($Level -eq 'WARNING') { $script:Warnings++ } }
            $script:Uploads=0
            $script:UploadedPaths=@()
            $script:Warnings=0
        }
        Test-OfflineCase 'Weekly snapshot keeps first CSV and manifest capture time on rerun' {
            $root=Join-Path $testRoot 'weekly-fresh'
            $source=Join-Path $testRoot 'weekly-source.csv'
            Set-Content -LiteralPath $source -Value 'Id,Value`n1,first' -Encoding utf8
            & $weeklyModule {param($s,$r)Save-SmartM365WeeklyInventoryHistory -SourceFiles @($s) -HistoryRootPath $r -UploadChangedFilesOnly -RetentionWeeks 0} $source $root
            $snapshot=Join-Path $root '2026-W39/weekly-source.csv'
            $manifestPath=Get-SmartM365JsonReadPath (Join-Path $root '2026-W39/manifest.json')
            $first=(Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json)
            $firstHash=(Get-FileHash -LiteralPath $manifestPath).Hash
            Set-Content -LiteralPath $source -Value 'Id,Value`n1,second' -Encoding utf8
            # Preferred manifests are republished on every run by design; unchanged CSVs must not be.
            $csvUploadsBefore=& $weeklyModule {@($script:UploadedPaths | Where-Object { $_ -like '*.csv' }).Count}
            & $weeklyModule {param($s,$r)Save-SmartM365WeeklyInventoryHistory -SourceFiles @($s) -HistoryRootPath $r -UploadChangedFilesOnly -RetentionWeeks 0} $source $root
            Assert-Offline ((Get-Content -LiteralPath $snapshot -Raw) -match 'first') 'Second run replaced the first weekly CSV.'
            Assert-Offline ((Get-FileHash -LiteralPath $manifestPath).Hash -eq $firstHash) 'Skipped snapshot rewrote manifest provenance.'
            Assert-Offline ((& $weeklyModule {@($script:UploadedPaths | Where-Object { $_ -like '*.csv' }).Count}) -eq $csvUploadsBefore) 'Skipped snapshot reuploaded unchanged history CSV.'
            Assert-Offline ($first.SnapshotTimestampStatus -eq 'Recorded' -and $first.SnapshotCreatedAtUtc -and $first.FileSnapshotCreatedAtUtc.'weekly-source.csv') 'New snapshot lacks explicit provenance.'
        }
        Test-OfflineCase 'Legacy weekly snapshot remains untouched and its capture time stays unknown' {
            $root=Join-Path $testRoot 'weekly-legacy'
            $week=Join-Path $root '2026-W39'
            New-Item -Path $week -ItemType Directory -Force | Out-Null
            $source=Join-Path $testRoot 'weekly-legacy.csv'
            Set-Content -LiteralPath $source -Value 'Id,Value`n1,new' -Encoding utf8
            $snapshot=Join-Path $week 'weekly-legacy.csv'
            Set-Content -LiteralPath $snapshot -Value 'Id,Value`n1,old' -Encoding utf8
            $manifestPath=Join-Path $week 'manifest.json'
            # Legacy manifest under the current contract: identity fields present, no recorded capture time.
            $legacyManifest=[ordered]@{UpdatedAt='2026-09-22T00:00:00Z';Week='2026-W39';HistoryLabel='SmartM365 inventory';HistoryRootPath=$root;Files=@('weekly-legacy.csv')}
            Set-Content -LiteralPath $manifestPath -Value ($legacyManifest | ConvertTo-Json) -Encoding utf8
            $firstHash=(Get-FileHash -LiteralPath $manifestPath).Hash
            $warningsBefore=& $weeklyModule {$script:Warnings}
            & $weeklyModule {param($s,$r)Save-SmartM365WeeklyInventoryHistory -SourceFiles @($s) -HistoryRootPath $r -UploadChangedFilesOnly -RetentionWeeks 0} $source $root
            # The .json.txt transition renames the legacy manifest byte for byte; its content must not change.
            Assert-Offline ((Get-FileHash -LiteralPath (Get-SmartM365JsonReadPath $manifestPath)).Hash -eq $firstHash) 'Legacy manifest was mutated on a skipped run.'
            Assert-Offline ((Get-Content -LiteralPath $snapshot -Raw) -match 'old') 'Legacy weekly CSV was overwritten.'
            Assert-Offline ((& $weeklyModule {$script:Warnings}) -gt $warningsBefore) 'Unknown legacy capture time was not reported.'
        }
        Test-OfflineCase 'Weekly CSV without manifest never inherits a later capture time' {
            $root=Join-Path $testRoot 'weekly-no-manifest'
            $week=Join-Path $root '2026-W39'
            New-Item -Path $week -ItemType Directory -Force | Out-Null
            $source=Join-Path $testRoot 'weekly-no-manifest.csv'
            Set-Content -LiteralPath $source -Value 'Id,Value`n1,new' -Encoding utf8
            $snapshot=Join-Path $week 'weekly-no-manifest.csv'
            Set-Content -LiteralPath $snapshot -Value 'Id,Value`n1,old' -Encoding utf8
            & $weeklyModule {param($s,$r)Save-SmartM365WeeklyInventoryHistory -SourceFiles @($s) -HistoryRootPath $r -UploadChangedFilesOnly -RetentionWeeks 0} $source $root
            $manifest=Get-Content -LiteralPath (Get-SmartM365JsonReadPath (Join-Path $week 'manifest.json')) -Raw | ConvertFrom-Json
            Assert-Offline ($manifest.SnapshotTimestampStatus -eq 'UnknownLegacy' -and -not $manifest.SnapshotCreatedAtUtc) 'Existing CSV without manifest was assigned an invented capture date.'
            Assert-Offline ((Get-Content -LiteralPath $snapshot -Raw) -match 'old') 'Existing manifest-less CSV was overwritten.'
        }
    }
    finally { Remove-Module $weeklyModule -Force }

    foreach ($variant in @('Core', 'WindowsPowerShell5')) {
        $relative = if ($variant -eq 'Core') { 'Modules/SmartM365.Core/SmartM365.Core.psm1' } else { 'Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1' }
        $module = Import-OfflineFunctions (Join-Path $SourceRoot $relative) @(
            'Add-SmartM365TenantKeyToCsvData', 'Write-SmartM365CsvAtomically', 'Add-SmartM365CsvRowsAtomically',
            'Copy-SmartM365FileAtomically', 'Write-SmartM365TextAtomically',
            'Get-SmartM365CoreContextValue', 'Assert-SmartM365CsvDataCompleteness',
            'Get-SmartM365CsvValidationRule', 'Get-SmartM365CsvValidationBaseName',
            'Publish-SmartM365Csv', 'Export-SmartM365Csv', 'Get-SmartM365MaxItemsValue', 'Test-SmartM365MaxItemsMode',
            'Get-SmartM365MaxItemsSuffix', 'Add-SmartM365MaxItemsSuffixToCsvPath',
            'Add-SmartM365MaxItemsSuffixToBaseName', 'Limit-SmartM365RowsForMaxItems',
            'Get-SmartM365MailTenantName', 'Format-SmartM365MailSubject',
            'Get-SmartM365ScriptVersionFromFile', 'Get-SmartM365MailScriptContext', 'Add-SmartM365MailExecutionFooter',
            'ExportAndCopyCsv', 'ExportAndCopyCsvFromConvert'
        )
        if ($variant -eq 'Core') {
            Remove-Module $module
            $module = Import-OfflineFunctions (Join-Path $SourceRoot $relative) @(
                'Add-SmartM365TenantKeyToCsvData', 'Write-SmartM365CsvAtomically', 'Add-SmartM365CsvRowsAtomically',
                'Copy-SmartM365FileAtomically', 'Write-SmartM365TextAtomically',
                'Get-SmartM365CoreContextValue', 'Assert-SmartM365CsvDataCompleteness',
                'Get-SmartM365CsvValidationRule', 'Get-SmartM365CsvValidationBaseName',
                'Publish-SmartM365Csv', 'Export-SmartM365Csv', 'Get-SmartM365MaxItemsValue', 'Test-SmartM365MaxItemsMode',
                'Get-SmartM365MaxItemsSuffix', 'Add-SmartM365MaxItemsSuffixToCsvPath',
                'Add-SmartM365MaxItemsSuffixToBaseName', 'Limit-SmartM365RowsForMaxItems',
                'Get-SmartM365MailTenantName', 'Format-SmartM365MailSubject',
                'Get-SmartM365ScriptVersionFromFile', 'Get-SmartM365MailScriptContext', 'Add-SmartM365MailExecutionFooter',
                'Export-SmartM365CsvStreamAtomically', 'ExportAndCopyCsv', 'ExportAndCopyCsvFromConvert'
            )
        }
        & $module {
            function script:WriteLog { param($Message, $Level) }
            function script:Invoke-SmartM365SharePointCsvUpload { throw 'Unexpected SharePoint call in offline test.' }
            function script:Invoke-SmartM365WeeklyInventoryHistoryForCsv { param($SourceFiles,$TimestampedPath) }
            function script:Start-Sleep { param($Seconds) }
            $script:SmartM365CoreTenantKey = 'synthetic-a'
            $script:SmartM365CoreOrganizationKey = 'synthetic-org'
            $script:SmartM365CoreEnvironmentKey = 'test'
            $script:SmartM365CoreTenantId = '00000000-0000-0000-0000-000000000001'
            $global:SmartM365MailTenantName = 'CONTOSO'
            $global:SmartM365ScriptFileName = 'SmartM365-Synthetic-Inventory.ps1'
            $global:SmartM365ScriptVersion = '9.8.7'
        }
        $loggerModule = Import-OfflineFunctions (Join-Path $SourceRoot $relative) @('Format-SmartM365LogLine','WriteLog')
        try {
            & $loggerModule {
                function script:Invoke-SmartM365TeamsNotificationFromLog {
                    param($Message,$Level)
                    $script:LastTeamsLevel = $Level
                }
            }
            Test-OfflineCase "$variant WARN level is normalized and counted" {
                $logPath = Join-Path $testRoot "$variant-warn.log"
                $observed = & $loggerModule {
                    param($Path)
                    $global:SmartM365WarningCount = 0
                    $global:SmartM365ErrorCount = 0
                    $global:LogTextFile = $Path
                    $script:LastTeamsLevel = ''
                    WriteLog -Message 'Synthetic warning' -Level 'WARN'
                    [pscustomobject]@{
                        WarningCount = $global:SmartM365WarningCount
                        ErrorCount   = $global:SmartM365ErrorCount
                        TeamsLevel   = $script:LastTeamsLevel
                        LogLine      = [string](Get-Content -LiteralPath $Path -Tail 1)
                    }
                } $logPath
                Assert-Offline ($observed.WarningCount -eq 1 -and $observed.ErrorCount -eq 0) 'WARN did not increment the warning counter exactly once.'
                Assert-Offline ($observed.TeamsLevel -ceq 'WARNING') 'WARN was not normalized before the Teams notification hook.'
                Assert-Offline ($observed.LogLine -match '\[WARNING\] Synthetic warning$') 'WARN was not persisted with the WARNING level.'
            }
        }
        finally {
            $global:LogTextFile = $null
            $global:SmartM365WarningCount = 0
            $global:SmartM365ErrorCount = 0
            Remove-Module $loggerModule -Force
        }
        Test-OfflineCase "$variant mail subjects use the tenant prefix" {
            $general = & $module { Format-SmartM365MailSubject -Subject 'SMART 365 - [CRITICAL] Microsoft Teams Inventory - 2026-09-21T12:12:41Z' }
            $legacy = & $module { Format-SmartM365MailSubject -Subject '[SmartM365] EXO quarantine report - ERROR' }
            $limited = & $module { Format-SmartM365MailSubject -Subject '[MAXITEMS-5 TEST] SMART365 - [WARNING] WinUpdate Feature Update' }
            Assert-Offline ($general -ceq '[SMART 365] - [CONTOSO] - [CRITICAL] Microsoft Teams Inventory - 2026-09-21T12:12:41Z') 'General subject prefix normalization changed.'
            Assert-Offline ($legacy -ceq '[SMART 365] - [CONTOSO] - EXO quarantine report - ERROR') 'Legacy SmartM365 subject was not normalized.'
            Assert-Offline ($limited -ceq '[SMART 365] - [CONTOSO] - [MAXITEMS-5 TEST] - [WARNING] WinUpdate Feature Update') 'MAXITEMS subject no longer begins with the tenant prefix.'
        }
        Test-OfflineCase "$variant orchestrator mail subject uses its dedicated prefix" {
            $subject = & $module { Format-SmartM365MailSubject -Subject '[SmartM365 Orchestrator][prod] Runtime update detected: 1.5.14 -> 1.5.15' -Orchestrator }
            $critical = & $module { Format-SmartM365MailSubject -Subject '[CRITICAL][SmartM365 Orchestrator][prod] Authenticode rejected job JobA' -Orchestrator }
            $again = & $module { param($s) Format-SmartM365MailSubject -Subject $s } $subject
            Assert-Offline ($subject -ceq '[SMART 365] - [CONTOSO] - [ Orchestrator] - Runtime update detected: 1.5.14 -> 1.5.15') 'Orchestrator subject prefix normalization changed.'
            Assert-Offline ($critical -ceq '[SMART 365] - [CONTOSO] - [ Orchestrator] - [CRITICAL] Authenticode rejected job JobA') 'Critical orchestrator subject was not normalized.'
            Assert-Offline ($again -ceq $subject) 'Subject normalization is not idempotent.'
        }
        Test-OfflineCase "$variant mail footer contains script name and version once" {
            $html = & $module { Add-SmartM365MailExecutionFooter -BodyHtml '<html><body><p>Fixture</p></body></html>' }
            $again = & $module { param($body) Add-SmartM365MailExecutionFooter -BodyHtml $body } $html
            Assert-Offline ($html -match 'SmartM365MailExecutionFooter:v1') 'Execution footer marker is missing.'
            Assert-Offline ($html -match 'SmartM365-Synthetic-Inventory\.ps1' -and $html -match '9\.8\.7') 'Execution footer lacks script name or version.'
            Assert-Offline (($again | Select-String -Pattern 'SmartM365MailExecutionFooter:v1' -AllMatches).Matches.Count -eq 1) 'Execution footer is not idempotent.'
            Assert-Offline ($html.IndexOf('SmartM365MailExecutionFooter:v1') -gt $html.IndexOf('<p>Fixture</p>')) 'Execution footer is not at the bottom of the mail body.'
        }
        Test-OfflineCase "$variant identity-first CSV roundtrip" {
            $path = Join-Path $testRoot "$variant-roundtrip.csv"
            $label = 'quoted "value", ' + [char]0x00e9 + "`nsecond line"
            & $module { param($p,$s) Write-SmartM365CsvAtomically -Path $p -Data @([pscustomobject]@{ Id = '001'; Label = $s }) -Columns @('Label','Id') } $path $label
            $row = @(Import-Csv -LiteralPath $path -Encoding UTF8)
            Assert-Offline ($row.Count -eq 1 -and $row[0].Label -ceq $label -and $row[0].Id -ceq '001') 'CSV multiline, Unicode or string identity changed.'
            Assert-Offline (($row[0].PSObject.Properties.Name -join ',') -eq 'TenantKey,OrganizationKey,EnvironmentKey,TenantId,Label,Id') 'Column contract changed.'
        }
        Test-OfflineCase "$variant empty explicit schema" {
            $path = Join-Path $testRoot "$variant-empty.csv"
            & $module { param($p) Write-SmartM365CsvAtomically -Path $p -Data @() -Columns @('Id','Value') } $path
            Assert-Offline ((Get-Content -LiteralPath $path -TotalCount 1) -eq '"TenantKey","OrganizationKey","EnvironmentKey","TenantId","Id","Value"') 'Empty header changed.'
        }
        Test-OfflineCase "$variant public exporter accepts an empty explicit schema" {
            $history = Join-Path $testRoot "$variant-empty-export-history.csv"
            $latest = Join-Path $testRoot "$variant-empty-export-latest.csv"
            & $module { param($h,$l) Export-SmartM365Csv -TimestampedPath $h -LatestPath $l -Data @() -Columns @('Id','Value') -NoSharePointUpload } $history $latest | Out-Null
            $expectedHeader = '"TenantKey","OrganizationKey","EnvironmentKey","TenantId","Id","Value"'
            Assert-Offline ((Get-Content -LiteralPath $history -TotalCount 1) -ceq $expectedHeader) 'Empty history export header changed.'
            Assert-Offline ((Get-Content -LiteralPath $latest -TotalCount 1) -ceq $expectedHeader) 'Empty latest export header changed.'
            Assert-Offline ((Get-FileHash -LiteralPath $history).Hash -eq (Get-FileHash -LiteralPath $latest).Hash) 'Empty history and latest exports differ.'
        }
        foreach ($field in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId')) {
            Test-OfflineCase "$variant reject conflicting $field and preserve last" {
                $path = Join-Path $testRoot "$variant-$field.csv"
                [IO.File]::WriteAllText($path, 'LAST VALID SYNTHETIC EXPORT')
                $hash = (Get-FileHash -LiteralPath $path).Hash
                $row = [pscustomobject]@{ Id = 'fixture' }
                $row | Add-Member -NotePropertyName $field -NotePropertyValue 'synthetic-other'
                $caught = $false
                try { & $module { param($p,$r) Write-SmartM365CsvAtomically -Path $p -Data @($r) } $path $row }
                catch { $caught = $_.Exception.Message -match 'identity.*conflict|conflict.*identity' }
                Assert-Offline $caught 'Conflicting identity was silently relabelled.'
                Assert-Offline ((Get-FileHash -LiteralPath $path).Hash -eq $hash) 'Last valid file changed.'
            }
        }
        Test-OfflineCase "$variant mixed rows without context" {
            $caught = $false
            try {
                & $module {
                    Add-SmartM365TenantKeyToCsvData -TenantKey '' -OrganizationKey '' -EnvironmentKey '' -TenantId '' -Data @(
                        [pscustomobject]@{TenantKey='synthetic-a'; OrganizationKey='org'; EnvironmentKey='test'; Id='1'},
                        [pscustomobject]@{TenantKey='synthetic-b'; OrganizationKey='org'; EnvironmentKey='test'; Id='2'}
                    )
                } | Out-Null
            } catch { $caught = $_.Exception.Message -match 'identity.*conflict|conflict.*identity' }
            Assert-Offline $caught 'Mixed tenants accepted when deriving identity from rows.'
        }
        Test-OfflineCase "$variant dictionary identity inference" {
            $value = & $module { Add-SmartM365TenantKeyToCsvData -TenantKey '' -OrganizationKey '' -EnvironmentKey '' -TenantId '' -Data @(@{TenantKey='synthetic-a';OrganizationKey='org';EnvironmentKey='test';Id='1'}) }
            Assert-Offline ($value.Data[0].TenantKey -eq 'synthetic-a' -and $value.Data[0].Id -eq '1') 'Dictionary identity was not recognized.'
        }
        Test-OfflineCase "$variant matching and missing identity accepted" {
            $value = & $module { Add-SmartM365TenantKeyToCsvData -Data @([pscustomobject]@{TenantKey='SYNTHETIC-A';Id='1'}, [pscustomobject]@{TenantKey='';Id='2'}) }
            Assert-Offline ($value.Data.Count -eq 2 -and $value.Data[1].TenantKey -eq 'synthetic-a') 'Valid identity inheritance changed.'
        }
        Test-OfflineCase "$variant dictionary conflicting identity rejected" {
            $caught = $false
            try { & $module { Add-SmartM365TenantKeyToCsvData -Data @(@{TenantKey='synthetic-other'; Id='1'}) } | Out-Null }
            catch { $caught = $_.Exception.Message -match 'identity.*conflict|conflict.*identity' }
            Assert-Offline $caught 'Dictionary conflict accepted.'
        }
        Test-OfflineCase "$variant failed serializer preserves last under Continue" {
            $path = Join-Path $testRoot "$variant-write.csv"
            [IO.File]::WriteAllText($path, 'LAST VALID SYNTHETIC EXPORT')
            $hash = (Get-FileHash -LiteralPath $path).Hash
            # A partial staging file followed by a non-terminating cmdlet error.
            & $module {
                function script:Export-Csv {
                    [CmdletBinding()]param([Parameter(ValueFromPipeline)]$InputObject, $Path, $Encoding, $Delimiter, [switch]$NoTypeInformation)
                    process { [IO.File]::WriteAllText($Path, 'PARTIAL'); Write-Error 'Synthetic disk failure' }
                }
            }
            try {
                $caught = $false
                try { & $module { param($p) $ErrorActionPreference='Continue'; Write-SmartM365CsvAtomically -Path $p -Data @([pscustomobject]@{Id='1'}) } $path 2>$null }
                catch { $caught = $_.Exception.Message -match 'Synthetic disk failure' }
                Assert-Offline $caught 'Serializer failure did not terminate publication.'
                Assert-Offline ((Get-FileHash -LiteralPath $path).Hash -eq $hash) 'Partial staging file replaced last valid CSV.'
            } finally { & $module { Remove-Item Function:Export-Csv } }
        }
        Test-OfflineCase "$variant locked destination preserves last" {
            $path = Join-Path $testRoot "$variant-locked.csv"
            [IO.File]::WriteAllText($path, 'LAST VALID SYNTHETIC EXPORT')
            $lock = [IO.File]::Open($path, 'Open', 'Read', 'Read')
            try {
                $caught = $false
                try { & $module { param($p) Write-SmartM365CsvAtomically -Path $p -Data @([pscustomobject]@{Id='1'}) } $path }
                catch { $caught = $true }
                Assert-Offline $caught 'Locked replacement unexpectedly succeeded.'
            } finally { $lock.Dispose() }
            Assert-Offline ([IO.File]::ReadAllText($path) -eq 'LAST VALID SYNTHETIC EXPORT') 'Locked file changed.'
        }
        Test-OfflineCase "$variant successful DATA-ALL DATA-LAST byte parity" {
            $history = Join-Path $testRoot "$variant-history.csv"
            $latest = Join-Path $testRoot "$variant-latest.csv"
            & $module { param($h,$l) Publish-SmartM365Csv -TimestampedPath $h -LatestPath $l -Data @([pscustomobject]@{Id='001'; Date='2026-01-01T12:00:00Z'; Missing=$null}) -RetentionMaxCsv 0 -NoSharePointUpload } $history $latest | Out-Null
            Assert-Offline ((Get-FileHash $history).Hash -eq (Get-FileHash $latest).Hash) 'Latest and historical bytes differ.'
            $row = Import-Csv $latest
            Assert-Offline ($row.Date -ceq '2026-01-01T12:00:00Z' -and $row.Missing -ceq '') 'Timestamp or missing value changed.'
        }
        Test-OfflineCase "$variant publication conflict preserves both copies" {
            $history = Join-Path $testRoot "$variant-pair-history.csv"
            $latest = Join-Path $testRoot "$variant-pair-latest.csv"
            [IO.File]::WriteAllText($history, 'HISTORY'); [IO.File]::WriteAllText($latest, 'LATEST')
            $caught = $false
            try { & $module { param($h,$l) Publish-SmartM365Csv -TimestampedPath $h -LatestPath $l -Data @([pscustomobject]@{TenantKey='synthetic-other';Id='001'}) -RetentionMaxCsv 0 -NoSharePointUpload } $history $latest | Out-Null }
            catch { $caught = $_.Exception.Message -match 'identity.*conflict|conflict.*identity' }
            Assert-Offline $caught 'Publication identity conflict accepted.'
            Assert-Offline ([IO.File]::ReadAllText($history) -eq 'HISTORY' -and [IO.File]::ReadAllText($latest) -eq 'LATEST') 'Publication changed a prior copy.'
        }
        Test-OfflineCase "$variant MAXITEMS protects canonical files" {
            $history = Join-Path $testRoot "$variant-max-history.csv"
            $latest = Join-Path $testRoot "$variant-max-latest.csv"
            [IO.File]::WriteAllText($latest, 'CANONICAL')
            $global:SmartM365MaxItems = 1
            try {
                $published = & $module { param($h,$l) Publish-SmartM365Csv -TimestampedPath $h -LatestPath $l -Data @([pscustomobject]@{Id='1'},[pscustomobject]@{Id='2'}) -RetentionMaxCsv 0 -NoSharePointUpload } $history $latest
                Assert-Offline ($published.LatestPath -like '*_MAXITEMS-1.csv') 'Limited run used canonical name.'
                Assert-Offline ([IO.File]::ReadAllText($latest) -eq 'CANONICAL') 'Canonical latest changed.'
                Assert-Offline (@(Import-Csv $published.LatestPath).Count -eq 1) 'MaxItems row limit changed.'
            } finally { $global:SmartM365MaxItems = 0 }
        }
        Test-OfflineCase "$variant tenant-neutral semicolon contract" {
            $path = Join-Path $testRoot "$variant-neutral.csv"
            & $module { param($p) Write-SmartM365CsvAtomically -Path $p -Data @([pscustomobject]@{Date='2026-01-01';Value='1.25'}) -NoTenantKey -Delimiter ';' } $path
            $row = Import-Csv $path -Delimiter ';'
            Assert-Offline (($row.PSObject.Properties.Name -join ',') -eq 'Date,Value' -and $row.Value -ceq '1.25') 'Neutral/delimiter contract changed.'
        }
        Test-OfflineCase "$variant atomic copy byte parity" {
            $source = Join-Path $testRoot "$variant-copy-source.csv"
            $destination = Join-Path $testRoot "$variant-copy-destination.csv"
            [IO.File]::WriteAllText($source, "synthetic,source`r`n1,2", [Text.UTF8Encoding]::new($false))
            [IO.File]::WriteAllText($destination, 'LAST VALID SYNTHETIC EXPORT')
            & $module { param($s,$d) Copy-SmartM365FileAtomically -SourcePath $s -DestinationPath $d } $source $destination
            Assert-Offline ((Get-FileHash $source).Hash -eq (Get-FileHash $destination).Hash) 'Atomic copy changed bytes.'
        }
        Test-OfflineCase "$variant locked atomic copy preserves last" {
            $source = Join-Path $testRoot "$variant-locked-copy-source.csv"
            $destination = Join-Path $testRoot "$variant-locked-copy-destination.csv"
            [IO.File]::WriteAllText($source, 'NEW SYNTHETIC EXPORT')
            [IO.File]::WriteAllText($destination, 'LAST VALID SYNTHETIC EXPORT')
            $lock = [IO.File]::Open($destination, 'Open', 'Read', 'Read')
            try {
                $caught = $false
                try { & $module { param($s,$d) Copy-SmartM365FileAtomically -SourcePath $s -DestinationPath $d } $source $destination }
                catch { $caught = $true }
                Assert-Offline $caught 'Locked atomic copy unexpectedly succeeded.'
            } finally { $lock.Dispose() }
            Assert-Offline ([IO.File]::ReadAllText($destination) -eq 'LAST VALID SYNTHETIC EXPORT') 'Locked atomic copy changed last valid file.'
        }
        foreach ($legacyFunction in @('ExportAndCopyCsv','ExportAndCopyCsvFromConvert')) {
            Test-OfflineCase "$variant $legacyFunction blocks incomplete DATA-LAST" {
                $output = Join-Path $testRoot "$variant-$legacyFunction-output"
                $globalOutput = Join-Path $testRoot "$variant-$legacyFunction-global"
                [void][IO.Directory]::CreateDirectory($output)
                [void][IO.Directory]::CreateDirectory($globalOutput)
                $latest = Join-Path $globalOutput 'Synthetic_Legacy.csv'
                [IO.File]::WriteAllText($latest, 'LAST VALID SYNTHETIC EXPORT')
                $lock = [IO.File]::Open($latest, 'Open', 'Read', 'Read')
                try {
                    $caught = $false
                    try {
                        & $module {
                            param($name,$o,$g)
                            & $name -BaseFileName 'Synthetic_Legacy' -OutputPath $o -GlobalPath $g -Data @([pscustomobject]@{Id='1'}) -SkipWeeklyHistory
                        } $legacyFunction $output $globalOutput
                    }
                    catch { $caught = $true }
                    Assert-Offline $caught 'Incomplete DATA-LAST publication was acknowledged.'
                } finally { $lock.Dispose() }
                Assert-Offline ([IO.File]::ReadAllText($latest) -eq 'LAST VALID SYNTHETIC EXPORT') 'Incomplete publication replaced last valid DATA-LAST.'
            }
        }
        if ($variant -eq 'Core') {
            Test-OfflineCase 'Core streaming later-row identity conflict preserves last' {
                $path = Join-Path $testRoot 'Core-stream-conflict.csv'
                [IO.File]::WriteAllText($path, 'LAST VALID SYNTHETIC EXPORT')
                $caught = $false
                try {
                    & $module {
                        param($p)
                        Export-SmartM365CsvStreamAtomically -Path $p -Data @(
                            [pscustomobject]@{TenantKey='synthetic-a';Id='1'},
                            [pscustomobject]@{TenantKey='synthetic-other';Id='2'}
                        )
                    } $path
                }
                catch { $caught = $_.Exception.Message -match 'identity.*conflict|conflict.*identity' }
                Assert-Offline $caught 'Later streaming identity conflict was silently relabelled.'
                Assert-Offline ([IO.File]::ReadAllText($path) -eq 'LAST VALID SYNTHETIC EXPORT') 'Streaming conflict changed last valid file.'
            }
        }
        Test-OfflineCase "$variant atomic CSV append preserves schema" {
            $path = Join-Path $testRoot "$variant-append.csv"
            & $module { param($p) Write-SmartM365CsvAtomically -Path $p -Data @([pscustomobject]@{Id='1';Stamp='2026-01-01T00:00:00Z'}) } $path
            $header = Get-Content -LiteralPath $path -TotalCount 1
            & $module { param($p) Add-SmartM365CsvRowsAtomically -Path $p -Data @([pscustomobject]@{Id='2';Stamp='2026-01-02T00:00:00Z'}) } $path
            $rows = @(Import-Csv -LiteralPath $path)
            Assert-Offline ($rows.Count -eq 2 -and $rows[1].Id -eq '2') 'Atomic append lost or changed rows.'
            Assert-Offline ((Get-Content -LiteralPath $path -TotalCount 1) -ceq $header) 'Atomic append changed the header schema.'
        }
        Test-OfflineCase "$variant mismatched CSV append preserves history" {
            $path = Join-Path $testRoot "$variant-append-schema-mismatch.csv"
            & $module { param($p) Write-SmartM365CsvAtomically -Path $p -Data @([pscustomobject]@{Id='1';Stamp='2026-01-01T00:00:00Z'}) } $path
            $hash = (Get-FileHash -LiteralPath $path).Hash
            $caught = $false
            try { & $module { param($p) Add-SmartM365CsvRowsAtomically -Path $p -Data @([pscustomobject]@{Stamp='2026-01-02T00:00:00Z';Id='2'}) } $path }
            catch { $caught = $_.Exception.Message -match 'schema.*header' }
            Assert-Offline $caught 'Mismatched append schema was accepted.'
            Assert-Offline ((Get-FileHash -LiteralPath $path).Hash -eq $hash) 'Schema mismatch changed existing CSV history.'
        }
        Test-OfflineCase "$variant locked CSV append preserves history" {
            $path = Join-Path $testRoot "$variant-append-locked.csv"
            & $module { param($p) Write-SmartM365CsvAtomically -Path $p -Data @([pscustomobject]@{Id='1'}) } $path
            $hash = (Get-FileHash -LiteralPath $path).Hash
            $lock = [IO.File]::Open($path, 'Open', 'Read', 'Read')
            try {
                $caught = $false
                try { & $module { param($p) Add-SmartM365CsvRowsAtomically -Path $p -Data @([pscustomobject]@{Id='2'}) } $path }
                catch { $caught = $true }
                Assert-Offline $caught 'Locked CSV history append unexpectedly succeeded.'
            } finally { $lock.Dispose() }
            Assert-Offline ((Get-FileHash -LiteralPath $path).Hash -eq $hash) 'Locked CSV history was modified.'
        }
        Test-OfflineCase "$variant atomic text replacement" {
            $path = Join-Path $testRoot "$variant-manifest.json"
            & $module { param($p) Write-SmartM365TextAtomically -Path $p -Content '{"status":"complete"}' -Encoding UTF8 } $path
            $manifest = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
            Assert-Offline ($manifest.status -eq 'complete') 'Atomic text manifest content changed.'
        }
        Test-OfflineCase "$variant locked text replacement preserves manifest" {
            $path = Join-Path $testRoot "$variant-manifest-locked.json"
            [IO.File]::WriteAllText($path, '{"status":"last-valid"}')
            $lock = [IO.File]::Open($path, 'Open', 'Read', 'Read')
            try {
                $caught = $false
                try { & $module { param($p) Write-SmartM365TextAtomically -Path $p -Content '{"status":"new"}' -Encoding UTF8 } $path }
                catch { $caught = $true }
                Assert-Offline $caught 'Locked manifest replacement unexpectedly succeeded.'
            } finally { $lock.Dispose() }
            Assert-Offline ([IO.File]::ReadAllText($path) -eq '{"status":"last-valid"}') 'Locked manifest was modified.'
        }
        Remove-Module $module
    }
    Test-OfflineCase 'Every SmartInventory mail sink uses the shared normalization contract' {
        $inventoryRoot = Join-Path $SourceRoot 'SmartInventory'
        $mailSinkNames = @('Send-SmartM365Mail', 'SendEmailHtmlReport', 'Send-CoreSmartM365Mail')
        $mailFiles = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $directBypasses = New-Object 'System.Collections.Generic.List[string]'
        foreach ($file in Get-ChildItem -LiteralPath $inventoryRoot -Recurse -File -Filter '*.ps1') {
            $tokens = $null
            $parseErrors = $null
            $ast = [Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$parseErrors)
            Assert-Offline (@($parseErrors).Count -eq 0) "SmartInventory mail audit could not parse $($file.FullName)."
            foreach ($command in $ast.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] }, $true)) {
                $commandName = $command.GetCommandName()
                if ($commandName -in $mailSinkNames) { [void]$mailFiles.Add($file.FullName) }
                if ($commandName -in @('Send-MailMessage', 'Send-SmartM365GraphMail')) {
                    $directBypasses.Add(('{0}:{1}:{2}' -f $file.FullName, $command.Extent.StartLineNumber, $commandName))
                }
            }
        }
        # 23 since PreparedEvidence/SmartM365-WorkplaceEvidence-Prepare.ps1 sends its run mail.
        Assert-Offline ($mailFiles.Count -eq 23) "Expected 23 SmartInventory mail-capable scripts, found $($mailFiles.Count)."
        Assert-Offline ($directBypasses.Count -eq 0) ("Direct mail transports bypass shared normalization: {0}" -f ($directBypasses -join '; '))
        foreach ($mailFile in $mailFiles) {
            $mailSource = [IO.File]::ReadAllText($mailFile)
            Assert-Offline ($mailSource -match '(?im)^\s*\.VERSION\s*\r?\n\s*[^\r\n]+') "Mail-capable script has no .VERSION metadata: $mailFile"
        }

        $orchestratorSource = [IO.File]::ReadAllText((Join-Path $inventoryRoot 'Orchestrator/SmartM365-Inventory-Orchestrator.ps1'))
        Assert-Offline (($orchestratorSource | Select-String -Pattern 'Format-SmartM365MailSubject -Subject \$Subject -Orchestrator' -AllMatches).Matches.Count -ge 2) 'Orchestrator mail paths do not enforce the dedicated subject prefix.'
        Assert-Offline ($orchestratorSource -match 'ConvertTo-SmartM365EmailBody -BodyHtml \$HtmlBody') 'The direct orchestrator SMTP helper does not apply the common footer.'

        foreach ($relativeModulePath in @('Modules/SmartM365.Core/SmartM365.Core.psm1', 'Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1')) {
            $moduleSource = [IO.File]::ReadAllText((Join-Path $SourceRoot $relativeModulePath))
            Assert-Offline (($moduleSource | Select-String -Pattern '\$Subject = Format-SmartM365MailSubject -Subject \$Subject' -AllMatches).Matches.Count -ge 2) "$relativeModulePath does not normalize both shared and direct Graph mail subjects."
            Assert-Offline ($moduleSource -match 'if \(\$BodyHtml -match ''SmartM365MailBranding:v1''\) \{ return Add-SmartM365MailExecutionFooter') "$relativeModulePath does not append execution metadata to pre-branded mail."
        }
    }
    $distributed = Import-OfflineFunctions (Join-Path $SourceRoot 'SmartInventory/Orchestrator/SmartM365.Orchestrator.Distributed.psm1') @(
        'Resolve-DistributedJsonPath', 'Write-JsonAtomically', 'ConvertTo-SafeFileName', 'Enter-SmartM365OrchestratorConcurrencyLease',
        'Set-SmartM365OrchestratorConcurrencyLease', 'Exit-SmartM365OrchestratorConcurrencyLease',
        'Enter-SmartM365OrchestratorOccurrenceClaim', 'Set-SmartM365OrchestratorOccurrenceClaim'
    )
    Test-OfflineCase 'Orchestrator locked lease update must not report success' {
        $lease = & $distributed { param($p) Enter-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $p -ConcurrencyKey 'Synthetic' -JobName 'JobA' -Occurrence ([datetime]'2026-01-01') -OwnerServer 'SERVER-A' } (Join-Path $testRoot 'leases')
        $hash = (Get-FileHash -LiteralPath $lease.LeasePath).Hash
        $lock = [IO.File]::Open($lease.LeasePath, 'Open', 'Read', 'Read')
        try {
            $caught = $false
            try { & $distributed { param($l) $ErrorActionPreference='Continue'; Set-SmartM365OrchestratorConcurrencyLease -LeasePath $l.LeasePath -LeaseId $l.Lease.LeaseId -OwnerServer 'SERVER-A' -SafeUntilUtc ([datetime]::UtcNow.AddHours(5)) } $lease 2>$null | Out-Null }
            catch { $caught = $true }
            Assert-Offline $caught 'Lease update acknowledged although persistence failed.'
        } finally { $lock.Dispose() }
        Assert-Offline ((Get-FileHash -LiteralPath $lease.LeasePath).Hash -eq $hash) 'Locked lease changed.'
    }
    Test-OfflineCase 'Orchestrator locked claim update must fail' {
        $claim = & $distributed { param($p) Enter-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $p -JobName 'JobA' -Occurrence ([datetime]'2026-01-01') -OwnerServer 'SERVER-A' -PlanId 'synthetic' } (Join-Path $testRoot 'claims')
        $lock = [IO.File]::Open($claim.ClaimPath, 'Open', 'Read', 'Read')
        try {
            $caught = $false
            try { & $distributed { param($c) $ErrorActionPreference='Continue'; Set-SmartM365OrchestratorOccurrenceClaim -ClaimPath $c.ClaimPath -OwnerServer 'SERVER-A' -Status Running } $claim 2>$null | Out-Null }
            catch { $caught = $true }
            Assert-Offline $caught 'Claim update acknowledged although persistence failed.'
        } finally { $lock.Dispose() }
    }
    Remove-Module $distributed
}
finally {
    # Delete only this run's verified synthetic temporary root.
    $resolved = [IO.Path]::GetFullPath($testRoot)
    $expectedParent = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($expectedParent, [StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'SmartInventory-Offline-*') { throw 'Unsafe cleanup root.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
if ($ResultPath) { $results | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $ResultPath -Encoding UTF8 }
$results | Format-Table Name, Passed, Error -AutoSize
$failed = @($results | Where-Object { -not $_.Passed }).Count
Write-Output ("Cases={0}; Passed={1}; Failed={2}; PowerShell={3}" -f $results.Count, ($results.Count-$failed), $failed, $PSVersionTable.PSVersion)
if ($failed) { exit 1 }
