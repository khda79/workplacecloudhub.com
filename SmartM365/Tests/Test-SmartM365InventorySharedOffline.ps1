<#
.SYNOPSIS
Synthetic regression tests for shared inventory identity and atomic persistence.
.VERSION
1.1.7
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
                    ProfileKey = 'prod'; OrganizationKey = 'emeis'; EnvironmentKey = 'prod'
                    TenantKey = 'emeis-prod'; TenantId = '00000000-0000-0000-0000-000000000001'
                }
                Initialize-SmartM365TenantContext -Tenant prod | Out-Null
            }
            Assert-Offline ($global:SmartM365MailTenantName -ceq 'EMEIS') 'PSCustomObject tenant context did not use the organization fallback.'
        }
        Test-OfflineCase 'Tenant context accepts dictionary MailClientName override' {
            & $tenantContextModule {
                $script:SyntheticEffectiveConfig = @{
                    ProfileKey = 'prod'; OrganizationKey = 'emeis'; EnvironmentKey = 'prod'
                    TenantKey = 'emeis-prod'; TenantId = '00000000-0000-0000-0000-000000000001'
                    MailClientName = 'EMEIS France'
                }
                Initialize-SmartM365TenantContext -Tenant prod | Out-Null
            }
            Assert-Offline ($global:SmartM365MailTenantName -ceq 'EMEIS France') 'Dictionary tenant context ignored MailClientName.'
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
            $global:SmartM365MailTenantName = 'EMEIS'
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
            Assert-Offline ($general -ceq '[SMART 365] - [EMEIS] - [CRITICAL] Microsoft Teams Inventory - 2026-09-21T12:12:41Z') 'General subject prefix normalization changed.'
            Assert-Offline ($legacy -ceq '[SMART 365] - [EMEIS] - EXO quarantine report - ERROR') 'Legacy SmartM365 subject was not normalized.'
            Assert-Offline ($limited -ceq '[SMART 365] - [EMEIS] - [MAXITEMS-5 TEST] - [WARNING] WinUpdate Feature Update') 'MAXITEMS subject no longer begins with the tenant prefix.'
        }
        Test-OfflineCase "$variant orchestrator mail subject uses its dedicated prefix" {
            $subject = & $module { Format-SmartM365MailSubject -Subject '[SmartM365 Orchestrator][prod] Runtime update detected: 1.5.14 -> 1.5.15' -Orchestrator }
            $critical = & $module { Format-SmartM365MailSubject -Subject '[CRITICAL][SmartM365 Orchestrator][prod] Authenticode rejected job JobA' -Orchestrator }
            $again = & $module { param($s) Format-SmartM365MailSubject -Subject $s } $subject
            Assert-Offline ($subject -ceq '[SMART 365] - [EMEIS] - [ Orchestrator] - Runtime update detected: 1.5.14 -> 1.5.15') 'Orchestrator subject prefix normalization changed.'
            Assert-Offline ($critical -ceq '[SMART 365] - [EMEIS] - [ Orchestrator] - [CRITICAL] Authenticode rejected job JobA') 'Critical orchestrator subject was not normalized.'
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
        Assert-Offline ($mailFiles.Count -eq 22) "Expected 22 SmartInventory mail-capable scripts, found $($mailFiles.Count)."
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
        'Write-JsonAtomically', 'ConvertTo-SafeFileName', 'Enter-SmartM365OrchestratorConcurrencyLease',
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAUZQOVpTs64KUZ
# +gGyzNwxsGJDnXptSuHOe8TIgThBzaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIICIlsEa+txCA2cFGfNJSPIKBY87fmVP+W056vfZCguYMA0GCSqG
# SIb3DQEBAQUABIIBgA2dEgCiSm3FIaZWERMG5f13ywAvyddyUV9+OgztHUIeQAF/
# Vp+p/VwEO+W85o/ifxziuI46xq5QOPWV/cIne+d+O47qQf3Egb/iwclNLKjWFxFn
# VFGrXfKpEcXbTC+SpgRy1/pv0NGXcEgAnO1Ois4bPvkrZcD1lklP+oUuD/pZZ6Mn
# u+DNikIVik1Mh2fMQ8h40oK21rn9c5XsOqiHc5DPCTzGvBr1yi2QXSTMb23EvmXN
# qdiHlhkkW77VVpC7qgrCU8BmL01nZRWwDgPG6rNrIYX5KBPdmLIuIIIU59pURTJd
# dg2oQdUCT6/kHJFisSp12iRN95P42FcnDrVOmhVEFZAZzax1vjHMpQZZAoDV1nLx
# 6P/sEXPMSUwaRJ29ZOruWD4je05CS2mwvt/Ut28jeec+EZ8cwLJaTUbZlQHVQIR7
# OGg/i9Q9yt4eCMVPG0COjKyVVX6DA1HEKyIM8QnTW5CBwxBtwcbdrvR+VYGLMOl1
# +3/O3ZHIQcKR08CvpaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjIxOTMw
# MTdaMC8GCSqGSIb3DQEJBDEiBCCUjvoPWEXhIccml70MIvAyTMobTnfK1b6hazZm
# 6gs3pjANBgkqhkiG9w0BAQEFAASCAgBnw8XdEF7mV8VLSiJ/dn5edVPL7yT2EcjT
# usDmoC2REcHaUPmtvp3YZbFPeQfEOL6raZkC09s0h9rYkhtx4VRZGyFI8q7rVctH
# U6545OXV9j9hRZcGdMan0ft6JlwC+QptLVsFlHVkL1yjenqMEJ/RhUJk2v6nNZfM
# DRuXCnlBpIOhhk3m24v6yFKfc7iqdag+7HJULV5SjI3/2fCkBVk5FmwKA8/+mhBS
# AKw+J5jwkMZaKqW4xHm3BG+9uWJ254ksvrMCGNRpetDzNBv+J0pkJD328ZUS7Mi2
# ajVIWu3JFJCy5XITOdMxMuKbnu6Xfe5tMAitCf6HH2qWek7BylTfGQeQlJW8DM+g
# aCLt+Xt8SQX5wo3SrQMHixxP+B4TyfujLNFs92xyZ8+ibbDC2xZM8PoHiGA4wT23
# mceIWVj8VR0xX9IKVYZHStp6JQM2Dy72v9zystC1XJnPWGae1P1t4InITVFP2Zmc
# wIHc5VgttnNznRpbjQW5O99KIEmrq/yuiJESFXWex4qb7BGNslDoD1+4FFIy53SZ
# TRTgABFBTeBlbhOsB9Cunhcml+ML+qDlsktK+lC2atBySMZsKl4Dt6R4g434FGuX
# E/EzZl8tgqrsR860+jnigmSaBcDxnJSOr9wJN9t2pBiqEQ2pg1oBXAS6exNfVraD
# Txyx6uw9NQ==
# SIG # End signature block
