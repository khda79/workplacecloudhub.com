<#
.SYNOPSIS
Synthetic regression tests for shared inventory identity and atomic persistence.
.VERSION
1.1.0
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
    foreach ($variant in @('Core', 'WindowsPowerShell5')) {
        $relative = if ($variant -eq 'Core') { 'Modules/SmartM365.Core/SmartM365.Core.psm1' } else { 'Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1' }
        $module = Import-OfflineFunctions (Join-Path $SourceRoot $relative) @(
            'Add-SmartM365TenantKeyToCsvData', 'Write-SmartM365CsvAtomically', 'Add-SmartM365CsvRowsAtomically',
            'Copy-SmartM365FileAtomically', 'Write-SmartM365TextAtomically',
            'Get-SmartM365CoreContextValue', 'Assert-SmartM365CsvDataCompleteness',
            'Get-SmartM365CsvValidationRule', 'Get-SmartM365CsvValidationBaseName',
            'Publish-SmartM365Csv', 'Get-SmartM365MaxItemsValue', 'Test-SmartM365MaxItemsMode',
            'Get-SmartM365MaxItemsSuffix', 'Add-SmartM365MaxItemsSuffixToCsvPath',
            'Add-SmartM365MaxItemsSuffixToBaseName', 'Limit-SmartM365RowsForMaxItems',
            'ExportAndCopyCsv', 'ExportAndCopyCsvFromConvert'
        )
        if ($variant -eq 'Core') {
            Remove-Module $module
            $module = Import-OfflineFunctions (Join-Path $SourceRoot $relative) @(
                'Add-SmartM365TenantKeyToCsvData', 'Write-SmartM365CsvAtomically', 'Add-SmartM365CsvRowsAtomically',
                'Copy-SmartM365FileAtomically', 'Write-SmartM365TextAtomically',
                'Get-SmartM365CoreContextValue', 'Assert-SmartM365CsvDataCompleteness',
                'Get-SmartM365CsvValidationRule', 'Get-SmartM365CsvValidationBaseName',
                'Publish-SmartM365Csv', 'Get-SmartM365MaxItemsValue', 'Test-SmartM365MaxItemsMode',
                'Get-SmartM365MaxItemsSuffix', 'Add-SmartM365MaxItemsSuffixToCsvPath',
                'Add-SmartM365MaxItemsSuffixToBaseName', 'Limit-SmartM365RowsForMaxItems',
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCIM7iRXIXv16FI
# hYZ1yv2y6kTDMTxYC5+ZCaGJkAlmraCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKj4sjssP2cptJwauw1AFGQI6wpLEUc4a5bFLEuw8JoYMA0GCSqG
# SIb3DQEBAQUABIIBgJVq7rxRy5Nc1NLyzlEq22HTbLkSjcu4Zq4m1W2uciLv/wjr
# SkKn/vcMLmmH9TwiFNPzvI8bCgF0Kx1PxMEGb1Zo2Ah3/c+/pK9/XauxQgmx2SWY
# key7MGdHuT+N29MUz30q5y7+YJYK/7759QLTkZiUakIXa5c7jNBhK//JClKTO3lY
# ciIvGhGP0IJHZoHhT02gFAhfXDJb7ciiT17WH7PZQ9v9iIKSMB+8ldkbDXAn2cjn
# AUy/at0Y3GtrT/wMgXnBBX4quzOiKKUUxQPiba0DAMViilATEG7czOkzOhsA5jJS
# pItlsrFybCZ4l+b6rnVW0k0vZ3otGY1EqFbWXLiDWg5P8T7wyf+PUeKtPVsiC662
# OJuUIyP+zvfhEtUEi/4tNRfkWHjWAPf79h/t4YmDP9ufffadXrgF3nHyX8DRpO5P
# lRbiED5mo29/A4vaB9wGR+V9mAjOSLvUnfcq0e8Gx+5FBeUpK0O9MuodJkd+pO+d
# 3g/EeyP5UOsA6PNjvKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTQwODAw
# MzBaMC8GCSqGSIb3DQEJBDEiBCA/N6QQxTW2N17VIe6ru0Nbvi2hHT28SDqvrxE8
# uwSl6zANBgkqhkiG9w0BAQEFAASCAgAo2QD5pR+LQAPvzDuGc6i+vNIuBwffkZEh
# 1X4BkjSJASfv3NL+QFvpCi+4uiempc6vDbZ9QsbZLE7Iv+tEKVLV6r/edl2pIQV6
# 6FmRfCwNac93eFaZ1N+yZxMZEGN3A9TxScVzLg+Eaurp04ka8F5jQQe3HVFykM+B
# aH4zzCQqfrIQLQXTiiJ/Om/qAl/CNmijB1uBxODIgSjsiWc6DMgNmAR/Mm3vDpOW
# Jj7ROm6h2x0wMshxK33EphNa19uhD6hMsdXX3UShH/MhbywwgpYdGSSWqKM3BNBY
# B3r+VtlRtRhM6KkHVY+alU2HpC7tnQB3R53/M6G2XUqJOg+smzEu+AwBdYSE15Kx
# s3WFBmjCXrFN519D0nbim5nRpTNKDsTTt+aFWra8GSggHNxrD/T+96+kxIFh7Oyp
# xKoqmQ6c5UCLkHadIcN9lZyrzxmdMugEkDC4NBSQQaUywiAA+FD7X3j60R1YmXyC
# iiVpWdtbsDzkPsCit5f4dFSowlw56v+ZImluqpU6VlKwO++Oz7SpPvgTsSw4CnvC
# FDXuqxgTSSWS+L6650c5JacvKhflqhvoKOa5zA/j0F1FJEYgZh+uhYIH1wlIXqwS
# v/c038ROI9M1Z4eNVy5/+h6viDRJ4ZLQ4vPgwrVzNpQ0CX6Dw1vqCjrLijAl92Ft
# RHHzrzjAhQ==
# SIG # End signature block
