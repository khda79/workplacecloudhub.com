<#
.SYNOPSIS
Runs offline SmartWorkplaceCMDB orchestrator and launcher tests.

.VERSION
1.1.15
#>
[CmdletBinding()]
param()

$ScriptVersion = '1.1.15'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$script:Passed = 0
$script:Failed = 0

function Invoke-SmartWorkplaceCMDBOrchestratorTest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$Test
    )
    try {
        & $Test
        $script:Passed++
        Write-Information "PASS $Name" -InformationAction Continue
    }
    catch {
        $script:Failed++
        Write-Information "FAIL $Name - $($_.Exception.Message)" `
            -InformationAction Continue
    }
}

function Assert-SmartWorkplaceCMDBOrchestratorTrue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][bool]$Condition,
        [Parameter(Mandatory)][string]$Message
    )
    if (-not $Condition) {
        throw $Message
    }
}

function Assert-SmartWorkplaceCMDBOrchestratorThrow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Body,
        [Parameter(Mandatory)][string]$ExpectedText
    )
    try {
        & $Body
    }
    catch {
        if ($_.Exception.Message -notlike "*$ExpectedText*") {
            throw "Unexpected exception: $($_.Exception.Message)"
        }
        return
    }
    throw "Expected exception containing '$ExpectedText'."
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$orchestrator = Join-Path $projectRoot 'Orchestration\SmartWorkplaceCMDB-Orchestrator.ps1'
$fixtureRoot = Join-Path $PSScriptRoot 'Fixtures'
$modulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$contractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json'
$rawContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
$adContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.activedirectory.tables.json'
$hardwareContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.hardware.tables.json'
$launcherRoot = Join-Path $projectRoot 'Launchers\Cloud'
Import-Module $modulePath -Force

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) (
    'SmartWorkplaceCMDB-Orchestrator-Tests-' +
    [guid]::NewGuid().ToString('N')
)
$identity = @{
    Tenant = 'test'
    OrganizationKey = 'contoso'
    EnvironmentKey = 'prod'
    TenantKey = 'contoso-prod'
    TenantId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
    NoConfigWrite = $true
}

try {
    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Default validation stays read-only' {
        $validateRoot = Join-Path $tempRoot 'Validate'
        $result = & $orchestrator @identity `
            -DataRootPath $validateRoot `
            -FixtureRootPath $fixtureRoot `
            -ValidateOnly
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            ($result.Status -eq 'Validated' -and
                $result.Mode -eq 'Validate' -and
                $result.StepCount -eq 16 -and
                -not (Test-Path $validateRoot)) `
            'Default validation mode wrote output or returned invalid status.'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Timestamp console and show lifecycle banners' {
        $validateRoot = Join-Path $tempRoot 'ConsoleValidation'
        $captured = @(& $orchestrator @identity `
                -DataRootPath $validateRoot `
                -FixtureRootPath $fixtureRoot `
                -Pipeline ActiveDirectory `
                -ValidateOnly 6>&1)
        $result = @($captured | Where-Object {
                $_ -is [psobject] -and
                $null -ne $_.PSObject.Properties['ScriptVersion'] -and
                $null -ne $_.PSObject.Properties['Status']
            } | Select-Object -Last 1)
        $messages = @($captured |
            Where-Object { $_ -is [System.Management.Automation.InformationRecord] } |
            ForEach-Object { [string]$_.MessageData })
        $operationalMessages = @($messages | Where-Object {
                $_ -match '^\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\] '
            })
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            ($result.Count -eq 1 -and
                $result[0].Status -eq 'Validated' -and
                $result[0].ScriptVersion -eq '1.1.16' -and
                $messages -contains ' SmartWorkplaceCMDB by WorkplaceCloudHub' -and
                $messages -contains ' Version : 1.1.16' -and
                $messages -contains ' SmartWorkplaceCMDB execution summary' -and
                $messages -contains ' Status   : Validated' -and
                @($operationalMessages | Where-Object {
                        $_ -like '*Active Directory collection*'
                    }).Count -ge 2 -and
                @($operationalMessages | Where-Object {
                        $_ -like '*orchestration validated*'
                    }).Count -eq 1 -and
                -not (Test-Path $validateRoot)) `
            'Console timestamps or lifecycle banners are missing or invalid.'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Run complete offline fixture pipeline' {
        $runtimeRoot = Join-Path $tempRoot 'Runtime'
        $script:FullResult = & $orchestrator @identity `
            -DataRootPath $runtimeRoot `
            -FixtureRootPath $fixtureRoot
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            ($script:FullResult.Status -eq 'Completed' -and
                $script:FullResult.Mode -eq 'Fixture' -and
                $script:FullResult.StepCount -eq 34 -and
                $script:FullResult.FailedStepCount -eq 0 -and
                -not $script:FullResult.SummaryEmailEligible -and
                $script:FullResult.SummaryEmailStatus -eq 'NotApplicable' -and
                [string]::IsNullOrWhiteSpace($script:FullResult.SummaryEmailHtmlPath)) `
            'Full fixture orchestration status is invalid.'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Count collector warnings in the completion summary' {
        $warningRoot = Join-Path $tempRoot 'WarningSummary'
        $warningFixtureRoot = Join-Path $tempRoot 'WarningFixtures'
        New-Item -ItemType Directory -Path $warningFixtureRoot -Force | Out-Null
        $fixtureObject = Get-Content -Raw -LiteralPath (Join-Path $fixtureRoot 'IntuneAnalytics.sample.json') | ConvertFrom-Json
        $alertRows = [object[]]$fixtureObject.windowsUpdateAlerts
        $exactDuplicate = ConvertFrom-Json (ConvertTo-Json -InputObject $alertRows[0] -Depth 8)
        $fixtureObject.windowsUpdateAlerts = [object[]]@($alertRows + $exactDuplicate)
        $warningFixturePath = Join-Path $warningFixtureRoot 'IntuneAnalytics.sample.json'
        ConvertTo-Json -InputObject $fixtureObject -Depth 12 | Set-Content -LiteralPath $warningFixturePath -Encoding UTF8
        $result = & $orchestrator @identity `
            -DataRootPath $warningRoot `
            -FixtureRootPath $warningFixtureRoot `
            -Pipeline IntuneAnalytics
        $orchestratorLog = Get-Content -Raw -LiteralPath $result.OrchestratorLogPath
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            ($result.Status -eq 'CompletedWithWarnings' -and
                $result.WarningCount -ge 1 -and
                $orchestratorLog -like ("*Warnings : {0}*" -f $result.WarningCount)) `
            'Collector warnings were not reflected in the completion status and count.'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Finalize without rerunning collection steps' {
        # The preceding data is synthetic, but FinalizeOnly intentionally accepts
        # only a live-owned root. Change only the test root marker to exercise
        # the live finalization path without weakening production isolation.
        $markerPath = Join-Path $script:FullResult.DataRootPath `
            '.collection-root.json.txt'
        $marker = Get-Content -Raw -LiteralPath $markerPath | ConvertFrom-Json
        $marker.Kind = 'Live'
        $marker | ConvertTo-Json -Depth 8 | Set-Content `
            -LiteralPath $markerPath -Encoding UTF8
        $devicePath = Join-Path $script:FullResult.LatestOutputRootPath `
            'CMDB\CMDB_Devices.csv'
        $beforeHash = (Get-FileHash -LiteralPath $devicePath -Algorithm SHA256).Hash
        $beforeWriteTime = (Get-Item -LiteralPath $devicePath).LastWriteTimeUtc
        $result = & $orchestrator @identity `
            -DataRootPath $script:FullResult.DataRootPath `
            -FinalizeOnly
        $afterHash = (Get-FileHash -LiteralPath $devicePath -Algorithm SHA256).Hash
        $afterWriteTime = (Get-Item -LiteralPath $devicePath).LastWriteTimeUtc
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            ($result.Status -eq 'Completed' -and
                $result.Mode -eq 'Finalize' -and
                $result.FinalizationOnly -and
                $result.StepCount -eq 0 -and
                $result.LoggingEnabled -and
                $beforeHash -eq $afterHash -and
                $beforeWriteTime -eq $afterWriteTime) `
            'Finalization reran or modified a collection output.'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Publish summary and terminal logs after finalization' {
        $content = Get-Content -Raw -LiteralPath $orchestrator
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            ($content -match "summaryResult\.HistoryPath" -and
                $content -match "summaryResult\.LatestPath" -and
                $content -match "summaryResult\.HtmlPath" -and
                $content -match "ReuseLatestSnapshot" -and
                $content -match "terminalFiles" -and
                $content -match "Terminal log synchronization") `
            'The post-summary or terminal SharePoint synchronization contract is incomplete.'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Publish all contracts and report' {
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath $script:FullResult.LatestOutputRootPath `
            -ContractPath $contractPath)
        $rawResults = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath $script:FullResult.LatestOutputRootPath `
            -ContractPath $rawContractPath)
        $adResults = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath $script:FullResult.LatestOutputRootPath `
            -ContractPath $adContractPath)
        $hardwareResults = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath $script:FullResult.LatestOutputRootPath `
            -ContractPath $hardwareContractPath)
        $reportPath = Join-Path `
            $script:FullResult.LatestOutputRootPath `
            'SmartWorkplaceCMDB-Overview.html'
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            ($results.Count -eq 37 -and
                @($results | Where-Object Status -ne 'Valid').Count -eq 0 -and
                $rawResults.Count -eq 27 -and
                @($rawResults | Where-Object Status -ne 'Valid').Count -eq 0 -and
                $adResults.Count -eq 6 -and
                @($adResults | Where-Object Status -ne 'Valid').Count -eq 0 -and
                $hardwareResults.Count -eq 1 -and
                @($hardwareResults | Where-Object Status -ne 'Valid').Count -eq 0 -and
                (Test-Path $reportPath -PathType Leaf)) `
            'Full fixture orchestration did not publish the complete model.'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Write auditable step log' {
        $rows = @(Import-Csv $script:FullResult.LogPath)
        $stepLogs = @($rows | ForEach-Object { $_.LogPath })
        $stepTranscripts = @($rows | ForEach-Object { $_.TranscriptPath })
        $invalidLogLines = 0
        foreach ($stepLog in $stepLogs) {
            if (-not (Test-Path -LiteralPath $stepLog -PathType Leaf)) {
                $invalidLogLines++
                continue
            }
            $invalidLogLines += @(Get-Content -LiteralPath $stepLog |
                Where-Object { $_ -notmatch '^\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3}\] \[(DEBUG|INFO|WARN|ERROR|OUTPUT)\] ' }).Count
        }
        $orchestratorLog = Get-Content `
            -LiteralPath $script:FullResult.OrchestratorLogPath -Raw
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            ($rows.Count -eq 34 -and
                @($rows | Where-Object Status -ne 'Completed').Count -eq 0 -and
                @($rows | Group-Object Sequence |
                    Where-Object Count -gt 1).Count -eq 0 -and
                $stepLogs.Count -eq 34 -and
                @($stepLogs | Where-Object {
                        [string]::IsNullOrWhiteSpace($_) -or
                        -not $_.StartsWith($script:FullResult.StepLogRootPath, [StringComparison]::OrdinalIgnoreCase)
                    }).Count -eq 0 -and
                $invalidLogLines -eq 0 -and
                (Test-Path -LiteralPath $script:FullResult.OrchestratorLogPath -PathType Leaf) -and
                $orchestratorLog -like '*SmartWorkplaceCMDB by WorkplaceCloudHub*' -and
                $orchestratorLog -like '*SmartWorkplaceCMDB execution summary*' -and
                $orchestratorLog -like '*Status   : Completed*' -and
                $stepTranscripts.Count -eq 34 -and
                @($stepTranscripts | Where-Object {
                        [string]::IsNullOrWhiteSpace($_) -or
                        -not $_.StartsWith($script:FullResult.StepTranscriptRootPath, [StringComparison]::OrdinalIgnoreCase) -or
                        -not $_.EndsWith('.transcript.txt', [StringComparison]::OrdinalIgnoreCase) -or
                        -not (Test-Path -LiteralPath $_ -PathType Leaf)
                    }).Count -eq 0 -and
                @($stepTranscripts | Where-Object {
                        (Get-Content -LiteralPath $_ -Raw) -notlike '*Started step*Completed step*'
                    }).Count -eq 0) `
            'Orchestrator step log or transcript is incomplete or inconsistent.'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Purge logs by age and per-script count' {
        $retentionRoot = Join-Path $tempRoot 'Retention'
        $retentionLogRoot = Join-Path $retentionRoot 'LOG-ALL'
        $orchestratorLogFolder = Join-Path $retentionLogRoot 'Orchestration\Logs'
        $runCsvFolder = Join-Path $retentionLogRoot 'Orchestration\Runs'
        $collectorLogFolder = Join-Path $retentionLogRoot `
            'Jobs\SmartWorkplaceCMDB-EntraUsers-Collect'
        foreach ($folder in @($orchestratorLogFolder, $runCsvFolder, $collectorLogFolder)) {
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
        }
        [ordered]@{
            Version=1
            Kind='Fixture'
            TenantKey='contoso-prod'
            OrganizationKey='contoso'
            EnvironmentKey='prod'
            TenantId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
        } | ConvertTo-Json | Set-Content -LiteralPath (
            Join-Path $retentionRoot '.collection-root.json.txt') -Encoding UTF8
        $oldLog = Join-Path $orchestratorLogFolder `
            'SmartWorkplaceCMDB-Orchestrator_HOST_20000101-000000000.log'
        Set-Content -LiteralPath $oldLog -Value 'old' -Encoding UTF8
        (Get-Item -LiteralPath $oldLog).LastWriteTime = (Get-Date).AddDays(-10)
        foreach ($number in 1..5) {
            $seed = Join-Path $collectorLogFolder (
                'SmartWorkplaceCMDB-EntraUsers-Collect_HOST_20260912-01010{0}000_01.log' -f $number)
            Set-Content -LiteralPath $seed -Value "seed $number" -Encoding UTF8
            (Get-Item -LiteralPath $seed).LastWriteTime = (Get-Date).AddMinutes(-10 + $number)
            $transcriptSeed = Join-Path $collectorLogFolder (
                'SmartWorkplaceCMDB-EntraUsers-Collect_HOST_20260912-01010{0}000_01.transcript.txt' -f $number)
            Set-Content -LiteralPath $transcriptSeed -Value "transcript $number" -Encoding UTF8
            (Get-Item -LiteralPath $transcriptSeed).LastWriteTime = (Get-Date).AddMinutes(-10 + $number)
        }
        $loggingConfigPath = Join-Path $tempRoot 'logging.local.json'
        [ordered]@{
            ConfigVersion='0.5.1'
            Logging=[ordered]@{
                Enabled=$true
                OrchestratorLogRetentionDays=7
                StepLogRetentionDays=7
                RunCsvRetentionDays=7
                MaxOrchestratorLogs=3
                MaxStepLogsPerScript=3
                MaxRunCsvFiles=3
            }
        } | ConvertTo-Json -Depth 10 | Set-Content `
            -LiteralPath $loggingConfigPath -Encoding UTF8
        $result = & $orchestrator @identity `
            -TenantConfigPath $loggingConfigPath `
            -DataRootPath $retentionRoot `
            -FixtureRootPath $fixtureRoot `
            -Pipeline EntraUsers
        $rows = @(Import-Csv -LiteralPath $result.LogPath)
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            (-not (Test-Path -LiteralPath $oldLog) -and
                @(Get-ChildItem -LiteralPath $collectorLogFolder -Filter '*.log' -File).Count -le 3 -and
                @(Get-ChildItem -LiteralPath $collectorLogFolder -Filter '*.transcript.txt' -File).Count -le 3 -and
                $rows.Count -eq 2 -and
                $result.OrchestratorLogRetentionDays -eq 7 -and
                $result.MaxStepLogsPerScript -eq 3) `
            'Log retention did not enforce configured age and count safeguards.'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Run bounded individual pipeline' {
        $boundedRoot = Join-Path $tempRoot 'Bounded'
        $result = & $orchestrator @identity `
            -DataRootPath $boundedRoot `
            -FixtureRootPath $fixtureRoot `
            -Pipeline EntraUsers `
            -MaxItems 1
        $rows = @(Import-Csv (
                Join-Path `
                    $result.LatestOutputRootPath `
                    'CMDB\CMDB_Users.csv'
            ))
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            ($result.StepCount -eq 2 -and
                $rows.Count -eq 1 -and
                $result.DataRootPath.StartsWith(([IO.Path]::GetFullPath($boundedRoot) + '\TestRuns\'), [StringComparison]::OrdinalIgnoreCase)) `
            'Bounded individual pipeline is invalid.'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Reject unsafe mode combinations' {
        Assert-SmartWorkplaceCMDBOrchestratorThrow {
            & $orchestrator @identity `
                -Collect `
                -ValidateOnly | Out-Null
        } 'cannot be used together'
        Assert-SmartWorkplaceCMDBOrchestratorThrow {
            & $orchestrator @identity `
                -FixtureRootPath $fixtureRoot `
                -Pipeline Full `
                -MaxItems 1 | Out-Null
        } 'requires an individual source pipeline'
        Assert-SmartWorkplaceCMDBOrchestratorThrow {
            & $orchestrator @identity `
                -FinalizeOnly `
                -Pipeline EntraUsers | Out-Null
        } 'requires -Pipeline Full'
        Assert-SmartWorkplaceCMDBOrchestratorThrow {
            & $orchestrator @identity `
                -FinalizeOnly `
                -Collect | Out-Null
        } '-FinalizeOnly cannot be combined'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Reject incomplete summary mail configuration before collection' {
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
        $configPath = Join-Path $tempRoot 'mail-invalid.local.json'
        [ordered]@{
            ConfigVersion='0.5.1';ProfileKey='test';OrganizationKey='contoso'
            EnvironmentKey='prod';TenantKey='contoso-prod'
            MicrosoftGraph=[ordered]@{
                TenantId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
                ClientId='bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
                CertificateThumbprint='ABCDEF'
            }
            Notifications=[ordered]@{
                Enabled=$true;SendMailMode='Graph';From='sender@example.invalid';To=''
            }
        } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $configPath -Encoding UTF8
        Assert-SmartWorkplaceCMDBOrchestratorThrow {
            & $orchestrator @identity -TenantConfigPath $configPath `
                -DataRootPath (Join-Path $tempRoot 'MailInvalid') `
                -FixtureRootPath $fixtureRoot -ValidateOnly | Out-Null
        } 'Notifications.From and Notifications.To are required'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Reject missing fixtures before output' {
        $missingFixtureRoot = Join-Path $tempRoot 'MissingFixtures'
        New-Item -ItemType Directory -Path $missingFixtureRoot -Force | Out-Null
        $missingOutputRoot = Join-Path $tempRoot 'MissingOutput'
        Assert-SmartWorkplaceCMDBOrchestratorThrow {
            & $orchestrator @identity `
                -DataRootPath $missingOutputRoot `
                -FixtureRootPath $missingFixtureRoot `
                -Pipeline EntraUsers `
                -ValidateOnly | Out-Null
        } 'Required fixture is missing'
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            (-not (Test-Path $missingOutputRoot)) `
            'Missing fixture validation created output.'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Validate centralized Cloud launchers' {
        $expected = [ordered]@{
            'Start-SmartWorkplaceCMDB-Full-Validate.cmd' = '-ValidateOnly'
            'Start-SmartWorkplaceCMDB-Full-Collect.cmd' = '-Collect'
            'Start-SmartWorkplaceCMDB-Full-Finalize.cmd' = '-FinalizeOnly'
            'Start-SmartWorkplaceCMDB-EntraUsers.cmd' = '-Pipeline EntraUsers'
            'Start-SmartWorkplaceCMDB-EntraGroups.cmd' = '-Pipeline EntraGroups'
            'Start-SmartWorkplaceCMDB-EntraDevices.cmd' = '-Pipeline EntraDevices'
            'Start-SmartWorkplaceCMDB-IntuneDevices.cmd' = '-Pipeline IntuneDevices'
            'Start-SmartWorkplaceCMDB-M365SubscribedSkus.cmd' = '-Pipeline M365SubscribedSkus'
            'Start-SmartWorkplaceCMDB-M365UserLicenses.cmd' = '-Pipeline M365UserLicenses'
            'Start-SmartWorkplaceCMDB-ExchangeOnlineMailboxes.cmd' = '-Pipeline ExchangeOnlineMailboxes'
            'Start-SmartWorkplaceCMDB-CuratedOnly.cmd' = '-Pipeline CuratedOnly'
        }
        $files = @(Get-ChildItem $launcherRoot -Filter '*.cmd' -File)
        $forbiddenPattern = @(('Smart' + 'M365'),
            ('SmartWorkplace' + 'Dashboard')) -join '|'
        $invalid = 0
        foreach ($entry in $expected.GetEnumerator()) {
            $path = Join-Path $launcherRoot $entry.Key
            if (-not (Test-Path $path -PathType Leaf)) {
                $invalid++
                continue
            }
            $content = Get-Content $path -Raw
            if ($content -notlike '*PowerShell\7\pwsh.exe*' -or
                $content -notlike '*SmartWorkplaceCMDB-Orchestrator.ps1*' -or
                $content -notlike "*$($entry.Value)*" -or
                $content -notlike '*%**' -or
                $content -match $forbiddenPattern) {
                $invalid++
            }
        }
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            ($files.Count -eq $expected.Count -and $invalid -eq 0) `
            'Centralized Cloud launcher contract is invalid.'
    }
}
finally {
    $resolved = [IO.Path]::GetFullPath($tempRoot)
    $tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    if ($resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and
        (Test-Path $resolved)) {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

Write-Information (
    "SmartWorkplaceCMDB orchestrator tests completed. Version={0}; Passed={1}; Failed={2}" -f
    $ScriptVersion,
    $script:Passed,
    $script:Failed
) -InformationAction Continue
if ($script:Failed -gt 0) {
    exit 1
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBGe7MkRsjam/o5
# WwAxb4m4KUQkjv9TK4V4uPqbBsle36CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBB7pcHTCt4n/LnqWcW/44V
# BUH3lCacR5+hG69TxP8ZxzANBgkqhkiG9w0BAQEFAASCAYCPpJAjKeI2PZeM6uX1
# He/umHjbnbWNfzwT2F11saOqWXXZw5+wyV09tCkip5oOWS9XcCph33IG5UJ9yvSU
# bFA67eiRl+N8KVr4l1v8zuD0Ey0bx2d2wS1uiPJPzCx29rb2AGQWayTl+z3PXzKy
# DtaW0c7ZHJCxy7HCC+RQwund767huohJAymeK3OumhjpymrDWemyHegl/A29vuTV
# /RDEsrQRwjSkuZVrCVehnjpxiqOrgz9UYPcW/yujYIOoGM7hO1qdoctvPJBKqMRj
# xZWe0GeCfWRxrtkP4oOKf1IIxp0806Gr34aWa5abPIVFD9Fml3yLnICWiZuXPnAQ
# KNrCFtNlTtVuykJS6BNHzhGsqP3VlojHisqxLCy/U3/qbFnWhXCGdgtodmA8Rwjw
# B0/pntYqhPHVlSj2NhmIXBauxu8d1M9eZ+NvouAmpeuvzGwZAPRuM/WGU46rjJSx
# xPOOE6pMsO0YWhPdhKl15S6iu4QnDLqOYFNQksqTKfiKcnc=
# SIG # End signature block
