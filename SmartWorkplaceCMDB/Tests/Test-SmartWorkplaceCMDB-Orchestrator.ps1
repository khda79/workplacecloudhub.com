<#
.SYNOPSIS
Runs offline SmartWorkplaceCMDB orchestrator and launcher tests.

.VERSION
1.1.6
#>
[CmdletBinding()]
param()

$ScriptVersion = '1.1.6'
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
                $result.StepCount -eq 12 -and
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
                $result[0].ScriptVersion -eq '1.1.5' -and
                $messages -contains ' SmartWorkplaceCMDB by WorkplaceCloudHub' -and
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
                $script:FullResult.StepCount -eq 25 -and
                $script:FullResult.FailedStepCount -eq 0 -and
                -not $script:FullResult.SummaryEmailEligible -and
                $script:FullResult.SummaryEmailStatus -eq 'NotApplicable' -and
                [string]::IsNullOrWhiteSpace($script:FullResult.SummaryEmailHtmlPath)) `
            'Full fixture orchestration status is invalid.'
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
            ($results.Count -eq 24 -and
                @($results | Where-Object Status -ne 'Valid').Count -eq 0 -and
                $rawResults.Count -eq 15 -and
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
            ($rows.Count -eq 25 -and
                @($rows | Where-Object Status -ne 'Completed').Count -eq 0 -and
                @($rows | Group-Object Sequence |
                    Where-Object Count -gt 1).Count -eq 0 -and
                $stepLogs.Count -eq 25 -and
                @($stepLogs | Where-Object {
                        [string]::IsNullOrWhiteSpace($_) -or
                        -not $_.StartsWith($script:FullResult.StepLogRootPath, [StringComparison]::OrdinalIgnoreCase)
                    }).Count -eq 0 -and
                $invalidLogLines -eq 0 -and
                (Test-Path -LiteralPath $script:FullResult.OrchestratorLogPath -PathType Leaf) -and
                $orchestratorLog -like '*SmartWorkplaceCMDB by WorkplaceCloudHub*' -and
                $orchestratorLog -like '*SmartWorkplaceCMDB execution summary*' -and
                $orchestratorLog -like '*Status   : Completed*' -and
                $stepTranscripts.Count -eq 25 -and
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
            Join-Path $retentionRoot '.collection-root.json') -Encoding UTF8
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBro6yGxm60xLg/
# az9yGaxnjpY4J0uvB7N5LHfVUlqnlKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC1+QpDjGM/tgjBES9X/vWy
# tIE7Exqb6THyJs5fEM1K0zANBgkqhkiG9w0BAQEFAASCAYAHPw1HkMEuyCMxff3C
# 2Cp5Db8wmjiEAfiXKNg8wslhapazRo1LDBadX4dqnEwayBcCkFCMjcKV+tLh8EEx
# cXzYD6Bf1q4I/yYLk239KHcU+Nmkna438G48xZxQCTGC8/C0LQCPsMkY82scu7dc
# D0fJ1QcKGw/ZlJpGZVmfFbI3fA6HpDstRqc1eZHTuvqQicYNh6JW55jF4IqtZPR0
# 8a9Hc0fDeLsvpuyRPZG+323NC4GwrG22Sm9qj/WZXgvUnm9Ij53cJ2DqXscvJHSB
# U+WvE7fJ1oG9mbn0Gek0pE/0VTgwm4+r/LaKAKKkt93812PDXbu53cr+T0Twfr6b
# tXMRr62eU6mTnVVY07S2lflMgZ8yliPJRV3+uZtPgzdv7ioke3R9fJRf1VdYzwE9
# JRCGtVv6oIolIlYAwIEp8OfhBoxdjwr/5zUV9+8m5iJuv/V+KQvikDRCG0W853Ke
# OLEvFuB+014Vx/0/lRJ/IVswUIGb/i9RjBkD4UBYz5xxzdA=
# SIG # End signature block
