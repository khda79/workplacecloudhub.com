<#
.SYNOPSIS
Validates Intune update reporting and Endpoint Analytics with synthetic data.

.VERSION
1.0.5
#>
[CmdletBinding()]
param()

$ScriptVersion = '1.0.5'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$passed = 0
$failed = 0

function Invoke-IntuneAnalyticsTest {
    param([string]$Name, [scriptblock]$Test)
    try {
        & $Test
        $script:passed++
        Write-Information "PASS $Name" -InformationAction Continue
    }
    catch {
        $script:failed++
        Write-Information "FAIL $Name - $($_.Exception.Message)" -InformationAction Continue
    }
}

function Assert-IntuneAnalyticsTrue {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$collector = Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneAnalytics-Collect.ps1'
$normalizer = Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneAnalytics-Normalize.ps1'
$fixture = Join-Path $PSScriptRoot 'Fixtures\IntuneAnalytics.sample.json'
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("SmartWorkplaceCMDB-IntuneAnalytics-{0}" -f [guid]::NewGuid().ToString('N'))
$identity = @{
    Tenant = 'audit'
    OrganizationKey = 'contoso'
    EnvironmentKey = 'test'
    TenantKey = 'contoso-test'
    TenantId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
    DataRootPath = $tempRoot
    NoConfigWrite = $true
}
$parseTokens = $null
$parseErrors = $null
$collectorAst = [System.Management.Automation.Language.Parser]::ParseFile($collector, [ref]$parseTokens, [ref]$parseErrors)
if (@($parseErrors).Count -gt 0) { throw 'Intune analytics collector has PowerShell parse errors.' }
foreach ($definition in @($collectorAst.EndBlock.Statements | Where-Object {
    $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
    $_.Name -in @('Get-IntuneExportErrorDetails', 'Invoke-IntuneExportRead', 'Invoke-IntuneReportExport', 'Get-UpgradeEligibilityGraphRows', 'Publish-IntuneAnalyticsBatch')
})) {
    . ([scriptblock]::Create($definition.Extent.Text))
}
try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    Invoke-IntuneAnalyticsTest 'Validate without creating output' {
        & $collector @identity -InputJsonPath $fixture -ValidateOnly | Out-Null
        Assert-IntuneAnalyticsTrue (-not (Test-Path (Join-Path $tempRoot 'DATA-LAST'))) 'ValidateOnly created output.'
    }
    Invoke-IntuneAnalyticsTest 'Collect independent report snapshots' {
        $result = & $collector @identity -InputJsonPath $fixture
        Assert-IntuneAnalyticsTrue ($result.WindowsUpdateRowCount -eq 2 -and $result.EndpointAnalyticsDeviceCount -eq 2 -and $result.UpgradeEligibilityDeviceCount -eq 2) 'Unexpected report row counts.'
        $raw = Join-Path $tempRoot 'DATA-LAST\Raw\Intune'
        Assert-IntuneAnalyticsTrue (@(Import-Csv (Join-Path $raw 'Intune_WindowsUpdateAlerts.csv')).Count -eq 2) 'Windows update report snapshot is invalid.'
        Assert-IntuneAnalyticsTrue (@(Import-Csv (Join-Path $raw 'Intune_EndpointAnalyticsDeviceScores.csv')).Count -eq 2) 'Endpoint Analytics snapshot is invalid.'
        $eligibility = @(Import-Csv (Join-Path $raw 'Intune_EndpointAnalyticsUpgradeEligibility.csv'))
        Assert-IntuneAnalyticsTrue ($eligibility.Count -eq 2) 'Upgrade eligibility snapshot is invalid.'
        Assert-IntuneAnalyticsTrue (@($eligibility.UpgradeEligibility | Sort-Object) -join ',' -eq 'capable,notCapable') 'Upgrade eligibility values were not preserved canonically.'
        Assert-IntuneAnalyticsTrue (@($eligibility.DeviceIdSource | Sort-Object -Unique) -join ',' -eq 'ReportedDeviceId') 'Upgrade eligibility device ID provenance is invalid.'
    }
    Invoke-IntuneAnalyticsTest 'Normalize dedicated Power BI facts' {
        & $normalizer @identity | Out-Null
        $powerBi = Join-Path $tempRoot 'DATA-LAST\PowerBI'
        $updates = @(Import-Csv (Join-Path $powerBi 'FactWindowsUpdateAlert.csv'))
        $analytics = @(Import-Csv (Join-Path $powerBi 'FactEndpointAnalyticsDevice.csv'))
        $eligibility = @(Import-Csv (Join-Path $powerBi 'FactEndpointAnalyticsUpgradeEligibility.csv'))
        Assert-IntuneAnalyticsTrue ($updates.Count -eq 2 -and $updates[0].TenantUpdateAlertKey -match '\|update-alert\|') 'Windows update fact grain is invalid.'
        Assert-IntuneAnalyticsTrue ($analytics.Count -eq 2 -and $analytics[0].EndpointAnalyticsScore -eq '75') 'Endpoint Analytics fact grain is invalid.'
        Assert-IntuneAnalyticsTrue ($eligibility.Count -eq 2 -and $eligibility[0].TenantUpgradeEligibilityDeviceKey -match '\|upgrade-eligibility\|') 'Upgrade eligibility fact grain is invalid.'
    }
    Invoke-IntuneAnalyticsTest 'Collapse repeated update evidence without losing distinct states' {
        $duplicateIdentity = @{} + $identity
        $duplicateIdentity.DataRootPath = Join-Path $tempRoot 'DuplicateUpdateEvidence'
        $duplicateFixture = Join-Path $tempRoot 'duplicate-update-evidence.json'
        $fixtureObject = Get-Content -Raw -LiteralPath $fixture | ConvertFrom-Json
        $alertRows = [object[]]$fixtureObject.windowsUpdateAlerts
        $exactDuplicate = ConvertFrom-Json (ConvertTo-Json -InputObject $alertRows[0] -Depth 8)
        $distinctEvidence = ConvertFrom-Json (ConvertTo-Json -InputObject $alertRows[0] -Depth 8)
        $distinctEvidence.currentDeviceUpdateStatus = 'rollbackInitiated'
        $distinctEvidence.latestAlertMessage = 'Feature update rollback initiated'
        $fixtureObject.windowsUpdateAlerts = [object[]]@($alertRows + $exactDuplicate + $distinctEvidence)
        ConvertTo-Json -InputObject $fixtureObject -Depth 12 | Set-Content -LiteralPath $duplicateFixture -Encoding UTF8
        $collectorWarnings = @()
        & $collector @duplicateIdentity -InputJsonPath $duplicateFixture `
            -WarningVariable collectorWarnings | Out-Null
        $warningText = @($collectorWarnings | ForEach-Object { [string]$_ }) -join "`n"
        $rawPath = Join-Path $duplicateIdentity.DataRootPath 'DATA-LAST\Raw\Intune\Intune_WindowsUpdateAlerts.csv'
        $rawRows = @(Import-Csv -LiteralPath $rawPath)
        & $normalizer @duplicateIdentity | Out-Null
        $factRows = @(Import-Csv (Join-Path $duplicateIdentity.DataRootPath 'DATA-LAST\PowerBI\FactWindowsUpdateAlert.csv'))
        $collidingEvidence = @($factRows | Where-Object {
                $_.DeviceId -eq $alertRows[0].deviceId -and
                $_.PolicyId -eq $alertRows[0].policyId -and
                $_.EventDateTimeUTC -eq '2026-09-10T08:00:00.0000000Z'
            })
        Assert-IntuneAnalyticsTrue ($rawRows.Count -eq 3) 'Repeated update evidence was not collapsed at collection time.'
        Assert-IntuneAnalyticsTrue ($factRows.Count -eq 3 -and $collidingEvidence.Count -eq 2) 'Distinct update evidence sharing an event key was lost.'
        Assert-IntuneAnalyticsTrue (@($collidingEvidence.TenantUpdateAlertKey | Sort-Object -Unique).Count -eq 2) 'Distinct update evidence keys are not unique.'
        Assert-IntuneAnalyticsTrue ($warningText -like '*1 repeated Windows update evidence row(s)*1 event key(s) with distinct evidence*') 'Update evidence reconciliation warning is missing.'
    }
    Invoke-IntuneAnalyticsTest 'Preserve multiple legitimate events for one device and policy' {
        $eventIdentity = @{} + $identity
        $eventIdentity.DataRootPath = Join-Path $tempRoot 'MultipleUpdateEvents'
        $eventFixture = Join-Path $tempRoot 'multiple-update-events.json'
        $fixtureObject = Get-Content -Raw -LiteralPath $fixture | ConvertFrom-Json
        $alertRows = [object[]]$fixtureObject.windowsUpdateAlerts
        $laterEvent = ConvertFrom-Json (ConvertTo-Json -InputObject $alertRows[0] -Depth 8)
        $laterEvent.eventDateTimeUTC = '2026-09-12T08:00:00Z'
        $laterEvent.lastWUScanTimeUTC = '2026-09-12T07:30:00Z'
        $laterEvent.aggregateState = 'success'
        $laterEvent.currentDeviceUpdateStatus = 'installed'
        $laterEvent.latestAlertMessage = 'Feature update installed'
        $fixtureObject.windowsUpdateAlerts = [object[]]@($alertRows + $laterEvent)
        ConvertTo-Json -InputObject $fixtureObject -Depth 12 | Set-Content -LiteralPath $eventFixture -Encoding UTF8
        & $collector @eventIdentity -InputJsonPath $eventFixture | Out-Null
        & $normalizer @eventIdentity | Out-Null
        $factRows = @(Import-Csv (Join-Path $eventIdentity.DataRootPath 'DATA-LAST\PowerBI\FactWindowsUpdateAlert.csv'))
        $devicePolicyRows = @($factRows | Where-Object {
                $_.DeviceId -eq $alertRows[0].deviceId -and $_.PolicyId -eq $alertRows[0].policyId
            })
        Assert-IntuneAnalyticsTrue ($factRows.Count -eq 3 -and $devicePolicyRows.Count -eq 2) 'A legitimate update event was collapsed.'
    }
    Invoke-IntuneAnalyticsTest 'Keep update evidence keys stable across collection times' {
        $firstIdentity = @{} + $identity
        $firstIdentity.DataRootPath = Join-Path $tempRoot 'StableKeyFirst'
        $secondIdentity = @{} + $identity
        $secondIdentity.DataRootPath = Join-Path $tempRoot 'StableKeySecond'
        & $collector @firstIdentity -InputJsonPath $fixture | Out-Null
        & $normalizer @firstIdentity | Out-Null
        Start-Sleep -Milliseconds 25
        & $collector @secondIdentity -InputJsonPath $fixture | Out-Null
        & $normalizer @secondIdentity | Out-Null
        $firstKeys = @(Import-Csv (Join-Path $firstIdentity.DataRootPath 'DATA-LAST\PowerBI\FactWindowsUpdateAlert.csv') | ForEach-Object TenantUpdateAlertKey | Sort-Object)
        $secondKeys = @(Import-Csv (Join-Path $secondIdentity.DataRootPath 'DATA-LAST\PowerBI\FactWindowsUpdateAlert.csv') | ForEach-Object TenantUpdateAlertKey | Sort-Object)
        Assert-IntuneAnalyticsTrue (($firstKeys -join "`n") -ceq ($secondKeys -join "`n")) 'Update evidence keys changed with SourceCollectedDateTime.'
    }
    Invoke-IntuneAnalyticsTest 'Treat the Intune minus-one score sentinel as unavailable' {
        $sentinelIdentity = @{} + $identity
        $sentinelIdentity.DataRootPath = Join-Path $tempRoot 'Sentinel'
        $sentinelFixture = Join-Path $tempRoot 'sentinel.json'
        $fixtureObject = Get-Content -Raw -LiteralPath $fixture | ConvertFrom-Json
        $fixtureObject.endpointAnalyticsDeviceScores[0].startupPerformanceScore = -1
        ConvertTo-Json -InputObject $fixtureObject -Depth 12 | Set-Content -LiteralPath $sentinelFixture -Encoding UTF8
        & $collector @sentinelIdentity -InputJsonPath $sentinelFixture | Out-Null
        $scores = @(Import-Csv (Join-Path $sentinelIdentity.DataRootPath 'DATA-LAST\Raw\Intune\Intune_EndpointAnalyticsDeviceScores.csv'))
        Assert-IntuneAnalyticsTrue ($scores[0].StartupPerformanceScore -eq '') 'The minus-one score sentinel was not published as unavailable.'
    }
    Invoke-IntuneAnalyticsTest 'Reconcile duplicate Endpoint Analytics device rows' {
        $duplicateIdentity = @{} + $identity
        $duplicateIdentity.DataRootPath = Join-Path $tempRoot 'DuplicateDevice'
        $duplicateFixture = Join-Path $tempRoot 'duplicate-device.json'
        $fixtureObject = Get-Content -Raw -LiteralPath $fixture | ConvertFrom-Json
        $scoreRows = [object[]]$fixtureObject.endpointAnalyticsDeviceScores
        $duplicate = ConvertFrom-Json (ConvertTo-Json -InputObject $scoreRows[0] -Depth 8)
        $duplicate.deviceName = 'Z Device Alias'
        $duplicate.endpointAnalyticsScore = 80
        $duplicate.startupPerformanceScore = -1
        $fixtureObject.endpointAnalyticsDeviceScores = [object[]]@($scoreRows + $duplicate)
        ConvertTo-Json -InputObject $fixtureObject -Depth 12 | Set-Content -LiteralPath $duplicateFixture -Encoding UTF8
        & $collector @duplicateIdentity -InputJsonPath $duplicateFixture | Out-Null
        $scores = @(Import-Csv (Join-Path $duplicateIdentity.DataRootPath 'DATA-LAST\Raw\Intune\Intune_EndpointAnalyticsDeviceScores.csv'))
        $reconciled = @($scores | Where-Object DeviceId -eq $scoreRows[0].deviceId)[0]
        Assert-IntuneAnalyticsTrue ($scores.Count -eq 2) 'Duplicate Endpoint Analytics devices were not consolidated.'
        Assert-IntuneAnalyticsTrue ($reconciled.DeviceName -eq $scoreRows[0].deviceName -and $reconciled.EndpointAnalyticsScore -eq '80' -and $reconciled.StartupPerformanceScore -eq '70') 'Duplicate Endpoint Analytics values were not reconciled deterministically.'
    }
    Invoke-IntuneAnalyticsTest 'Canonicalize documented numeric upgrade eligibility values' {
        $numericIdentity = @{} + $identity
        $numericIdentity.DataRootPath = Join-Path $tempRoot 'NumericEligibility'
        $numericFixture = Join-Path $tempRoot 'numeric-eligibility.json'
        $fixtureObject = Get-Content -Raw -LiteralPath $fixture | ConvertFrom-Json
        $fixtureObject.endpointAnalyticsUpgradeEligibility[0].upgradeEligibility = 2
        $fixtureObject.endpointAnalyticsUpgradeEligibility[1].upgradeEligibility = 3
        ConvertTo-Json -InputObject $fixtureObject -Depth 12 | Set-Content -LiteralPath $numericFixture -Encoding UTF8
        & $collector @numericIdentity -InputJsonPath $numericFixture | Out-Null
        $eligibility = @(Import-Csv (Join-Path $numericIdentity.DataRootPath 'DATA-LAST\Raw\Intune\Intune_EndpointAnalyticsUpgradeEligibility.csv'))
        Assert-IntuneAnalyticsTrue (@($eligibility.UpgradeEligibility | Sort-Object) -join ',' -eq 'capable,notCapable') 'Numeric Graph enum values were not mapped to capable and notCapable.'
    }
    Invoke-IntuneAnalyticsTest 'Reject conflicting eligibility for the same Intune device ID' {
        $conflictIdentity = @{} + $identity
        $conflictIdentity.DataRootPath = Join-Path $tempRoot 'ConflictingEligibility'
        $conflictFixture = Join-Path $tempRoot 'conflicting-eligibility.json'
        $fixtureObject = Get-Content -Raw -LiteralPath $fixture | ConvertFrom-Json
        $eligibilityRows = [object[]]$fixtureObject.endpointAnalyticsUpgradeEligibility
        $duplicate = ConvertFrom-Json (ConvertTo-Json -InputObject $eligibilityRows[0] -Depth 8)
        $duplicate.id = 'metric-device-conflict'
        $duplicate.upgradeEligibility = 'capable'
        $fixtureObject.endpointAnalyticsUpgradeEligibility = [object[]]@($eligibilityRows + $duplicate)
        ConvertTo-Json -InputObject $fixtureObject -Depth 12 | Set-Content -LiteralPath $conflictFixture -Encoding UTF8
        $thrown = $false
        try { & $collector @conflictIdentity -InputJsonPath $conflictFixture | Out-Null } catch { $thrown = $true }
        Assert-IntuneAnalyticsTrue $thrown 'Conflicting eligibility values did not stop publication.'
        Assert-IntuneAnalyticsTrue (-not (Test-Path (Join-Path $conflictIdentity.DataRootPath 'DATA-LAST\Raw\Intune\Intune_EndpointAnalyticsUpgradeEligibility.csv'))) 'Conflicting eligibility produced a latest snapshot.'
    }
    Invoke-IntuneAnalyticsTest 'Bound both report families independently' {
        $bounded = @{} + $identity
        $bounded.DataRootPath = Join-Path $tempRoot 'Bounded'
        $result = & $collector @bounded -InputJsonPath $fixture -MaxItems 1
        foreach ($path in @($result.PublishedPath)) {
            Assert-IntuneAnalyticsTrue (@(Import-Csv -LiteralPath $path).Count -eq 1) "$(Split-Path $path -Leaf) was not bounded."
        }
    }
    Invoke-IntuneAnalyticsTest 'Honor Retry-After while reading the same export job' {
        $script:readAttempt = 0
        $script:fakeClock = [datetime]'2026-09-22T12:00:00Z'
        $script:readDelays = @()
        $value = Invoke-IntuneExportRead -ReportName 'EADeviceScoresV2' -Operation 'job status' `
            -RetryWindowSeconds 100 -RequestScript {
                $script:readAttempt++
                if ($script:readAttempt -eq 1) {
                    $error503 = [Exception]::new('503 ServiceUnavailable')
                    $error503.Data['Retry-After'] = '40'
                    throw $error503
                }
                return 'completed'
            } -ClockScript { $script:fakeClock } -SleepScript {
                param($seconds)
                $script:readDelays += $seconds
                $script:fakeClock = $script:fakeClock.AddSeconds($seconds)
            } -WarningAction SilentlyContinue
        Assert-IntuneAnalyticsTrue ($value -eq 'completed' -and $script:readAttempt -eq 2) 'The same read was not retried.'
        Assert-IntuneAnalyticsTrue ($script:readDelays.Count -eq 1 -and $script:readDelays[0] -eq 40) 'Retry-After was not honored.'
    }
    Invoke-IntuneAnalyticsTest 'Stop when Retry-After exceeds the read window' {
        $script:readAttempt = 0
        $script:fakeClock = [datetime]'2026-09-22T12:00:00Z'
        $thrown = $false
        try {
            Invoke-IntuneExportRead -ReportName 'EADeviceScoresV2' -Operation 'archive download' `
                -RetryWindowSeconds 100 -RequestScript {
                    $script:readAttempt++
                    $error503 = [Exception]::new('503 ServiceUnavailable')
                    $error503.Data['Retry-After'] = '120'
                    throw $error503
                } -ClockScript { $script:fakeClock } -SleepScript { throw 'Sleep must not be called.' } | Out-Null
        }
        catch { $thrown = $_.Exception.Message -like '*Retry window exhausted*' }
        Assert-IntuneAnalyticsTrue ($thrown -and $script:readAttempt -eq 1) 'An early retry violated Retry-After.'
    }
    Invoke-IntuneAnalyticsTest 'Fail fast on Graph authorization errors' {
        $script:readAttempt = 0
        $thrown = $false
        try {
            Invoke-IntuneExportRead -ReportName 'EADeviceScoresV2' -Operation 'job status' -RequestScript {
                $script:readAttempt++
                throw [Exception]::new('403 Forbidden')
            } -SleepScript { throw 'Sleep must not be called.' } | Out-Null
        }
        catch { $thrown = $_.Exception.Message -like '*Status=403*' }
        Assert-IntuneAnalyticsTrue ($thrown -and $script:readAttempt -eq 1) 'An authorization failure was retried.'
    }
    Invoke-IntuneAnalyticsTest 'Do not retry a failed export-job POST' {
        $script:postCalls = 0
        function Invoke-MgGraphRequest {
            param($Method, $Uri, $Body, $ContentType, $OutputType, $ErrorAction)
            $script:postCalls++
            throw [Exception]::new('503 ServiceUnavailable')
        }
        $thrown = $false
        try {
            Invoke-IntuneReportExport -ReportName 'EADeviceScoresV2' -Select @('DeviceId') | Out-Null
        }
        catch { $thrown = $_.Exception.Message -like '*POST outcome is unknown*' }
        Assert-IntuneAnalyticsTrue ($thrown -and $script:postCalls -eq 1) 'Export job creation was repeated after an ambiguous 503.'
        Remove-Item Function:\Invoke-MgGraphRequest
    }
    Invoke-IntuneAnalyticsTest 'Do not publish empty readiness after two transient route failures' {
        function Invoke-GraphCollection { param($Uri) throw [Exception]::new('503 ServiceUnavailable') }
        $thrown = $false
        try { Get-UpgradeEligibilityGraphRows | Out-Null }
        catch { $thrown = $_.Exception.Message -like '*must not be published as an empty snapshot*' }
        Assert-IntuneAnalyticsTrue $thrown 'Two transient Work From Anywhere failures were treated as an empty report.'
        Remove-Item Function:\Invoke-GraphCollection
    }
    Invoke-IntuneAnalyticsTest 'Rollback all three snapshots if batch promotion fails' {
        Import-Module (Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1') -Force
        $contractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
        $contract = Get-SmartWorkplaceCMDBTableContract -Path $contractPath
        $names = @('Intune_WindowsUpdateAlerts.csv', 'Intune_EndpointAnalyticsDeviceScores.csv', 'Intune_EndpointAnalyticsUpgradeEligibility.csv')
        $rawRoot = Join-Path $tempRoot 'DATA-LAST\Raw\Intune'
        $tables = @{}
        $latestPaths = @{}
        $rowsByTable = @{}
        $before = @{}
        foreach ($name in $names) {
            $tables[$name] = @($contract.tables | Where-Object name -eq $name)[0]
            $latestPaths[$name] = Join-Path $rawRoot $name
            $rowsByTable[$name] = @(Import-Csv -LiteralPath $latestPaths[$name])
            $before[$name] = @((Get-FileHash -LiteralPath $latestPaths[$name]).Hash, (Get-FileHash -LiteralPath ($latestPaths[$name] + '.status.json.txt')).Hash)
        }
        $paths = [pscustomobject]@{
            DataRootPath = $tempRoot
            DataAllRootPath = Join-Path $tempRoot 'DATA-ALL'
            LatestOutputRootPath = Join-Path $tempRoot 'DATA-LAST'
            TenantKey = $identity.TenantKey
            OrganizationKey = $identity.OrganizationKey
            EnvironmentKey = $identity.EnvironmentKey
            TenantId = $identity.TenantId
        }
        $thrown = $false
        try {
            Publish-IntuneAnalyticsBatch -Paths $paths -TableNames $names -Tables $tables `
                -LatestPaths $latestPaths -RowsByTable $rowsByTable -ContractPath $contractPath `
                -Fixture -BeforePromotionScript { param($destination, $index) if ($index -eq 3) { throw 'Injected promotion failure.' } } | Out-Null
        }
        catch { $thrown = $_.Exception.Message -like '*Injected promotion failure*' }
        Assert-IntuneAnalyticsTrue $thrown 'The injected promotion failure did not stop the batch.'
        foreach ($name in $names) {
            $after = @((Get-FileHash -LiteralPath $latestPaths[$name]).Hash, (Get-FileHash -LiteralPath ($latestPaths[$name] + '.status.json.txt')).Hash)
            Assert-IntuneAnalyticsTrue (($before[$name] -join '|') -eq ($after -join '|')) "The failed batch changed '$name' or its state."
        }
    }
}
finally {
    if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
Write-Information "SmartWorkplaceCMDB Intune analytics tests completed. Version=$ScriptVersion; Passed=$passed; Failed=$failed" -InformationAction Continue
if ($failed -gt 0) { exit 1 }

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC+DFBZqivfEQM/
# qbAGBaN3dIZMW57T3CxfJvSYB8Wq26CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBbmESLdM0/tQC/ClRVrZYd
# 3UKozqx7exJ3g6R5wKdyaTANBgkqhkiG9w0BAQEFAASCAYCZPBRDGFTjWGtoBd18
# v6OUIpVJZW75V5XcDwTCfocfIxYkppikM1qbycPnuPCIv1lseiHy24MpprSetU6c
# VCn0BaJ/CwYPnoqCjAbHaUhckI2+S6glNxBCeUOrDrZTduVM4Fkrt5J/9soblicW
# IvgGGRTb1dmakJvB3K+zWtL0tkSAOP6igzrtMoxtFK/Hkl595ZEOn8jDpCnMj9hu
# uyFmoX2mjQUp23K/KEXxIDfNgcTO+0o5CWppAMNZHYFuAx4CekGFyOc34YXXU4kv
# 0KrCDOnc24dko6iuslue7mZ1lruM5+dIXsKteZdB2q9DLx3hpnvmI9Fd3fVMJgjQ
# LdrOFQ28do0vWmFcxBFd/QICjbSnl0GxyCS9KORnCWak/jx2geBCilD+ItvWuFOu
# fwZgqWZL+5TUBn2L5IYxM1SFyfXQ1zek+hrwKNqkZlGLSzvD2FxL5MLQuJZLlxRs
# KzkF3a+7XxdW5u3WGUZ8K3r0gq32R4TludVRtJTbAgcsg/w=
# SIG # End signature block
