#requires -Version 5.1

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Assert-True {
    param([bool]$Condition,[string]$Message)
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

function Assert-Equal {
    param($Actual,$Expected,[string]$Message)
    if ([string]$Actual -ne [string]$Expected) {
        throw "ASSERT FAILED: $Message. Expected='$Expected'; Actual='$Actual'."
    }
}

$toolkitRoot = Split-Path -Parent $PSScriptRoot
$launcherPath = Join-Path $toolkitRoot 'Scripts\SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1'
$endpointPath = Join-Path $toolkitRoot 'Scripts\SmartM365-Invoke-IntuneHybridJoinRepair.ps1'

$tokens = $null
$parseErrors = $null
$launcherAst = [System.Management.Automation.Language.Parser]::ParseFile($launcherPath,[ref]$tokens,[ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw "Launcher parsing failed: $($parseErrors[0].Message)" }

foreach ($functionName in @('Get-ComputerListKey','Get-PostCycleCloudRefreshRows','Merge-ScopedInventoryMap','Test-AdInventoryRefreshDue','Get-AdaptiveCycleDelaySeconds')) {
    $functions = @($launcherAst.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName
    },$true))
    Assert-Equal $functions.Count 1 "helper $functionName exists exactly once"
    . ([scriptblock]::Create($functions[0].Extent.Text))
}

$cloudCandidates = @(Get-PostCycleCloudRefreshRows -Rows @(
    [pscustomobject]@{ Computer='FR-01'; Status='ADMIN_SHARE_UNREACHABLE' },
    [pscustomobject]@{ Computer='FR-02'; Status='SKIPPED_BY_STATUS_BACKOFF' },
    [pscustomobject]@{ Computer='FR-03'; Status='DRYRUN_READY' },
    [pscustomobject]@{ Computer='FR-04'; Status='WAITING_FOR_INTUNE_ENROLLMENT' }
))
Assert-Equal $cloudCandidates.Count 1 'only cloud-change-capable rows enter the post-cycle Graph scope'
Assert-Equal $cloudCandidates[0].Computer 'FR-04' 'the actionable device remains in scope'

$existingMap = @{
    'FR-SCOPE' = [pscustomobject]@{ Marker='stale' }
    'FR-KEEP' = [pscustomobject]@{ Marker='keep' }
    '__SMARTM365_INVENTORY_CHECKED__' = [pscustomobject]@{}
}
$refreshedMap = @{
    'FR-SCOPE' = [pscustomobject]@{ Marker='fresh' }
    '__SMARTM365_INVENTORY_CHECKED__' = [pscustomobject]@{}
}
$mergedMap = Merge-ScopedInventoryMap -ExistingMap $existingMap -RefreshedMap $refreshedMap -ScopedComputers @('fr-scope.contoso.test','FR-ABSENT')
Assert-Equal $mergedMap.Count 2 'scoped merge retains unrelated inventory and removes scoped absences'
Assert-Equal $mergedMap['FR-SCOPE'].Marker 'fresh' 'scoped refreshed device replaces stale data'
Assert-Equal $mergedMap['FR-KEEP'].Marker 'keep' 'unrelated cached device is preserved'
Assert-True (-not $mergedMap.ContainsKey('__SMARTM365_INVENTORY_CHECKED__')) 'inventory sentinel is not merged into the complete map'

$now = [datetime]::SpecifyKind([datetime]'2026-07-29T12:00:00',[DateTimeKind]::Utc)
Assert-True (Test-AdInventoryRefreshDue -LastRefreshUtc $null -FreshnessHours 12 -NowUtc $now) 'missing AD refresh timestamp is due'
Assert-True (-not (Test-AdInventoryRefreshDue -LastRefreshUtc $now.AddHours(-11) -FreshnessHours 12 -NowUtc $now)) 'fresh AD cache is reused'
Assert-True (Test-AdInventoryRefreshDue -LastRefreshUtc $now.AddHours(-12) -FreshnessHours 12 -NowUtc $now) 'AD cache refreshes when TTL expires'

$allBackoffRows = @(
    [pscustomobject]@{ Status='SKIPPED_BY_STATUS_BACKOFF'; EffectiveStatus='SKIPPED_BY_STATUS_BACKOFF'; BackoffUntilUtc=$now.AddMinutes(90).ToString('o') },
    [pscustomobject]@{ Status='SKIPPED_BY_TECH_RUN_GUARD_STARTED_NO_RESULT'; EffectiveStatus='SKIPPED_BY_TECH_RUN_GUARD_STARTED_NO_RESULT'; BackoffUntilUtc=$now.AddMinutes(120).ToString('o') }
)
$adaptiveDelay = Get-AdaptiveCycleDelaySeconds -Rows $allBackoffRows -MinimumDelaySeconds 60 -NowUtc $now
Assert-Equal $adaptiveDelay 5400 'all-backoff cycle waits for the earliest future expiry'
$mixedDelay = Get-AdaptiveCycleDelaySeconds -Rows @($allBackoffRows[0],[pscustomobject]@{ Status='ACTIONABLE'; EffectiveStatus='ACTIONABLE'; BackoffUntilUtc='' }) -MinimumDelaySeconds 60 -NowUtc $now
Assert-Equal $mixedDelay 60 'an actionable row preserves the configured minimum delay'
$invalidDelay = Get-AdaptiveCycleDelaySeconds -Rows @([pscustomobject]@{ Status='SKIPPED_BY_STATUS_BACKOFF'; EffectiveStatus='SKIPPED_BY_STATUS_BACKOFF'; BackoffUntilUtc='invalid' }) -MinimumDelaySeconds 60 -NowUtc $now
Assert-Equal $invalidDelay 60 'invalid backoff evidence safely falls back to the configured delay'

$launcherText = Get-Content -LiteralPath $launcherPath -Raw
foreach ($requiredText in @('$LauncherVersion = "2.10.77"','CloudRefreshScope.txt','Merge-ScopedInventoryMap','remoteStagingScript','Move-Item -LiteralPath $remoteStagingScript')) {
    Assert-True $launcherText.Contains($requiredText) "launcher source contains $requiredText"
}

$endpointTokens = $null
$endpointErrors = $null
$endpointAst = [System.Management.Automation.Language.Parser]::ParseFile($endpointPath,[ref]$endpointTokens,[ref]$endpointErrors)
if ($endpointErrors.Count -gt 0) { throw "Endpoint parsing failed: $($endpointErrors[0].Message)" }
$dsregFunctions = @($endpointAst.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-ParsedDsregStatusSnapshot'
},$true))
Assert-Equal $dsregFunctions.Count 1 'dsreg snapshot helper exists exactly once'
$dsregText = $dsregFunctions[0].Extent.Text
Assert-True (-not $dsregText.Contains('$env:TEMP')) 'dsreg snapshot no longer uses the shared Windows temp directory'
Assert-True $dsregText.Contains('$OutputDirPath') 'dsreg snapshot writes directly into the run evidence directory'
Assert-True (Get-Content -LiteralPath $endpointPath -Raw).Contains('$ScriptVersion = "2.10.39"') 'endpoint version is incremented'

Write-Output 'PASS: Intune Hybrid Join orchestrator optimization tests completed.'
