<#
.VERSION
1.0.0
#>
#requires -Version 5.1
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $root 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
Import-Module (Join-Path $root 'Campaigns\HybridJoin\SmartEndpointMigration.Campaign.HybridJoin.psd1') -MinimumVersion 1.0.0 -Force
$module = Get-Module SmartEndpointMigration.Campaign.HybridJoin
& $module {
    $id = '11111111-1111-1111-1111-111111111111'
    $tenant = '22222222-2222-2222-2222-222222222222'
    $row = [pscustomobject]@{ Dsreg_DeviceId=$id; Dsreg_TenantId=$tenant; IntuneEnrolled=$true }
    $cloud = [pscustomobject]@{ AzureADDeviceId=$id; InventoryTenantId=$tenant; DeviceName='OTHER-NAME' }
    function Assert-State($actual,$expected) { if ($actual -ne $expected) { throw "Expected $expected, got $actual" }; Write-Host "PASS: $expected" }
    Assert-State (Get-HybridJoinIntuneConfirmation -Row $row -CloudRows @($cloud)).State 'Confirmed'
    $wrong = [pscustomobject]@{ AzureADDeviceId='33333333-3333-3333-3333-333333333333'; InventoryTenantId=$tenant; DeviceName='SAME-NAME' }
    Assert-State (Get-HybridJoinIntuneConfirmation -Row $row -CloudRows @($wrong)).State 'NotFound'
    Assert-State (Get-HybridJoinIntuneConfirmation -Row $row -CloudRows @($cloud,$cloud)).State 'Ambiguous'
    $row.IntuneEnrolled = $false
    Assert-State (Get-HybridJoinIntuneConfirmation -Row $row -CloudRows @($cloud)).State 'CloudOnly'
    $row.IntuneEnrolled = $true
    $cloud.InventoryTenantId = '44444444-4444-4444-4444-444444444444'
    Assert-State (Get-HybridJoinIntuneConfirmation -Row $row -CloudRows @($cloud)).State 'Unknown'
    $row.Dsreg_DeviceId = ''
    Assert-State (Get-HybridJoinIntuneConfirmation -Row $row -CloudRows @($cloud)).State 'Unknown'
    if (Test-HybridJoinAlreadyEnrolled ([pscustomobject]@{ IntuneInventoryPresent=$true; Status='SUCCESS' })) { throw 'Name-only result was archived.' }
    if (Test-HybridJoinAlreadyEnrolled ([pscustomobject]@{ IntuneConfirmationState='Confirmed'; Status='GPUPDATE_TIMEOUT' })) { throw 'A failed run was archived.' }
    if (Test-HybridJoinAlreadyEnrolled ([pscustomobject]@{ IntuneConfirmationState='Confirmed'; Status='ENROLLED_DETECTED_POST_CYCLE'; UserLicensePreflightState='NoAssignedLicense' })) { throw 'An unlicensed device was archived.' }
    Write-Host 'PASS: name-only result is not archived'
}

# Preflight never treats missing or ambiguous evidence as a license absence.
$probe=[pscustomobject]@{ Sid='S-1-5-21-42'; Tenants=@('22222222-2222-2222-2222-222222222222'); AutoEnroll=1; CredentialType=1 }
$user=[pscustomobject]@{ InventoryTenantId=$probe.Tenants[0]; InventoryScope='AllUsers'; InventoryComplete='True'; OnPremisesSecurityIdentifier=$probe.Sid; UserPrincipalName='user@example.test'; AssignedSkuIds=''; CollectedUtc=[datetime]::UtcNow.ToString('o') }
$decision=Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user)
if (-not $decision.Block -or $decision.UserPrincipalName -ne 'user@example.test' -or $decision.CollectedUtc -ne $user.CollectedUtc) { throw 'Confirmed unlicensed user was not guarded or identified.' }
$user.CollectedUtc=[datetime]::UtcNow.AddHours(-25).ToString('o')
if ((Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user)).Block) { throw 'Stale user evidence was treated as an absent license.' }
$user.CollectedUtc='2026-10-30T21:00:00'
if ((Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user)).Block) { throw 'Timezone-unspecified evidence was treated as an absent license.' }
$user.CollectedUtc=[datetime]::UtcNow.AddHours(1).ToString('o')
if ((Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user)).Block) { throw 'Future user evidence was treated as an absent license.' }
$user.CollectedUtc=[datetime]::UtcNow.ToString('o')
$missingColumn=[pscustomobject]@{ InventoryTenantId=$probe.Tenants[0]; InventoryScope='AllUsers'; InventoryComplete='True'; OnPremisesSecurityIdentifier=$probe.Sid; CollectedUtc=$user.CollectedUtc }
if ((Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($missingColumn)).Block) { throw 'Missing license field was treated as an absent license.' }
$user.AssignedSkuIds=$null
if ((Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user)).Block) { throw 'Null license field was treated as an absent license.' }
$user.AssignedSkuIds='sku'
if ((Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user)).Block) { throw 'User with an assigned SKU was incorrectly blocked.' }
$user.AssignedSkuIds=''
if ((Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user,$user)).Block) { throw 'Ambiguous SID was incorrectly blocked.' }
$probe.CredentialType=2
if (-not (Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user)).Block) { throw 'An unlicensed user was not guarded for device credential enrollment.' }
$probeWithoutPolicy=[pscustomobject]@{ Sid=$probe.Sid; Tenants=$probe.Tenants }
if (-not (Get-SemtUserLicensePreflightDecision -Evidence $probeWithoutPolicy -Users @($user)).Block) { throw 'An unlicensed user was not guarded without MDM policy data.' }
Write-Host 'PASS: preflight guards no-SKU evidence and leaves ambiguous cases unknown'

# Both audit and repair check licenses before copying or starting the endpoint.
$worker = Join-Path $root 'Scripts\SmartM365-EndpointMigration-PsExecWorker.ps1'
$tokens=$null; $errors=$null
$workerAst=[System.Management.Automation.Language.Parser]::ParseFile($worker,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw 'Worker parse failed.' }
$preflightIf=$workerAst.Find({param($n) $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Extent.Text.StartsWith("if (`$campaign -eq 'HybridJoin'") -and $n.Extent.Text.Contains('Invoke-SemtUserLicensePreflight')},$true)
if (-not $preflightIf) { throw 'Worker license preflight branch missing.' }
$runPreflight=[scriptblock]::Create($preflightIf.Extent.Text)
function Invoke-SemtUserLicensePreflight { param($ComputerName,$InventoryPath); return [pscustomobject]@{State='NoAssignedLicense';Detail='No user SKU';Block=$true} }
function Write-SemtLogLine { param($Path,$Message) }
$campaign='HybridJoin'; $target='fixture.invalid'; $packageRoot=$root; $logPath='unused'; $access=[pscustomobject]@{AdminShareReachable=$true}
foreach ($arguments in @(@('-AuditOnly'),@('-RepairFixture'))) {
    $contract=[pscustomobject]@{Flags=[pscustomobject]@{DryRun=$false}; Endpoint=[pscustomobject]@{Arguments=@($arguments)}}
    $result=@{}
    & $runPreflight
    if ($result.UserLicensePreflightState -ne 'NoAssignedLicense') { throw "Preflight omitted for $arguments" }
}
Write-Host 'PASS: audit and repair both run the license preflight'
$argumentFunction=$workerAst.Find({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-WorkerEndpointArguments'},$true)
if (-not $argumentFunction) { throw 'Worker endpoint argument guard missing.' }
. ([scriptblock]::Create($argumentFunction.Extent.Text))
$blocked=[pscustomobject]@{Block=$true;Detail='No user license'}
$allowed=[pscustomobject]@{Block=$false}
$forced=@(Get-WorkerEndpointArguments -Arguments @('-AllowDsregLeave') -LicensePreflight $blocked)
if ($forced -notcontains '-AuditOnly') { throw 'Unlicensed repair was not forced into audit mode.' }
$audit=@(Get-WorkerEndpointArguments -Arguments @('-AuditOnly') -LicensePreflight $blocked)
if (@($audit | Where-Object { $_ -eq '-AuditOnly' }).Count -ne 1) { throw 'Audit mode was not preserved exactly once.' }
$repair=@(Get-WorkerEndpointArguments -Arguments @('-AllowDsregLeave') -LicensePreflight $allowed)
if ($repair -contains '-AuditOnly') { throw 'Licensed repair was unexpectedly changed into audit.' }
$rejected=$false
try { [void](Get-WorkerEndpointArguments -Arguments @('-AuditOnly:$false') -LicensePreflight $blocked) } catch { $rejected=$true }
if (-not $rejected) { throw 'Conflicting audit argument was not rejected.' }
Write-Host 'PASS: unlicensed audit and repair both run the endpoint in audit-only mode'
$actionIf=$workerAst.Find({param($n) $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Extent.Text.StartsWith('if ($null -ne $licensePreflight -and $licensePreflight.Block -and [string]$result.RemoteStatus')},$true)
if (-not $actionIf) { throw 'License action after audit is missing.' }
$applyAction=[scriptblock]::Create($actionIf.Extent.Text)
$licensePreflight=$blocked
$result=@{RemoteStatus='AUDIT_HYBRID_JOIN_UNHEALTHY';RemoteNextAction='FIX_HYBRID_JOIN';NextAction=''}
& $applyAction
if ($result.NextAction -ne 'ASSIGN_INTUNE_LICENSE_THEN_RETRY' -or $result.RemoteNextAction -ne 'FIX_HYBRID_JOIN') { throw 'License action lost the device diagnostic.' }
$result=@{RemoteStatus='';NextAction='CHECK_TARGET_EXECUTION_ACCESS'}
& $applyAction
if ($result.NextAction -ne 'CHECK_TARGET_EXECUTION_ACCESS') { throw 'Audit execution failure was hidden by the license result.' }
Write-Host 'PASS: license action preserves audit diagnostics and execution failures'

# A blocked computer is not retried during a Loop Run or removed from its Wave.
$enginePath=Join-Path $root 'Scripts\SmartM365-Invoke-EndpointMigrationWithPsExec.ps1'
$tokens=$null; $errors=$null
$engineAst=[System.Management.Automation.Language.Parser]::ParseFile($enginePath,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw 'Engine parse failed.' }
$filterIf=$engineAst.Find({param($n) $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Extent.Text.StartsWith("if (`$Campaign -eq 'HybridJoin' -and `$licenseBlockedThisRun.Count")},$true)
if (-not $filterIf) { throw 'License-blocked Loop Run filter missing.' }
$runFilter=[scriptblock]::Create('while ($true) { ' + $filterIf.Extent.Text + '; break }; [pscustomobject]@{Computers=@($computers); Reason=$stopReason}')
$Campaign='HybridJoin'; $licenseBlockedThisRun=@{ 'PC-UNLICENSED' = $true }
$computers=@('PC-UNLICENSED','PC-LICENSED'); $stopReason='Finished'; $filtered=& $runFilter
if (@($filtered.Computers).Count -ne 1 -or $filtered.Computers[0] -ne 'PC-LICENSED') { throw 'Loop Run retried a license-blocked computer.' }
$computers=@('PC-UNLICENSED'); $stopReason='Finished'; $filtered=& $runFilter
if ($filtered.Reason -ne 'BlockedUserLicense' -or @($filtered.Computers).Count -ne 0) { throw 'Loop Run did not stop when all remaining computers are license-blocked.' }
$licenseBlockedThisRun=@{ 'PC1.FIRST.EXAMPLE' = $true }
$computers=@('pc1.first.example','pc1.second.example'); $stopReason='Finished'; $filtered=& $runFilter
if (@($filtered.Computers).Count -ne 1 -or $filtered.Computers[0] -ne 'pc1.second.example') { throw 'Distinct domain names were incorrectly treated as the same blocked computer.' }
Write-Host 'PASS: Loop Run leaves unlicensed computers pending without retrying them'
$markIf=$engineAst.Find({param($n) $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Extent.Text.StartsWith("if ([string](Get-SemtPropertyValue -InputObject `$row -Name 'UserLicensePreflightState'")},$true)
if (-not $markIf) { throw 'Engine license-preflight retry guard missing.' }
$markLicenseBlocked=[scriptblock]::Create($markIf.Extent.Text)
$licenseBlockedThisRun=@{}
$row=[pscustomobject]@{ComputerName='PC-NOLICENSE';Status='AUDIT_HYBRID_JOIN_UNHEALTHY';UserLicensePreflightState='NoAssignedLicense'}
& $markLicenseBlocked
if (-not $licenseBlockedThisRun.ContainsKey('PC-NOLICENSE')) { throw 'Audited unlicensed device was not marked to stop retries.' }
Write-Host 'PASS: audit result still stops Loop retries for an unlicensed device'

# An audited unlicensed device retains the diagnostic and does not trigger post-cycle Graph calls.
& $module {
    $row=[pscustomobject]@{Status='AUDIT_HYBRID_JOIN_UNHEALTHY'; NextAction='ASSIGN_INTUNE_LICENSE_THEN_RETRY'; UserLicensePreflightState='NoAssignedLicense'; UserLicensePreflightDetail='No user SKU'; InteractiveUserSid=''; Dsreg_TenantId=''}
    $context=[pscustomobject]@{PackageRoot=(Join-Path ([IO.Path]::GetTempPath()) ('SEMT-absent-'+[guid]::NewGuid().ToString('N'))); DryRun=$false; Log={param($text,$color)}; CycleNumber=1}
    Invoke-HybridJoinPostCycle -Context $context -Rows @($row)
    if ($row.Status -ne 'AUDIT_HYBRID_JOIN_UNHEALTHY' -or $row.NextAction -ne 'ASSIGN_INTUNE_LICENSE_THEN_RETRY' -or $row.UserLicenseState -ne 'NoAssignedLicense' -or $row.UserLicenseDetail -ne 'No user SKU') { throw 'Audit diagnostic or license result was not preserved.' }
    if ($row.PSObject.Properties['PostCycleIntuneInventoryChecked']) { throw 'License-blocked device triggered cloud confirmation.' }
    $already=[pscustomobject]@{Status='AUDIT_SUCCESS_ALREADY_INTUNE'; NextAction='ASSIGN_INTUNE_LICENSE_THEN_RETRY'; UserLicensePreflightState='NoAssignedLicense'; UserLicensePreflightDetail='No user SKU'; InteractiveUserSid=''; Dsreg_TenantId=''}
    Invoke-HybridJoinPostCycle -Context $context -Rows @($already)
    if ($already.Status -ne 'AUDIT_SUCCESS_ALREADY_INTUNE' -or $already.NextAction -ne 'ASSIGN_INTUNE_LICENSE_THEN_RETRY') { throw 'Unlicensed already-enrolled audit was incorrectly reclassified.' }
    # A local Intune trace with an unhealthy join stays a review item; it is never reclassified as enrolled.
    $trace=[pscustomobject]@{Status='INTUNE_TRACE_DEVICE_JOIN_UNHEALTHY'; NextAction='REVIEW_INTUNE_TRACE_AND_DEVICE_JOIN'; UserLicensePreflightState='Unknown'; InteractiveUserSid=''; Dsreg_TenantId=''}
    Invoke-HybridJoinPostCycle -Context ([pscustomobject]@{PackageRoot=$context.PackageRoot; DryRun=$true; Log={param($text,$color)}; CycleNumber=1}) -Rows @($trace)
    if ($trace.Status -ne 'INTUNE_TRACE_DEVICE_JOIN_UNHEALTHY' -or $trace.NextAction -ne 'REVIEW_INTUNE_TRACE_AND_DEVICE_JOIN' -or (Get-HybridJoinNextAction -Status 'INTUNE_TRACE_DEVICE_JOIN_UNHEALTHY') -ne 'REVIEW_INTUNE_TRACE_AND_DEVICE_JOIN') { throw 'Local Intune trace with unhealthy join was not kept for review.' }
}
Write-Host 'PASS: unlicensed audit preserves the device diagnostic and skips post-cycle Graph confirmation'
Write-Host 'PASS: local Intune trace with unhealthy join stays a review item'
$hjEndpointText = [IO.File]::ReadAllText((Join-Path $root 'Campaigns\HybridJoin\Scripts\SmartM365-Invoke-EndpointMigration-HybridJoin.ps1'))
if ($hjEndpointText -match '\$status = "INTUNE_ENROLLED"' -or $hjEndpointText -notmatch '"INTUNE_TRACE_DEVICE_JOIN_UNHEALTHY" \{ return "REVIEW_INTUNE_TRACE_AND_DEVICE_JOIN" \}' -or $hjEndpointText -notmatch '\$status = "INTUNE_TRACE_DEVICE_JOIN_UNHEALTHY"') { throw 'Endpoint still reports a local Intune trace with unhealthy join as INTUNE_ENROLLED.' }
Write-Host 'PASS: endpoint reports a local Intune trace with unhealthy join for review, not as enrolled'

# Exercise the target attempt contract, not just the technician-side classifier.
$endpoint = Join-Path $root 'Campaigns\HybridJoin\Scripts\SmartM365-Invoke-EndpointMigration-HybridJoin.ps1'
$tokens=$null; $errors=$null
$ast=[System.Management.Automation.Language.Parser]::ParseFile($endpoint,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw 'Endpoint parse failed.' }
$definition=$ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Write-SemtHybridJoinAttemptResult' },$true)
. ([scriptblock]::Create($definition.Extent.Text))
function Write-SemtAttemptResult { param($CampaignDataRoot,$Data) return $Data }
$script:SemtCampaign='HybridJoin'; $script:SemtAttemptStartUtc=[datetime]::UtcNow
$script:SemtCsvPath=''; $script:SemtTranscriptPath=''; $script:SemtTargetGuardStatus=''
$script:RetryAfterRebootAction=''; $script:RetryAfterRebootDetail=''; $script:RetryAfterRebootAttempt=''; $script:RetryAfterRebootMaxAttemptsResult=''; $script:RetryAfterRebootTaskNameResult=''
$ScriptVersion='1.0.0'; $RunId='test'; $CorrelationId='test'; $ComputerName='PC'; $RunLogPath=''; $DataRoot=''; $RetryAfterRebootTaskRun=$false
$script:dsregDeviceId='11111111-1111-1111-1111-111111111111'; $script:dsregTenantId='22222222-2222-2222-2222-222222222222'; $script:interactiveUserSid='S-1-5-21-1'; $script:mdmEnrollmentRecentEvidence='event'
$attempt=Write-SemtHybridJoinAttemptResult -State Completed -Status SUCCESS
if ($attempt.Dsreg_DeviceId -ne $script:dsregDeviceId -or $attempt.Dsreg_TenantId -ne $script:dsregTenantId -or $attempt.InteractiveUserSid -ne $script:interactiveUserSid -or $attempt.MdmEnrollmentRecentEvidence -ne 'event') { throw 'Attempt identity evidence was not projected.' }
Write-Host 'PASS: endpoint attempt carries identity and MDM evidence'
# Unhealthy join with a local Intune trace: the launcher proves absence from Intune by device ID.
& $module {
    $id = '55555555-5555-5555-5555-555555555555'
    $options = [ordered]@{ 'HybridJoin.RepairUnhealthyJoinWithIntuneTrace' = [pscustomobject]@{ Value = $true } }
    $newContext = {
        param($intuneRows,$intuneSource)
        $intuneMap = @{}
        foreach ($r in @($intuneRows)) { $intuneMap[(Get-SemtComputerKey -ComputerName $r.DeviceName)] = $r }
        [pscustomobject]@{ Options = $options; Inventory = @{
            Entra = [pscustomobject]@{ Source = 'Shared'; Map = @{ (Get-SemtComputerKey -ComputerName 'PC-TRACE') = [pscustomobject]@{ DeviceId = $id; EntraRegisteredState = 'Registered' } } }
            Intune = [pscustomobject]@{ Source = $intuneSource; Map = $intuneMap } } }
    }
    $absent = & $newContext @([pscustomobject]@{ DeviceName = 'OTHER'; AzureADDeviceId = '66666666-6666-6666-6666-666666666666' }) 'RunScopedRefresh'
    $arguments = @(Get-HybridJoinPerComputerArguments -Context $absent -Computer 'PC-TRACE')
    if (($arguments -join ' ') -ne "-IntuneAbsentDeviceId $id") { throw "Absence proof not passed: $($arguments -join ' ')" }
    if (Get-HybridJoinIntuneAbsenceProof -Context (& $newContext @([pscustomobject]@{ DeviceName = 'RENAMED'; AzureADDeviceId = $id }) 'Shared') -Computer 'PC-TRACE') { throw 'An Intune record with the device ID was ignored.' }
    if (Get-HybridJoinIntuneAbsenceProof -Context (& $newContext @([pscustomobject]@{ DeviceName = 'PC-TRACE'; AzureADDeviceId = '' }) 'Shared') -Computer 'PC-TRACE') { throw 'An Intune record with the same name was ignored.' }
    if (Get-HybridJoinIntuneAbsenceProof -Context (& $newContext @() 'SharedNotFresh') -Computer 'PC-TRACE') { throw 'A stale Intune inventory was accepted as proof.' }
    $options['HybridJoin.RepairUnhealthyJoinWithIntuneTrace'].Value = $false
    if (@(Get-HybridJoinPerComputerArguments -Context $absent -Computer 'PC-TRACE').Count -ne 0) { throw 'Absence proof passed while the option is off.' }
    if ((Get-HybridJoinNextAction -Status 'INTUNE_TRACE_REMOVE_FAILED') -ne 'REVIEW_INTUNE_TRACE_AND_DEVICE_JOIN') { throw 'INTUNE_TRACE_REMOVE_FAILED is not a review item.' }
    Write-Host 'PASS: Intune absence proof requires fresh inventories and no record by name or device ID'
}

# Endpoint trace: kept between runs, repair only when every condition holds.
foreach ($name in @('Write-AtomicJsonFile','ConvertTo-SemtTargetJsonDateText','Read-SemtTargetJson','Read-IntuneTraceRepairState','Update-IntuneTraceRepairState','Clear-IntuneTraceRepairState','Get-IntuneTraceRepairDecision')) {
    $definition=$ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name },$true)
    if (-not $definition) { throw "Endpoint function missing: $name" }
    . ([scriptblock]::Create($definition.Extent.Text))
}
function Write-RunLog { param($Message) }
$traceRoot = Join-Path ([IO.Path]::GetTempPath()) ('semt-trace-' + [guid]::NewGuid().ToString('N'))
try {
    $script:IntuneTraceStatePath = Join-Path $traceRoot 'IntuneTraceRepair.json.txt'
    $deviceId = '77777777-7777-7777-7777-777777777777'
    $first = Update-IntuneTraceRepairState -DeviceId $deviceId -IntuneEnrollmentIds 'ID-A' -AzureAdJoined 'YES' -DeviceAuthStatus '' -KeySignTest 'FAILED'
    $decision = Get-IntuneTraceRepairDecision -Trace $first -LocalDeviceId $deviceId -CloudAbsentDeviceId $deviceId -Enabled $true -DelayHours 24
    if (-not $first.IsNew -or $decision.Proceed -or $decision.Reason -notmatch 'recorded on this run') { throw 'A new trace allowed an immediate repair.' }
    $saved = Get-Content -LiteralPath $script:IntuneTraceStatePath -Raw | ConvertFrom-Json
    $saved.FirstSeenUtc = (Get-Date).ToUniversalTime().AddHours(-25).ToString('o')
    $saved | ConvertTo-Json | Set-Content -LiteralPath $script:IntuneTraceStatePath -Encoding UTF8
    $old = Update-IntuneTraceRepairState -DeviceId $deviceId -IntuneEnrollmentIds 'ID-A' -AzureAdJoined 'YES' -DeviceAuthStatus '' -KeySignTest 'FAILED'
    if ($old.IsNew -or $old.AgeHours -lt 24.9 -or $old.SeenCount -ne 2) { throw "Trace age or count not kept: $($old.AgeHours) h, $($old.SeenCount)" }
    if (-not (Get-IntuneTraceRepairDecision -Trace $old -LocalDeviceId $deviceId -CloudAbsentDeviceId $deviceId -Enabled $true -DelayHours 24).Proceed) { throw 'An old trace with proof was not repaired.' }
    if ((Get-IntuneTraceRepairDecision -Trace $old -LocalDeviceId $deviceId -CloudAbsentDeviceId $deviceId -Enabled $false -DelayHours 24).Proceed) { throw 'Repair ran with the option off.' }
    if ((Get-IntuneTraceRepairDecision -Trace $old -LocalDeviceId $deviceId -CloudAbsentDeviceId '' -Enabled $true -DelayHours 24).Proceed) { throw 'Repair ran without the cloud proof.' }
    if ((Get-IntuneTraceRepairDecision -Trace $old -LocalDeviceId $deviceId -CloudAbsentDeviceId '88888888-8888-8888-8888-888888888888' -Enabled $true -DelayHours 24).Proceed) { throw 'Repair ran with a proof for another device ID.' }
    if ((Get-IntuneTraceRepairDecision -Trace $old -LocalDeviceId $deviceId -CloudAbsentDeviceId $deviceId -Enabled $true -DelayHours 48).Proceed) { throw 'Repair ran before the delay.' }
    if (-not (Update-IntuneTraceRepairState -DeviceId $deviceId -IntuneEnrollmentIds 'ID-B' -AzureAdJoined 'YES' -DeviceAuthStatus '' -KeySignTest 'FAILED').IsNew) { throw 'Changed enrollment IDs kept the old delay.' }
    Clear-IntuneTraceRepairState -Reason 'test'
    if (Test-Path -LiteralPath $script:IntuneTraceStatePath) { throw 'Trace was not removed.' }
}
finally { Remove-Item -LiteralPath $traceRoot -Recurse -Force -ErrorAction SilentlyContinue }
$hjEndpointText = [IO.File]::ReadAllText($endpoint)
if ($hjEndpointText -notmatch 'if \(\$intuneTraceRemoveFailed\) \{' -or $hjEndpointText -notmatch 'elseif \(\$intuneEnrolled\) \{' -or $hjEndpointText -notmatch 'if \(\$traceDecision\.Proceed\) \{' -or $hjEndpointText -notmatch "Clear-IntuneTraceRepairState -Reason 'device join healthy'") { throw 'Endpoint trace and repair gate are not wired.' }
Write-Host 'PASS: endpoint keeps the Intune trace and repairs only after the delay with a matching cloud proof'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB495+1rANDk/RC
# OahNPAWGcm3DOYhYOS6aABm22+8hi6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICXj9B4cvzN1K0gE3Zfnm+NYqT0JGxNZ/ALfjmrscjwTMA0GCSqG
# SIb3DQEBAQUABIIBgDGHg/MuizIx7ETCMfHUk7JkSwoC/NDZhOy/WvfP01B0Dl6k
# vn2MYKQm6HBazdgF8saInZG2XLjeyiMc1YUXr9mtUvMYFAQLUdal25h35q7ppjaG
# Dol3FvsfN1qXywfR8J/HowLB8cTwZAZ08RX2MZpUxGVAlXKUc2DH0+vlDXFcBwds
# ITqVdFQZ3DsxtZffkpWhcrtRhJ73oBNSOxvnEQRfmV/EaOteSFHQkf2Glf8qnqpo
# DC7fgU2KIfmFZyAilxovEwW/jl9sAoajPoD7Xx1ZjKqayVE6OM/L/T/WHCiwODvL
# GzOhadE3mqkJA4TJ8Jp3h6mMWPs2VWSGhAWIZIH8YOK7wy2Dja6fqQ6CxEbhe+CH
# mmNpbVOft31C+FuDexBIKoiwQMShuJBmq+Z8xvJXmMhsP9u95J5QPMT21G/UDGTQ
# 3vlTKpXPZro4GKCCwI7Bw2FI+HbN95wdS2vQ1mFFOwBnTI0TTAuDUMFF75eRaD3h
# 1c/b+qXukQkMU9+v2qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDUxOTU2
# MzNaMC8GCSqGSIb3DQEJBDEiBCAURlKPPgf3rtA615giyE+AQnJfu05mPF3rFjLh
# B6r86TANBgkqhkiG9w0BAQEFAASCAgA+NaQLmlzmpOe6PIYtgs1Q28rQA4SKdzPZ
# KOxFLHGka8Le0kCzWi2BaZ00r3MzV9x42VLfC+/2+ffCOsW3T88Q+Q7gzgceaKIg
# 9Ongb7IMPyAstn5yCvJSVtamq13XDWeoKbxWpVQdGpuSjtQ8Yfxn2BFXLU9jQFJc
# Rhg+ynmEs5TdkRlB/O9Ado4Fc2a2c1P5CCM6E1N503FYqEuvGBsjb4fHjVAAoMw8
# OzHgQkvCMKP0lGj+drXp8kkg/r//fTU79yo1jDVK+umMWsCYx6CqVlPaXazgL3Qv
# JqtfNSf52ubBJmL7mjuy5/fcVom1IAA6aUIWOP93G8mLbv/dgCCR6shqwmn8jdj8
# G/6f6Hg0GRuFx0kSIBdONnORNl2PA/2rM+y/EmEgLKTVQ2R+82MCZrGiEC4T1qcs
# CvsP+3WZFyYB7oUzTgDTTIBFcsxuy6u675xE7pUTZx3C5N+oLN6CKLHnsAulhq06
# dKblounpRmdes/YeoU9F6WBn1ixjuYEgKu4Jj/UuhrmnmnhRh39ydJJkVENI74rT
# t2ERxwW8h853Cpp44GABVjCqLCedOW6oUEPI4WGNkcK3OlWZNOg3ZFv9ulsgDuE5
# fiS3a+KkWhH/SGKlkw5x46L16HzEjzx47hM6sx2UFCQnBWrRfFfAIHWJmJ92QLB4
# 4aTL2rg23w==
# SIG # End signature block
