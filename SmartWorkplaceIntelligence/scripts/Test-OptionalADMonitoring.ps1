[CmdletBinding()]
param([string]$TestRoot = (Join-Path ([IO.Path]::GetTempPath()) 'SmartM365-OfflineTests'))
# Synthetic local files only. No collectors, tenant calls, production data or uploads.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$fixture = Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
$raw = Join-Path $fixture 'raw'
$last = Join-Path $raw 'DATA-LAST'
New-Item -ItemType Directory -Path $last,(Join-Path $fixture 'config') -Force | Out-Null
$adPath = Join-Path $last 'AD_HealthCheck.csv'
$requiredPath = Join-Path $last 'Required.csv'
$contract = Join-Path $fixture 'config/prepared-source-contract.json.txt'
@{currentFiles=@('Required.csv');optionalCurrentFiles=@(@{file='AD_HealthCheck.csv';maxAgeHours=48});mappingFiles=@();dailyFiles=@();history=@()} |
    ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $contract
[pscustomobject]@{TenantKey='synthetic-test';Value=1} | Export-Csv -LiteralPath $requiredPath -NoTypeInformation
$module = Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force -PassThru
Import-Module (Join-Path $PSScriptRoot 'OptionalMonitoringEvidence.psm1') -Force
& $module {param($root) $script:ProductRoot=$root} $fixture
$script:checks=0
function Check([bool]$Pass,[string]$Message) { if (-not $Pass) { throw $Message }; $script:checks++ }
function Reject([scriptblock]$Action,[string]$Pattern) {
    $message=''; try { & $Action | Out-Null } catch { $message=$_.Exception.Message }
    Check ($message -match $Pattern) "Expected '$Pattern', got '$message'"
}
$now = [datetime]::UtcNow
function Write-ADFixture([double]$AgeHours=1,[string[]]$Statuses=@('OK','Warning','Critical','NotMeasured'),[string]$Tenant='synthetic-test') {
    @($Statuses | ForEach-Object {
        [pscustomobject]@{TenantKey=$Tenant;Forest='example.invalid';Domain='example.invalid';Category='Replication';Check=('Check-'+$_);Status=$_;RunDateUtc=$now.AddHours(-$AgeHours).ToString('O')}
    }) | Export-Csv -LiteralPath $adPath -NoTypeInformation
}
function Read-Monitoring { Get-OptionalADHealthEvidence -Path $adPath -AsOfUtc $now }
function Get-Plan { @(Get-PreparedSourcePlan -DataRoot $raw -SourceContractPath $contract -MetadataOnly) }
$runParameters = @{DataRoot=$raw;OutputRoot=(Join-Path $raw 'DATA-POWERBI');WorkRoot=(Join-Path $fixture 'work');TenantKey='synthetic-test';AccountClassificationConfigPath=(Join-Path $fixture 'unused.json.txt');ValidateOnly=$true}

# Optional source selection and capture; all required and tenant guards stay strict.
$issues=[Collections.Generic.List[object]]::new()
$plan=@(Get-PreparedSourcePlan -DataRoot $raw -SourceContractPath $contract -OptionalSourceIssues $issues)
Check ($plan.Count -eq 1 -and $issues.Count -eq 1) 'Missing monitoring blocked selection or lost its audit'
Write-ADFixture
$plan=Get-Plan
$adEntry=@($plan | Where-Object IsOptional)[0]
Check ($plan.Count -eq 2 -and $adEntry.MaxAgeHours -eq 48) 'Fresh optional monitoring was not selected'
[IO.File]::SetLastWriteTimeUtc($adPath,$now.AddHours(-49))
# An optional source past its target is kept as the latest available data and recorded as stale.
$staleIssues=[Collections.Generic.List[object]]::new()
Check (@(Get-PreparedSourcePlan -DataRoot $raw -SourceContractPath $contract -MetadataOnly -OptionalSourceIssues $staleIssues).Count -eq 2 -and @($staleIssues | Where-Object Phase -eq 'StaleUsed').Count -eq 1 -and $staleIssues[0].Reason -like '*48 h target*') 'Stale optional monitoring was not kept and recorded as stale'
$overrideIssues=[Collections.Generic.List[object]]::new()
$null=@(Get-PreparedSourcePlan -DataRoot $raw -SourceContractPath $contract -MetadataOnly -AgeOverrides @{'AD_HealthCheck.csv'=168} -OptionalSourceIssues $overrideIssues)
Check (@($overrideIssues | Where-Object { $_.Phase -eq 'StaleUsed' -and $_.Reason -like '*48 h target*' }).Count -eq 1) 'Global age override weakened optional monitoring freshness'
$staleEntry=@((Get-Plan) | Where-Object IsOptional)[0]
$staleCopy=& $module {param($entry,$root) Copy-PreparedStableSource -Entry $entry -SnapshotRoot $root -Attempts 1 -RetryDelayMs 0} $staleEntry (Join-Path $fixture 'stale-capture')
Check ($null -ne $staleCopy -and (Test-Path (Join-Path (Join-Path $fixture 'stale-capture') $staleEntry.Relative))) 'Stale optional monitoring was not captured'
[IO.File]::SetLastWriteTimeUtc($adPath,$now.AddMinutes(10))
Check ((Get-Plan).Count -eq 1) 'Future optional publication was accepted'
Write-ADFixture
Remove-Item -LiteralPath $adPath
$copyIssues=[Collections.Generic.List[object]]::new()
$snapshot=Join-Path $fixture 'capture'
$copy=& $module {param($entry,$root,$issues) Copy-PreparedStableSource -Entry $entry -SnapshotRoot $root -Attempts 1 -RetryDelayMs 0 -OptionalSourceIssues $issues} $adEntry $snapshot $copyIssues
Check ($null -eq $copy -and $copyIssues.Count -eq 1 -and -not (Test-Path (Join-Path $snapshot $adEntry.Relative))) 'Disappearing optional source blocked capture or retained partial bytes'
$result=Invoke-PreparedEvidencePipeline @runParameters
Check ($result.SourceFiles -eq 1 -and -not $result.Publication -and $result.OptionalSourceIssues.Count -eq 1) 'Missing optional source blocked ValidateOnly'
$capture=Get-Content -LiteralPath (Join-Path $result.DiagnosticPath 'capture.json.txt') -Raw | ConvertFrom-Json
Check ($capture.OptionalSourceIssues.Count -eq 1) 'Skipped monitoring was not retained in capture diagnostics'
Remove-Item -LiteralPath $requiredPath
Reject { Get-Plan } 'Required.csv'
[pscustomobject]@{TenantKey='synthetic-test';Value=1} | Export-Csv -LiteralPath $requiredPath -NoTypeInformation
[IO.File]::SetLastWriteTimeUtc($requiredPath,$now.AddHours(-169))
Reject { Get-Plan } 'publication-age'
[IO.File]::SetLastWriteTimeUtc($requiredPath,$now)
Write-ADFixture -Tenant 'other-test'
Reject { Invoke-PreparedEvidencePipeline @runParameters } 'TenantKey'

# Business acquisition freshness: modifying/copying a stale file never refreshes it.
Write-ADFixture
$monitoring=Read-Monitoring
Check ($monitoring.Available -and $monitoring.Rows.Count -eq 3 -and $monitoring.UnmeasuredChecks -eq 1) 'Measured coverage includes NotMeasured checks'
Write-ADFixture -AgeHours 48
Check ((Read-Monitoring).Available) 'Exactly 48-hour acquisition was incorrectly excluded'
Write-ADFixture -AgeHours 49
Check (-not (Read-Monitoring).Available) 'Recently copied stale acquisition was accepted'
Write-ADFixture -AgeHours -1
Check (-not (Read-Monitoring).Available) 'Future acquisition was accepted'
Write-ADFixture -Statuses @('NotMeasured')
Check (-not (Read-Monitoring).Available) 'Unmeasured checks became healthy evidence'
Write-ADFixture -Statuses @('Unexpected')
Check (-not (Read-Monitoring).Available) 'Unknown status became a healthy result'
Write-ADFixture
$rows=@(Import-Csv $adPath); $rows[0].RunDateUtc=''; $rows | Export-Csv $adPath -NoTypeInformation
Check (-not (Read-Monitoring).Available) 'Missing acquisition timestamp fell back to modification time'
Write-ADFixture
$rows=@(Import-Csv $adPath); $rows[1].RunDateUtc=$now.AddHours(-49).ToString('O'); $rows | Export-Csv $adPath -NoTypeInformation
Check (-not (Read-Monitoring).Available) 'Mixed observations used the newest timestamp'
'Forest,Domain,Category,Check,Status,RunDateUtc' | Set-Content $adPath
Check (-not (Read-Monitoring).Available) 'Header-only monitoring became observed evidence'
'Forest,Domain,Category,Check,Status,RunDateUtc','example.invalid,example.invalid,Replication,Check,OK' | Set-Content $adPath
Check (-not (Read-Monitoring).Available) 'Malformed monitoring became observed evidence'

# Execute both real readers; unchanged output schemas must pass prepared validation.
$device=Join-Path $fixture 'DeviceInventoryEvidence.csv'
$users=Join-Path $fixture 'UserInventoryEvidence.csv'
$security=Join-Path $fixture 'SecurityControlEvidence.csv'
$trust=Join-Path $fixture 'DataTrustEvidence.csv'
[pscustomobject]@{'Compliance State'='Compliant';'Secure Boot State'='Enabled'} | Export-Csv $device -NoTypeInformation
[pscustomobject]@{'MFA Registration State'='Registered'} | Export-Csv $users -NoTypeInformation
[pscustomobject]@{AzureADDeviceId='test-device';DeviceName='test-device';'BitLocker/Encryption'='Pass';CodeIntegrity='Pass'} | Export-Csv (Join-Path $last 'Intune_Devices_Compliance_Policies.csv') -NoTypeInformation
[pscustomobject]@{Status='OK'} | Export-Csv (Join-Path $last 'M365_Entra_AzureADConnect_SyncHealth.csv') -NoTypeInformation
foreach ($case in 'fresh','missing','stale','unmeasured') {
    if ($case -eq 'missing') { Remove-Item $adPath }
    elseif ($case -eq 'stale') { Write-ADFixture -AgeHours 49 }
    elseif ($case -eq 'unmeasured') { Write-ADFixture -Statuses @('NotMeasured') }
    else { Write-ADFixture }
    & (Join-Path $PSScriptRoot 'New-SecurityEvidence.ps1') -DataRoot $last -DeviceEvidencePath $device -UserEvidencePath $users -OutputPath $security -AsOfUtc $now | Out-Null
    $controls=@(Import-Csv $security)
    $control=@($controls | Where-Object 'Control Sort' -eq 5)[0]
    Check ($controls.Count -eq 11) "Control count changed: $case"
    if ($case -eq 'fresh') {
        Check ($control.'Covered Entities' -eq '3' -and $control.'Healthy Entities' -eq '1' -and $control.'Affected Entities' -eq '2') 'Measured AD counts changed'
        Check ([double]$control.'Health Rate' -lt 0.34) 'NotMeasured was counted as healthy'
    } else {
        Check ($control.'Evidence Status' -eq 'Not collected' -and $control.'Gap State' -eq 'Not assessed') "Unavailable monitoring became healthy: $case"
        foreach ($field in 'Covered Entities','Healthy Entities','Affected Entities','Health Rate','Snapshot Date') { Check ($control.$field -eq '') "Unavailable AD value fabricated: $field / $case" }
    }
    & (Join-Path $PSScriptRoot 'New-DataTrustEvidence.ps1') -DataRoot $raw -OutputPath $trust -AsOfUtc $now | Out-Null
    $adTrust=@(Import-Csv $trust | Where-Object 'File Name' -eq 'AD_HealthCheck.csv')[0]
    Check ($adTrust.'Decision Readiness' -eq $(if($case -eq 'fresh'){'Ready'}else{'Optional unavailable'})) "Data Trust still blocks optional monitoring: $case"
    foreach ($table in 'Security Control Evidence','Data Trust Evidence') {
        $validation=& (Get-Process -Id $PID).Path -NoProfile -File (Join-Path $PSScriptRoot 'Test-PreparedEvidence.ps1') -Root $fixture -Tables $table
        Check ($LASTEXITCODE -eq 0 -and ($validation | ConvertFrom-Json).Passed) "Prepared reader schema/type validation failed: $case / $table"
    }
}
Remove-Item (Join-Path $last 'M365_Entra_AzureADConnect_SyncHealth.csv')
Reject { & (Join-Path $PSScriptRoot 'New-SecurityEvidence.ps1') -DataRoot $last -DeviceEvidencePath $device -UserEvidencePath $users -OutputPath $security } 'Required private evidence'
$jobs=(Get-Content (Join-Path $PSScriptRoot '../../SmartM365/SmartInventory/Orchestrator/Orchestrator-Jobs.json.template') -Raw | ConvertFrom-Json).Jobs
Check ('AD-HealthCheck' -notin ($jobs | Where-Object Name -eq 'WorkplaceEvidence-Prepare').DependsOn) 'Workplace preparation still waits for AD health'
Check (($jobs | Where-Object Name -eq 'AD-HealthCheck').Enabled) 'AD monitoring was disabled'
Write-Host "PASS: $script:checks optional AD monitoring checks. Synthetic local evidence only; no collectors or live writes."
