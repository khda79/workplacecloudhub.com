<#
.SYNOPSIS
Offline tests for the orchestrator launch prerequisites: application certificate (absent, expired,
without private key, expiring soon), runtime and tenant data shares, cache, capability withdrawal,
election exclusion, alerts, recovery and no impact when everything is ready.
.VERSION
1.3
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Distributed.psm1') -Force
$script:passed = 0
function Check($Value, $Message) { if (-not $Value) { throw $Message }; $script:passed++ }

function New-TestCertificate {
    param([int]$NotAfterDays, [int]$NotBeforeDays = -1, [switch]$NoPrivateKey)
    $rsa = [Security.Cryptography.RSA]::Create(2048)
    $request = [Security.Cryptography.X509Certificates.CertificateRequest]::new('CN=SmartM365 prerequisite test', $rsa, [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $certificate = $request.CreateSelfSigned([datetimeoffset]::UtcNow.AddDays($NotBeforeDays), [datetimeoffset]::UtcNow.AddDays($NotAfterDays))
    if ($NoPrivateKey) { return [Security.Cryptography.X509Certificates.X509Certificate2]::new($certificate.RawData) }
    return $certificate
}
function New-Candidate($Certificate) { [pscustomobject]@{ StorePath = 'Cert:\CurrentUser\My'; Certificate = $Certificate } }

$thumbprint = ('AB' * 20)
$global:PrereqTestCandidates = @()
$global:PrereqTestProviderCalls = 0
$global:PrereqTestProvider = { param($requested) $global:PrereqTestProviderCalls++; @($global:PrereqTestCandidates) }
$root = Join-Path ([IO.Path]::GetTempPath()) ('launch-prereq-' + [guid]::NewGuid().ToString('N'))
$runtime = Join-Path $root 'Runtime'
$dataAll = Join-Path $root 'DATA-ALL'
$dataLast = Join-Path $root 'DATA-LAST'
$logAll = Join-Path $root 'LOG-ALL'
$readOnly = Join-Path $root 'ReadOnlyShare'
foreach ($folder in @($runtime, $dataAll, $dataLast, $logAll, $readOnly)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
$scriptPath = Join-Path $runtime 'Collector.ps1'
[IO.File]::WriteAllText($scriptPath, '# synthetic collector')
$identity = [Security.Principal.WindowsIdentity]::GetCurrent().User
$denyWrite = [Security.AccessControl.FileSystemAccessRule]::new($identity, [Security.AccessControl.FileSystemRights]'CreateFiles,AppendData', [Security.AccessControl.InheritanceFlags]'ContainerInherit,ObjectInherit', [Security.AccessControl.PropagationFlags]::None, [Security.AccessControl.AccessControlType]::Deny)
$acl = Get-Acl -LiteralPath $readOnly
$acl.AddAccessRule($denyWrite)
Set-Acl -LiteralPath $readOnly -AclObject $acl

try {
    # ---------- Certificate readiness ----------
    $now = [datetime]::UtcNow
    $absent = Get-SmartM365OrchestratorCachedCertificateReadiness -Thumbprint $thumbprint -Cache @{} -CertificateProvider $global:PrereqTestProvider -NowUtc $now
    Check (-not $absent.Ready -and $absent.Reason -eq 'NotFound') 'An absent certificate was not reported NotFound.'
    Check ($absent.Detail -notmatch $thumbprint) 'The full thumbprint leaked into the readiness detail.'

    $global:PrereqTestCandidates = @(New-Candidate (New-TestCertificate -NotBeforeDays -30 -NotAfterDays -1))
    $expired = Get-SmartM365OrchestratorCachedCertificateReadiness -Thumbprint $thumbprint -Cache @{} -CertificateProvider $global:PrereqTestProvider -NowUtc $now
    Check (-not $expired.Ready -and $expired.Reason -eq 'Expired' -and $expired.Detail -match 'expired on') 'An expired certificate was not reported Expired.'

    $global:PrereqTestCandidates = @(New-Candidate (New-TestCertificate -NotAfterDays 200 -NoPrivateKey))
    $noKey = Get-SmartM365OrchestratorCachedCertificateReadiness -Thumbprint $thumbprint -Cache @{} -CertificateProvider $global:PrereqTestProvider -NowUtc $now
    Check (-not $noKey.Ready -and $noKey.Reason -eq 'NoPrivateKey') 'A certificate without private key was reported usable.'

    $global:PrereqTestCandidates = @(New-Candidate (New-TestCertificate -NotAfterDays 10))
    $soon = Get-SmartM365OrchestratorCachedCertificateReadiness -Thumbprint $thumbprint -Cache @{} -ExpiryWarningDays 30 -CertificateProvider $global:PrereqTestProvider -NowUtc $now
    Check ($soon.Ready -and $soon.ExpiresSoon -and $soon.DaysRemaining -in 9, 10) 'A certificate expiring in 10 days was not flagged with a 30-day threshold.'
    $notSoon = Get-SmartM365OrchestratorCachedCertificateReadiness -Thumbprint $thumbprint -Cache @{} -ExpiryWarningDays 5 -CertificateProvider $global:PrereqTestProvider -NowUtc $now
    Check ($notSoon.Ready -and -not $notSoon.ExpiresSoon) 'A certificate outside the warning threshold was flagged.'

    # An expired copy next to a valid one does not block; the valid one is used.
    $global:PrereqTestCandidates = @((New-Candidate (New-TestCertificate -NotBeforeDays -30 -NotAfterDays -1)), (New-Candidate (New-TestCertificate -NotAfterDays 200)))
    $mixed = Get-SmartM365OrchestratorCachedCertificateReadiness -Thumbprint $thumbprint -Cache @{} -CertificateProvider $global:PrereqTestProvider -NowUtc $now
    Check ($mixed.Ready -and -not $mixed.ExpiresSoon) 'A valid certificate next to an expired copy was not used.'
    Check (-not (Get-SmartM365OrchestratorCachedCertificateReadiness -Thumbprint '' -Cache @{} -NowUtc $now).Ready) 'An empty thumbprint was reported ready.'

    # ---------- Cache ----------
    $cache = @{}
    $global:PrereqTestProviderCalls = 0
    $null = Get-SmartM365OrchestratorCachedCertificateReadiness -Thumbprint $thumbprint -Cache $cache -CacheMinutes 5 -CertificateProvider $global:PrereqTestProvider -NowUtc $now
    $null = Get-SmartM365OrchestratorCachedCertificateReadiness -Thumbprint $thumbprint -Cache $cache -CacheMinutes 5 -CertificateProvider $global:PrereqTestProvider -NowUtc $now.AddMinutes(4)
    Check ($global:PrereqTestProviderCalls -eq 1) 'The certificate check was repeated inside the cache period.'
    $null = Get-SmartM365OrchestratorCachedCertificateReadiness -Thumbprint $thumbprint -Cache $cache -CacheMinutes 5 -CertificateProvider $global:PrereqTestProvider -NowUtc $now.AddMinutes(6)
    Check ($global:PrereqTestProviderCalls -eq 2) 'The certificate check was not repeated after the cache period.'

    # ---------- Paths ----------
    $dataChecks = @(
        [pscustomobject]@{ Name = 'Tenant DATA-ALL root'; Path = $dataAll; Access = 'ReadWrite'; PathType = 'Container'; Capability = 'SharedRuntime' }
        [pscustomobject]@{ Name = 'Tenant DATA-LAST root'; Path = $dataLast; Access = 'ReadWrite'; PathType = 'Container'; Capability = 'SharedRuntime' }
        [pscustomobject]@{ Name = 'Tenant LOG-ALL root'; Path = $logAll; Access = 'ReadWrite'; PathType = 'Container'; Capability = 'SharedRuntime' }
    )
    $scriptCheck = [pscustomobject]@{ Name = 'Job script'; Path = $scriptPath; Access = 'Read'; PathType = 'Leaf'; Capability = '' }
    $ready = Test-SmartM365OrchestratorLaunchPrerequisite -RequiredCapabilities @('SharedRuntime', 'Graph') -CertificateThumbprint $thumbprint -PathChecks (@($scriptCheck) + $dataChecks) -Cache @{} -CertificateProvider $global:PrereqTestProvider -NowUtc $now
    Check ($ready.Ready -and @($ready.Failures).Count -eq 0) 'A fully ready job was blocked.'
    Check (@(Get-ChildItem -LiteralPath $dataAll, $dataLast, $logAll -Force).Count -eq 0) 'Write probes left files behind.'

    $global:PrereqTestProviderCalls = 0
    $adOnly = Test-SmartM365OrchestratorLaunchPrerequisite -RequiredCapabilities @('SharedRuntime', 'AD') -CertificateThumbprint $thumbprint -PathChecks @($scriptCheck) -Cache @{} -CertificateProvider $global:PrereqTestProvider -NowUtc $now
    Check ($adOnly.Ready -and $null -eq $adOnly.Certificate -and $global:PrereqTestProviderCalls -eq 0) 'The certificate was checked for a job without app-only cloud capability.'

    $missingScript = [pscustomobject]@{ Name = 'Job script'; Path = (Join-Path $root 'MissingRuntime\Collector.ps1'); Access = 'Read'; PathType = 'Leaf'; Capability = '' }
    $runtimeMissing = Test-SmartM365OrchestratorLaunchPrerequisite -RequiredCapabilities @('SharedRuntime') -PathChecks @($missingScript) -Cache @{} -NowUtc $now
    Check (-not $runtimeMissing.Ready -and $runtimeMissing.Failures[0].JobSpecific -and @($runtimeMissing.Failures[0].Capabilities).Count -eq 0 -and $runtimeMissing.Failures[0].Detail -match 'not found') 'A missing runtime script was not reported as a job-specific failure.'

    $readOnlyCheck = [pscustomobject]@{ Name = 'Tenant DATA-LAST root'; Path = $readOnly; Access = 'ReadWrite'; PathType = 'Container'; Capability = 'SharedRuntime' }
    $readOnlyResult = Test-SmartM365OrchestratorLaunchPrerequisite -PathChecks @($readOnlyCheck) -Cache @{} -NowUtc $now
    Check (-not $readOnlyResult.Ready -and -not $readOnlyResult.Failures[0].JobSpecific -and @($readOnlyResult.Failures[0].Capabilities) -contains 'SharedRuntime') 'A read-only data share was not reported as a SharedRuntime failure.'
    $readOnlyRead = Test-SmartM365OrchestratorLaunchPrerequisite -PathChecks @([pscustomobject]@{ Name = 'Read share'; Path = $readOnly; Access = 'Read'; PathType = 'Container' }) -Cache @{} -NowUtc $now
    Check $readOnlyRead.Ready 'A readable share failed a Read check.'

    # Path cache: a share that disappears is seen only after the cache period.
    $volatile = Join-Path $root 'Volatile'
    New-Item -ItemType Directory -Path $volatile | Out-Null
    $volatileCheck = [pscustomobject]@{ Name = 'Volatile'; Path = $volatile; Access = 'Read'; PathType = 'Container' }
    $pathCache = @{}
    Check (Test-SmartM365OrchestratorLaunchPrerequisite -PathChecks @($volatileCheck) -Cache $pathCache -CacheMinutes 5 -NowUtc $now).Ready 'The volatile folder was not ready.'
    Remove-Item -LiteralPath $volatile -Force
    Check (Test-SmartM365OrchestratorLaunchPrerequisite -PathChecks @($volatileCheck) -Cache $pathCache -CacheMinutes 5 -NowUtc $now.AddMinutes(2)).Ready 'The path cache was not used inside its period.'
    Check (-not (Test-SmartM365OrchestratorLaunchPrerequisite -PathChecks @($volatileCheck) -Cache $pathCache -CacheMinutes 5 -NowUtc $now.AddMinutes(6)).Ready) 'A removed folder stayed ready after the cache period.'

    # Server-level capability probe includes the tenant data roots.
    $capabilityDocument = Get-SmartM365OrchestratorServerCapability -ServerName 'SERVER-A' -SharedDataFolderPath $dataAll -ProbeMode Static -DataPathChecks @($readOnlyCheck)
    $sharedRuntime = @($capabilityDocument.Results | Where-Object Name -eq 'SharedRuntime')[0]
    Check (-not $sharedRuntime.Ready -and 'SharedRuntime' -notin @($capabilityDocument.ReadyCapabilities)) 'The periodic probe ignored a read-only data root.'
    Check ($null -ne $capabilityDocument.CertificateStatus -and $capabilityDocument.CertificateStatus.Reason -eq 'NotConfigured') 'The capability document does not publish the certificate status.'

    # ---------- Election: BlockedJobs ----------
    $electionJobs = @(
        [pscustomobject]@{ Name = 'BlockedHere'; Enabled = $true; AssignmentMode = 'Elected'; DependsOn = @(); RequiredCapabilities = @('SharedRuntime'); RequiredGraphAppRoles = @(); EstimatedDurationMinutes = 600; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('01:00'); DaysOfWeek = @() } }
        [pscustomobject]@{ Name = 'Other'; Enabled = $true; AssignmentMode = 'Elected'; DependsOn = @(); RequiredCapabilities = @('SharedRuntime'); RequiredGraphAppRoles = @(); EstimatedDurationMinutes = 1; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('02:00'); DaysOfWeek = @() } }
    )
    $electionCapabilities = @(
        [pscustomobject]@{ ServerName = 'SERVER-A'; ReadyCapabilities = @('SharedRuntime'); GraphAppRoles = @(); BlockedJobs = @('BlockedHere') }
        [pscustomobject]@{ ServerName = 'SERVER-B'; ReadyCapabilities = @('SharedRuntime'); GraphAppRoles = @() }
    )
    foreach ($weight in @(1.0, 100.0)) {
        $plan = Get-SmartM365OrchestratorElectionPlan -Jobs $electionJobs -ServerCapabilities $electionCapabilities -ServerWeights @{ 'SERVER-A' = $weight; 'SERVER-B' = 1.0 }
        Check ([string]($plan.Assignments | Where-Object JobName -eq 'BlockedHere').OwnerServer -eq 'SERVER-B') 'A job blocked on a server was still assigned to it.'
    }
    $onlyBlocked = Get-SmartM365OrchestratorElectionPlan -Jobs $electionJobs -ServerCapabilities @($electionCapabilities[0])
    Check (@($onlyBlocked.UnassignedGroups | Where-Object { 'BlockedHere' -in @($_.Jobs) }).Count -eq 1 -and [string]($onlyBlocked.Assignments | Where-Object JobName -eq 'Other').OwnerServer -eq 'SERVER-A') 'The only server lost unrelated jobs or kept the blocked one.'

    # ---------- Orchestrator gate, alerts and recovery ----------
    $orchestratorPath = Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'
    $ast = [Management.Automation.Language.Parser]::ParseFile($orchestratorPath, [ref]$null, [ref]$null)
    function Get-FunctionText([string]$Name) { $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true).Extent.Text }
    $launchPhase = Get-FunctionText 'Invoke-LaunchPhase'
    $gateIndex = $launchPhase.IndexOf('Test-OrchestratorLaunchPrerequisiteGate')
    Check ($gateIndex -gt 0 -and $gateIndex -lt $launchPhase.IndexOf('Enter-SmartM365OrchestratorMaintenanceGate') -and $gateIndex -lt $launchPhase.IndexOf('Enter-SmartM365OrchestratorConcurrencyLease') -and $gateIndex -lt $launchPhase.IndexOf('Enter-SmartM365OrchestratorOccurrenceClaim') -and $gateIndex -lt $launchPhase.IndexOf('Start-InventoryJob')) 'The prerequisite gate does not run before the gate, lease, claim and launch.'
    Check ((Get-Content -LiteralPath $orchestratorPath -Raw).Contains('Invoke-OrchestratorPrerequisiteMonitoring -Now $now')) 'Prerequisite monitoring is not wired in the scheduler loop.'

    $names = @('Get-OrchestratorPipelineWaitStartUtc', 'Get-OrchestratorPrerequisiteBlockTable', 'Get-OrchestratorTenantDataPathCheck', 'Get-OrchestratorJobPrerequisitePathCheck', 'Test-OrchestratorJobLaunchPrerequisite', 'Register-OrchestratorPrerequisiteFailure', 'Publish-OrchestratorEffectiveCapability', 'Test-OrchestratorLaunchPrerequisiteGate', 'Send-OrchestratorPrerequisiteAlert', 'Update-OrchestratorCertificateExpiryWarning', 'Invoke-OrchestratorPrerequisiteMonitoring', 'ConvertTo-OrchestratorUtcTime', 'Resolve-OrchestratorJobPath', 'Resolve-OrchestratorJobArguments')
    $stubs = @'
function Write-OrchestratorLog { param([string]$Message, [string]$Level = 'INFO') $script:Logs.Add([pscustomobject]@{ Level = $Level; Message = $Message }) }
function Add-JobRunCsvRow { param($JobName, $ScheduledTime, $StartTime, $EndTime, $DurationSec, $ExitCode, $Status, $RetryCount, $LogPath) $script:Rows.Add([pscustomobject]@{ JobName = $JobName; Status = $Status; ScheduledTime = $ScheduledTime }) }
function Set-OrchestratorPipelineJobStatus { param($BatchId, $JobName, $Status, $Attempt, $Detail) $script:PipelineStatuses.Add([pscustomobject]@{ JobName = $JobName; Status = $Status; Detail = $Detail }) }
function Send-OrchestratorMail { param($Subject, $HtmlBody, [switch]$IsError) $script:Mails.Add([pscustomobject]@{ Subject = $Subject; IsError = [bool]$IsError }); return $true }
function Write-FileAtomically { param($Path, $Content) [IO.File]::WriteAllText($Path, $Content) }
function ConvertTo-HtmlText { param($Text) [Net.WebUtility]::HtmlEncode([string]$Text) }
function Get-SmartM365ScriptConfigValue { param($Config, $Name, $DefaultValue) $DefaultValue }
function Get-OrchestratorJobLaunchDefinition { param($Job) $null }
function Test-SmartM365OrchestratorLaunchPrerequisite { SmartM365.Orchestrator.Distributed\Test-SmartM365OrchestratorLaunchPrerequisite @args -CertificateProvider $global:PrereqTestProvider }
function Get-SmartM365OrchestratorCachedCertificateReadiness { SmartM365.Orchestrator.Distributed\Get-SmartM365OrchestratorCachedCertificateReadiness @args -CertificateProvider $global:PrereqTestProvider }
'@
    $gate = New-Module -ScriptBlock ([scriptblock]::Create((($names | ForEach-Object { Get-FunctionText $_ }) -join "`n") + "`n" + $stubs))
    $capabilityPath = Join-Path $root 'Orchestrator-Capabilities.json'
    function Reset-Gate {
        param([object[]]$TenantDataRoots = @())
        & $gate {
            param($CapabilityPath, $DataRoots)
            $script:Logs = [Collections.Generic.List[object]]::new(); $script:Rows = [Collections.Generic.List[object]]::new()
            $script:PipelineStatuses = [Collections.Generic.List[object]]::new(); $script:Mails = [Collections.Generic.List[object]]::new()
            $script:Settings = [pscustomobject]@{
                LaunchPrerequisiteChecksEnabled = $true; PrerequisiteCacheMinutes = 5; PrerequisitePathTimeoutSeconds = 10; PrerequisiteAlertRepeatHours = 24
                CertificateExpiryWarningDays = 30; DependencyWaitTimeoutMinutes = 60; CapabilitiesPath = $CapabilityPath; TenantDataRoots = @($DataRoots)
            }
            $script:State = @{ Jobs = @{} }
            $script:PrerequisiteCache = @{}; $script:PrerequisiteBlockedOccurrences = @{}; $script:ForcedPending = @(); $script:LastCertificateReadiness = $null
            $script:ServerCapabilitiesBase = [pscustomobject]@{ ServerName = 'SERVER-A'; GeneratedAtUtc = '2026-10-21T00:00:00.0000000Z'; ReadyCapabilities = @('SharedRuntime', 'Graph', 'EXO', 'AD'); BlockedJobs = @(); PrerequisiteBlocks = @() }
            $script:ServerCapabilities = $script:ServerCapabilitiesBase
            $script:LastElectionPlanRefreshUtc = [datetime]::UtcNow
            $script:Manifest = [pscustomobject]@{ JobsByName = @{} }
            $script:Tenant = 'test'; $script:ScriptName = 'SmartM365-Inventory-Orchestrator'; $script:ScriptVersion = 'test'
            $global:Thumbprint = 'AB' * 20
        } $capabilityPath $TenantDataRoots
    }
    function New-Job([string]$Name, [string]$Mode, [string[]]$Capabilities, [string]$Script = $scriptPath) {
        $job = [pscustomobject]@{ Name = $Name; AssignmentMode = $Mode; RequiredCapabilities = $Capabilities; ScriptPath = $Script; LauncherPath = ''; RequiredPaths = @(); DependencyWaitTimeoutMinutes = 0 }
        & $gate { param($j) $script:Manifest.JobsByName[$j.Name] = $j } $job
        return $job
    }
    function Invoke-Gate($Job, [datetime]$Occurrence, [string]$Reason = 'due', [switch]$IsForced, $PipelineInfo = $null, [datetime]$Now = (Get-Date)) {
        & $gate { param($j, $o, $r, $f, $p, $n) Test-OrchestratorLaunchPrerequisiteGate -Job $j -Occurrence $o -Reason $r -IsForced:$f -PipelineInfo $p -Now $n } $Job $Occurrence $Reason ([bool]$IsForced) $PipelineInfo $Now
    }
    function Get-GateValue([scriptblock]$Script) { & $gate $Script }
    $dataRoots = @([pscustomobject]@{ Name = 'Tenant DATA-LAST root'; Path = $dataLast })

    # Everything ready: no row, no block, capabilities unchanged.
    Reset-Gate -TenantDataRoots $dataRoots
    $global:PrereqTestCandidates = @(New-Candidate (New-TestCertificate -NotAfterDays 200))
    $graphJob = New-Job 'GraphJob' 'Elected' @('SharedRuntime', 'Graph')
    $occurrence = (Get-Date).Date.AddHours(1)
    Check (Invoke-Gate $graphJob $occurrence) 'A ready elected job was blocked.'
    Check ((Get-GateValue { $script:Rows.Count + $script:Mails.Count + @($script:Logs | Where-Object Level -eq 'ERROR').Count + $script:State.PrerequisiteBlocks.Count }) -eq 0) 'A ready launch produced rows, mails, errors or blocks.'
    Check (-not (Test-Path -LiteralPath $capabilityPath)) 'A ready launch rewrote the capability document.'
    & $gate { param($n) Invoke-OrchestratorPrerequisiteMonitoring -Now $n } (Get-Date)
    Check ((Get-GateValue { $script:Mails.Count }) -eq 0) 'Monitoring alerted while everything was ready.'

    # Certificate absent: elected job not launched, cloud capabilities withdrawn, one row per occurrence.
    Reset-Gate -TenantDataRoots $dataRoots
    $global:PrereqTestCandidates = @()
    $graphJob = New-Job 'GraphJob' 'Elected' @('SharedRuntime', 'Graph')
    $now = Get-Date
    Check (-not (Invoke-Gate $graphJob $occurrence -Now $now)) 'An elected Graph job was launched without certificate.'
    $capabilities = Get-GateValue { $script:ServerCapabilities }
    Check ('Graph' -notin @($capabilities.ReadyCapabilities) -and 'EXO' -notin @($capabilities.ReadyCapabilities) -and 'SharedRuntime' -in @($capabilities.ReadyCapabilities) -and 'AD' -in @($capabilities.ReadyCapabilities)) 'The certificate failure did not withdraw exactly the cloud capabilities.'
    Check ((Get-GateValue { $script:LastElectionPlanRefreshUtc }) -eq [datetime]::MinValue) 'The election plan refresh was not forced.'
    $published = Get-Content -LiteralPath $capabilityPath -Raw | ConvertFrom-Json
    Check ('Graph' -notin @($published.ReadyCapabilities) -and @($published.PrerequisiteBlocks).Count -eq 1) 'The withdrawn capability was not published for the election.'
    Check ((Get-GateValue { $script:ServerCapabilitiesBase.ReadyCapabilities }) -contains 'Graph') 'The probed capabilities were modified instead of overlaid.'
    Check (-not (Invoke-Gate $graphJob $occurrence -Now $now.AddMinutes(1))) 'The blocked job was launched on the next tick.'
    $rows = Get-GateValue { @($script:Rows) }
    Check ($rows.Count -eq 1 -and $rows[0].Status -eq 'BlockedPrerequisite') 'The blocked occurrence was not recorded exactly once as BlockedPrerequisite.'
    Check (@($rows | Where-Object Status -in 'Success', 'Failed', 'CompletedWithWarnings').Count -eq 0) 'A blocked occurrence was recorded as executed.'
    Check (@(Get-GateValue { @($script:Logs) } | Where-Object { $_.Level -eq 'ERROR' -and $_.Message -match 'BlockedPrerequisite' -and $_.Message -match 'reassign' }).Count -eq 1) 'The block reason was not logged clearly.'
    Check (Invoke-Gate (New-Job 'AdJob' 'Elected' @('SharedRuntime', 'AD')) $occurrence -Now $now) 'A job without cloud capability was blocked by the certificate.'

    # Alert sent once, repeated after PrerequisiteAlertRepeatHours, then recovery.
    & $gate { param($n) Invoke-OrchestratorPrerequisiteMonitoring -Now $n } $now
    & $gate { param($n) Invoke-OrchestratorPrerequisiteMonitoring -Now $n } $now.AddMinutes(10)
    $mails = Get-GateValue { @($script:Mails) }
    Check ($mails.Count -eq 1 -and $mails[0].IsError -and $mails[0].Subject -match 'prerequisites missing') 'The missing prerequisite alert was not sent exactly once.'
    & $gate { param($n) Invoke-OrchestratorPrerequisiteMonitoring -Now $n } $now.AddHours(25)
    Check ((Get-GateValue { $script:Mails.Count }) -eq 2) 'The alert was not repeated after the repeat period.'
    $global:PrereqTestCandidates = @(New-Candidate (New-TestCertificate -NotAfterDays 200))
    & $gate { param($n) Invoke-OrchestratorPrerequisiteMonitoring -Now $n } $now.AddHours(26)
    $mails = Get-GateValue { @($script:Mails) }
    Check ($mails.Count -eq 3 -and $mails[2].Subject -match 'recovered' -and -not $mails[2].IsError) 'No recovery mail followed the restored certificate.'
    Check ((Get-GateValue { $script:State.PrerequisiteBlocks.Count }) -eq 0 -and 'Graph' -in @((Get-GateValue { $script:ServerCapabilities }).ReadyCapabilities)) 'The restored certificate did not republish the cloud capabilities.'
    Check (Invoke-Gate $graphJob $occurrence -Now $now.AddHours(26)) 'The recovered job stayed blocked.'

    # Pinned job whose runtime script is missing: job-specific block, occurrence stays due.
    Reset-Gate -TenantDataRoots $dataRoots
    $pinned = New-Job 'PinnedJob' 'Pinned' @('SharedRuntime') (Join-Path $root 'MissingRuntime\Collector.ps1')
    Check (-not (Invoke-Gate $pinned $occurrence)) 'A pinned job was launched without its runtime script.'
    $capabilities = Get-GateValue { $script:ServerCapabilities }
    Check (@($capabilities.BlockedJobs) -contains 'PinnedJob' -and @($capabilities.ReadyCapabilities).Count -eq 4) 'A missing script withdrew a server capability or was not job-specific.'
    Check (@(Get-GateValue { @($script:Logs) } | Where-Object { $_.Message -match 'stays due' }).Count -eq 1) 'The pinned outcome was not explained.'
    [IO.File]::WriteAllText((New-Item -ItemType Directory -Path (Join-Path $root 'MissingRuntime') -Force).FullName + '\Collector.ps1', '# restored')
    & $gate { param($n) Invoke-OrchestratorPrerequisiteMonitoring -Now $n } (Get-Date).AddMinutes(6)
    Check (@((Get-GateValue { $script:ServerCapabilities }).BlockedJobs).Count -eq 0) 'The restored script did not lift the job block.'

    # Read-only data root: SharedRuntime withdrawn for every job.
    Reset-Gate -TenantDataRoots @([pscustomobject]@{ Name = 'Tenant DATA-LAST root'; Path = $readOnly })
    $global:PrereqTestCandidates = @(New-Candidate (New-TestCertificate -NotAfterDays 200))
    Check (-not (Invoke-Gate (New-Job 'AdJob' 'Elected' @('SharedRuntime', 'AD')) $occurrence)) 'A job was launched with a read-only data root.'
    Check ('SharedRuntime' -notin @((Get-GateValue { $script:ServerCapabilities }).ReadyCapabilities)) 'A read-only data root did not withdraw SharedRuntime.'

    # Forced and pipeline requests.
    Reset-Gate -TenantDataRoots $dataRoots
    $global:PrereqTestCandidates = @()
    & $gate { $script:ForcedPending = @('GraphJob', 'Other') }
    Check (-not (Invoke-Gate (New-Job 'GraphJob' 'Pinned' @('Graph')) (Get-Date) -Reason 'forced' -IsForced)) 'A forced job was launched without certificate.'
    Check ((@(Get-GateValue { $script:ForcedPending }) -join ',') -eq 'Other') 'The refused -Force request was not consumed.'
    Reset-Gate -TenantDataRoots $dataRoots
    $pipelineJob = New-Job 'GraphJob' 'Elected' @('Graph')
    $recent = @{ BatchId = 'batch-1'; CreatedAtUtc = [datetime]::UtcNow.AddMinutes(-5); Attempt = 0 }
    Check (-not (Invoke-Gate $pipelineJob $occurrence -Reason 'pipeline' -PipelineInfo $recent)) 'A pipeline job was launched without certificate.'
    $statuses = Get-GateValue { @($script:PipelineStatuses) }
    Check ($statuses.Count -eq 1 -and $statuses[0].Status -eq 'Pending' -and $statuses[0].Detail -match 'Waiting for launch prerequisites') 'The pipeline request did not show the prerequisite wait.'
    $old = @{ BatchId = 'batch-1'; CreatedAtUtc = [datetime]::UtcNow.AddMinutes(-90); Attempt = 0 }
    Check (-not (Invoke-Gate $pipelineJob $occurrence -Reason 'pipeline' -PipelineInfo $old)) 'A timed-out pipeline job was launched.'
    Check ((Get-GateValue { @($script:PipelineStatuses) })[-1].Status -eq 'BlockedPrerequisite') 'The pipeline request was not closed as BlockedPrerequisite after the wait timeout.'

    # Expiry warning: once per day.
    Reset-Gate -TenantDataRoots $dataRoots
    $global:PrereqTestCandidates = @(New-Candidate (New-TestCertificate -NotAfterDays 10))
    # Fixed time of day: two checks 45 minutes apart must stay on the same day (failed after 23:15).
    $day = (Get-Date).Date.AddHours(10)
    & $gate { param($n) Invoke-OrchestratorPrerequisiteMonitoring -Now $n } $day
    & $gate { param($n) Invoke-OrchestratorPrerequisiteMonitoring -Now $n } $day.AddMinutes(45)
    $mails = Get-GateValue { @($script:Mails) }
    Check ($mails.Count -eq 1 -and $mails[0].Subject -match 'expires in (9|10) day') 'The certificate expiry warning was not sent exactly once.'
    Check ((Get-GateValue { $script:LastCertificateReadiness }).ExpiresSoon) 'The expiry state was not kept for the heartbeat.'
    & $gate { $script:Settings.CertificateExpiryWarningDays = 0 }
    & $gate { param($n) $script:State.CertificateExpiryWarning.LastSentDate = ''; Invoke-OrchestratorPrerequisiteMonitoring -Now $n } $day.AddDays(1)
    Check ((Get-GateValue { $script:Mails.Count }) -eq 1) 'The expiry warning was sent while disabled.'

    # Checks disabled: no gate.
    Reset-Gate -TenantDataRoots $dataRoots
    $global:PrereqTestCandidates = @()
    & $gate { $script:Settings.LaunchPrerequisiteChecksEnabled = $false }
    Check (Invoke-Gate (New-Job 'GraphJob' 'Elected' @('Graph')) $occurrence) 'Disabled checks still blocked a launch.'
}
finally {
    $acl = Get-Acl -LiteralPath $readOnly
    [void]$acl.RemoveAccessRule($denyWrite)
    Set-Acl -LiteralPath $readOnly -AclObject $acl
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Variable -Name PrereqTestCandidates, PrereqTestProviderCalls, PrereqTestProvider, Thumbprint -Scope Global -ErrorAction SilentlyContinue
}
"Launch prerequisite offline tests passed: $script:passed checks."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCebsAMD+Jfq3OW
# l4FdbWuQPZOcGU5COh5iWbOv7cXCk6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMy6onL0uJsspX0x3IwWdjGY7fQHtnCSAitUy/IX68JqMA0GCSqG
# SIb3DQEBAQUABIIBgEjDQurHOPYCkohNKOi433QckxSf0J8UJwqQ1YS+7Nz+Ecyy
# NSPqeYmJkwUHqFTer0am+M20CSHcs21z5eYh3SvcgMe7fQ8v+hq9h2RxnQoc6Mvt
# qryQSvcQeI42uhf0NK0MdvhKA+AS28GEdnSQCoRymRDH5rGZLjF+YViSWbgs9lbq
# //TJpAPZhgb1aY0SASjBh2+OXZbG29qNssd2u73z7sIMn8B/FoDVV6kCwUdycMB9
# z35LwOctzHSp5Zvb66daf/XFKyDlhTMPHrmV+l7fxFcz3e6BlP9geqeZhKR99FvX
# fGigHys9nJcxAc/Ife+GwKStmNqDE6VhFd3EHKTLA1kBkm/DKSdXQ4SYKgUNWF6+
# t+TTkvszPeXK0N9ORQFwZAdqtxwzjer9/vEg7X0ccFTWYUa+qxBihJZN5FiEbzmG
# twJCEgl12hv7UweTfgeZfErqIli6i3qfrnRLALYKDbbiXv2NfK+GhXqyJBHkp1/k
# BUck33+WaF69gdKecKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgxMDA2
# NDRaMC8GCSqGSIb3DQEJBDEiBCDbyZEYuUC/Zz12LmS50cOtqqdjd654XznbTIha
# H/Pf5jANBgkqhkiG9w0BAQEFAASCAgB6SMaFLSBKTC0CXKyJhWbxgzPsaG1q/EeN
# gcT2dK30C1Q1I1pwjl/ZpbggPePVPVAdqhsDiOlp65dMdwOM/+lqSQc698Sx43+F
# SP9rF7xcFbQMXOiqY+/cY2G46KVs+jvQ3f4VJ10uc6xodwKpAvuO+4tHZe74+htt
# KwAasEGBm8wVNSoOtAEmBUhPWrJXM5CJlGvdboqSmHo1M3tDRUHxJB7yxIQThO3d
# MyyvklAqc1XPBAKq5PN6WJFikGeEx5cH85LNs/yfdMwUQsgg2YOSQ5+aKetEiHOp
# 2jtaB+RgKqUTUw2UP+Zt9MZlj5jXqicaYpErQhM6h1k2NncxyaSElVowPioan/b6
# HMIAIVh1FbE2CBVY6mAgaebrVa0+0WrsRaPb9TwSR28K8LogjOJMezQC09i6ny3c
# QVO9ihumc6AuocpPPewvSaAMVvBeqF4AQaS1/ZbGcllDOt75XESrS1JDgoJcGPfj
# moCjhv4NoChP7hLy0s0MxKuo/ngoSsLb57yxB41pByfMuj0CRGJca2uVHRX0Tb+s
# X+ZJgTgqh+pcuSZIennXFuI12l62nvwO558IVL/oGnm0Q4bT2NEcrfuHIcjjzurZ
# cuJrr0jwopaGrqiydNL3PBMkTuSlDeBv5W8LfvKYJR8IBoZ4SebQFCB4zsCSb9qK
# Lo1pq3SjZw==
# SIG # End signature block
