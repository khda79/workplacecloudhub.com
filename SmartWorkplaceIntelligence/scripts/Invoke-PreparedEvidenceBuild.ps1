# Version 1.9.0: directory cleanup candidates (AD users, AD workstations, user mailboxes); earlier: optional Microsoft 365 usage signals, hybrid identity findings and device readiness gap facts.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$DataRoot,
    [Parameter(Mandatory)][string]$OutputRoot,
    [ValidateSet('All','Estate','Workforce','Devices','Applications','IntuneOperations','Collaboration','Content','Mailbox','HybridIdentity','DirectoryCleanup','DailyTrends','Security','Backup','Trust','Executive','SharePointMigration')]
    [string]$Only = 'All',
    [switch]$SkipWorkforceHistory,
    [switch]$ReuseStagedWorkforce,
    [string]$AccountClassificationConfigPath,
    [ValidatePattern('^[A-Fa-f0-9]{64}$')][string]$AccountClassificationExpectedSHA256,
    [Nullable[long]]$AccountClassificationExpectedBytes,
    [string]$AdCountryRulesPath,
    [string]$EndpointProtectionPolicyPath,
    [Nullable[double]]$DeviceSystemMaxObservationAgeHours,
    [string]$DeviceSystemOffsetlessGraphTimeZoneId,
    [string]$TenantKey,
    [string]$IntuneSourcePolicyPath,
    [string]$NativeEvidenceAuditPath,
    [string]$PythonCommand='python',
    [switch]$MailboxMigrationExceptionsEnabled,
    [string]$MailboxMigrationBindingsPath,
    [switch]$SharePointMigrationEnabled,
    [string]$SharePointMigrationWorkbookFileName='SmartWorkplaceIntelligence-Migration-to-SPO.xlsx',
    [string]$SharePointMigrationBindingsFileName='',
    [string]$SharePointMigrationTargetTenantUrl='',
    [string]$SharePointMigrationPreviousHistoryPath='',
    [int]$SharePointMigrationReviewMaxAgeDays=7,
    [int]$SharePointMigrationMaxSourceAgeHours=168,
    [string]$SharePointMigrationReportingTimeZoneId='UTC',
    [string]$DailyTrendsQualificationPath,
    [double]$DailyTrendsPartialThresholdPercent = 1,
    [int]$DailyTrendsPersistentAbsenceObservations = 7,
    # Hours after the last successful acquisition before the kept Office activation report is warned about.
    [ValidateRange(1,8760)][int]$OfficeActivationWarningAgeHours = 72,
    # TargetAddress domains accepted in addition to the accepted remote routing domain (private setting).
    [string[]]$HybridAllowedTargetDomains = @(),
    # Optional reference clock passed to every generator; omitted means each generator uses the current time.
    [Nullable[datetimeoffset]]$AsOfUtc,
    # 1 runs the generators one after another with live output; higher values run independent
    # generators concurrently, each in its own child process, and print each log when it completes.
    [ValidateRange(1,8)][int]$MaxParallelGenerators = 1
)

# Offline preparation only. No collection, upload, Desktop control or promotion.
# Each child process releases its memory when its domain completes.
$ErrorActionPreference = 'Stop'
$DataRoot = (Resolve-Path -LiteralPath $DataRoot).ProviderPath
$last = Join-Path $DataRoot 'DATA-LAST'
if (-not (Test-Path -LiteralPath $last -PathType Container)) { throw 'DATA-LAST is required.' }
$OutputRoot = [IO.Path]::GetFullPath($OutputRoot)
if ($OutputRoot.TrimEnd('\') -eq $DataRoot.TrimEnd('\') -or $OutputRoot.TrimEnd('\') -eq $last.TrimEnd('\')) {
    throw 'Outputs must use a dedicated staging directory, never the raw source root.'
}
New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
$shell = (Get-Process -Id $PID).Path
if ($PSVersionTable.PSVersion.Major -lt 7) { throw 'PowerShell 7 is required.' }
function OutPath([string]$Name) { Join-Path $OutputRoot ($Name + '.csv') }
# Generators in their sequential order. DependsOn names the tasks whose staged outputs a
# generator reads; a dependency excluded by -Only is treated as already staged.
$script:tasks = [Collections.Generic.List[object]]::new()
function Add-Task([string]$Id, [string]$Domain, [string]$Script, [object]$Arguments, [string[]]$DependsOn = @(), [scriptblock]$BeforeStart, [switch]$Native) {
    if ($Only -ne 'All' -and $Only -ne $Domain) { return }
    $script:tasks.Add([pscustomobject]@{Id=$Id;Domain=$Domain;Script=$Script;Arguments=$Arguments;DependsOn=@($DependsOn);BeforeStart=$BeforeStart;Native=[bool]$Native})
}
$clockArgs=@();$clockDateArgs=@();$clockUtcArgs=@()
if ($null -ne $AsOfUtc) {
    $reference=$AsOfUtc.ToUniversalTime()
    $clockArgs=@('-AsOfUtc',$reference.ToString('o',[Globalization.CultureInfo]::InvariantCulture))
    $clockDateArgs=@('-AsOfDate',$reference.UtcDateTime.ToString('yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture))
    $clockUtcArgs=@('-AsOfUtc',$reference.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ss.fffffff',[Globalization.CultureInfo]::InvariantCulture))
}
$countryArgs=@();if($AdCountryRulesPath){$countryArgs=@('-AdCountryRulesPath',$AdCountryRulesPath)}
$classificationStart = $null
if (($AccountClassificationExpectedSHA256 -or $null -ne $AccountClassificationExpectedBytes) -and -not $AccountClassificationConfigPath) {
    throw 'An expected account classification capture requires its explicit file path.'
}
if ($Only -in @('All','Workforce','HybridIdentity','DirectoryCleanup') -and -not $ReuseStagedWorkforce -and $AccountClassificationConfigPath) {
    Import-Module (Join-Path $PSScriptRoot 'PreparedClassification.psm1')
    $classificationCheck = @{Path=$AccountClassificationConfigPath;Stage='BuildStart'}
    if ($AccountClassificationExpectedSHA256) { $classificationCheck.ExpectedSHA256=$AccountClassificationExpectedSHA256 }
    if ($null -ne $AccountClassificationExpectedBytes) { $classificationCheck.ExpectedBytes=$AccountClassificationExpectedBytes }
    $classificationStart = Test-PreparedClassificationCapture @classificationCheck
}
Add-Task 'Estate' 'Estate' 'New-EnterpriseEstateEvidence.ps1' (@('-DataRoot',$DataRoot,'-OutputRoot',$OutputRoot)+$countryArgs)
if ($Only -eq 'All') {
    if (-not $NativeEvidenceAuditPath) { throw 'Unified preparation requires a receipt-qualified native capture audit.' }
    Add-Task 'Native' 'All' $null @{LastRoot=$last;TenantKey=$TenantKey;AuditPath=$NativeEvidenceAuditPath;OutputRoot=$OutputRoot;PythonCommand=$PythonCommand} -Native
}
$deviceArgs=@('-DataRoot',$DataRoot,'-OutputPath',(OutPath 'DeviceInventoryEvidence'),'-SignalsOutputPath',(OutPath 'DeviceRiskEvidenceSummary'),'-DirectorySummaryOutputPath',(OutPath 'DeviceDirectorySummary'),'-SystemEvidenceOutputPath',(OutPath 'DeviceSystemEvidence'),'-ReadinessGapOutputPath',(OutPath 'DeviceReadinessGapEvidence'))
if ($null -ne $DeviceSystemMaxObservationAgeHours) { $deviceArgs+=@('-DeviceSystemMaxObservationAgeHours',$DeviceSystemMaxObservationAgeHours) }
if ($DeviceSystemOffsetlessGraphTimeZoneId) { $deviceArgs+=@('-DeviceSystemOffsetlessGraphTimeZoneId',$DeviceSystemOffsetlessGraphTimeZoneId) }
$deviceArgs+=$clockArgs
# Current workforce may use the current staged console evidence, never a future snapshot in history.
Add-Task 'DeviceInventory' 'Devices' 'New-DeviceInventoryEvidence.ps1' $deviceArgs
$wf = @('-DataRoot',$DataRoot,'-UserOutputPath',(OutPath 'UserInventoryEvidence'),'-HistoryOutputPath',(OutPath 'UserActivityHistoryEvidence'),'-TrendOutputPath',(OutPath 'WorkforceTrendEvidence'),'-IdentityOutputPath',(OutPath 'IdentityReconciliationEvidence'),'-SignalsOutputPath',(OutPath 'WorkforceOperationalSignals'))
if ($Only -eq 'All') { $wf+=@('-SystemEvidencePath',(OutPath 'DeviceSystemEvidence')) }
if ($AccountClassificationConfigPath) { $wf += @('-AccountClassificationConfigPath',$AccountClassificationConfigPath) }
if ($SkipWorkforceHistory) { $wf += '-SkipHistory' }
$wf+=@('-OfficeActivationWarningAgeHours',[string]$OfficeActivationWarningAgeHours)
$wf+=$clockArgs
if ($ReuseStagedWorkforce) {
    foreach ($name in @('UserInventoryEvidence','IdentityReconciliationEvidence','WorkforceOperationalSignals')) {
        if (-not (Test-Path -LiteralPath (OutPath $name))) { throw "Missing staged workforce output: $name" }
    }
} else {
    $beforeWorkforce = $null
    if ($null -ne $classificationStart) {
        $beforeWorkforce = { $null = Test-PreparedClassificationCapture -Path $AccountClassificationConfigPath -Stage 'BeforeWorkforce' -ExpectedSHA256 $classificationStart.SHA256 -ExpectedBytes $classificationStart.Bytes }
    }
    Add-Task 'Workforce' 'Workforce' 'New-WorkforceIdentityEvidence.ps1' $wf -DependsOn 'DeviceInventory' -BeforeStart $beforeWorkforce
}
if ($EndpointProtectionPolicyPath) {
    Add-Task 'EndpointProtection' 'Devices' 'New-EndpointProtectionEvidence.ps1' (@('-DataRoot',$DataRoot,'-InventoryEvidencePath',(OutPath 'DeviceInventoryEvidence'),'-SystemEvidencePath',(OutPath 'DeviceSystemEvidence'),'-PolicyPath',$EndpointProtectionPolicyPath,'-OutputRoot',$OutputRoot)+$clockArgs) -DependsOn 'DeviceInventory'
}
Add-Task 'WindowsMigration' 'Devices' 'New-WindowsMigrationEvidence.ps1' (@('-DataRoot',$DataRoot,'-PreparedInventoryPath',(OutPath 'DeviceInventoryEvidence'),'-OutputRoot',$OutputRoot)+$countryArgs) -DependsOn 'DeviceInventory'
Add-Task 'DeviceLifecycle' 'Devices' 'New-DeviceLifecycleExperienceTrendEvidence.ps1' @('-DataRoot',$DataRoot,'-WindowsOutputPath',(OutPath 'WindowsLifecycleTrendEvidence'),'-EndpointOutputPath',(OutPath 'EndpointExperienceTrendEvidence'))
Add-Task 'Applications' 'Applications' 'New-ApplicationInventoryEvidence.ps1' @('-DataRoot',$DataRoot,'-InventoryOutputPath',(OutPath 'ApplicationInventoryEvidence'),'-TrendOutputPath',(OutPath 'ApplicationTrendEvidence'))
Add-Task 'Collaboration' 'Collaboration' 'New-CollaborationEvidence.ps1' (@('-DataRoot',$DataRoot,'-AdoptionOutputPath',(OutPath 'CollaborationAdoptionEvidence'),'-ObjectOutputPath',(OutPath 'CollaborationObjectEvidence'),'-TrendOutputPath',(OutPath 'CollaborationTrendEvidence'))+$clockArgs)
Add-Task 'M365UsageSignals' 'Collaboration' 'New-M365UsageSignalsEvidence.ps1' (@('-DataRoot',$DataRoot,'-EmailClientOutputPath',(OutPath 'EmailClientUsageEvidence'),'-OneDriveActivityOutputPath',(OutPath 'OneDriveActivityEvidence'),'-GroupActivityOutputPath',(OutPath 'GroupActivityEvidence'),'-CopilotDailyOutputPath',(OutPath 'CopilotDailyAdoptionEvidence'))+$clockArgs)
Add-Task 'Content' 'Content' 'New-ContentStorageEvidence.ps1' (@('-DataRoot',$DataRoot,'-ObjectOutputPath',(OutPath 'ContentStorageEvidence'),'-TrendOutputPath',(OutPath 'ContentStorageTrendEvidence'))+$clockArgs)
Add-Task 'Mailbox' 'Mailbox' 'New-MailboxEvidence.ps1' @('-DataRoot',$last,'-UserEvidencePath',(OutPath 'UserInventoryEvidence'),'-OutputPath',(OutPath 'MailboxEvidence')) -DependsOn 'Workforce'
$hybridArgs=@('-DataRoot',$last,'-UserEvidencePath',(OutPath 'UserInventoryEvidence'),'-OutputPath',(OutPath 'HybridIdentityFindings'))
if ($AccountClassificationConfigPath) { $hybridArgs+=@('-AccountClassificationConfigPath',$AccountClassificationConfigPath) }
if (@($HybridAllowedTargetDomains).Count) { $hybridArgs+=@('-AllowedTargetDomains',(@($HybridAllowedTargetDomains) -join ',')) }
$hybridArgs+=$clockArgs
$beforeHybrid = $null
if ($null -ne $classificationStart) {
    $beforeHybrid = { $null = Test-PreparedClassificationCapture -Path $AccountClassificationConfigPath -Stage 'BeforeHybridIdentity' -ExpectedSHA256 $classificationStart.SHA256 -ExpectedBytes $classificationStart.Bytes }
}
Add-Task 'HybridIdentity' 'HybridIdentity' 'New-HybridIdentityEvidence.ps1' $hybridArgs -DependsOn 'Workforce' -BeforeStart $beforeHybrid
$cleanupArgs=@('-DataRoot',$DataRoot,'-TenantKey',$TenantKey,'-UserInventoryPath',(OutPath 'UserInventoryEvidence'),'-AllDevicesPath',(OutPath 'AllDevicesEvidence'),
    '-UserOutputPath',(OutPath 'UserCleanupEvidence'),'-ComputerOutputPath',(OutPath 'ComputerCleanupEvidence'),'-MailboxOutputPath',(OutPath 'MailboxCleanupEvidence'),'-QualificationOutputPath',(OutPath 'DirectoryCleanupSourceQualification'))
if ($AccountClassificationConfigPath) { $cleanupArgs+=@('-AccountClassificationConfigPath',$AccountClassificationConfigPath) }
$cleanupArgs+=$clockArgs
$beforeCleanup = $null
if ($null -ne $classificationStart) {
    $beforeCleanup = { $null = Test-PreparedClassificationCapture -Path $AccountClassificationConfigPath -Stage 'BeforeDirectoryCleanup' -ExpectedSHA256 $classificationStart.SHA256 -ExpectedBytes $classificationStart.Bytes }
}
Add-Task 'DirectoryCleanup' 'DirectoryCleanup' 'New-DirectoryCleanupEvidence.ps1' $cleanupArgs -DependsOn 'Workforce','Estate' -BeforeStart $beforeCleanup
if ($MailboxMigrationExceptionsEnabled) {
    $exceptionArgs=@('-DataRoot',$DataRoot,'-OutputRoot',$OutputRoot,'-TenantKey',$TenantKey,'-NativeEvidenceAuditPath',$NativeEvidenceAuditPath,'-PythonCommand',$PythonCommand,'-Enabled')
    if ($MailboxMigrationBindingsPath) { $exceptionArgs+=@('-BindingsPath',$MailboxMigrationBindingsPath) }
    Add-Task -Id 'MailboxExceptions' -Domain 'Mailbox' -Script 'New-MailboxMigrationExceptions.ps1' -Arguments $exceptionArgs -DependsOn 'Native'
}
$dailyTrendArgs=@('-DataRoot',$DataRoot,'-OutputRoot',$OutputRoot,'-PythonCommand',$PythonCommand,'-PartialThresholdPercent',[string]$DailyTrendsPartialThresholdPercent,'-PersistentAbsenceObservations',[string]$DailyTrendsPersistentAbsenceObservations)
if($DailyTrendsQualificationPath){$dailyTrendArgs+=@('-QualificationPath',$DailyTrendsQualificationPath)}
Add-Task 'DailyTrends' 'DailyTrends' 'New-DailyTrendsEvidence.ps1' $dailyTrendArgs
$migrationArgs=@('-DataRoot',$DataRoot,'-OutputRoot',$OutputRoot,'-TenantKey',$TenantKey,'-PythonCommand',$PythonCommand,
    '-WorkbookFileName',$SharePointMigrationWorkbookFileName,'-ReviewMaxAgeDays',[string]$SharePointMigrationReviewMaxAgeDays,
    '-MaxSourceAgeHours',[string]$SharePointMigrationMaxSourceAgeHours,'-ReportingTimeZoneId',$SharePointMigrationReportingTimeZoneId)+$clockArgs
if($SharePointMigrationEnabled){$migrationArgs+='-Enabled'}
if($SharePointMigrationBindingsFileName){$migrationArgs+=@('-BindingsFileName',$SharePointMigrationBindingsFileName)}
if($SharePointMigrationTargetTenantUrl){$migrationArgs+=@('-TargetTenantUrl',$SharePointMigrationTargetTenantUrl)}
if($SharePointMigrationPreviousHistoryPath){$migrationArgs+=@('-PreviousHistoryPath',$SharePointMigrationPreviousHistoryPath)}
if($IntuneSourcePolicyPath){$migrationArgs+=@('-SourcePolicyPath',$IntuneSourcePolicyPath)}
Add-Task 'SharePointMigration' 'SharePointMigration' 'New-SharePointMigrationEvidence.ps1' $migrationArgs -DependsOn 'Native'
Add-Task 'Security' 'Security' 'New-SecurityEvidence.ps1' (@('-DataRoot',$last,'-ExtendedEvidenceRoot',$last,'-UserEvidencePath',(OutPath 'UserInventoryEvidence'),'-DeviceEvidencePath',(OutPath 'DeviceInventoryEvidence'),'-OutputPath',(OutPath 'SecurityControlEvidence'))+$clockUtcArgs) -DependsOn 'Workforce','DeviceInventory'
Add-Task 'Backup' 'Backup' 'New-BackupResilienceEvidence.ps1' (@('-DataRoot',$last,'-OutputRoot',$OutputRoot)+$clockDateArgs)
$intuneArgs=@('-DataRoot',$DataRoot,'-OutputRoot',$OutputRoot)
if($TenantKey){$intuneArgs+=@('-TenantKey',$TenantKey)}
if($IntuneSourcePolicyPath){$intuneArgs+=@('-SourcePolicyPath',$IntuneSourcePolicyPath)}
$intuneArgs+=$clockArgs
Add-Task 'IntuneOperations' 'IntuneOperations' 'New-IntuneOperationalEvidence.ps1' $intuneArgs
$trustArgs=@('-DataRoot',$DataRoot,'-OutputPath',(OutPath 'DataTrustEvidence'))
if($Only -eq 'All'){$trustArgs+=@('-IntuneOperationalAuditPath',(Join-Path $OutputRoot 'IntuneOperationalAudit.json.txt'))}
if($MailboxMigrationExceptionsEnabled){$trustArgs+=@('-MailboxMigrationExceptionsAuditPath',(Join-Path $OutputRoot 'MailboxMigrationExceptionsAudit.json.txt'))}
$trustArgs+=@('-OfficeActivationWarningAgeHours',[string]$OfficeActivationWarningAgeHours)+$clockDateArgs+$clockUtcArgs
Add-Task 'Trust' 'Trust' 'New-DataTrustEvidence.ps1' $trustArgs -DependsOn 'IntuneOperations','MailboxExceptions'
Add-Task 'Executive' 'Executive' 'New-ExecutiveTrendEvidence.ps1' @('-DataRoot',$DataRoot,'-OutputPath',(OutPath 'ExecutiveKPITrends'))

function Write-TaskStart($Task) {
    $started = [datetime]::UtcNow
    if (-not $Task.Native) { Write-Host ('[{0:O}] Starting {1}' -f $started, $Task.Domain) }
    $started
}
function Write-TaskEnd($Task, [datetime]$Started) {
    if (-not $Task.Native) { Write-Host ('[{0:O}] Completed {1}: {2:n1} seconds' -f [datetime]::UtcNow, $Task.Domain, ([datetime]::UtcNow - $Started).TotalSeconds) }
}
if ($MaxParallelGenerators -le 1) {
    # Sequential: the former order with live output.
    foreach ($task in $script:tasks) {
        if ($task.BeforeStart) { & $task.BeforeStart }
        $started = Write-TaskStart $task
        if ($task.Native) {
            Import-Module (Join-Path $PSScriptRoot 'NativeEvidence.psm1') -Force
            $nativeArgs = $task.Arguments
            Invoke-NativeEvidenceStage @nativeArgs
        } else {
            $taskArgs = @($task.Arguments)
            & $shell -NoProfile -File (Join-Path $PSScriptRoot $task.Script) @taskArgs
            if ($LASTEXITCODE -ne 0) { throw "Preparation failed: $($task.Domain) (exit $LASTEXITCODE). Nothing was published." }
        }
        Write-TaskEnd $task $started
    }
} else {
    # Parallel: ready tasks start in the order below (critical path and longest generators first);
    # each runs in its own child process and its log is printed in one block when it completes.
    $priority = @('DeviceInventory','Workforce','Native','IntuneOperations','Applications','DeviceLifecycle','Estate','Trust','Mailbox','HybridIdentity','DirectoryCleanup','Security','WindowsMigration','EndpointProtection','Executive','Backup','Collaboration','M365UsageSignals','Content','DailyTrends')
    $taskIds = @($script:tasks.Id)
    $pending = [Collections.Generic.List[object]]::new()
    foreach ($id in $priority) { foreach ($task in $script:tasks) { if ($task.Id -ceq $id) { $pending.Add($task) } } }
    foreach ($task in $script:tasks) { if (-not $pending.Contains($task)) { $pending.Add($task) } }
    $running = @{}
    $completed = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $worker = {
        param($Shell, $Root, $Task)
        $lines = [Collections.Generic.List[string]]::new()
        $exitCode = 0
        try {
            if ($Task.Native) {
                Import-Module (Join-Path $Root 'NativeEvidence.psm1') -Force
                $nativeArgs = $Task.Arguments
                # A stringified WarningRecord loses its prefix; keep it so the entry point counts it.
                Invoke-NativeEvidenceStage @nativeArgs *>&1 | ForEach-Object {
                    if ($_ -is [Management.Automation.WarningRecord]) { $lines.Add('WARNING: ' + $_.Message) } else { $lines.Add([string]$_) }
                }
            } else {
                $taskArgs = @($Task.Arguments)
                & $Shell -NoProfile -File (Join-Path $Root $Task.Script) @taskArgs 2>&1 | ForEach-Object { $lines.Add([string]$_) }
                $exitCode = $LASTEXITCODE
            }
        } catch { $lines.Add($_.Exception.Message); $exitCode = 1 }
        [pscustomobject]@{ExitCode=$exitCode;Lines=$lines.ToArray()}
    }
    try {
        while ($pending.Count -gt 0 -or $running.Count -gt 0) {
            foreach ($task in @($pending)) {
                if ($running.Count -ge $MaxParallelGenerators) { break }
                if (@($task.DependsOn | Where-Object { $_ -in $taskIds -and -not $completed.Contains($_) }).Count -gt 0) { continue }
                if ($task.BeforeStart) { & $task.BeforeStart }
                $started = Write-TaskStart $task
                $job = Start-ThreadJob -ScriptBlock $worker -ArgumentList $shell,$PSScriptRoot,$task
                $running[$job.Id] = [pscustomobject]@{Task=$task;Job=$job;Started=$started}
                [void]$pending.Remove($task)
            }
            if ($running.Count -eq 0) { throw 'Preparation failed: generator dependencies cannot be satisfied. Nothing was published.' }
            $finished = Wait-Job -Job @($running.Values | ForEach-Object Job) -Any
            $entry = $running[$finished.Id]
            $running.Remove($finished.Id)
            $result = @(Receive-Job -Job $finished -Wait -AutoRemoveJob -ErrorAction Stop) | Where-Object { $null -ne $_ -and $_.PSObject.Properties['ExitCode'] } | Select-Object -Last 1
            if ($null -eq $result) { throw "Preparation failed: $($entry.Task.Domain) returned no result. Nothing was published." }
            foreach ($line in $result.Lines) { foreach ($part in ($line -split '\r?\n')) { Write-Host $part } }
            if ($result.ExitCode -ne 0) { throw "Preparation failed: $($entry.Task.Domain) (exit $($result.ExitCode)). Nothing was published." }
            Write-TaskEnd $entry.Task $entry.Started
            [void]$completed.Add($entry.Task.Id)
        }
    } finally {
        # On failure, stop the generators still running; this script publishes nothing.
        foreach ($entry in @($running.Values)) {
            Stop-Job -Job $entry.Job -ErrorAction SilentlyContinue
            Remove-Job -Job $entry.Job -Force -ErrorAction SilentlyContinue
        }
    }
}
Write-Host 'Preparation finished. Outputs are staging only; schema/data validation and promotion are separate steps.'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD8Fm5Nrwv3N45P
# VlZtDPo6W99tK1oUZeEZkOW6RNwN3qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICGqpcY9J96Tm2cngMsBf+9BMxvOLZ05QeZKHxMNUcVHMA0GCSqG
# SIb3DQEBAQUABIIBgJFzz6+eJ5Rb1sAy//bfT4kGGdaUlV0UD/RBgvZC6Lw1QEej
# Ob0nJ1OS/MZopBUI69YOyZKyx8Zk6tkXn91D01mQ3A+nr6hwnoOpSBIb76kApUwo
# 2bTIuuCaYBMRjiSbkCBWT93uaNhYOv7Nf1wCVnoVW59BvjXlIJCqhzlLhVWl1YRj
# V4rj4Q3lt7M9LHkz/mVNbpYUChbOpkasf908lRiZNmm5zEsxWYpIabUz38uyCSoJ
# SChkHw3MhvKCJBWo1dXyAT15SWAdkO50x9DcDrkIJPSQgn4lpVvCATXCOa6dew6w
# GbLwfk9DCVKbGoW24dD0Dci0CDbhn/yYLsOuh+DQw+//Z126Q6O8D84SMVdiGQER
# aF5E5+wVzPtUFLkBctqyB8ZPWoGqU8h6DUGZwKYYxl1HipZd3tHln+m0mmwpjHnl
# aFAmWs0aP1GB/Xlt9PTBmBtacAuLi/mhZHzQgQN0Ceze4p94pG5pAprvDnFq6/3g
# 7GGFVAXvLfv9LHpGFaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgwOTE4
# MDVaMC8GCSqGSIb3DQEJBDEiBCAv1AS/qMl5q0Fggv2gx4B66bcoBlOq8YpThZRX
# EDNesjANBgkqhkiG9w0BAQEFAASCAgBaHepB2qM7NTJ06NMT6MzUiy0sVr7mTkOT
# QcqxoOydJvPJKK4LLQp/BpfuKVqiSQjrMxaANmOg7iTAk3r1x2ehlqFGz1UTSnOW
# a8kdhpjKq4LdSkl25wSwhE3U0vpTYRCPfZlV6FAe8QjkM3Cy28iXqD2Od4hG44Sp
# tPobxOif1KtIXAhc77nbM7TnbUyfpcXmXbjGFPSE7CNkKO7rmCiYyjj85V8he3pb
# MsmiAfHt6Jx3isRBwn5TkllPHdHLZgN4mJUfrKcoRsWi1AjS9W3XXiybZDMVgYW/
# yhW+0R5F/KwtGDX7+P1jhPUyEQ2IagIhA23vzBXC0sr6erQdb+9bdDzuSsesUgVj
# VNOwJGZ4czLrTKAdoiCjKD1qez2LWdj+X5n9IEEAVtBY9ESNa0QEEVE1Pyj0qXjW
# k6WiSZlJ9co661EQRIGng8XNNGGXaKzsxzlZscfR6VxwMOjIKBV9srg+W0NsqKWc
# oLukvSo7y4W2pVQnkajVObDD77ZA3VEhxIuRu6rPAccyevCA1pMzXiQxpXr3zUeG
# SGzAGHjtsE/nwnv18NVZw8or3fr1VhLY8aeSo+OOkujDiYZsYaMr4kqhFfOuFek3
# WHceS9TL88sparO5ovPv7B7Mc7Xt+NtFELXoaDncBoHTQ3ul+ORbR8XEjgC5XUMv
# rQSUZZC2qg==
# SIG # End signature block
