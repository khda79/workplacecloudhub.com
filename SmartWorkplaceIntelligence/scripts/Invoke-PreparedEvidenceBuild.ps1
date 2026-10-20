# Version 1.5.0: unified SPO migration planning follows native source qualification.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$DataRoot,
    [Parameter(Mandatory)][string]$OutputRoot,
    [ValidateSet('All','Estate','Workforce','Devices','Applications','IntuneOperations','Collaboration','Content','Mailbox','DailyTrends','Security','Backup','Trust','Executive','SharePointMigration')]
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
if ($Only -in @('All','Workforce') -and -not $ReuseStagedWorkforce -and $AccountClassificationConfigPath) {
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
$deviceArgs=@('-DataRoot',$DataRoot,'-OutputPath',(OutPath 'DeviceInventoryEvidence'),'-SignalsOutputPath',(OutPath 'DeviceRiskEvidenceSummary'),'-DirectorySummaryOutputPath',(OutPath 'DeviceDirectorySummary'),'-SystemEvidenceOutputPath',(OutPath 'DeviceSystemEvidence'))
if ($null -ne $DeviceSystemMaxObservationAgeHours) { $deviceArgs+=@('-DeviceSystemMaxObservationAgeHours',$DeviceSystemMaxObservationAgeHours) }
if ($DeviceSystemOffsetlessGraphTimeZoneId) { $deviceArgs+=@('-DeviceSystemOffsetlessGraphTimeZoneId',$DeviceSystemOffsetlessGraphTimeZoneId) }
$deviceArgs+=$clockArgs
# Current workforce may use the current staged console evidence, never a future snapshot in history.
Add-Task 'DeviceInventory' 'Devices' 'New-DeviceInventoryEvidence.ps1' $deviceArgs
$wf = @('-DataRoot',$DataRoot,'-UserOutputPath',(OutPath 'UserInventoryEvidence'),'-HistoryOutputPath',(OutPath 'UserActivityHistoryEvidence'),'-TrendOutputPath',(OutPath 'WorkforceTrendEvidence'),'-IdentityOutputPath',(OutPath 'IdentityReconciliationEvidence'),'-SignalsOutputPath',(OutPath 'WorkforceOperationalSignals'))
if ($Only -eq 'All') { $wf+=@('-SystemEvidencePath',(OutPath 'DeviceSystemEvidence')) }
if ($AccountClassificationConfigPath) { $wf += @('-AccountClassificationConfigPath',$AccountClassificationConfigPath) }
if ($SkipWorkforceHistory) { $wf += '-SkipHistory' }
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
Add-Task 'Content' 'Content' 'New-ContentStorageEvidence.ps1' (@('-DataRoot',$DataRoot,'-ObjectOutputPath',(OutPath 'ContentStorageEvidence'),'-TrendOutputPath',(OutPath 'ContentStorageTrendEvidence'))+$clockArgs)
Add-Task 'Mailbox' 'Mailbox' 'New-MailboxEvidence.ps1' @('-DataRoot',$last,'-UserEvidencePath',(OutPath 'UserInventoryEvidence'),'-OutputPath',(OutPath 'MailboxEvidence')) -DependsOn 'Workforce'
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
$trustArgs+=$clockDateArgs+$clockUtcArgs
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
    $priority = @('DeviceInventory','Workforce','Native','IntuneOperations','Applications','DeviceLifecycle','Estate','Trust','Mailbox','Security','WindowsMigration','EndpointProtection','Executive','Backup','Collaboration','Content','DailyTrends')
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDIeRR1jNAE2NC/
# C1R613LBeVKOv0RLItSfnQkV5c9EjqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBOqDK1KiNveJSfWJWC07lz
# gAvhsnZE/1aKCRsQfRc07DANBgkqhkiG9w0BAQEFAASCAYCEmvbDxzeBln7YSsVI
# SCg0qmwsI3K//UPSRLYKdrLP7tNYRJf35YCcqXwnrFIRrxbrq+Tu08b2WqQspUb/
# TCM2tkcuZAThb2Ggt3zhTHJKsIhykPfG1twnbUEYlTGM2npO0KUWvEnluGtjKNyr
# 0+BPONOWJZc4Km0BwEP67B/Pv+Nz0/Cq8bl+WM0r4GxMP4iZptGTKvCSmodkhdSq
# 4F6RF8om9IP5B03NDqsYwjhMDpmTZeN3+KHjr2UchjzNRYcfYWWyxz2UZ6I7zoEv
# a7KQWZ083zg6B/LOG9kJDQH/CdkI7eKyTeAveeXyp+gP7JUg3oSZcYqV3F9pjHxH
# cSIYUUr7E2SSU9G6C+dUOfBPrAQVeRBhCVV10Qlw8QfoAKCBmYKw1n5V9KZvT1TQ
# jKtWvjC1YDSEAcwcA4Z8mkSxka3D5ocQAjHO22xI+Q19yc2YOA1rH6t0Hl4YlqQV
# 5lftW97+dpb7QsuyPDAjC0+pb+jZ3F32qSZ7KHtADeFjRFE=
# SIG # End signature block
