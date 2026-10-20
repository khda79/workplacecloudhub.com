# Version 1.2.0: a native-stage warning keeps its WARNING prefix in the parallel worker.
[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot)
# Runs the real build dispatcher with simulated generators only; no production preparation.
$ErrorActionPreference = 'Stop'
$fixture = Join-Path ([IO.Path]::GetFullPath($TestRoot)) ('scheduler-' + [guid]::NewGuid().ToString('N'))
$script:checks = 0
function Check([bool]$Pass, [string]$Message) {
    if (-not $Pass) { throw $Message }; $script:checks++
}
$harness = Join-Path $fixture 'harness'
$scripts = Join-Path $harness 'SmartWorkplaceIntelligence/scripts'
$common = Join-Path $harness 'SmartM365/SmartInventory/Common'
New-Item -ItemType Directory -Path $scripts,$common -Force | Out-Null
foreach ($name in @('Invoke-PreparedEvidenceBuild.ps1','PreparedClassification.psm1')) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path $scripts $name)
}
foreach ($name in @('SmartM365.AccountClassification.psd1','SmartM365.AccountClassification.psm1')) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot "../../SmartM365/SmartInventory/Common/$name") -Destination (Join-Path $common $name)
}
# Every simulated generator records its start/end ticks, sleeps, writes a deterministic output
# from its own arguments and fails only when the control file names it.
$stub = @'
$name = [IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$root = $env:PREPARED_SCHEDULER_STUB_ROOT
$control = Get-Content -LiteralPath (Join-Path $root 'control.json.txt') -Raw | ConvertFrom-Json
[IO.File]::WriteAllText((Join-Path $root "$name.start"), [string][datetime]::UtcNow.Ticks)
if ($control.Mutate -eq $name) { [IO.File]::WriteAllText($control.RulesPath, $control.Changed) }
Start-Sleep -Milliseconds ([int]$control.SleepMilliseconds)
Write-Host "stub output $name"
[IO.File]::WriteAllText((Join-Path $root "$name.out"), ($args -join '|'))
[IO.File]::WriteAllText((Join-Path $root "$name.end"), [string][datetime]::UtcNow.Ticks)
if ($control.Fail -eq $name) { exit 3 }
'@
$generators = @('New-EnterpriseEstateEvidence','New-DeviceInventoryEvidence','New-WorkforceIdentityEvidence','New-EndpointProtectionEvidence','New-WindowsMigrationEvidence','New-DeviceLifecycleExperienceTrendEvidence','New-ApplicationInventoryEvidence','New-CollaborationEvidence','New-ContentStorageEvidence','New-MailboxEvidence','New-DailyTrendsEvidence','New-SharePointMigrationEvidence','New-SecurityEvidence','New-BackupResilienceEvidence','New-IntuneOperationalEvidence','New-DataTrustEvidence','New-ExecutiveTrendEvidence')
foreach ($name in $generators) { [IO.File]::WriteAllText((Join-Path $scripts ($name + '.ps1')), $stub) }
[IO.File]::WriteAllText((Join-Path $scripts 'New-MailboxMigrationExceptions.ps1'), $stub)
[IO.File]::WriteAllText((Join-Path $scripts 'NativeEvidence.psm1'),@'
function Invoke-NativeEvidenceStage {
    param($LastRoot,$TenantKey,$AuditPath,$OutputRoot,$PythonCommand)
    $root = $env:PREPARED_SCHEDULER_STUB_ROOT
    $control = Get-Content -LiteralPath (Join-Path $root 'control.json.txt') -Raw | ConvertFrom-Json
    [IO.File]::WriteAllText((Join-Path $root 'Native.start'), [string][datetime]::UtcNow.Ticks)
    Start-Sleep -Milliseconds ([int]$control.SleepMilliseconds)
    Write-Host 'stub output Native'
    if ($control.NativeWarning) { Write-Warning 'Native stage warning' }
    [IO.File]::WriteAllText((Join-Path $root 'Native.out'), "$TenantKey|$PythonCommand")
    [IO.File]::WriteAllText((Join-Path $root 'Native.end'), [string][datetime]::UtcNow.Ticks)
    if ($control.Fail -eq 'Native') { throw 'Native Workplace evidence validation failed. Nothing was published.' }
}
Export-ModuleMember -Function Invoke-NativeEvidenceStage
'@)
$valid = '{"SchemaVersion":"1.0","RuleVersion":"SYNTHETIC_DO_NOT_LOG","Population":{"Human":["human-a"],"NonHuman":[]}}'
$changed = $valid.Replace('human-a','human-b')
$shell = (Get-Process -Id $PID).Path
$sequentialOrder = @('New-EnterpriseEstateEvidence','Native','New-DeviceInventoryEvidence','New-WorkforceIdentityEvidence','New-EndpointProtectionEvidence','New-WindowsMigrationEvidence','New-DeviceLifecycleExperienceTrendEvidence','New-ApplicationInventoryEvidence','New-CollaborationEvidence','New-ContentStorageEvidence','New-MailboxEvidence','New-DailyTrendsEvidence','New-SharePointMigrationEvidence','New-SecurityEvidence','New-BackupResilienceEvidence','New-IntuneOperationalEvidence','New-DataTrustEvidence','New-ExecutiveTrendEvidence')
$dependencies = @{
    'New-WorkforceIdentityEvidence'=@('New-DeviceInventoryEvidence')
    'New-EndpointProtectionEvidence'=@('New-DeviceInventoryEvidence')
    'New-WindowsMigrationEvidence'=@('New-DeviceInventoryEvidence')
    'New-MailboxEvidence'=@('New-WorkforceIdentityEvidence')
    'New-SecurityEvidence'=@('New-WorkforceIdentityEvidence','New-DeviceInventoryEvidence')
    'New-DataTrustEvidence'=@('New-IntuneOperationalEvidence')
    'New-SharePointMigrationEvidence'=@('Native')
}
function Invoke-Case([string]$Name, [int]$Parallel, [hashtable]$Overrides, [string[]]$Extra = @()) {
    $case = Join-Path $fixture $Name
    $raw = Join-Path $case 'raw'; $staging = Join-Path $case 'staging'; $events = Join-Path $case 'events'
    New-Item -ItemType Directory -Path (Join-Path $raw 'DATA-LAST'),$staging,$events -Force | Out-Null
    $rules = Join-Path $case 'AccountClassification.json.txt'
    [IO.File]::WriteAllText($rules,$valid)
    $control = @{SleepMilliseconds=400;Fail='';Mutate='';RulesPath=$rules;Changed=$changed;NativeWarning=$false}
    foreach ($key in $Overrides.Keys) { $control[$key] = $Overrides[$key] }
    $control | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $events 'control.json.txt')
    $arguments = @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $scripts 'Invoke-PreparedEvidenceBuild.ps1'),'-DataRoot',$raw,'-OutputRoot',$staging,
        '-AccountClassificationConfigPath',$rules,'-AccountClassificationExpectedSHA256',(Get-FileHash $rules).Hash,'-AccountClassificationExpectedBytes',[string](Get-Item $rules).Length,
        '-NativeEvidenceAuditPath',(Join-Path $case 'unused-audit.json.txt'),'-TenantKey','synthetic-test','-EndpointProtectionPolicyPath',(Join-Path $case 'unused-policy.json.txt'),
        '-MaxParallelGenerators',[string]$Parallel) + $Extra
    $env:PREPARED_SCHEDULER_STUB_ROOT = $events
    try { $output = (& $shell @arguments 2>&1 | Out-String); $exitCode = $LASTEXITCODE }
    finally { Remove-Item Env:\PREPARED_SCHEDULER_STUB_ROOT -ErrorAction SilentlyContinue }
    $spans = @{}
    foreach ($file in Get-ChildItem -LiteralPath $events -Filter '*.start') {
        $endPath = [IO.Path]::ChangeExtension($file.FullName,'.end')
        $spans[$file.BaseName] = [pscustomobject]@{Start=[long](Get-Content -LiteralPath $file.FullName -Raw);End=$(if (Test-Path -LiteralPath $endPath) {[long](Get-Content -LiteralPath $endPath -Raw)} else {[long]::MaxValue})}
    }
    $outputs = @{}
    foreach ($file in Get-ChildItem -LiteralPath $events -Filter '*.out') { $outputs[$file.BaseName] = (Get-Content -LiteralPath $file.FullName -Raw).Replace($case,'<case>') }
    [pscustomobject]@{Output=$output;ExitCode=$exitCode;Spans=$spans;Outputs=$outputs}
}
function Get-MaxConcurrency($Spans) {
    $max = 0
    foreach ($span in $Spans.Values) {
        $count = @($Spans.Values | Where-Object { $_.Start -le $span.Start -and $_.End -gt $span.Start }).Count
        if ($count -gt $max) { $max = $count }
    }
    $max
}
try {
    $exceptions = Invoke-Case -Name 'mailbox-exceptions' -Parallel 3 -Overrides @{} -Extra @('-MailboxMigrationExceptionsEnabled')
    Check ($exceptions.ExitCode -eq 0) "Exception snapshot scheduling failed: $($exceptions.Output)"
    Check ($exceptions.Spans['New-MailboxMigrationExceptions'].Start -ge $exceptions.Spans['Native'].End) 'Exception matching started before native qualification completed'
    Check ($exceptions.Spans['New-DataTrustEvidence'].Start -ge $exceptions.Spans['New-MailboxMigrationExceptions'].End) 'Trust started before the exception audit existed'
    $exceptionFailed = Invoke-Case -Name 'mailbox-exceptions-failed' -Parallel 3 -Overrides @{Fail='New-MailboxMigrationExceptions'} -Extra @('-MailboxMigrationExceptionsEnabled')
    Check ($exceptionFailed.ExitCode -ne 0 -and -not $exceptionFailed.Spans.ContainsKey('New-DataTrustEvidence')) 'Trust started after exception matching failed'
    # Sequential: the former order, one generator at a time.
    $sequential = Invoke-Case 'sequential' 1 @{}
    Check ($sequential.ExitCode -eq 0) "Sequential build failed: $($sequential.Output)"
    $order = @($sequential.Spans.GetEnumerator() | Sort-Object { $_.Value.Start } | ForEach-Object Key)
    Check (($order -join ',') -ceq ($sequentialOrder -join ',')) "Sequential order changed: $($order -join ',')"
    Check ((Get-MaxConcurrency $sequential.Spans) -eq 1) 'Sequential mode overlapped generators'
    Check ($sequential.Output -match 'stage=BuildStart;' -and $sequential.Output -match 'stage=BeforeWorkforce;') 'Sequential classification checkpoints missing'

    # Parallel: same generators and outputs, dependencies respected, at most three at a time.
    $parallel = Invoke-Case 'parallel' 3 @{}
    Check ($parallel.ExitCode -eq 0) "Parallel build failed: $($parallel.Output)"
    Check ($parallel.Spans.Count -eq $sequentialOrder.Count) 'Parallel mode did not run every generator once'
    $concurrency = Get-MaxConcurrency $parallel.Spans
    Check ($concurrency -ge 2 -and $concurrency -le 3) "Unexpected parallel concurrency: $concurrency"
    foreach ($task in $dependencies.Keys) {
        foreach ($dependency in $dependencies[$task]) {
            Check ($parallel.Spans[$task].Start -ge $parallel.Spans[$dependency].End) "$task started before $dependency completed"
        }
    }
    foreach ($name in $sequentialOrder) {
        Check ($parallel.Outputs[$name] -ceq $sequential.Outputs[$name]) "Generator arguments differ between modes: $name"
        Check ($parallel.Output -match [regex]::Escape("stub output $name")) "Generator log not replayed: $name"
    }
    Check ($parallel.Output -match 'stage=BeforeWorkforce;') 'Parallel classification checkpoint missing'
    Check (@([regex]::Matches($parallel.Output,'Completed ')).Count -eq ($sequentialOrder.Count - 1)) 'Completion lines missing'

    # A native-stage warning keeps its WARNING prefix when the parallel worker replays it.
    $nativeWarning = Invoke-Case -Name 'native-warning' -Parallel 3 -Overrides @{NativeWarning=$true}
    Check ($nativeWarning.ExitCode -eq 0 -and $nativeWarning.Output -match '(?m)^WARNING: Native stage warning\s*$') 'Parallel native warning lost its WARNING prefix'

    # A failed generator stops the build; its dependants never start.
    $failed = Invoke-Case 'failed' 3 @{Fail='New-IntuneOperationalEvidence'}
    Check ($failed.ExitCode -ne 0 -and $failed.Output -match 'Preparation failed: IntuneOperations \(exit 3\)') 'Generator failure not reported'
    Check (-not $failed.Spans.ContainsKey('New-DataTrustEvidence')) 'Trust started after IntuneOperations failed'
    $nativeFailed = Invoke-Case 'native-failed' 3 @{Fail='Native'}
    Check ($nativeFailed.ExitCode -ne 0 -and $nativeFailed.Output -match 'Native Workplace evidence validation failed') 'Native failure not reported'

    # Rules changed while devices run: Workforce never starts, in both modes.
    foreach ($parallelism in 1,3) {
        $tampered = Invoke-Case "tampered-$parallelism" $parallelism @{Mutate='New-DeviceInventoryEvidence'}
        Check ($tampered.ExitCode -ne 0 -and $tampered.Output -match 'BeforeWorkforce') "Changed rules accepted with parallelism $parallelism"
        Check (-not $tampered.Spans.ContainsKey('New-WorkforceIdentityEvidence')) "Workforce started after the rules changed (parallelism $parallelism)"
        Check ($tampered.Output -notmatch 'SYNTHETIC_DO_NOT_LOG|human-a|human-b|Population') 'Dispatcher logs leaked private rules'
    }

    # -Only runs one domain; dependencies outside it are treated as already staged.
    $only = Invoke-Case 'only' 3 @{} @('-Only','Mailbox')
    Check ($only.ExitCode -eq 0 -and $only.Spans.Count -eq 1 -and $only.Spans.ContainsKey('New-MailboxEvidence')) 'Only mode ran other generators'
    Write-Host "PASS: $script:checks build scheduler checks (sequential order, dependencies, concurrency, failures, classification checkpoint). Simulated generators only."
} finally {
    Remove-Item -LiteralPath $fixture -Recurse -Force -ErrorAction SilentlyContinue
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA2gUmTfhc+4f1L
# gDAJKam6pzMx5zE7267lGsPGafBn6KCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDg6RhE0+sd1DcXDjutD8fB
# NOo3qdkqavWSroQOul/CSDANBgkqhkiG9w0BAQEFAASCAYA3ANqE4kdo8bw369n3
# 2IKGA5eixF/2I2D15Jg8D0BeFHEsapbwQmySGenZ8sUkMLvSMSkkTPTcoEtvTJds
# jZOpk0wyCZkx5z/yDbOIKuZBQ3pfYRv9s88Zbc4cRBvxf9NjVZLyu5FZUZ/kOGYO
# yBv794wXjFFAIvoQbbrh66VE+aaUYJ8mODg9uX0zZqJTaArjkph219Rx2H4kzi9k
# ZM6S53qhOqkZKp8fqxZqD5wLECAJqCSPKsQdjMEZb6m9EKleSsV38hCGIG+Tia0o
# ZAug7PT04eo2qe6O10vf3+32Y/lbkyQ+T7uCHkFCXFyEa4b8jVqCEqItnHZbjIaO
# OUXc4vIYigwQYjysQedvtTXxWRnrIXMyXejykOfyytNqk6d14BkU8dB0wQR/aE+w
# n9Rpkt8dPL33t+jCuFaiE+iuWVaTWOaLqJImKNKqT33398fR1MAH7iizPxWK82h9
# +OtQj/+RlGTlD+XQeZczA0SEC+t/RafdB9R0R2al9HdxeA0=
# SIG # End signature block
