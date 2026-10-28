# Version 1.4.0: optional M365 usage signals generator after Collaboration; 1.3.0: HybridIdentity generator after Workforce with its own classification checkpoint; 1.2.0: native-stage warning prefix.
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
$generators = @('New-EnterpriseEstateEvidence','New-DeviceInventoryEvidence','New-WorkforceIdentityEvidence','New-EndpointProtectionEvidence','New-WindowsMigrationEvidence','New-DeviceLifecycleExperienceTrendEvidence','New-ApplicationInventoryEvidence','New-CollaborationEvidence','New-M365UsageSignalsEvidence','New-ContentStorageEvidence','New-MailboxEvidence','New-HybridIdentityEvidence','New-DailyTrendsEvidence','New-SharePointMigrationEvidence','New-SecurityEvidence','New-BackupResilienceEvidence','New-IntuneOperationalEvidence','New-DataTrustEvidence','New-ExecutiveTrendEvidence')
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
$sequentialOrder = @('New-EnterpriseEstateEvidence','Native','New-DeviceInventoryEvidence','New-WorkforceIdentityEvidence','New-EndpointProtectionEvidence','New-WindowsMigrationEvidence','New-DeviceLifecycleExperienceTrendEvidence','New-ApplicationInventoryEvidence','New-CollaborationEvidence','New-M365UsageSignalsEvidence','New-ContentStorageEvidence','New-MailboxEvidence','New-HybridIdentityEvidence','New-DailyTrendsEvidence','New-SharePointMigrationEvidence','New-SecurityEvidence','New-BackupResilienceEvidence','New-IntuneOperationalEvidence','New-DataTrustEvidence','New-ExecutiveTrendEvidence')
$dependencies = @{
    'New-WorkforceIdentityEvidence'=@('New-DeviceInventoryEvidence')
    'New-EndpointProtectionEvidence'=@('New-DeviceInventoryEvidence')
    'New-WindowsMigrationEvidence'=@('New-DeviceInventoryEvidence')
    'New-MailboxEvidence'=@('New-WorkforceIdentityEvidence')
    'New-HybridIdentityEvidence'=@('New-WorkforceIdentityEvidence')
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
    Check ($sequential.Output -match 'stage=BuildStart;' -and $sequential.Output -match 'stage=BeforeWorkforce;' -and $sequential.Output -match 'stage=BeforeHybridIdentity;') 'Sequential classification checkpoints missing'
    Check ($sequential.Outputs['New-HybridIdentityEvidence'] -match [regex]::Escape('-UserEvidencePath|<case>') -and $sequential.Outputs['New-HybridIdentityEvidence'] -match 'HybridIdentityFindings\.csv' -and $sequential.Outputs['New-HybridIdentityEvidence'] -match '-AccountClassificationConfigPath') 'HybridIdentity arguments changed'
    Check ($sequential.Outputs['New-DeviceInventoryEvidence'] -match 'DeviceReadinessGapEvidence\.csv') 'DeviceInventory gap output argument missing'

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
    Check ($parallel.Output -match 'stage=BeforeWorkforce;' -and $parallel.Output -match 'stage=BeforeHybridIdentity;') 'Parallel classification checkpoint missing'
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
        Check (-not $tampered.Spans.ContainsKey('New-WorkforceIdentityEvidence') -and -not $tampered.Spans.ContainsKey('New-HybridIdentityEvidence')) "Workforce or HybridIdentity started after the rules changed (parallelism $parallelism)"
        Check ($tampered.Output -notmatch 'SYNTHETIC_DO_NOT_LOG|human-a|human-b|Population') 'Dispatcher logs leaked private rules'
    }

    # -Only runs one domain; dependencies outside it are treated as already staged.
    $only = Invoke-Case 'only' 3 @{} @('-Only','Mailbox')
    Check ($only.ExitCode -eq 0 -and $only.Spans.Count -eq 1 -and $only.Spans.ContainsKey('New-MailboxEvidence')) 'Only mode ran other generators'
    # Rules changed after Workforce: the hybrid findings never start on different rules.
    $hybridTampered = Invoke-Case 'hybrid-tampered' 1 @{Mutate='New-MailboxEvidence'}
    Check ($hybridTampered.ExitCode -ne 0 -and $hybridTampered.Output -match 'BeforeHybridIdentity' -and -not $hybridTampered.Spans.ContainsKey('New-HybridIdentityEvidence')) 'HybridIdentity started on changed rules'
    $hybridOnly = Invoke-Case 'only-hybrid' 3 @{} @('-Only','HybridIdentity','-HybridAllowedTargetDomains','partner.example,other.example')
    Check ($hybridOnly.ExitCode -eq 0 -and $hybridOnly.Spans.Count -eq 1 -and $hybridOnly.Output -match 'stage=BeforeHybridIdentity;' -and $hybridOnly.Outputs['New-HybridIdentityEvidence'] -match [regex]::Escape('-AllowedTargetDomains|partner.example,other.example')) 'HybridIdentity only mode or allowed domains changed'
    Write-Host "PASS: $script:checks build scheduler checks (sequential order, dependencies, concurrency, failures, classification checkpoint). Simulated generators only."
} finally {
    Remove-Item -LiteralPath $fixture -Recurse -Force -ErrorAction SilentlyContinue
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBjvuP/3yiYRVma
# dCCOaIbF98Mni5b+7UH6HmkkC86uF6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHVL0jYLfvBKdxP8/Kk+93BOIDajQWh0jq+g6q+PgyVPMA0GCSqG
# SIb3DQEBAQUABIIBgA1vbGFqX9gcQqgs3X58hniOx/8kQYY7YPZqOy0wEY4+rHwU
# 7CHw/b+ZmJ/zRL6r/unolanUPiVgZih2TmFDYsvU1fN0KT+A80usQWgG4TqW0vNW
# 2AsUUYyuRnH+p9KTS9VkANWG+DQnJAilTQstTQb1c5yiV/BOajxHGJyNjbjSSIDC
# kkOT/IvqfrZj/9CylA7ZLrLIvmwJEVleYANW9Sj+KMJaUBea896veTIfC2m85cJ0
# bmtwInNVJ9+NYwl+THGsl7lOGlmwMjUJkEOPnOuLJQXwQVIae3WQd501pH5jpQkm
# rfATQmCP5e+1CKs4fK5m86jGPMQrT4YeNwctiHIfT7Vrb+Yjw68HuWVCNFfb8I1J
# nXMOXLPnts4r8VUPsv5hF/Akf8YzH2xK0Kb5w72jz9UULfQ7ot/rxWK+5pV0t6mx
# 1eA68yDT1bEHkkNR/p6bR43SK1p5YC4aKpt+mEgwEm8mEKwVTYfHeYvJy3UouNof
# 6b+H8rFptGqNaQjpd6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMjU3
# NTZaMC8GCSqGSIb3DQEJBDEiBCAkQU3NOppfZpi/3kYW1LmZX0taNKid9QTEwgqz
# QqI21jANBgkqhkiG9w0BAQEFAASCAgBkXVNUjznDOBU8sj3HFZ9QqxR9T3hNy3dA
# 15/CJdT4KZJaoSB8qaw9e7xkxeJ2i2/D3Ic/8u9z/Aa0c9kX/I5bn2M1VCqjesuJ
# LyF0paNBcF5B14a0AYs6EcTTnw4JbrKOWB+Ax+ZZhb4BaWWBckMEMHbhYJksrkYS
# qJ3KIedfli0zxybu8zOGvtg6eb/Uh3itS+zz3LVarCnR0f/xFzFMCugrvEYwiEg6
# 0vSs7BiHdTj8Vh0r1a1T8tQV/N/rHdQlRKSuBYgVk3HLBdYpwV29ki7G9S7TQ2Bn
# TPpZEzb+KnWYYt9LOa6nGpkrFr5BdMloSxqwcLi8Ah6g33kEpxmgPcY8fN3Tyweo
# Qp1KScrdhlp+TyhcNhHuPW00w0lwkAqSCYU0J8iJwBHVysRaiJMUQx4iZu5SzJTC
# EHMtlssbkw9Zn1mE19QDAT2V/DaANTUyoZOjrDthPAHyiAM+xQ4l7PJfd8dKkS65
# t1XuybEGTzfAfHw03ggFijBiI9mtVUnHViNmi40TlAHlmIvagzmy3PsAbIljX1GA
# 8XVe6UyTqyRkoJQ1SWYta9JSZOhiPvYw18W/wUN6TEINFcYpH0BHsTsOaCUIQwNn
# YkSR0b8NRiBBc5/vFR68ALta40iTnMr1o0AMSPh/Ik1fF798dOlejPC8BdZHJsse
# 7gIJFPDtDw==
# SIG # End signature block
