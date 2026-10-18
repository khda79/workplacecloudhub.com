# Version 1.0.0
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
$generators = @('New-EnterpriseEstateEvidence','New-DeviceInventoryEvidence','New-WorkforceIdentityEvidence','New-EndpointProtectionEvidence','New-WindowsMigrationEvidence','New-DeviceLifecycleExperienceTrendEvidence','New-ApplicationInventoryEvidence','New-CollaborationEvidence','New-ContentStorageEvidence','New-MailboxEvidence','New-DailyTrendsEvidence','New-SecurityEvidence','New-BackupResilienceEvidence','New-IntuneOperationalEvidence','New-DataTrustEvidence','New-ExecutiveTrendEvidence')
foreach ($name in $generators) { [IO.File]::WriteAllText((Join-Path $scripts ($name + '.ps1')), $stub) }
[IO.File]::WriteAllText((Join-Path $scripts 'NativeEvidence.psm1'),@'
function Invoke-NativeEvidenceStage {
    param($LastRoot,$TenantKey,$AuditPath,$OutputRoot,$PythonCommand)
    $root = $env:PREPARED_SCHEDULER_STUB_ROOT
    $control = Get-Content -LiteralPath (Join-Path $root 'control.json.txt') -Raw | ConvertFrom-Json
    [IO.File]::WriteAllText((Join-Path $root 'Native.start'), [string][datetime]::UtcNow.Ticks)
    Start-Sleep -Milliseconds ([int]$control.SleepMilliseconds)
    Write-Host 'stub output Native'
    [IO.File]::WriteAllText((Join-Path $root 'Native.out'), "$TenantKey|$PythonCommand")
    [IO.File]::WriteAllText((Join-Path $root 'Native.end'), [string][datetime]::UtcNow.Ticks)
    if ($control.Fail -eq 'Native') { throw 'Native Workplace evidence validation failed. Nothing was published.' }
}
Export-ModuleMember -Function Invoke-NativeEvidenceStage
'@)
$valid = '{"SchemaVersion":"1.0","RuleVersion":"SYNTHETIC_DO_NOT_LOG","Population":{"Human":["human-a"],"NonHuman":[]}}'
$changed = $valid.Replace('human-a','human-b')
$shell = (Get-Process -Id $PID).Path
$sequentialOrder = @('New-EnterpriseEstateEvidence','Native','New-DeviceInventoryEvidence','New-WorkforceIdentityEvidence','New-EndpointProtectionEvidence','New-WindowsMigrationEvidence','New-DeviceLifecycleExperienceTrendEvidence','New-ApplicationInventoryEvidence','New-CollaborationEvidence','New-ContentStorageEvidence','New-MailboxEvidence','New-DailyTrendsEvidence','New-SecurityEvidence','New-BackupResilienceEvidence','New-IntuneOperationalEvidence','New-DataTrustEvidence','New-ExecutiveTrendEvidence')
$dependencies = @{
    'New-WorkforceIdentityEvidence'=@('New-DeviceInventoryEvidence')
    'New-EndpointProtectionEvidence'=@('New-DeviceInventoryEvidence')
    'New-WindowsMigrationEvidence'=@('New-DeviceInventoryEvidence')
    'New-MailboxEvidence'=@('New-WorkforceIdentityEvidence')
    'New-SecurityEvidence'=@('New-WorkforceIdentityEvidence','New-DeviceInventoryEvidence')
    'New-DataTrustEvidence'=@('New-IntuneOperationalEvidence')
}
function Invoke-Case([string]$Name, [int]$Parallel, [hashtable]$Overrides, [string[]]$Extra = @()) {
    $case = Join-Path $fixture $Name
    $raw = Join-Path $case 'raw'; $staging = Join-Path $case 'staging'; $events = Join-Path $case 'events'
    New-Item -ItemType Directory -Path (Join-Path $raw 'DATA-LAST'),$staging,$events -Force | Out-Null
    $rules = Join-Path $case 'AccountClassification.json.txt'
    [IO.File]::WriteAllText($rules,$valid)
    $control = @{SleepMilliseconds=400;Fail='';Mutate='';RulesPath=$rules;Changed=$changed}
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
