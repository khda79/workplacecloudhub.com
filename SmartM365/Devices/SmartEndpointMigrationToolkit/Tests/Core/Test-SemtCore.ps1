<#
.SYNOPSIS
    Offline synthetic tests for the Smart Endpoint Migration Toolkit engine and shared target guards.
.DESCRIPTION
    Covers option precedence and inherit markers, endpoint argument emission, configuration
    templates, computer lists, Wave creation and launchers, the shared worker cap (strictest
    limit, stale reservation recovery, cross-process wait), the Wave lock, run guard history
    isolation per Campaign, worker contract versioning, correlation and foreign evidence,
    shared inventory isolation, long/space/Unicode paths, the target guard region (identical in
    both endpoint scripts), read-only legacy detection and legacy setup media reuse with
    synthetic folders. No network, PsExec, endpoint, scheduled task, registry write or setup.exe.
    The technician state is redirected with SEMT_RUNTIME_ROOT_OVERRIDE to a temp folder.
    Prints one PASS/FAIL line per assertion and "Summary: Passed=N; Failed=M". Exits 1 on failure.
    Runs on Windows PowerShell 5.1 and PowerShell 7.
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:Passed = 0
$script:Failed = 0
$script:PackageRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent

function Write-TestResult {
    param([bool]$Condition, [string]$Message, [string]$Detail = '')
    if ($Condition) { $script:Passed++; Write-Host ("PASS: {0}" -f $Message) }
    else {
        $script:Failed++
        if ([string]::IsNullOrWhiteSpace($Detail)) { Write-Host ("FAIL: {0}" -f $Message) } else { Write-Host ("FAIL: {0} ({1})" -f $Message,$Detail) }
    }
}

function Test-Step {
    param([string]$Message, [scriptblock]$Check)
    try { Write-TestResult -Condition ([bool](& $Check)) -Message $Message }
    catch { Write-TestResult -Condition $false -Message $Message -Detail $_.Exception.Message }
}

function Test-Throws {
    param([string]$Message, [scriptblock]$Action, [string]$Pattern = '.')
    try { & $Action | Out-Null; Write-TestResult -Condition $false -Message $Message -Detail 'no exception' }
    catch { Write-TestResult -Condition ($_.Exception.Message -match $Pattern) -Message $Message -Detail $_.Exception.Message }
}

function Get-FunctionText {
    param([string]$Path, [string[]]$Names)
    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    foreach ($name in $Names) {
        $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
        if (-not $definition) { throw "Function not found: $name in $Path" }
        $definition.Extent.Text
    }
}

function Get-RegionText {
    param([string]$Path, [string]$Region)
    $text = [IO.File]::ReadAllText($Path)
    $pattern = '(?s)#region ' + [regex]::Escape($Region) + '.*?#endregion ' + [regex]::Escape($Region)
    $match = [regex]::Match($text, $pattern)
    if (-not $match.Success) { throw "Region '$Region' not found in $Path" }
    return ($match.Value -replace "`r`n", "`n")
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SEMT-CoreTest-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$previousOverride = $env:SEMT_RUNTIME_ROOT_OVERRIDE
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
$env:SEMT_RUNTIME_ROOT_OVERRIDE = Join-Path $tempRoot 'Runtime'

try {
    Import-Module (Join-Path $script:PackageRoot 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
    $module = Get-Module SmartEndpointMigration.Core
    $catalog = Import-SemtOptionCatalog -PackageRoot $script:PackageRoot

    #region Runtime and catalog
    Test-Step 'runtime root honours SEMT_RUNTIME_ROOT_OVERRIDE (test isolation)' { (Get-SemtRuntimeRoot) -eq $env:SEMT_RUNTIME_ROOT_OVERRIDE }
    Test-Step 'option catalog loads with unique names' { $catalog.Options.Count -gt 50 -and $catalog.ByName.Count -eq $catalog.Options.Count }
    Test-Step 'every catalog option has Type, Category, Label, Description and SEMT_* EnvVar' {
        @($catalog.Options | Where-Object { -not $_.Type -or -not $_.Category -or -not $_.Label -or -not $_.Description -or [string]$_.EnvVar -notmatch '^SEMT_' }).Count -eq 0
    }
    Test-Step 'campaign env vars use SEMT_HYBRIDJOIN_* / SEMT_WINDOWS11_* prefixes' {
        @($catalog.Options | Where-Object { ($_.Scope -eq 'HybridJoin' -and $_.EnvVar -notmatch '^SEMT_HYBRIDJOIN_') -or ($_.Scope -eq 'Windows11Upgrade' -and $_.EnvVar -notmatch '^SEMT_WINDOWS11_') }).Count -eq 0
    }
    Test-Step 'no Batch/LOT wording in option labels or descriptions' { @($catalog.Options | Where-Object { ('{0} {1}' -f $_.Label,$_.Description) -match '\b(LOT|LOTs|Batch|batch)\b' }).Count -eq 0 }
    Test-Step 'D11: Hybrid Join AllowRebootAfterDsregLeave defaults to false' { (Get-SemtOptionDefault -Definition $catalog.ByName['HybridJoin.AllowRebootAfterDsregLeave'] -Campaign HybridJoin) -eq $false }
    #endregion

    #region Option precedence
    $resolved = Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign HybridJoin
    $default = $resolved['ThrottleLimit'].Value
    Test-Step 'default layer used when no layer sets the option' { $resolved['ThrottleLimit'].Source -eq 'Default' }
    $resolved = Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign HybridJoin -GlobalLayer @{ ThrottleLimit = 3 } -WaveLayer @{ ThrottleLimit = 4 }
    Test-Step 'Wave overrides global' { $resolved['ThrottleLimit'].Value -eq 4 -and $resolved['ThrottleLimit'].Source -eq 'Wave' }
    $resolved = Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign HybridJoin -GlobalLayer @{ ThrottleLimit = 3 } -WaveLayer @{ ThrottleLimit = 4 } -EnvironmentLayer @{ ThrottleLimit = 5 } -GuiLayer @{ ThrottleLimit = 6 }
    Test-Step 'GUI captured choices override Wave, environment and global' { $resolved['ThrottleLimit'].Value -eq 6 -and $resolved['ThrottleLimit'].Source -eq 'Gui' }
    $resolved = Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign HybridJoin -GlobalLayer @{ ThrottleLimit = 3 } -WaveLayer @{ ThrottleLimit = 4 } -GuiLayer @{ ThrottleLimit = 6 } -CliLayer @{ ThrottleLimit = 7 }
    Test-Step 'explicit CLI overrides everything' { $resolved['ThrottleLimit'].Value -eq 7 -and $resolved['ThrottleLimit'].Source -eq 'Cli' }
    foreach ($marker in @('', '__INHERIT__', '__USE_GLOBAL__', 'USE_GLOBAL', '__USE_DEFAULT__')) {
        $resolved = Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign HybridJoin -GlobalLayer @{ ThrottleLimit = 3 } -WaveLayer @{ ThrottleLimit = $marker }
        Test-Step ("Wave inherit marker '{0}' keeps the global value" -f $marker) { $resolved['ThrottleLimit'].Value -eq 3 -and $resolved['ThrottleLimit'].Source -eq 'Global' }
    }
    Test-Throws 'out-of-range value is rejected with its source' { Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign HybridJoin -WaveLayer @{ ThrottleLimit = 100000 } } 'Wave|range|between|maximum'
    Test-Throws 'unknown option is rejected' { Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign HybridJoin -WaveLayer @{ NotAnOption = 1 } } 'SEMT_OPTION_UNKNOWN'
    Test-Step 'option of the other Campaign is ignored, not applied' {
        $r = Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign HybridJoin -GlobalLayer @{ 'Windows11Upgrade.AllowReboot' = $true }
        -not $r.Contains('Windows11Upgrade.AllowReboot')
    }
    Test-Step 'dry run with unlimited cycles is limited to one cycle by rule' {
        $r = Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign HybridJoin -CliLayer @{ DryRun = $true; MaxCycles = 0 }
        $r['MaxCycles'].Value -eq 1 -and $r['MaxCycles'].Source -eq 'Rule:DryRunSingleCycle'
    }
    Test-Step 'Windows 11 DirectSetup never re-enables AllowReboot when it is off' {
        $r = Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign Windows11Upgrade -GuiLayer @{ 'Windows11Upgrade.DirectSetupUpgrade' = $true; 'Windows11Upgrade.AllowReboot' = $false }
        $emitted = @(ConvertTo-SemtEndpointArguments -Catalog $catalog -Campaign Windows11Upgrade -Options $r)
        $r['Windows11Upgrade.AllowReboot'].Value -eq $false -and $emitted -notcontains '-AllowReboot' -and $emitted -contains '-DirectSetupUpgrade'
    }
    Test-Step 'environment layer reads SEMT_* variables only for the Campaign' {
        $env:SEMT_THROTTLE_LIMIT_TEST_GUARD = '1'
        $name = [string]$catalog.ByName['ThrottleLimit'].EnvVar
        $previous = [Environment]::GetEnvironmentVariable($name)
        try {
            [Environment]::SetEnvironmentVariable($name, '9')
            $layer = Get-SemtEnvironmentOptionLayer -Catalog $catalog -Campaign HybridJoin
            [int]$layer['ThrottleLimit'] -eq 9
        }
        finally { [Environment]::SetEnvironmentVariable($name, $previous); Remove-Item Env:\SEMT_THROTTLE_LIMIT_TEST_GUARD -ErrorAction SilentlyContinue }
    }
    $resolved = Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign Windows11Upgrade -CliLayer @{ AuditOnly = $true }
    $arguments = @(ConvertTo-SemtEndpointArguments -Catalog $catalog -Campaign Windows11Upgrade -Options $resolved)
    Test-Step 'endpoint arguments: switch emitted when true' { $arguments -contains '-AuditOnly' }
    $w11Endpoint = Join-Path $script:PackageRoot 'Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'
    $hjEndpoint = Join-Path $script:PackageRoot 'Campaigns\HybridJoin\Scripts\SmartM365-Invoke-EndpointMigration-HybridJoin.ps1'
    foreach ($pair in @(@('HybridJoin',$hjEndpoint),@('Windows11Upgrade',$w11Endpoint))) {
        $tokens = $null; $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($pair[1], [ref]$tokens, [ref]$errors)
        $parameters = @($ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
        $mapped = @(Get-SemtCampaignOptionDefinitions -Catalog $catalog -Campaign $pair[0] | ForEach-Object { [string](Get-SemtPropertyValue -InputObject $_ -Name 'EndpointParameter' -Default '') } | Where-Object { $_ })
        $missing = @($mapped | Where-Object { $parameters -notcontains $_ })
        Write-TestResult -Condition ($missing.Count -eq 0) -Message ("{0}: every EndpointParameter exists on the endpoint script" -f $pair[0]) -Detail ($missing -join ', ')
    }
    #endregion

    #region Templates and configuration
    Test-Step 'configuration templates are in sync with the catalog (-Check)' {
        $hostPath = (Get-Process -Id $PID).Path
        $null = & $hostPath -NoProfile -ExecutionPolicy Bypass -File (Join-Path $script:PackageRoot 'Release\SmartM365-EndpointMigration-Update-ConfigTemplates.ps1') -Check 2>&1
        $LASTEXITCODE -eq 0
    }
    Test-Step 'self-healing adds missing keys and keeps existing values' {
        $config = Join-Path $tempRoot 'heal.local.json'
        $template = Join-Path $tempRoot 'heal.local.json.template'
        Set-Content -LiteralPath $template -Value '{ "SchemaVersion": 1, "Options": { "ThrottleLimit": null, "MaxCycles": null } }' -Encoding UTF8
        Set-Content -LiteralPath $config -Value '{ "SchemaVersion": 1, "Options": { "ThrottleLimit": 4 } }' -Encoding UTF8
        [void](Sync-SemtJsonWithTemplate -Path $config -TemplatePath $template)
        $data = Read-SemtJsonFile -Path $config
        [int]$data.Options.ThrottleLimit -eq 4 -and $data.Options.PSObject.Properties['MaxCycles']
    }
    #endregion

    #region Computer lists and Waves
    $packageCopy = Join-Path $tempRoot 'Pkg'
    New-Item -ItemType Directory -Path (Join-Path $packageCopy 'Waves\HybridJoin') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $packageCopy 'Waves\Windows11Upgrade') -Force | Out-Null
    foreach ($campaignName in @('HybridJoin','Windows11Upgrade')) {
        $templateTarget = Join-Path $packageCopy ("Campaigns\{0}\Templates" -f $campaignName)
        New-Item -ItemType Directory -Path $templateTarget -Force | Out-Null
        Copy-Item -Path (Join-Path $script:PackageRoot ("Campaigns\{0}\Templates\*" -f $campaignName)) -Destination $templateTarget -Recurse -Force
    }
    $wave = New-SemtWave -PackageRoot $packageCopy -Campaign HybridJoin -Name 'WAVE-TEST-01' -Computers @('PC-0001','pc-0001.contoso.example','PC-0002','# comment','') -Source Manual
    Test-Step 'New-SemtWave creates Computers.txt, Wave.local.json and the four launchers' {
        (Test-Path (Join-Path $wave.Path 'Computers.txt')) -and (Test-Path (Join-Path $wave.Path 'Wave.local.json')) -and @(Get-ChildItem -LiteralPath $wave.Path -Filter 'Start-Wave-*.cmd').Count -eq 4
    }
    Test-Step 'Wave computer list skips comments and blank lines and removes case-insensitive duplicates' {
        $listed = @(Get-SemtComputerList -Path $wave.ComputersPath)
        $listed -notcontains '# comment' -and @($listed | Where-Object { $_ -ieq 'PC-0001' }).Count -eq 1 -and $listed -contains 'PC-0002'
    }
    Test-Step 'short-name key matches FQDN and short name (used for target and run guard correlation)' { $fqdn = 'pc-0001.contoso.example'; $short = 'PC-0001'; (Get-SemtComputerKey -ComputerName $fqdn) -eq (Get-SemtComputerKey -ComputerName $short) }
    Test-Throws 'existing Wave is never overwritten' { New-SemtWave -PackageRoot $packageCopy -Campaign HybridJoin -Name 'WAVE-TEST-01' -Computers @('PC-9') -Source Manual } 'exist'
    Test-Throws 'invalid Wave name is rejected' { New-SemtWave -PackageRoot $packageCopy -Campaign HybridJoin -Name 'bad name&' -Computers @('PC-9') -Source Manual } '.'
    Test-Step 'Waves are isolated per Campaign' { @(Get-SemtWaves -PackageRoot $packageCopy -Campaign Windows11Upgrade).Count -eq 0 -and @(Get-SemtWaves -PackageRoot $packageCopy -Campaign HybridJoin).Count -eq 1 }
    $wrapper = [IO.File]::ReadAllText((Join-Path $wave.Path 'Start-Wave-Once-IgnoreRunGuard.cmd'))
    Test-Step 'Wave launcher calls the shared Wave runner with Once and IgnoreRunGuard' { $wrapper -match 'Run-EndpointMigrationWave\.cmd" "%~dp0\." Once IgnoreRunGuard' }
    Test-Step 'Wave launchers are ASCII with CRLF line endings' {
        $bytes = [IO.File]::ReadAllBytes((Join-Path $wave.Path 'Start-Wave-Loop.cmd'))
        @($bytes | Where-Object { $_ -gt 126 }).Count -eq 0 -and ([Text.Encoding]::ASCII.GetString($bytes) -match "`r`n")
    }
    Test-Step 'no Batch/LOT wording in Wave launchers' { $wrapper -notmatch '\b(LOT|Batch)\b' }
    Test-Step 'Move-SemtComputersFromList archives completed computers' {
        $moved = Move-SemtComputersFromList -ComputerListPath $wave.ComputersPath -ComputerNames @('PC-0001') -ArchiveFileName 'Computers.Completed.txt' -Label 'Completed'
        $left = @(Get-SemtComputerList -Path $wave.ComputersPath)
        $left.Count -eq 1 -and $left[0] -eq 'PC-0002' -and (Test-Path (Join-Path $wave.Path 'Computers.Completed.txt'))
    }
    #endregion

    #region Shared worker cap
    $gateRoot = Join-Path $tempRoot 'Gate'
    $gate = Initialize-SemtWorkerGate -Limit 2 -LeaseTimeoutMinutes 120 -HeartbeatStaleMinutes 20 -Root $gateRoot
    $lease1 = Request-SemtWorkerLease -Gate $gate -Campaign HybridJoin -RunId R1 -WaveName W1 -Computer PC-1 -CycleNumber 1
    $lease2 = Request-SemtWorkerLease -Gate $gate -Campaign Windows11Upgrade -RunId R2 -WaveName W2 -Computer PC-2 -CycleNumber 1
    Test-Step 'one shared cap counts both Campaigns' { (Get-SemtWorkerGateSnapshot -Gate $gate).Active.Count -eq 2 -and $lease1 -and $lease2 }
    $stale = Join-Path $gate.LeaseRoot 'lease_999999_stale.json'
    Write-SemtJsonFileAtomic -Path $stale -InputObject ([pscustomobject]@{ SchemaVersion = 1; Campaign = 'HybridJoin'; Computer = 'PC-STALE'; Limit = 2; LauncherProcessId = 999999; LauncherProcessStartUtc = '2020-01-01T00:00:00.0000000Z'; CreatedUtc = (Get-SemtUtcNowText); HeartbeatUtc = (Get-SemtUtcNowText) })
    $snapshot = & $module { param($g) Get-SemtWorkerGateSnapshotUnlocked -Gate $g -RemoveStale } $gate
    Test-Step 'reservation of a dead launcher is recovered as stale' { @($snapshot.Removed | Where-Object { $_.Reason -eq 'LauncherProcessExited' }).Count -eq 1 -and -not (Test-Path -LiteralPath $stale) }
    $oldLease = Join-Path $gate.LeaseRoot 'lease_heartbeat_old.json'
    $selfStart = & $module { Get-SemtProcessStartUtc -ProcessId $PID }
    Write-SemtJsonFileAtomic -Path $oldLease -InputObject ([pscustomobject]@{ SchemaVersion = 1; Campaign = 'HybridJoin'; Computer = 'PC-OLD'; Limit = 2; LauncherProcessId = $PID; LauncherProcessStartUtc = $selfStart; CreatedUtc = (Get-Date).ToUniversalTime().AddMinutes(-30).ToString('o'); HeartbeatUtc = (Get-Date).ToUniversalTime().AddMinutes(-25).ToString('o') })
    $snapshot = & $module { param($g) Get-SemtWorkerGateSnapshotUnlocked -Gate $g -RemoveStale } $gate
    Test-Step 'reservation without heartbeat is recovered as stale' { @($snapshot.Removed | Where-Object { $_.Reason -eq 'HeartbeatStale' }).Count -eq 1 }
    Remove-SemtWorkerLease -LeasePath $lease2 | Out-Null
    $strict = Join-Path $gate.LeaseRoot 'lease_strict.json'
    Write-SemtJsonFileAtomic -Path $strict -InputObject ([pscustomobject]@{ SchemaVersion = 1; Campaign = 'Windows11Upgrade'; Computer = 'PC-STRICT'; Limit = 1; LauncherProcessId = $PID; LauncherProcessStartUtc = $selfStart; CreatedUtc = (Get-SemtUtcNowText); HeartbeatUtc = (Get-SemtUtcNowText) })
    Test-Step 'strictest active limit wins when launchers use different limits' { (Get-SemtWorkerGateSnapshot -Gate $gate).StrictestActiveLimit -eq 1 }
    Remove-Item -LiteralPath $strict -Force
    Test-Step 'cap full: a second process waits, then starts when a slot is released' {
        $fill = Request-SemtWorkerLease -Gate $gate -Campaign Windows11Upgrade -RunId R3 -WaveName W3 -Computer PC-3 -CycleNumber 1
        $job = Start-Job -ScriptBlock {
            param($ModulePath, $Root, $Override)
            $env:SEMT_RUNTIME_ROOT_OVERRIDE = $Override
            Import-Module $ModulePath -MinimumVersion 1.0.0 -Force
            $g = Initialize-SemtWorkerGate -Limit 2 -LeaseTimeoutMinutes 120 -HeartbeatStaleMinutes 20 -Root $Root
            Request-SemtWorkerLease -Gate $g -Campaign HybridJoin -RunId R4 -WaveName W4 -Computer PC-4 -CycleNumber 1 -PollSeconds 1
        } -ArgumentList (Join-Path $script:PackageRoot 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1'), $gateRoot, $env:SEMT_RUNTIME_ROOT_OVERRIDE
        try {
            $waited = $false
            for ($i = 0; $i -lt 20; $i++) { Start-Sleep -Seconds 1; if ($job.State -ne 'Running') { break } }
            $waited = ($job.State -eq 'Running')
            Remove-SemtWorkerLease -LeasePath $fill | Out-Null
            [void](Wait-Job -Job $job -Timeout 60)
            $leasePath = [string](Receive-Job -Job $job | Select-Object -Last 1)
            $waited -and $job.State -eq 'Completed' -and -not [string]::IsNullOrWhiteSpace($leasePath)
        }
        finally { Stop-Job -Job $job -ErrorAction SilentlyContinue; Remove-Job -Job $job -Force -ErrorAction SilentlyContinue }
    }
    #endregion

    #region Wave lock
    $lock = Enter-SemtWaveLock -Campaign HybridJoin -WavePath $wave.Path
    Test-Step 'a second Run of the same Wave is refused while the lock is held' {
        $job = Start-Job -ScriptBlock { param($ModulePath, $Override, $WavePath) $env:SEMT_RUNTIME_ROOT_OVERRIDE = $Override; Import-Module $ModulePath -MinimumVersion 1.0.0 -Force; try { Enter-SemtWaveLock -Campaign HybridJoin -WavePath $WavePath | Out-Null; 'ACQUIRED' } catch { 'REFUSED' } } -ArgumentList (Join-Path $script:PackageRoot 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1'), $env:SEMT_RUNTIME_ROOT_OVERRIDE, $wave.Path
        [void](Wait-Job $job -Timeout 60); $result = Receive-Job $job; Remove-Job $job -Force
        $result -contains 'REFUSED'
    }
    Exit-SemtWaveLock -WaveLock $lock
    Test-Step 'the Wave lock is released by Exit-SemtWaveLock' { $again = Enter-SemtWaveLock -Campaign HybridJoin -WavePath $wave.Path; Exit-SemtWaveLock -WaveLock $again; $true }
    #endregion

    #region Run guard history
    Test-Step 'run guard history is isolated per Campaign' {
        $hj = Get-SemtRunGuardHistoryPath -Campaign HybridJoin
        $w11 = Get-SemtRunGuardHistoryPath -Campaign Windows11Upgrade
        $hj -ne $w11 -and $hj -like '*HybridJoin*' -and $w11 -like '*Windows11Upgrade*'
    }
    #endregion

    #region Worker contract and correlation
    $contractPath = Join-Path $tempRoot 'contract.json'
    Write-SemtJsonFileAtomic -Path $contractPath -InputObject ([pscustomobject]@{ ContractVersion = 99; Campaign = 'HybridJoin'; EngineVersion = '1.0.0' })
    Test-Throws 'worker refuses a contract of another version' { Read-SemtWorkerContract -Path $contractPath } 'SEMT_WORKER_CONTRACT_VERSION_MISMATCH'
    $fullContract = [ordered]@{ ContractVersion = 1; ContractId = 'c1'; Campaign = 'HybridJoin'; RunId = 'R1'; CorrelationId = 'K1'; Computer = 'PC-1'; ConnectionTarget = 'PC-1'; CycleNumber = 1; PackageRoot = $script:PackageRoot; Endpoint = @{ a = 1 }; PsExec = @{ a = 1 }; Evidence = @{ a = 1 }; Flags = @{ a = 1 }; EngineVersion = '0.9.0' }
    Write-SemtJsonFileAtomic -Path $contractPath -InputObject ([pscustomobject]$fullContract)
    Test-Throws 'worker refuses a contract from another engine version' { Read-SemtWorkerContract -Path $contractPath } 'SEMT_WORKER_ENGINE_MISMATCH'
    $fullContract.Remove('CorrelationId'); $fullContract.EngineVersion = '1.0.0'
    Write-SemtJsonFileAtomic -Path $contractPath -InputObject ([pscustomobject]$fullContract)
    Test-Throws 'worker refuses a contract without CorrelationId' { Read-SemtWorkerContract -Path $contractPath } "SEMT_WORKER_CONTRACT_INVALID: missing 'CorrelationId'"
    $final = ConvertFrom-SemtFinalStatusLine -Lines @('noise','FINAL_STATUS=ENROLLED; EXIT_CODE=0; NEXT_ACTION=NONE; CORRELATION_ID=abc-1; DETAIL=Done; with semicolon')
    Test-Step 'final status line is parsed with its CorrelationId' { $final.Status -eq 'ENROLLED' -and $final.CorrelationId -eq 'abc-1' -and $final.Detail -like 'Done*' }
    Test-Step 'final status line without CorrelationId (legacy format) still parses' { (ConvertFrom-SemtFinalStatusLine -Lines @('FINAL_STATUS=X; EXIT_CODE=1; NEXT_ACTION=Y; DETAIL=Z')).CorrelationId -eq '' }
    $share = Join-Path $tempRoot 'TargetShare'
    New-Item -ItemType Directory -Path (Join-Path $share 'State\Results') -Force | Out-Null
    Write-SemtJsonFileAtomic -Path (Join-Path $share 'State\Results\attempt-1.json') -InputObject ([pscustomobject]@{ CorrelationId = 'attempt-OTHER'; State = 'Completed' })
    Test-Step 'foreign attempt result (other CorrelationId) is never trusted' { $r = Read-SemtRemoteAttemptResult -RemoteDataShare $share -CorrelationId 'attempt-1'; $r.State -eq 'Foreign' -and -not $r.Completed }
    Write-SemtJsonFileAtomic -Path (Join-Path $share 'State\Results\attempt-2.json') -InputObject ([pscustomobject]@{ CorrelationId = 'attempt-2'; State = 'Running' })
    Test-Step 'running attempt result is reported as not completed' { $r = Read-SemtRemoteAttemptResult -RemoteDataShare $share -CorrelationId 'attempt-2'; $r.State -eq 'Running' -and -not $r.Completed }
    Write-SemtJsonFileAtomic -Path (Join-Path $share 'State\Results\attempt-3.json') -InputObject ([pscustomobject]@{ CorrelationId = 'attempt-3'; State = 'Completed'; FinalStatus = 'ENROLLED' })
    Test-Step 'matching completed attempt result is accepted' { (Read-SemtRemoteAttemptResult -RemoteDataShare $share -CorrelationId 'attempt-3').Completed }
    Test-Step 'missing attempt result is reported, not invented' { -not (Read-SemtRemoteAttemptResult -RemoteDataShare $share -CorrelationId 'attempt-4').Exists }
    #endregion

    #region Inventory isolation
    $inventoryPackage = Join-Path $tempRoot 'InvPkg'
    New-Item -ItemType Directory -Path (Join-Path $inventoryPackage 'Inventory') -Force | Out-Null
    $sharedAd = Get-SemtSharedInventoryPath -PackageRoot $inventoryPackage -Kind AD
    Test-Step 'shared inventories live under <Package>\Inventory' { $sharedAd -eq (Join-Path $inventoryPackage 'Inventory\DevicesAD.csv') }
    Test-Step 'missing shared inventory is reported as Missing' { (Test-SemtSharedInventoryUsable -Path $sharedAd -Kind AD -MaxAgeHours 12).Reason -eq 'Missing' }
    Set-Content -LiteralPath $sharedAd -Value 'Name,Enabled' -Encoding UTF8
    (Get-Item -LiteralPath $sharedAd).LastWriteTime = (Get-Date).AddHours(-20)
    Test-Step 'stale shared inventory is not reused' { -not (Test-SemtSharedInventoryUsable -Path $sharedAd -Kind AD -MaxAgeHours 12).Usable }
    $sharedIntune = Get-SemtSharedInventoryPath -PackageRoot $inventoryPackage -Kind Intune
    [pscustomobject]@{ DeviceName = 'PC-1'; InventoryTenantId = 't1'; InventoryAuthenticationMode = 'AppOnly'; InventoryScope = 'AllManagedDevices' } | Export-Csv -LiteralPath $sharedIntune -NoTypeInformation -Encoding UTF8
    Test-Step 'Intune inventory with unverified provenance is rejected' { (Test-SemtSharedInventoryUsable -Path $sharedIntune -Kind Intune -MaxAgeHours 2).Reason -eq 'ProvenanceRejected' }
    [pscustomobject]@{ DeviceName = 'PC-1'; InventoryTenantId = 't1'; InventoryAuthenticationMode = 'DelegatedInteractive'; InventoryScope = 'AllManagedDevices' } | Export-Csv -LiteralPath $sharedIntune -NoTypeInformation -Encoding UTF8
    Test-Step 'fresh Intune inventory with verified provenance is reused' { (Test-SemtSharedInventoryUsable -Path $sharedIntune -Kind Intune -MaxAgeHours 2).Usable }
    Test-Throws 'a scoped export never overwrites a shared inventory' { Invoke-SemtInventoryExport -Kind AD -ExportScriptPath (Join-Path $tempRoot 'none.ps1') -OutputPath $sharedAd -LogPath (Join-Path $tempRoot 'inv.log') -ScopedComputers @('PC-1') -PackageRoot $inventoryPackage } '.'
    #endregion

    #region Paths
    $unicodeFolder = Join-Path $tempRoot ('Dossier avec espaces ' + [char]0x00E9 + [char]0x4E2D)
    New-Item -ItemType Directory -Path $unicodeFolder -Force | Out-Null
    Test-Step 'JSON written and read back under a space and Unicode path' {
        $path = Join-Path $unicodeFolder 'state.json'
        Write-SemtJsonFileAtomic -Path $path -InputObject ([pscustomobject]@{ Value = 'ok' })
        (Read-SemtJsonFile -Path $path).Value -eq 'ok'
    }
    Test-Step 'SHA256 works on a path longer than 260 characters' {
        $deep = $tempRoot
        while ($deep.Length -lt 300) { $deep = Join-Path $deep ('segment-' + ('x' * 30)) }
        $long = '\\?\' + $deep
        [void][IO.Directory]::CreateDirectory($long)
        $file = Join-Path $deep 'payload.bin'
        [IO.File]::WriteAllBytes('\\?\' + $file, [byte[]](1, 2, 3))
        (Get-SemtFileSha256 -Path $file) -eq '039058C6F2C0CB492C533B0A4D14EF77CC0F78ABCCCED5287D84A1A2011CFB81'
    }
    Test-Step 'file URIs in reports are escaped' { (ConvertTo-SemtFileUri -Path 'C:\Runs\A #1\r%e.html') -notmatch ' |#1' }
    #endregion

    #region GUI launcher helpers
    $guiPath = Join-Path $script:PackageRoot 'Scripts\SmartM365-EndpointMigration-GUI.ps1'
    foreach ($text in @(Get-FunctionText -Path $guiPath -Names @('Assert-CmdLiteralValue','ConvertTo-CmdArgument','ConvertTo-CmdWindowTitle','New-GuiLaunchCommandFile','Test-UnlimitedCycleConfirmationRequired'))) { . ([scriptblock]::Create($text)) }
    Test-Throws 'CMD launcher refuses percent signs (no variable expansion)' { ConvertTo-CmdArgument -Value 'C:\a%PATH%b' } 'cannot be represented'
    Test-Throws 'CMD launcher refuses quotes' { ConvertTo-CmdArgument -Value 'C:\a"b' } 'cannot be represented'
    Test-Step 'CMD launcher quotes paths with spaces' { (ConvertTo-CmdArgument -Value 'C:\Dossier avec espaces\x.ps1') -eq '"C:\Dossier avec espaces\x.ps1"' }
    Test-Step 'CMD window title neutralises metacharacters' { (ConvertTo-CmdWindowTitle -Value 'a&b|c<d>e^f%g') -notmatch '[&|<>^%]' }
    Test-Step 'generated GUI launcher keeps the exit code and pauses' {
        $previousTemp = $env:TEMP
        $env:TEMP = $tempRoot
        try { $launcher = New-GuiLaunchCommandFile -WorkingDirectory $unicodeFolder -Commands @('echo test') -NamePrefix 'SEMT-Test' -WindowTitle 'SEMT test' -PauseWhenDone }
        finally { $env:TEMP = $previousTemp }
        $content = [IO.File]::ReadAllText($launcher)
        $content -match 'setlocal DisableDelayedExpansion' -and $content -match 'exit /b %EXITCODE%' -and $content -match 'pause' -and $content -match 'chcp 65001'
    }
    Test-Step 'unlimited Loop requires an explicit confirmation' { (Test-UnlimitedCycleConfirmationRequired -RunMode Loop -MaxCycles 0) -and -not (Test-UnlimitedCycleConfirmationRequired -RunMode Once -MaxCycles 0) -and -not (Test-UnlimitedCycleConfirmationRequired -RunMode Loop -MaxCycles 3) }
    #endregion

    #region Shared target guards (endpoint)
    $hjRegion = Get-RegionText -Path $hjEndpoint -Region 'SEMT shared target guards v1'
    $w11Region = Get-RegionText -Path $w11Endpoint -Region 'SEMT shared target guards v1'
    Test-Step 'shared target guard region is identical in both endpoint scripts' { $hjRegion -ceq $w11Region }
    Test-Step 'target mutex has its own SEMT name (not a legacy name)' { $hjRegion -match "Global\\SmartM365_SEMT_TargetOperation" -and $hjRegion -notmatch "SemtTargetMutexName = 'Global\\SmartM365_(IntuneHybridJoin|Windows11Upgrade)Toolkit" }
    Test-Step 'legacy detection never writes, deletes or waits on legacy state' {
        $body = ([regex]::Match($hjRegion, '(?s)function Get-SemtLegacyEndpointActivity.*?\n}\n')).Value
        $body -notmatch '(Remove-Item|Set-Content|Write-SemtTargetJson|Move-Item|New-Item|\.WaitOne|Unregister-ScheduledTask|Stop-Process|SetValue|CreateSubKey|DeleteSubKey)'
    }
    . ([scriptblock]::Create($w11Region))
    $targetRoot = Join-Path $tempRoot 'TargetData'
    $hjData = Join-Path $targetRoot 'HybridJoin'
    $w11Data = Join-Path $targetRoot 'Windows11Upgrade'
    $guardHolder = Start-Job -ScriptBlock {
        param($Region, $DataRoot)
        . ([scriptblock]::Create($Region))
        $r = Enter-SemtTargetOperationGuard -Campaign Windows11Upgrade -CampaignDataRoot $DataRoot -RunId R-W11 -CorrelationId C-W11
        'ACQUIRED=' + $r.Acquired
        Start-Sleep -Seconds 20
        Exit-SemtTargetOperationGuard -CampaignDataRoot $DataRoot
    } -ArgumentList $w11Region, $w11Data
    try {
        $held = $false
        for ($i = 0; $i -lt 30; $i++) { Start-Sleep -Seconds 1; if (Test-Path (Join-Path $targetRoot 'State\TargetOperation.json')) { $held = $true; break } }
        Test-Step 'cross-campaign target guard: lease written by the holder' { $held }
        $blocked = Enter-SemtTargetOperationGuard -Campaign HybridJoin -CampaignDataRoot $hjData -RunId R-HJ -CorrelationId C-HJ
        Test-Step 'cross-campaign target guard blocks the other Campaign with an explicit status' { -not $blocked.Acquired -and $blocked.Status -eq 'TARGET_OPERATION_ACTIVE' -and $blocked.Detail -match 'Windows11Upgrade' }
    }
    finally { [void](Wait-Job $guardHolder -Timeout 60); Remove-Job $guardHolder -Force -ErrorAction SilentlyContinue }
    $acquired = Enter-SemtTargetOperationGuard -Campaign HybridJoin -CampaignDataRoot $hjData -RunId R-HJ -CorrelationId C-HJ
    Test-Step 'target guard is acquired once the holder released it' { $acquired.Acquired }
    Exit-SemtTargetOperationGuard -CampaignDataRoot $hjData
    Set-SemtTargetPendingOperation -Campaign Windows11Upgrade -CampaignDataRoot $w11Data -Reason 'SETUP_STARTED_REBOOT_PENDING' -Hours 168
    $pendingBlocked = Enter-SemtTargetOperationGuard -Campaign HybridJoin -CampaignDataRoot $hjData -RunId R-HJ -CorrelationId C-HJ2
    Test-Step 'pending Windows 11 setup blocks Hybrid Join (TARGET_PENDING_OPERATION)' { -not $pendingBlocked.Acquired -and $pendingBlocked.Status -eq 'TARGET_PENDING_OPERATION' }
    $sameCampaign = Enter-SemtTargetOperationGuard -Campaign Windows11Upgrade -CampaignDataRoot $w11Data -RunId R-W11 -CorrelationId C-W11b
    Test-Step 'pending marker does not block its own Campaign' { $sameCampaign.Acquired }
    Exit-SemtTargetOperationGuard -CampaignDataRoot $w11Data
    Test-Step 'pending marker is cleared only by its own Campaign' { -not (Clear-SemtTargetPendingOperation -Campaign HybridJoin -CampaignDataRoot $hjData) -and (Clear-SemtTargetPendingOperation -Campaign Windows11Upgrade -CampaignDataRoot $w11Data) }
    Write-SemtTargetJson -Path (Join-Path $targetRoot 'State\TargetPendingOperation.json') -Data ([pscustomobject]@{ Campaign = 'Windows11Upgrade'; Reason = 'old'; ExpiresUtc = (Get-Date).ToUniversalTime().AddHours(-1).ToString('o'); LastBootUtc = ''; ClearedByReboot = $true })
    $expired = Enter-SemtTargetOperationGuard -Campaign HybridJoin -CampaignDataRoot $hjData -RunId R-HJ -CorrelationId C-HJ3
    Test-Step 'expired pending marker no longer blocks' { $expired.Acquired }
    Exit-SemtTargetOperationGuard -CampaignDataRoot $hjData
    Set-Content -LiteralPath (Join-Path $targetRoot 'State\TargetPendingOperation.json') -Value '{ not json' -Encoding ASCII
    $unreadable = Enter-SemtTargetOperationGuard -Campaign HybridJoin -CampaignDataRoot $hjData -RunId R-HJ -CorrelationId C-HJ4
    Test-Step 'unreadable pending marker gives an explicit indeterminate status' { -not $unreadable.Acquired -and $unreadable.Status -eq 'TARGET_STATE_UNREADABLE' }
    Exit-SemtTargetOperationGuard -CampaignDataRoot $hjData
    Remove-Item -LiteralPath (Join-Path $targetRoot 'State\TargetPendingOperation.json') -Force
    Test-Step 'attempt results keep the newest 50 files' {
        for ($i = 0; $i -lt 55; $i++) { [void](Write-SemtAttemptResult -CampaignDataRoot $hjData -Data @{ CorrelationId = ('c-{0:000}' -f $i); State = 'Completed' }) }
        @(Get-ChildItem -LiteralPath (Join-Path $hjData 'State\Results') -Filter '*.json').Count -eq 50
    }
    #endregion

    #region Legacy detection (read-only, synthetic roots)
    $script:SemtLegacyHybridJoinRoot = Join-Path $tempRoot 'Legacy\IntuneHybridJoinToolkit'
    $script:SemtLegacyWindows11Root = Join-Path $tempRoot 'Legacy\Windows11UpgradeToolkit'
    $baseline = Get-SemtLegacyEndpointActivity -Campaign HybridJoin
    Test-Step 'legacy detection returns a known status on this computer' { $baseline.Status -in @('LEGACY_ABSENT','LEGACY_INSTALLED_IDLE','LEGACY_INTUNE_PACKAGE_ASSIGNED','LEGACY_ENDPOINT_RUNNING','LEGACY_SETUP_IN_PROGRESS','LEGACY_COPY_IN_PROGRESS','LEGACY_RETRY_AFTER_REBOOT_PENDING','LEGACY_STATE_UNREADABLE') }
    if ($baseline.Status -eq 'LEGACY_ABSENT') {
        New-Item -ItemType Directory -Path $script:SemtLegacyHybridJoinRoot -Force | Out-Null
        $idle = Get-SemtLegacyEndpointActivity -Campaign HybridJoin
        Test-Step 'installed but idle legacy toolkit does not block' { $idle.Status -eq 'LEGACY_INSTALLED_IDLE' -and -not $idle.Blocking }
        $lockDir = Join-Path $script:SemtLegacyWindows11Root 'Locks\Win11-fr-FR'
        New-Item -ItemType Directory -Path $lockDir -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $lockDir 'lock.json') -Value ('{ "PID": ' + $PID + ' }') -Encoding ASCII
        $copy = Get-SemtLegacyEndpointActivity -Campaign Windows11Upgrade
        Test-Step 'legacy media lock held by a live process blocks (LEGACY_COPY_IN_PROGRESS)' { $copy.Status -eq 'LEGACY_COPY_IN_PROGRESS' -and $copy.Blocking }
        Set-Content -LiteralPath (Join-Path $lockDir 'lock.json') -Value '{ "PID": 999999 }' -Encoding ASCII
        Test-Step 'legacy media lock of a dead process does not block' { -not (Get-SemtLegacyEndpointActivity -Campaign Windows11Upgrade).Blocking }
        Set-Content -LiteralPath (Join-Path $lockDir 'lock.json') -Value '{ broken' -Encoding ASCII
        $unknown = Get-SemtLegacyEndpointActivity -Campaign Windows11Upgrade
        Test-Step 'unreadable legacy state gives an explicit blocking status' { $unknown.Status -eq 'LEGACY_STATE_UNREADABLE' -and $unknown.Blocking }
        Test-Step 'legacy files were not modified by detection' { [IO.File]::ReadAllText((Join-Path $lockDir 'lock.json')).Trim() -eq '{ broken' }
    }
    else {
        Write-Host ("INFO: real legacy state present on this computer ({0}); synthetic legacy status checks skipped." -f $baseline.Status)
    }
    #endregion

    #region Legacy setup media reuse
    # Behaviour is covered by Tests\Windows11Upgrade\Test-SemtWindows11MediaResolution.ps1 and
    # Test-SemtWindows11MediaAdoption.ps1 (synthetic media, fault injection, real ACLs).
    Test-Step 'Windows 11 endpoint uses the shared media resolution before any copy' { $text = [IO.File]::ReadAllText($w11Endpoint); $text -match 'Resolve-SemtEndpointSetupMedia -CachePath \$cachePath' -and $text -match 'function Invoke-SemtSetupMediaAdoption' -and $text -notmatch 'Invoke-SemtLegacySetupMediaReuse' }
    #endregion
}
catch {
    Write-TestResult -Condition $false -Message 'test harness' -Detail ($_.Exception.Message + ' @ ' + $_.InvocationInfo.PositionMessage)
}
finally {
    $env:SEMT_RUNTIME_ROOT_OVERRIDE = $previousOverride
    try { if (Test-Path -LiteralPath $tempRoot) { [IO.Directory]::Delete('\\?\' + $tempRoot, $true) } } catch { Write-Host ("WARN: temp folder not fully removed: {0}" -f $tempRoot) }
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $script:Passed,$script:Failed)
if ($script:Failed -gt 0) { exit 1 }
exit 0
