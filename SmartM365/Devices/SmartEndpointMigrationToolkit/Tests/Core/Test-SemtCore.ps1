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
    foreach ($campaign in @('HybridJoin','Windows11Upgrade')) {
        Test-Step "legacy activity defaults to Warn for $campaign" { (Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign $campaign)['LegacyActivityPolicy'].Value -eq 'Warn' }
        Test-Step "legacy activity Block override remains available for $campaign" { (Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign $campaign -GlobalLayer @{ LegacyActivityPolicy = 'Block' })['LegacyActivityPolicy'].Value -eq 'Block' }
    }
    foreach ($endpoint in @('Campaigns\HybridJoin\Scripts\SmartM365-Invoke-EndpointMigration-HybridJoin.ps1','Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1')) {
        $content = Get-Content -LiteralPath (Join-Path $script:PackageRoot $endpoint) -Raw
        Test-Step "standalone endpoint defaults legacy activity to Warn: $endpoint" { $content -match "\[string\]\`$LegacyActivityPolicy\s*=\s*'Warn'" }
    }
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
        $config = Join-Path $tempRoot 'heal.local.json.txt'
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
    Test-Step 'New-SemtWave creates Computers.txt, Wave.local.json.txt and the four launchers' {
        (Test-Path (Join-Path $wave.Path 'Computers.txt')) -and (Test-Path (Join-Path $wave.Path 'Wave.local.json.txt')) -and @(Get-ChildItem -LiteralPath $wave.Path -Filter 'Start-Wave-*.cmd').Count -eq 4
    }
    Test-Step 'Wave configuration uses the new file and ignores an old JSON file' {
        $path = Join-Path $tempRoot 'old-wave'
        New-Item -ItemType Directory -Path $path -Force | Out-Null
        [IO.File]::WriteAllText((Join-Path $path 'Wave.local.json'), '{"Options":{"ThrottleLimit":123}}')
        $cfg = Read-SemtWaveConfig -PackageRoot $packageCopy -Campaign HybridJoin -WavePath $path -CreateWhenMissing
        $cfg.Path.EndsWith('Wave.local.json.txt') -and $cfg.Data.Options.ThrottleLimit -ne 123 -and (Test-Path (Join-Path $path 'Wave.local.json'))
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
    $configDir = Join-Path $packageCopy 'Config'
    New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    Copy-Item (Join-Path $script:PackageRoot 'Config\SmartEndpointMigrationToolkit.local.json.template') $configDir
    $externalWaves = Join-Path $tempRoot 'External Waves'
    $externalRuns = Join-Path $tempRoot 'External Runs'
    $runtimePath = Join-Path $configDir 'SmartEndpointMigrationToolkit.local.json.txt'
    Write-SemtJsonFileAtomic -Path $runtimePath -InputObject @{ Storage = @{ WavesPath = $externalWaves; RunsPath = $externalRuns } }
    $externalWave = New-SemtWave -PackageRoot $packageCopy -Campaign Windows11Upgrade -Name 'WAVE-EXTERNAL' -Computers @('PC-3')
    Test-Step 'configured Waves path is used and existing runtime config is healed' {
        $externalWave.Path -eq (Join-Path $externalWaves 'Windows11Upgrade\WAVE-EXTERNAL') -and
        (Read-SemtJsonFile -Path $runtimePath).Options -and @(Get-SemtWaves -PackageRoot $packageCopy -Campaign Windows11Upgrade).Count -eq 1
    }
    $externalRun = New-SemtRunContext -PackageRoot $packageCopy -Campaign Windows11Upgrade -WavePath $externalWave.Path -RunId 'TEST-RUN'
    Test-Step 'configured Runs path is used for Wave and Single PC Runs' {
        $single = New-SemtEphemeralWave -PackageRoot $packageCopy -Campaign HybridJoin -ComputerName PC-4 -RunId 'TEST-SINGLE'
        $externalRun.RunRoot -eq (Join-Path $externalRuns 'Windows11Upgrade\WAVE-EXTERNAL\TEST-RUN') -and $single.RunRoot -eq (Join-Path $externalRuns 'HybridJoin\_SinglePC\TEST-SINGLE')
    }
    $scriptDir = Join-Path $packageCopy 'Scripts'
    New-Item -ItemType Directory -Path $scriptDir -Force | Out-Null
    $stub = '@echo off' + "`r`n" + 'echo %~1^|%~2^|%~3 > "%SEMT_WRAPPER_RESULT%"' + "`r`n" + 'exit /b 7' + "`r`n"
    [IO.File]::WriteAllText((Join-Path $scriptDir 'Run-EndpointMigrationWave.cmd'),$stub,[Text.Encoding]::ASCII)
    $env:SEMT_WRAPPER_RESULT = Join-Path $tempRoot 'wrapper-result.txt'
    & (Join-Path $externalWave.Path 'Start-Wave-Once-IgnoreRunGuard.cmd') | Out-Null
    Test-Step 'external Wave launcher calls the package and propagates its exit code' {
        $LASTEXITCODE -eq 7 -and (Get-Content -LiteralPath $env:SEMT_WRAPPER_RESULT -Raw) -match 'WAVE-EXTERNAL.*\|Once\|IgnoreRunGuard'
    }
    Remove-Item Env:SEMT_WRAPPER_RESULT -ErrorAction SilentlyContinue
    Write-SemtJsonFileAtomic -Path $runtimePath -InputObject @{ Storage = @{ WavesPath = 'relative'; RunsPath = '' } }
    Test-Throws 'relative storage path is rejected' { Get-SemtStorageRoot -PackageRoot $packageCopy -Kind Waves } 'SEMT_STORAGE_PATH_INVALID'
    Write-SemtJsonFileAtomic -Path $runtimePath -InputObject @{ Storage = @{ WavesPath = ''; RunsPath = '' } }
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
    $stale = Join-Path $gate.LeaseRoot 'lease_999999_stale.json.txt'
    Write-SemtJsonFileAtomic -Path $stale -InputObject ([pscustomobject]@{ SchemaVersion = 1; Campaign = 'HybridJoin'; Computer = 'PC-STALE'; Limit = 2; LauncherProcessId = 999999; LauncherProcessStartUtc = '2020-01-01T00:00:00.0000000Z'; CreatedUtc = (Get-SemtUtcNowText); HeartbeatUtc = (Get-SemtUtcNowText) })
    $snapshot = & $module { param($g) Get-SemtWorkerGateSnapshotUnlocked -Gate $g -RemoveStale } $gate
    Test-Step 'reservation of a dead launcher is recovered as stale' { @($snapshot.Removed | Where-Object { $_.Reason -eq 'LauncherProcessExited' }).Count -eq 1 -and -not (Test-Path -LiteralPath $stale) }
    $oldLease = Join-Path $gate.LeaseRoot 'lease_heartbeat_old.json.txt'
    $selfStart = & $module { Get-SemtProcessStartUtc -ProcessId $PID }
    Write-SemtJsonFileAtomic -Path $oldLease -InputObject ([pscustomobject]@{ SchemaVersion = 1; Campaign = 'HybridJoin'; Computer = 'PC-OLD'; Limit = 2; LauncherProcessId = $PID; LauncherProcessStartUtc = $selfStart; CreatedUtc = (Get-Date).ToUniversalTime().AddMinutes(-30).ToString('o'); HeartbeatUtc = (Get-Date).ToUniversalTime().AddMinutes(-25).ToString('o') })
    $snapshot = & $module { param($g) Get-SemtWorkerGateSnapshotUnlocked -Gate $g -RemoveStale } $gate
    Test-Step 'reservation without heartbeat is recovered as stale' { @($snapshot.Removed | Where-Object { $_.Reason -eq 'HeartbeatStale' }).Count -eq 1 }
    Remove-SemtWorkerLease -LeasePath $lease2 | Out-Null
    $strict = Join-Path $gate.LeaseRoot 'lease_strict.json.txt'
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
    $contractPath = Join-Path $tempRoot 'contract.json.txt'
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
    Write-SemtJsonFileAtomic -Path (Join-Path $share 'State\Results\attempt-1.json.txt') -InputObject ([pscustomobject]@{ CorrelationId = 'attempt-OTHER'; State = 'Completed' })
    Test-Step 'foreign attempt result (other CorrelationId) is never trusted' { $r = Read-SemtRemoteAttemptResult -RemoteDataShare $share -CorrelationId 'attempt-1'; $r.State -eq 'Foreign' -and -not $r.Completed }
    Write-SemtJsonFileAtomic -Path (Join-Path $share 'State\Results\attempt-2.json.txt') -InputObject ([pscustomobject]@{ CorrelationId = 'attempt-2'; State = 'Running' })
    Test-Step 'running attempt result is reported as not completed' { $r = Read-SemtRemoteAttemptResult -RemoteDataShare $share -CorrelationId 'attempt-2'; $r.State -eq 'Running' -and -not $r.Completed }
    Write-SemtJsonFileAtomic -Path (Join-Path $share 'State\Results\attempt-3.json.txt') -InputObject ([pscustomobject]@{ CorrelationId = 'attempt-3'; State = 'Completed'; FinalStatus = 'ENROLLED' })
    Test-Step 'matching completed attempt result is accepted' { (Read-SemtRemoteAttemptResult -RemoteDataShare $share -CorrelationId 'attempt-3').Completed }
    Test-Step 'missing attempt result is reported, not invented' { -not (Read-SemtRemoteAttemptResult -RemoteDataShare $share -CorrelationId 'attempt-4').Exists }
    Test-Step 'full evidence collects external JSON as .json.txt without overwriting a namesake' {
        $folder = Join-Path $share 'Output'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        [IO.File]::WriteAllText((Join-Path $folder 'external.json'), '{"source":"external"}')
        [IO.File]::WriteAllText((Join-Path $folder 'collision.json'), '{"source":"old"}')
        [IO.File]::WriteAllText((Join-Path $folder 'collision.json.txt'), '{"source":"new"}')
        $dest = Join-Path $tempRoot 'FullJsonEvidence'
        [void](Copy-SemtCentralEvidence -RemoteDataShare $share -RemoteDataRoot 'C:\Data' -TargetPath $dest -CorrelationId 'attempt-3' -AttemptResult $null -AttemptStartedUtc ([datetime]::UtcNow) -Mode Full)
        (Test-Path (Join-Path $dest 'Output\external.json.txt')) -and
        -not (Test-Path (Join-Path $dest 'Output\external.json')) -and
        ([IO.File]::ReadAllText((Join-Path $dest 'Output\collision.json.txt')) -eq '{"source":"new"}') -and
        ((Get-Content (Join-Path $dest 'CentralLogCollection.skipped.txt') -Raw) -match 'conflicting')
    }
    # Why there is no result: data folder not visible, no Results folder, or no file for this attempt.
    $hiddenShare = Join-Path $tempRoot 'TargetShareNotVisible'
    $emptyShare = Join-Path $tempRoot 'TargetShareEmpty'
    New-Item -ItemType Directory -Path $emptyShare -Force | Out-Null
    Test-Step 'missing attempt result says why (data folder not visible, no Results folder, no file)' {
        (Read-SemtRemoteAttemptResult -RemoteDataShare $hiddenShare -CorrelationId 'attempt-4').Detail -match 'data folder is not visible' -and (Read-SemtRemoteAttemptResult -RemoteDataShare $emptyShare -CorrelationId 'attempt-4').Detail -match 'no State\\Results folder' -and (Read-SemtRemoteAttemptResult -RemoteDataShare $share -CorrelationId 'attempt-4').Detail -match 'no result file for this attempt'
    }
    Test-Step 'evidence collection records why nothing was collected when the data folder is not visible' {
        $evidenceTarget = Join-Path $tempRoot 'EvidenceNotVisible\Latest'
        $copied = Copy-SemtCentralEvidence -RemoteDataShare $hiddenShare -RemoteDataRoot 'C:\ProgramData\WCH\Test' -TargetPath $evidenceTarget -CorrelationId 'attempt-4' -AttemptResult $null -AttemptStartedUtc ([datetime]::UtcNow)
        $copied -eq 0 -and (Get-Content -LiteralPath (Join-Path $evidenceTarget 'CentralLogCollection.txt') -Raw) -match 'not visible through the administrative share'
    }
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
    (Get-Item -LiteralPath $sharedAd).LastWriteTime = (Get-Date).AddHours(-23)
    Test-Step 'AD inventory younger than 24 hours is reused' { (Test-SemtSharedInventoryUsable -Path $sharedAd -Kind AD -MaxAgeHours 24).Usable }
    (Get-Item -LiteralPath $sharedAd).LastWriteTime = (Get-Date).AddHours(-25)
    Test-Step 'AD inventory older than 24 hours is rejected' { -not (Test-SemtSharedInventoryUsable -Path $sharedAd -Kind AD -MaxAgeHours 24).Usable }
    $sharedIntune = Get-SemtSharedInventoryPath -PackageRoot $inventoryPackage -Kind Intune
    [pscustomobject]@{ DeviceName = 'PC-1'; InventoryTenantId = 't1'; InventoryAuthenticationMode = 'AppOnly'; InventoryScope = 'AllManagedDevices' } | Export-Csv -LiteralPath $sharedIntune -NoTypeInformation -Encoding UTF8
    Test-Step 'Intune inventory with unverified provenance is rejected' { (Test-SemtSharedInventoryUsable -Path $sharedIntune -Kind Intune -MaxAgeHours 2).Reason -eq 'ProvenanceRejected' }
    [pscustomobject]@{ DeviceName = 'PC-1'; InventoryTenantId = 't1'; InventoryAuthenticationMode = 'DelegatedInteractive'; InventoryScope = 'AllManagedDevices' } | Export-Csv -LiteralPath $sharedIntune -NoTypeInformation -Encoding UTF8
    Test-Step 'fresh Intune inventory with verified provenance is reused' { (Test-SemtSharedInventoryUsable -Path $sharedIntune -Kind Intune -MaxAgeHours 2).Usable }
    Test-Throws 'a scoped export never overwrites a shared inventory' { Invoke-SemtInventoryExport -Kind AD -ExportScriptPath (Join-Path $tempRoot 'none.ps1') -OutputPath $sharedAd -LogPath (Join-Path $tempRoot 'inv.log') -ScopedComputers @('PC-1') -PackageRoot $inventoryPackage } '.'
    # The shared AD inventory always covers the forest; the refusal happens before any AD query.
    $adExport = Join-Path $script:PackageRoot 'Scripts\SmartM365-EndpointMigration-Export-ADDevicesCsv.ps1'
    Test-Throws 'a single-domain AD export never writes the shared AD inventory' { & $adExport -Domain 'contoso.test' -ForceRefresh } 'SEMT_SHARED_AD_INVENTORY_DOMAIN_SCOPED'
    $engineText = [IO.File]::ReadAllText((Join-Path $script:PackageRoot 'Scripts\SmartM365-Invoke-EndpointMigrationWithPsExec.ps1'))
    # Graph module installed by PowerShellGet in a folder missing from PSModulePath (redirected Documents).
    $fakeModules = Join-Path $tempRoot 'Redirected Documents\WindowsPowerShell\Modules'
    $fakeLocation = Join-Path $fakeModules 'Microsoft.Graph.Authentication\9.9.9'
    New-Item -ItemType Directory -Path $fakeLocation -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $fakeLocation 'Microsoft.Graph.Authentication.psd1') -Value "@{ ModuleVersion = '9.9.9'; GUID = '00000000-0000-4000-8000-000000000099' }" -Encoding ASCII
    $savedModulePath = $env:PSModulePath
    $coreModule = Get-Module SmartEndpointMigration.Core
    try {
        $env:PSModulePath = Join-Path $PSHOME 'Modules'
        & $coreModule { param($Location) $script:FakeGraphLocation = $Location; function script:Get-InstalledModule { param($Name,[switch]$AllVersions,$ErrorAction) if ($script:FakeGraphLocation) { [pscustomobject]@{ Name = $Name; Version = [version]'9.9.9'; InstalledLocation = $script:FakeGraphLocation } } } } $fakeLocation
        $found = Find-SemtGraphAuthenticationModule
        Test-Step 'Graph module installed outside PSModulePath is found through its install record' { $null -ne $found -and [string]$found.Version -eq '9.9.9' -and @($env:PSModulePath -split ';') -contains $fakeModules }
        $env:PSModulePath = Join-Path $PSHOME 'Modules'
        & $coreModule { $script:FakeGraphLocation = Join-Path $env:TEMP ('missing-' + [guid]::NewGuid().ToString('N')) }
        Test-Step 'Graph module with a stale install record is reported missing, with a diagnostic' { $null -eq (Find-SemtGraphAuthenticationModule) -and (Get-SemtGraphModuleDiagnostic) -match 'PSModulePath: .*Install-Module Microsoft\.Graph\.Authentication -Scope AllUsers' }
    }
    finally {
        $env:PSModulePath = $savedModulePath
        & $coreModule { Remove-Item -Path Function:\Get-InstalledModule -ErrorAction SilentlyContinue; $script:FakeGraphLocation = $null }
    }
    # Graph paging (fictional 3-page response): one short progress line per page, never the page URL.
    foreach ($exporterName in @('Intune','Entra')) {
        $exporterPath = Join-Path $script:PackageRoot ("Scripts\SmartM365-EndpointMigration-Export-{0}DevicesCsv.ps1" -f $exporterName)
        $pagingCheck = [scriptblock]::Create((@(Get-FunctionText -Path $exporterPath -Names @('Get-GraphProperty','Invoke-GraphPagedRequest')) -join "`n") + @'

$script:HostLines = New-Object System.Collections.Generic.List[string]
function Write-Host { param([object]$Object,$ForegroundColor) $script:HostLines.Add([string]$Object) }
function Invoke-MgGraphRequest {
    param($Method,$Uri,$ErrorAction)
    $page = if ($Uri -match 'skiptoken=(\d+)') { [int]$Matches[1] } else { 1 }
    $result = @{ value = @(1..3 | ForEach-Object { @{ id = ('{0}-{1}' -f $page,$_) } }) }
    if ($page -lt 3) { $result['@odata.nextLink'] = 'https://graph.microsoft.com/v1.0/x?$select=a%2cb&$skiptoken=' + ($page + 1) }
    return $result
}
$all = @(Invoke-GraphPagedRequest -Uri 'https://graph.microsoft.com/v1.0/x?$top=3')
$loud = @($script:HostLines)
$script:HostLines.Clear()
[void](Invoke-GraphPagedRequest -Uri 'https://graph.microsoft.com/v1.0/x?$top=3' -Quiet)
[pscustomobject]@{ Count = $all.Count; Lines = $loud; QuietLines = $script:HostLines.Count }
'@)
        $paging = & $pagingCheck
        Test-Step ("{0} export: every Graph page is read and reported on one short line without its URL" -f $exporterName) { $paging.Count -eq 9 -and ($paging.Lines -join '|') -eq 'Graph page 1: 3 device(s) read so far.|Graph page 2: 6 device(s) read so far.|Graph page 3: 9 device(s) read so far.' }
        Test-Step ("{0} export: per-computer lookups are quiet" -f $exporterName) { $paging.QuietLines -eq 0 }
    }
    Test-Step 'Run inventory refresh searches every forest domain (AdDomain is a DNS suffix only)' { $engineText -match 'Invoke-SemtInventoryExport -Kind \$Kind' -and $engineText -notmatch "-Domain \(Opt 'AdDomain'\)" }
    #endregion

    #region Paths
    $unicodeFolder = Join-Path $tempRoot ('Dossier avec espaces ' + [char]0x00E9 + [char]0x4E2D)
    New-Item -ItemType Directory -Path $unicodeFolder -Force | Out-Null
    Test-Step 'JSON written and read back under a space and Unicode path' {
        $path = Join-Path $unicodeFolder 'state.json.txt'
        Write-SemtJsonFileAtomic -Path $path -InputObject ([pscustomobject]@{ Value = 'ok' })
        (Read-SemtJsonFile -Path $path).Value -eq 'ok'
    }
    # Run folders repeat the Wave name (Runs\<Campaign>\<Wave>\<RunId>): report names and temporary
    # files must not add to the path (Windows PowerShell 5.1 fails above 260 characters).
    $nearLimit = $tempRoot
    while ($nearLimit.Length -lt 200) { $nearLimit = Join-Path $nearLimit ('WAVE-AUTO-HJ-FR-TEST-' + ('x' * 20)) }
    $nearLimit = $nearLimit.Substring(0, 200).TrimEnd('\')
    $reportRoot = Join-Path $nearLimit 'Reports'
    New-Item -ItemType Directory -Path $reportRoot -Force | Out-Null
    Test-Step 'JSON atomic write of a 250-character path (short temporary name)' {
        $path = Join-Path $nearLimit ('S' * (249 - $nearLimit.Length) + '.json.txt')
        Write-SemtJsonFileAtomic -Path $path -InputObject ([pscustomobject]@{ Value = 'near-limit' })
        $path.Length -ge 249 -and (Read-SemtJsonFile -Path $path).Value -eq 'near-limit' -and @(Get-ChildItem -LiteralPath $nearLimit -Force -Filter '*.tmp').Count -eq 0
    }
    $reportContext = [pscustomobject]@{ Campaign = 'HybridJoin'; RunId = '20261029-102531-2cf749'; WaveName = 'WAVE-AUTO-HJ-FR-TEST-' + ('L' * 59); PackageRoot = $packageRoot; ReportRoot = $reportRoot; RunRoot = $nearLimit; SinglePc = $false; WavePath = $nearLimit; LauncherLogPath = (Join-Path $nearLimit 'Launcher.log') }
    $reportState = New-SemtReportState -RunContext $reportContext -Columns @('ComputerName','Status') -CampaignTitle 'Hybrid Join'
    Test-Step 'Run report name holds the Campaign and RunId only (the Wave is in the folder path)' { (Split-Path -Leaf $reportState.HtmlPath) -eq 'SEMT_HybridJoin_20261029-102531-2cf749.html' }
    Test-Step 'Run HTML report is written in a folder near the path limit and leaves no temporary file' {
        [void](Write-SemtRunHtmlReport -State $reportState -CycleNumber 1 -IsLive)
        (Test-Path -LiteralPath $reportState.HtmlPath -PathType Leaf) -and @(Get-ChildItem -LiteralPath $reportRoot -Force -Filter '*.tmp').Count -eq 0
    }
    # A virtual machine skipped by -SkipVirtualMachines must never consume the endpoint run guard:
    # Hybrid Join left LastRun.json.txt at RUNNING, Windows 11 checked the guard before the skip.
    $hjEndpointText = [IO.File]::ReadAllText((Join-Path $script:PackageRoot 'Campaigns\HybridJoin\Scripts\SmartM365-Invoke-EndpointMigration-HybridJoin.ps1'))
    $w11EndpointText = [IO.File]::ReadAllText((Join-Path $script:PackageRoot 'Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'))
    Test-Step 'Hybrid Join endpoint: a skipped virtual machine neither applies nor writes the run guard' {
        $function = $hjEndpointText.IndexOf('function Get-ComputerSystemSummary {')
        $early = $hjEndpointText.IndexOf('$script:SkipAsVirtualMachine = ($earlyComputerSystem.Success')
        $guard = $hjEndpointText.IndexOf('# 12-hour local run guard')
        $function -ge 0 -and $function -lt $early -and $early -lt $guard -and $hjEndpointText -match '(?s)if \(\$script:SkipAsVirtualMachine\) \{\s*Write-RunLog "\$RunGuardHours-hour run guard not applied' -and $hjEndpointText -match '(?s)if \(-not \$script:SkipAsVirtualMachine\) \{\s*\[PSCustomObject\]@\{\s*RunId\s*=\s*\$RunId'
    }
    Test-Step 'Windows 11 endpoint: a virtual machine is skipped before the run guard is checked' {
        $skip = $w11EndpointText.IndexOf("`$status = 'SKIPPED_VIRTUAL_MACHINE'")
        $guard = $w11EndpointText.IndexOf('if (-not $IgnoreRunGuard -and $RunGuardHours -gt 0')
        $skip -ge 0 -and $guard -gt $skip
    }
    foreach ($profileFunction in @(@('Campaigns\HybridJoin\SmartEndpointMigration.Campaign.HybridJoin.psm1','Get-HybridJoinCentralLogBucket'),@('Campaigns\Windows11Upgrade\SmartEndpointMigration.Campaign.Windows11Upgrade.psm1','Get-Windows11CentralLogBucket'))) {
        $profileAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:PackageRoot $profileFunction[0]), [ref]$null, [ref]$null)
        . ([scriptblock]::Create($profileAst.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $profileFunction[1] }, $true).Extent.Text))
    }
    Test-Step 'central logs: skipped virtual machines go to Skipped, run guards to TargetBusy (not Errors)' {
        (Get-HybridJoinCentralLogBucket -Result ([pscustomobject]@{ Status = 'SKIPPED_VIRTUAL_MACHINE' })) -eq 'Skipped' -and (Get-HybridJoinCentralLogBucket -Result ([pscustomobject]@{ Status = 'RUN_GUARD_ACTIVE' })) -eq 'TargetBusy' -and (Get-Windows11CentralLogBucket -Result ([pscustomobject]@{ RemoteStatus = 'SKIPPED_VIRTUAL_MACHINE'; Status = 'SKIPPED_VIRTUAL_MACHINE'; ExitCode = '0' })) -eq 'Skipped' -and (Get-Windows11CentralLogBucket -Result ([pscustomobject]@{ RemoteStatus = 'RUN_GUARD_ACTIVE'; ExitCode = '3' })) -eq 'TargetBusy'
    }
    Test-Step 'Run engine: a technician back-off skip shows the previous status, the next try time and the bypass' {
        $engineText -match "previous status \{2\};\{3\} a run mode 'ignore run guard' bypasses it" -and $engineText -match "next try after \{0:yyyy-MM-dd HH:mm\} \(local\)"
    }
    Test-Step 'Run engine: a stop requested during an inventory refresh does not start the next refresh' {
        $engineText -match '(?s)foreach \(\$kind in @\(\$script:CampaignProfile\.InventoryKinds\)\) \{\s*# A stop requested[^\n]*\n\s*if \(\(Get-SemtCancellationState\)\.Requested\)'
    }
    Test-Step 'target launches use RemoteSigned, never Bypass (PsExec, 64-bit relaunch, retry-after-reboot tasks)' {
        $remoteText = [IO.File]::ReadAllText((Join-Path $script:PackageRoot 'Modules\SmartEndpointMigration.Core\Functions\Semt.Remote.ps1'))
        $remoteText -match "'powershell\.exe','-NoProfile','-ExecutionPolicy','RemoteSigned','-File'" -and $remoteText -notmatch "'-ExecutionPolicy','Bypass'" -and $hjEndpointText -notmatch '(?i)ExecutionPolicy"?,?\s*"?Bypass' -and $w11EndpointText -notmatch '(?i)ExecutionPolicy Bypass' -and $hjEndpointText -match '-ExecutionPolicy RemoteSigned -File' -and $w11EndpointText -match '-ExecutionPolicy RemoteSigned -File'
    }
    # Retry-after-reboot task: Windows PowerShell 5.1 split a /TR value that starts and ends with a quote
    # ("C:\Program Files\..." became two arguments) and schtasks.exe failed with "Invalid argument/option".
    foreach ($endpointText in @($hjEndpointText,$w11EndpointText)) {
        $quotingAst = [System.Management.Automation.Language.Parser]::ParseInput($endpointText, [ref]$null, [ref]$null)
        . ([scriptblock]::Create($quotingAst.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'ConvertTo-SemtCommandLineArgument' }, $true).Extent.Text))
        Test-Step 'endpoint native arguments: quotes and spaces are escaped for schtasks.exe /TR' {
            (ConvertTo-SemtCommandLineArgument -Value '"C:\P\ps.exe" -File "C:\Program Files\r.ps1"') -eq '"\"C:\P\ps.exe\" -File \"C:\Program Files\r.ps1\""' -and (ConvertTo-SemtCommandLineArgument -Value '/Create') -eq '/Create' -and (ConvertTo-SemtCommandLineArgument -Value 'C:\a b\') -eq '"C:\a b\\"' -and (ConvertTo-SemtCommandLineArgument -Value '') -eq '""'
        }
    }
    Test-Step 'endpoints create the retry-after-reboot task through Invoke-SemtSchtasks (no direct schtasks.exe /Create)' {
        $hjEndpointText -match "Invoke-SemtSchtasks -Arguments @\('/Create','/TN',\`$script:RetryAfterRebootTaskName" -and $w11EndpointText -match 'Invoke-SemtSchtasks -Arguments \$createArgs' -and $hjEndpointText -notmatch '& schtasks\.exe /Create' -and $w11EndpointText -notmatch '& schtasks\.exe @createArgs'
    }
    Test-Step 'Run engine: report names omit the Wave name; a cycle HTML report failure never fails the Run' {
        $engineText -match "'SEMT_\{0\}_cycle\{1\}_\{2\}\.csv' -f \`$Campaign,\`$cycle" -and $engineText -match "'SEMT_\{0\}_\{1\}_final\.html' -f \`$Campaign,\`$script:RunId" -and $engineText -match 'HTML report update failed' -and $engineText -match 'Final report failed'
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
        for ($i = 0; $i -lt 30; $i++) { Start-Sleep -Seconds 1; if (Test-Path (Join-Path $targetRoot 'State\TargetOperation.json.txt')) { $held = $true; break } }
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
    Write-SemtTargetJson -Path (Join-Path $targetRoot 'State\TargetPendingOperation.json.txt') -Data ([pscustomobject]@{ Campaign = 'Windows11Upgrade'; Reason = 'old'; ExpiresUtc = (Get-Date).ToUniversalTime().AddHours(-1).ToString('o'); LastBootUtc = ''; ClearedByReboot = $true })
    $expired = Enter-SemtTargetOperationGuard -Campaign HybridJoin -CampaignDataRoot $hjData -RunId R-HJ -CorrelationId C-HJ3
    Test-Step 'expired pending marker no longer blocks' { $expired.Acquired }
    Exit-SemtTargetOperationGuard -CampaignDataRoot $hjData
    Set-Content -LiteralPath (Join-Path $targetRoot 'State\TargetPendingOperation.json.txt') -Value '{ not json' -Encoding ASCII
    $unreadable = Enter-SemtTargetOperationGuard -Campaign HybridJoin -CampaignDataRoot $hjData -RunId R-HJ -CorrelationId C-HJ4
    Test-Step 'unreadable pending marker gives an explicit indeterminate status' { -not $unreadable.Acquired -and $unreadable.Status -eq 'TARGET_STATE_UNREADABLE' }
    Exit-SemtTargetOperationGuard -CampaignDataRoot $hjData
    Remove-Item -LiteralPath (Join-Path $targetRoot 'State\TargetPendingOperation.json.txt') -Force
    Test-Step 'attempt results keep the newest 50 files' {
        for ($i = 0; $i -lt 55; $i++) { [void](Write-SemtAttemptResult -CampaignDataRoot $hjData -Data @{ CorrelationId = ('c-{0:000}' -f $i); State = 'Completed' }) }
        @(Get-ChildItem -LiteralPath (Join-Path $hjData 'State\Results') -Filter '*.json.txt').Count -eq 50
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

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCrCaJi+bKEg1e6
# hr+kQqCw6VV43umluImHm30qxIvQWKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBHLNBJyT9N7+EX7mnQp0r4
# JV2ORQfweCJvKVUpuUgZBTANBgkqhkiG9w0BAQEFAASCAYAxrmwvVVwBtn0hWlLG
# zCNhDDm9W/N2ZS9BKHRwpRx/SV81MHgNBzvG9OAdsZroCrvBUSykXFiEwTru8bDr
# WFBaxHtmKJB0Zi9aLELQ/TmX6z+b2hO6zZsl1mw9r5Rhkf7/aaFQDRsQks3AKYLW
# u+aUmoA2ALVaQOfGAnj9lv8qYqA/7Gyujr0bexrCpX2e+S7/2ROMWj5gJpc02W5i
# yFiZFYrmg9REWZ2XV/gC1N6nCyGW+CVqBudRfLKQFV3mI+q2V7DLnz9U/woS1nto
# HgVuAly9gHM8VHNykDqe84uBfjfdIv3QT5cx6Z+8z8PXSIcNg9BcckqxFFvh3XVJ
# fR0lPfbmf2gW60JNILv8ty1xRwMAe9vXtLuMzde9c8jD+k6EAy8V7Igx8G+A9q4N
# LpLBaN0QWyWdHjY3XmrWbfrImYVT3V0ROGgCUmSxyS8SSzd+7VZp9/G1DVs93jrg
# WnHjWwbedWKxn+/kWV2lGzMp8Ow+lzz3Qtp+HIRwNJkLDeM=
# SIG # End signature block
