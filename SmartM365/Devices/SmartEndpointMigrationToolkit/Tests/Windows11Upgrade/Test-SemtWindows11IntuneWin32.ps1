<#
.SYNOPSIS
    Offline synthetic tests for the Smart Endpoint Migration Toolkit Windows 11 Upgrade Intune Win32 package.
.DESCRIPTION
    Validates the Windows11Upgrade IntuneWin32 builder, installer, runner, detection template,
    setup media integrity helper, publisher metadata helpers, setup media manifest generator and
    root CMD without network, Microsoft Graph, Intune, IntuneWinAppUtil, scheduled tasks or
    setup.exe. Static checks use the PowerShell AST/tokenizer; synthetic checks run in a temp folder
    with a tiny fake media tree.
    IntuneWin32 assertions are ported from SmartM365/Devices/Windows11UpgradeToolkit/Tests/Test-SmartM365Windows11RunGuardAndCmd.ps1 1.7.11 (new file, not a copy).
    Prints one PASS/FAIL line per assertion and a final "Summary: Passed=N; Failed=M" line.
    Exits 1 when an assertion fails. Runs on Windows PowerShell 5.1 and PowerShell 7.
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

function Write-TestResult {
    param([bool]$Condition, [string]$Message, [string]$Detail = '')
    if ($Condition) {
        $script:Passed++
        Write-Host ("PASS: {0}" -f $Message)
    }
    else {
        $script:Failed++
        if ([string]::IsNullOrWhiteSpace($Detail)) { Write-Host ("FAIL: {0}" -f $Message) }
        else { Write-Host ("FAIL: {0} ({1})" -f $Message,$Detail) }
    }
}

function Test-Step {
    param([string]$Message, [scriptblock]$Check)
    try {
        $value = & $Check
        Write-TestResult -Condition ([bool]$value) -Message $Message
    }
    catch {
        Write-TestResult -Condition $false -Message $Message -Detail $_.Exception.Message
    }
}

function Get-ScriptAst {
    param([string]$Path)
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    return [pscustomobject]@{ Ast = $ast; Tokens = $tokens; Errors = $errors }
}

function Import-ScriptFunction {
    param([string]$Path, [string]$Name)
    $parsed = Get-ScriptAst -Path $Path
    if ($parsed.Errors.Count -gt 0) { throw ($parsed.Errors | ForEach-Object Message | Out-String) }
    $functionAst = $parsed.Ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name }, $true)
    if (-not $functionAst) { throw "Function not found: $Name in $Path" }
    Set-Item -Path ("Function:\global:{0}" -f $Name) -Value $functionAst.Body.GetScriptBlock()
}

function Get-CodeWithoutComments {
    param([string]$Path)
    $parsed = Get-ScriptAst -Path $Path
    $text = [System.IO.File]::ReadAllText($Path)
    $builder = New-Object System.Text.StringBuilder
    $position = 0
    foreach ($token in @($parsed.Tokens | Where-Object { $_.Kind -eq [System.Management.Automation.Language.TokenKind]::Comment })) {
        $start = $token.Extent.StartOffset
        if ($start -gt $position) { [void]$builder.Append($text.Substring($position, $start - $position)) }
        [void]$builder.Append(' ')
        $position = $token.Extent.EndOffset
    }
    if ($position -lt $text.Length) { [void]$builder.Append($text.Substring($position)) }
    return $builder.ToString()
}

function Get-AssignedStringValue {
    param([string]$Path, [string]$VariableName)
    $parsed = Get-ScriptAst -Path $Path
    $assignment = $parsed.Ast.Find({
            param($node)
            $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $node.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
            $node.Left.VariablePath.UserPath -eq $VariableName -and
            $node.Right -is [System.Management.Automation.Language.CommandExpressionAst] -and
            $node.Right.Expression -is [System.Management.Automation.Language.StringConstantExpressionAst]
        }, $true)
    if (-not $assignment) { throw "String assignment not found: `$$VariableName in $Path" }
    return [string]$assignment.Right.Expression.Value
}

function Get-ParameterDefault {
    param([string]$Path, [string]$ParameterName)
    $parsed = Get-ScriptAst -Path $Path
    $parameter = @($parsed.Ast.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq $ParameterName })
    if ($parameter.Count -eq 0 -or $null -eq $parameter[0].DefaultValue) { return $null }
    return $parameter[0].DefaultValue.SafeGetValue()
}

function Get-ArrayLiteralStrings {
    param([string]$Path, [string]$VariableName)
    $parsed = Get-ScriptAst -Path $Path
    $assignment = $parsed.Ast.Find({
            param($node)
            $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $node.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
            $node.Left.VariablePath.UserPath -eq $VariableName
        }, $true)
    if (-not $assignment) { throw "Assignment not found: `$$VariableName in $Path" }
    $strings = $assignment.Right.FindAll({ param($node) $node -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)
    return @($strings | ForEach-Object { [string]$_.Value })
}

$testsDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$productRoot = Split-Path -Parent (Split-Path -Parent $testsDir)
$campaignRoot = Join-Path $productRoot 'Campaigns\Windows11Upgrade'
$intuneRoot = Join-Path $campaignRoot 'IntuneWin32'
$sourceRoot = Join-Path $intuneRoot 'Source'
$builder = Join-Path $intuneRoot 'SmartM365-EndpointMigration-Build-Windows11IntunePackage.ps1'
$builderAll = Join-Path $intuneRoot 'SmartM365-EndpointMigration-Build-AllWindows11IntunePackages.ps1'
$publisher = Join-Path $intuneRoot 'SmartM365-EndpointMigration-Publish-Windows11IntuneApp.ps1'
$publisherAll = Join-Path $intuneRoot 'SmartM365-EndpointMigration-Publish-AllWindows11IntuneApps.ps1'
$installer = Join-Path $sourceRoot 'SmartM365-EndpointMigration-Install.ps1'
$runner = Join-Path $sourceRoot 'SmartM365-EndpointMigration-Run-IntuneUpgrade.ps1'
$detectTemplate = Join-Path $sourceRoot 'SmartM365-EndpointMigration-Detect-Template.ps1'
$integrityHelper = Join-Path $sourceRoot 'SmartM365-EndpointMigration-SetupMediaIntegrity.ps1'
$endpoint = Join-Path $campaignRoot 'Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'
$manifestGenerator = Join-Path $campaignRoot 'Scripts\SmartM365-EndpointMigration-New-SetupMediaManifest.ps1'
$rootCmd = Join-Path $productRoot 'New-SetupMediaManifest.cmd'
$intuneReadme = Join-Path $intuneRoot 'README.md'
$setupSourceReadme = Join-Path $campaignRoot 'SetupSource\README.md'
$gatesReadme = Join-Path $campaignRoot 'SetupSourceGates\README.md'
$sourceMapTemplate = Join-Path $campaignRoot 'Templates\SetupSourceMap.csv.template'
$provenancePath = Join-Path $productRoot 'Docs\PROVENANCE.json.txt'

$newRegistryRoot = 'SOFTWARE\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade\IntunePackages'
$newDataRoot = 'C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade'
$newTaskName = 'SmartM365 SEMT Windows 11 Upgrade - Intune'
$newPackagePrefix = 'SmartM365-SEMT-Windows11Upgrade-'

$changedScripts = @($builder, $builderAll, $publisher, $publisherAll, $installer, $runner, $detectTemplate, $integrityHelper, $manifestGenerator)
$allScopeFiles = $changedScripts + @($rootCmd, $intuneReadme, $setupSourceReadme, $gatesReadme, $sourceMapTemplate, $MyInvocation.MyCommand.Path)

$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('SEMT-W11-IntuneWin32-Test-{0}' -f [guid]::NewGuid().ToString('N'))
Write-Host ("Engine: {0} {1}; TestRoot={2}" -f $PSVersionTable.PSEdition,$PSVersionTable.PSVersion,$testRoot)

try {
    # ---------- Static: parse, headers, ASCII ----------
    $provenance = Get-Content -LiteralPath $provenancePath -Raw | ConvertFrom-Json
    foreach ($path in $changedScripts + @($MyInvocation.MyCommand.Path)) {
        $leaf = Split-Path -Leaf $path
        Test-Step "parse without errors: $leaf" { (Get-ScriptAst -Path $path).Errors.Count -eq 0 }
        $text = Get-Content -LiteralPath $path -Raw
        Test-Step "comment help has SYNOPSIS, DESCRIPTION, VERSION 1.0.0 and NOTES Author: $leaf" {
            ($text -match '(?s)^\s*<#.*?\.SYNOPSIS.*?\.DESCRIPTION.*?\.VERSION\s+1\.0\.0\s.*?\.NOTES\s+Author: https://github\.com/khda79/workplacecloudhub\.com.*?#>')
        }
        Test-Step "no signature block: $leaf" { -not $text.Contains('SIG # Begin ' + 'signature block') }
        $versionMatch = [regex]::Match($text, '(?m)^\s*\$(script:)?ScriptVersion\s*=\s*''(?<V>[^'']+)''')
        if ($versionMatch.Success) { Test-Step "ScriptVersion equals .VERSION: $leaf" { $versionMatch.Groups['V'].Value -eq '1.0.0' } }
        $entry = @($provenance.Files | Where-Object { ($_.Destination -replace '/', '\') -like ('*\' + $leaf) })
        if ($entry.Count -gt 0) {
            $expected = 'Provenance: copied from {0} {1} (see Docs\PROVENANCE.json.txt).' -f $entry[0].Source,$entry[0].SourceVersion
            Test-Step "DESCRIPTION carries the provenance line: $leaf" { $text.Contains($expected) }
        }
    }
    foreach ($path in $allScopeFiles) {
        $leaf = Split-Path -Leaf $path
        Test-Step "ASCII only: $leaf" { @([System.IO.File]::ReadAllBytes($path) | Where-Object { $_ -gt 127 }).Count -eq 0 }
    }

    # ---------- Static: identities ----------
    $oldIdentityPattern = 'SmartM365\\Windows11UpgradeToolkit\\IntunePackages|SmartM365 Windows 11 Upgrade Toolkit - Intune|SmartM365-Windows11UpgradeToolkit-|ProgramData\\SmartM365\\Windows11UpgradeToolkit|Windows11UpgradeToolkit-\$Language|W11UT'
    $oldScriptNamePattern = '(?<![\w-])(Install|Run-IntuneUpgrade|Detect-Template)\.ps1|SmartM365-Invoke-Windows11UpgradeRepair(\.ps1|_)|(?<![\w-])SmartM365-SetupMediaIntegrity\.ps1|New-SmartM365SetupMediaManifest\.ps1|Build-SmartM365Windows11IntunePackage\.ps1|Build-AllSmartM365Windows11IntunePackages\.ps1|Publish-SmartM365Windows11IntuneApp\.ps1|Publish-AllSmartM365Windows11IntuneApp\.ps1'
    foreach ($path in $changedScripts) {
        $leaf = Split-Path -Leaf $path
        $code = Get-CodeWithoutComments -Path $path
        # Read-only references to the previous toolkit are allowed only where its media is resolved:
        # the shared media resolution in the helper and the cache-only requirement rule.
        $allowedLegacyReads = @{
            'SmartM365-EndpointMigration-SetupMediaIntegrity.ps1' = @('SOFTWARE\SmartM365\Windows11UpgradeToolkit\IntunePackages', 'C:\ProgramData\SmartM365\Windows11UpgradeToolkit')
            'SmartM365-EndpointMigration-Publish-Windows11IntuneApp.ps1' = @('C:\ProgramData\SmartM365\Windows11UpgradeToolkit\SetupMedia')
        }
        if ($allowedLegacyReads.ContainsKey($leaf)) { foreach ($allowed in $allowedLegacyReads[$leaf]) { $code = $code.Replace($allowed, '<allowed-legacy-read>') } }
        Test-Step "no old registry root, task name, data root or PackageId prefix outside comments (allowed legacy reads excepted): $leaf" { $code -notmatch $oldIdentityPattern }
        Test-Step "no old script file name outside comments: $leaf" { $code -notmatch $oldScriptNamePattern }
    }
    foreach ($path in @($rootCmd, $intuneReadme, $setupSourceReadme, $gatesReadme)) {
        $leaf = Split-Path -Leaf $path
        Test-Step "no old script file name: $leaf" { (Get-Content -LiteralPath $path -Raw) -notmatch $oldScriptNamePattern }
    }

    Test-Step 'installer registry root is the new 64-bit SEMT root' { (Get-AssignedStringValue -Path $installer -VariableName 'registrySubKeyRoot') -eq $newRegistryRoot }
    Test-Step 'runner registry root is the new 64-bit SEMT root' { (Get-AssignedStringValue -Path $runner -VariableName 'registrySubKeyRoot') -eq $newRegistryRoot }
    Test-Step 'detection template registry key is the new SEMT root' { (Get-Content -LiteralPath $detectTemplate -Raw).Contains(('"{0}\$packageId"' -f $newRegistryRoot)) }
    foreach ($path in @($installer, $runner, $detectTemplate)) {
        $leaf = Split-Path -Leaf $path
        $code = Get-CodeWithoutComments -Path $path
        Test-Step "registry access uses RegistryView Registry64 only: $leaf" { ($code -match '\[Microsoft\.Win32\.RegistryView\]::Registry64') -and ($code -notmatch 'Registry32|RegistryView\]::Default|HKLM:\\SOFTWARE') }
    }
    foreach ($path in @($installer, $runner)) {
        $leaf = Split-Path -Leaf $path
        Test-Step "DataRoot default is the new SEMT data root: $leaf" { (Get-ParameterDefault -Path $path -ParameterName 'DataRoot') -eq $newDataRoot }
        Test-Step "TaskName default is the new SEMT task: $leaf" { (Get-ParameterDefault -Path $path -ParameterName 'TaskName') -eq $newTaskName }
    }
    $installerText = Get-Content -LiteralPath $installer -Raw
    $runnerText = Get-Content -LiteralPath $runner -Raw
    $detectText = Get-Content -LiteralPath $detectTemplate -Raw
    $builderText = Get-Content -LiteralPath $builder -Raw
    $publisherText = Get-Content -LiteralPath $publisher -Raw
    Test-Step 'installer lock owner has no old product name' { $installerText.Contains("Owner = 'SEMT_Windows11Upgrade_IntuneInstall'") }
    Test-Step 'installer and runner log names have no old product name' { $installerText.Contains("'SEMT-Intune-Install.log'") -and $installerText.Contains("'SEMT-Intune-Install-Robocopy.log'") -and $runnerText.Contains("'SEMT-Intune-Runner.log'") }
    Test-Step 'installer copies the new endpoint, runner and helper file names' {
        (Get-AssignedStringValue -Path $installer -VariableName 'endpointScriptFileName') -eq 'SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1' -and
        (Get-AssignedStringValue -Path $installer -VariableName 'runnerScriptFileName') -eq 'SmartM365-EndpointMigration-Run-IntuneUpgrade.ps1' -and
        (Get-AssignedStringValue -Path $installer -VariableName 'integrityHelperFileName') -eq 'SmartM365-EndpointMigration-SetupMediaIntegrity.ps1'
    }
    Test-Step 'runner launches the new endpoint file and reads the new helper from its Program Files folder' { $runnerText.Contains("Join-Path `$programRoot 'SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'") -and $runnerText.Contains("Join-Path `$programIntuneRoot 'SmartM365-EndpointMigration-SetupMediaIntegrity.ps1'") -and -not $runnerText.Contains("Join-Path `$DataRoot 'SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'") }
    Test-Step 'runner looks for the new endpoint log names' { $runnerText.Contains("-Filter 'SEMT_Windows11Upgrade_*.log'") }

    # ---------- Static: installer behavior kept ----------
    Test-Step 'installer validates the complete destination cache' { $installerText -match 'Test-SetupCacheReady[\s\S]*?Test-SmartM365SetupMediaIntegrity' }
    Test-Step 'installer writes Installed only after final cache validation' { $installerText.LastIndexOf("Set-PackageDetectionState -InstallState 'Installed'") -gt $installerText.LastIndexOf('Test-SetupCacheReady -Path $targetMediaRoot') }
    Test-Step 'installer marks an already-Windows 11 device Installed' { $installerText -match "Device is already Windows 11[\s\S]*?Set-PackageDetectionState -InstallState 'Installed'" }
    Test-Step 'installer clears stale repair metadata when Installed' { $installerText -match 'if \(\$InstallState -eq ''Installed''\)[\s\S]*?RepairReason[\s\S]*?RepairRequiredUtc' }
    Test-Step 'installer evaluates Windows 11 before media helper validation' { $installerText.IndexOf("if ((Get-OsFamily) -eq 'Windows11')") -lt $installerText.IndexOf('if (-not (Test-Path -LiteralPath $integrityHelperPackagePath') }
    Test-Step 'installer exits 1618 when the cache lock is held' { $installerText -match 'exit 1618' }
    Test-Step 'installer task retries every 2 hours for 30 days as SYSTEM' { ($installerText -match 'AddHours\(2\) -RepetitionInterval \(New-TimeSpan -Hours 2\) -RepetitionDuration \(New-TimeSpan -Days 30\)') -and ($installerText -match "New-ScheduledTaskPrincipal -UserId 'SYSTEM'") }
    Test-Step 'installer task argument runs the runner with DataRoot and TaskName' { $installerText.Contains('''-NoProfile -ExecutionPolicy Bypass -File "{0}" -DataRoot "{1}" -TaskName "{2}"'' -f $runner,$DataRoot,$TaskName') }

    Import-ScriptFunction -Path $installer -Name 'Set-PackageDetectionState'
    $global:SyntheticRegistryWrites = New-Object System.Collections.ArrayList
    function global:Set-Registry64String { param([string]$SubKey, [string]$Name, [string]$Value); [void]$global:SyntheticRegistryWrites.Add([pscustomobject]@{ SubKey = $SubKey; Name = $Name; Value = $Value }) }
    function global:Write-InstallLog { param([string]$Message, [string]$Level = 'INFO') }
    $manifest = [pscustomobject]@{ PackageId = 'SmartM365-SEMT-Windows11Upgrade-Win11-fr-FR'; PackageVersion = '1.0.0'; Language = 'fr-FR'; MediaId = 'Win11'; SetupCacheFolder = 'Win11-fr-FR' }
    $packageMode = 'WithMedia'
    $registrySubKey = '{0}\{1}' -f (Get-AssignedStringValue -Path $installer -VariableName 'registrySubKeyRoot'),$manifest.PackageId
    Set-PackageDetectionState -InstallState 'Installed'
    $writtenNames = @($global:SyntheticRegistryWrites | ForEach-Object Name | Select-Object -Unique)
    Test-Step 'installer detection state keeps the same registry value names' { @('PackageId','PackageVersion','Language','MediaId','SetupCacheFolder','PackageMode','InstallState','InstalledUtc','RepairReason','RepairRequiredUtc' | Where-Object { $writtenNames -notcontains $_ }).Count -eq 0 }
    Test-Step 'installer detection state is written under the new package key' { @($global:SyntheticRegistryWrites | Where-Object { $_.SubKey -ne ($newRegistryRoot + '\SmartM365-SEMT-Windows11Upgrade-Win11-fr-FR') }).Count -eq 0 }

    # ---------- Runner ----------
    $endpointArgs = @(Get-ArrayLiteralStrings -Path $runner -VariableName 'endpointArgs')
    $endpointArgNames = @($endpointArgs | Where-Object { $_ -like '-*' })
    foreach ($required in @('-CorrelationId','-DataRoot','-SetupCacheRoot','-ReuseLegacySetupMedia')) {
        Test-Step "runner passes $required to the endpoint" { $endpointArgNames -contains $required }
    }
    $expectedArgNames = @('-RunGuardHours','-DirectSetupUpgrade','-SkipVirtualMachines','-AllowReboot','-AllowSetupCompletionRebootWhenNoUser','-AllowSetupProfileRepair','-AllowDiskCleanup','-SetupExecutionMode','-SetupMediaId','-SetupLanguage','-SetupDynamicUpdate','-SetupCacheRoot','-SetupProcessHeartbeatSeconds','-SetupProcessTimeoutMinutes','-SetupMediaCopyTimeoutMinutes','-ForceRequiredRebootWhenUptimeOverDays','-DataRoot','-CorrelationId','-ReuseLegacySetupMedia','-SetupMediaPath')
    Test-Step 'runner keeps the existing switches, adds -SetupMediaPath and no implicit reboot switch' { (@(Compare-Object -ReferenceObject $expectedArgNames -DifferenceObject $endpointArgNames).Count -eq 0) }
    $endpointParameters = @((Get-ScriptAst -Path $endpoint).Ast.ParamBlock.Parameters | ForEach-Object { '-' + $_.Name.VariablePath.UserPath })
    Test-Step 'every runner argument is a parameter of the new endpoint script' { @($endpointArgNames | Where-Object { $endpointParameters -notcontains $_ }).Count -eq 0 }
    Test-Step 'runner creates a new GUID CorrelationId per run' { $runnerText -match '\$correlationId = \[guid\]::NewGuid\(\)\.ToString\(''N''\)' -and $runnerText.Contains("'-CorrelationId', `$correlationId") }
    Test-Step 'runner passes the runner DataRoot and cache root values' { $runnerText.Contains("'-DataRoot', `$DataRoot") -and $runnerText.Contains("'-SetupCacheRoot', `$setupCacheRoot") -and $runnerText.Contains("`$setupCacheRoot = Join-Path `$DataRoot 'SetupMedia'") }
    Test-Step 'runner marks an invalid setup cache as RepairRequired' { $runnerText -match "InstallState\s+-Value\s+'RepairRequired'" }
    Test-Step 'runner evaluates Windows 11 before endpoint and cache validation' { $runnerText.IndexOf("if (`$osFamily -eq 'Windows11')") -lt $runnerText.IndexOf('if (-not (Test-Path -LiteralPath $endpointScript') }
    Test-Step 'runner performs log maintenance before and after the endpoint run' { ([regex]::Matches($runnerText, "Invoke-RunnerLogMaintenance -Phase '(Startup|PostRun)'").Count) -eq 2 }
    Test-Step 'runner defaults keep 7 days and 200 endpoint log files' { ($runnerText -match '\$LogRetentionDays\s*=\s*7') -and ($runnerText -match '\$MaxEndpointLogFiles\s*=\s*200') }

    foreach ($functionName in @('Invoke-IntuneLogMaintenance','Set-PackageInstalledForWindows11','Get-EndpointAttemptOutcome')) { Import-ScriptFunction -Path $runner -Name $functionName }
    $global:SyntheticRegistryWrites.Clear()
    function global:Write-RunnerLog { param([string]$Message, [string]$Level = 'INFO') }
    $registrySubKeyRoot = Get-AssignedStringValue -Path $runner -VariableName 'registrySubKeyRoot'
    Set-PackageInstalledForWindows11 -Manifest ([pscustomobject]@{ PackageId = 'Synthetic-W11'; PackageVersion = '9.9.9' })
    Test-Step 'Windows 11 runner state writes InstallState Installed' { (($global:SyntheticRegistryWrites | Where-Object Name -eq 'InstallState' | Select-Object -Last 1).Value) -eq 'Installed' }
    Test-Step 'Windows 11 runner state records CompletionReason AlreadyWindows11' { (($global:SyntheticRegistryWrites | Where-Object Name -eq 'CompletionReason' | Select-Object -Last 1).Value) -eq 'AlreadyWindows11' }
    Test-Step 'Windows 11 runner state clears the repair timestamp' { (($global:SyntheticRegistryWrites | Where-Object Name -eq 'RepairRequiredUtc' | Select-Object -Last 1).Value) -eq '' }
    Test-Step 'Windows 11 runner state uses the new registry root' { @($global:SyntheticRegistryWrites | Where-Object { $_.SubKey -ne ($newRegistryRoot + '\Synthetic-W11') }).Count -eq 0 }

    $logRoot = Join-Path $testRoot 'intune-log-retention'
    New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
    $now = Get-Date
    foreach ($index in 0..7) { $p = Join-Path $logRoot ('Endpoint_recent_{0:D2}_stdout.log' -f $index); Set-Content -LiteralPath $p -Value $index -Encoding ASCII; (Get-Item -LiteralPath $p).LastWriteTime = $now.AddHours(-1 * $index) }
    foreach ($index in 0..1) { $p = Join-Path $logRoot ('Endpoint_old_{0:D2}_stderr.log' -f $index); Set-Content -LiteralPath $p -Value $index -Encoding ASCII; (Get-Item -LiteralPath $p).LastWriteTime = $now.AddDays(-8).AddHours(-1 * $index) }
    foreach ($name in @('SEMT-Intune-Runner.log','SEMT-Intune-Install.log','SEMT-Intune-Install-Robocopy.log')) { $p = Join-Path $logRoot $name; Set-Content -LiteralPath $p -Value $name -Encoding ASCII; (Get-Item -LiteralPath $p).LastWriteTime = $now.AddDays(-30) }
    $maintenance = Invoke-IntuneLogMaintenance -Path $logRoot -RetentionDays 7 -MaximumEndpointFiles 5
    Test-Step 'log maintenance scans only timestamped endpoint logs' { $maintenance.Scanned -eq 10 }
    Test-Step 'log maintenance removes endpoint logs older than 7 days and caps the count' { $maintenance.RemovedByAge -eq 2 -and $maintenance.RemovedByCount -eq 3 -and $maintenance.Remaining -eq 5 }
    Test-Step 'log maintenance preserves persistent SEMT logs' { @('SEMT-Intune-Runner.log','SEMT-Intune-Install.log','SEMT-Intune-Install-Robocopy.log' | Where-Object { -not (Test-Path -LiteralPath (Join-Path $logRoot $_)) }).Count -eq 0 }

    $outcomeRoot = Join-Path $testRoot 'outcome'
    $resultsRoot = Join-Path $outcomeRoot 'State\Results'
    $lastRunPath = Join-Path $outcomeRoot 'LastRun.json'
    New-Item -ItemType Directory -Path $resultsRoot -Force | Out-Null
    $cid = [guid]::NewGuid().ToString('N')
    [pscustomobject]@{ CorrelationId = $cid; State = 'Completed'; Status = 'SETUP_UPGRADE_STARTED'; ExitCode = '0'; NextAction = 'WAIT' } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $resultsRoot "$cid.json") -Encoding ASCII
    [pscustomobject]@{ CorrelationId = 'other'; Status = 'OLD'; ExitCode = 1; NextAction = 'OLD' } | ConvertTo-Json | Set-Content -LiteralPath $lastRunPath -Encoding ASCII
    $outcome = Get-EndpointAttemptOutcome -ResultsRoot $resultsRoot -LastRunPath $lastRunPath -CorrelationId $cid
    Test-Step 'runner trusts a Completed Results\<CorrelationId>.json first' { $outcome.Source -eq 'AttemptResult' -and $outcome.Completed -and $outcome.Status -eq 'SETUP_UPGRADE_STARTED' }
    $cidRunning = [guid]::NewGuid().ToString('N')
    [pscustomobject]@{ CorrelationId = $cidRunning; State = 'Running'; Status = 'RUNNING' } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $resultsRoot "$cidRunning.json") -Encoding ASCII
    $outcome = Get-EndpointAttemptOutcome -ResultsRoot $resultsRoot -LastRunPath $lastRunPath -CorrelationId $cidRunning
    Test-Step 'runner does not treat a Running attempt as completed' { -not $outcome.Completed -and $outcome.Status -like 'ATTEMPT_NOT_COMPLETED*' }
    $cidMissing = [guid]::NewGuid().ToString('N')
    $outcome = Get-EndpointAttemptOutcome -ResultsRoot $resultsRoot -LastRunPath $lastRunPath -CorrelationId $cidMissing
    Test-Step 'runner falls back to LastRun.json and flags a foreign attempt as not completed' { $outcome.Source -eq 'LastRun' -and -not $outcome.Completed -and $outcome.Status -eq 'OLD' }
    [pscustomobject]@{ CorrelationId = $cidMissing; Status = 'COMPLIANT'; ExitCode = 0; NextAction = 'NONE' } | ConvertTo-Json | Set-Content -LiteralPath $lastRunPath -Encoding ASCII
    $outcome = Get-EndpointAttemptOutcome -ResultsRoot $resultsRoot -LastRunPath $lastRunPath -CorrelationId $cidMissing
    Test-Step 'runner accepts LastRun.json of the same CorrelationId as completed' { $outcome.Source -eq 'LastRun' -and $outcome.Completed -and $outcome.Status -eq 'COMPLIANT' }

    # ---------- Detection template ----------
    Test-Step 'detection template keeps both placeholders' { $detectText.Contains("'__PACKAGE_ID__'") -and $detectText.Contains("'__PACKAGE_VERSION__'") }
    Test-Step 'detection prioritizes Windows 11 over registry state' { $detectText.IndexOf('Device is already Windows 11') -lt $detectText.IndexOf('Get-Registry64PackageState -SubKey') }
    Test-Step 'detection rejects RepairRequired state on Windows 10' { $detectText -match "@\('Installed', 'AlreadyWindows11'\)" }

    # ---------- Integrity helper ----------
    . $integrityHelper
    $mediaTest = Join-Path $testRoot 'setup-media-integrity'
    New-Item -ItemType Directory -Path (Join-Path $mediaTest 'sources'),(Join-Path $mediaTest 'boot') -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $mediaTest 'setup.exe') -Value 'synthetic setup' -Encoding ASCII
    Set-Content -LiteralPath (Join-Path $mediaTest 'sources\install.wim') -Value 'synthetic image' -Encoding ASCII
    Set-Content -LiteralPath (Join-Path $mediaTest 'boot\bcd') -Value 'synthetic boot data' -Encoding ASCII
    $rows = foreach ($relative in @('setup.exe','sources\install.wim','boot\bcd')) {
        $item = Get-Item -LiteralPath (Join-Path $mediaTest $relative)
        [pscustomobject]@{ RelativePath = $relative; Length = [int64]$item.Length; SHA256 = (Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256).Hash }
    }
    $rows | Export-Csv -LiteralPath (Join-Path $mediaTest 'SmartM365-SetupMediaManifest.sha256.csv') -NoTypeInformation -Encoding UTF8
    Test-Step 'integrity helper validates every manifest row' { (Test-SmartM365SetupMediaIntegrity -MediaRoot $mediaTest).Files -eq 3 }
    Remove-Item -LiteralPath (Join-Path $mediaTest 'boot\bcd') -Force
    $rejected = $false
    try { Test-SmartM365SetupMediaIntegrity -MediaRoot $mediaTest | Out-Null } catch { $rejected = $_.Exception.Message -match 'File missing.*boot\\bcd' }
    Test-Step 'integrity helper rejects a missing boot BCD file' { $rejected }
    Set-Content -LiteralPath (Join-Path $mediaTest 'boot\bcd') -Value 'synthetic boot data' -Encoding ASCII
    Add-Content -LiteralPath (Join-Path $mediaTest 'sources\install.wim') -Value 'corruption' -Encoding ASCII
    $rejected = $false
    try { Test-SmartM365SetupMediaIntegrity -MediaRoot $mediaTest | Out-Null } catch { $rejected = $_.Exception.Message -match '(Length|SHA256) mismatch.*sources\\install\.wim' }
    Test-Step 'integrity helper rejects a changed install image' { $rejected }
    Test-Step 'integrity helper hashes through long literal paths' { (ConvertTo-SmartM365LongLiteralPath -Path 'C:\Media\setup.exe') -eq '\\?\C:\Media\setup.exe' -and (ConvertTo-SmartM365LongLiteralPath -Path '\\srv\share\setup.exe') -eq '\\?\UNC\srv\share\setup.exe' }

    # ---------- Manifest generator ----------
    $generatorText = Get-Content -LiteralPath $manifestGenerator -Raw
    Test-Step 'manifest generator keeps SmartM365-SetupMediaManifest.sha256.csv' { (Get-ParameterDefault -Path $manifestGenerator -ParameterName 'ManifestFileName') -eq 'SmartM365-SetupMediaManifest.sha256.csv' }
    Import-ScriptFunction -Path $manifestGenerator -Name 'ConvertTo-LongLiteralPath'
    Test-Step 'manifest generator keeps long-path hashing' { ($generatorText -match 'OpenRead\(\(ConvertTo-LongLiteralPath') -and (ConvertTo-LongLiteralPath -Path '\\srv\share\a.wim') -eq '\\?\UNC\srv\share\a.wim' }

    # ---------- Builder staging with a fake tiny media folder, no IntuneWinAppUtil ----------
    $fakeCampaign = Join-Path $testRoot 'Product\Campaigns\Windows11Upgrade'
    $fakeIntune = Join-Path $fakeCampaign 'IntuneWin32'
    New-Item -ItemType Directory -Path (Join-Path $fakeIntune 'Source'),(Join-Path $fakeCampaign 'Scripts') -Force | Out-Null
    foreach ($file in @(Get-ChildItem -LiteralPath $intuneRoot -Filter '*.ps1' -File)) { Copy-Item -LiteralPath $file.FullName -Destination $fakeIntune }
    foreach ($file in @(Get-ChildItem -LiteralPath $sourceRoot -Filter '*.ps1' -File)) { Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $fakeIntune 'Source') }
    Copy-Item -LiteralPath $endpoint -Destination (Join-Path $fakeCampaign 'Scripts')
    Copy-Item -LiteralPath $manifestGenerator -Destination (Join-Path $fakeCampaign 'Scripts')
    $fakeMedia = Join-Path $fakeCampaign 'SetupSource\FR-fr'
    New-Item -ItemType Directory -Path (Join-Path $fakeMedia 'sources'),(Join-Path $fakeMedia 'boot') -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $fakeMedia 'setup.exe') -Value 'fake setup' -Encoding ASCII
    Set-Content -LiteralPath (Join-Path $fakeMedia 'sources\install.wim') -Value 'fake image' -Encoding ASCII
    Set-Content -LiteralPath (Join-Path $fakeMedia 'sources\lang.ini') -Value @('[Available UI Languages]','fr-FR = 3') -Encoding ASCII
    Set-Content -LiteralPath (Join-Path $fakeMedia 'boot\bcd') -Value 'fake bcd' -Encoding ASCII
    & (Join-Path $fakeCampaign 'Scripts\SmartM365-EndpointMigration-New-SetupMediaManifest.ps1') -SetupSourceRoot (Join-Path $fakeCampaign 'SetupSource') -Force *> $null
    $fakeManifest = Join-Path $fakeMedia 'SmartM365-SetupMediaManifest.sha256.csv'
    Test-Step 'manifest generator writes a manifest validated by the integrity helper' { (Test-Path -LiteralPath $fakeManifest -PathType Leaf) -and (Test-SmartM365SetupMediaIntegrity -MediaRoot $fakeMedia).Files -eq 4 }

    $prepRoot = Join-Path $testRoot 'ContentPrep'
    $buildResult = $null
    try { $buildResult = & (Join-Path $fakeIntune 'SmartM365-EndpointMigration-Build-Windows11IntunePackage.ps1') -Language fr-FR -SkipIntuneWinBuild -ContentPrepRoot $prepRoot 6> $null }
    catch { Write-TestResult -Condition $false -Message 'builder runs with -SkipIntuneWinBuild' -Detail $_.Exception.Message }
    $expectedPackageId = 'SmartM365-SEMT-Windows11Upgrade-Win11-fr-FR'
    Test-Step 'builder uses the SmartM365-SEMT-Windows11Upgrade- PackageId prefix' { $null -ne $buildResult -and $buildResult.PackageId -eq $expectedPackageId }
    Test-Step 'builder uses the new Intune app display name' { $null -ne $buildResult -and $buildResult.DisplayName -eq 'Smart Endpoint Migration Toolkit - Windows 11 Upgrade (fr-FR)' }
    Test-Step 'builder did not call IntuneWinAppUtil' { $null -ne $buildResult -and [string]::IsNullOrEmpty([string]$buildResult.IntuneWin) -and -not (Test-Path -LiteralPath $prepRoot) -and @(Get-ChildItem -LiteralPath $testRoot -Filter '*.intunewin' -Recurse -File).Count -eq 0 }
    $stagedSource = Join-Path $fakeIntune "Build\$expectedPackageId\Source"
    $stagedNames = @(Get-ChildItem -LiteralPath $stagedSource -File -ErrorAction SilentlyContinue | ForEach-Object Name | Sort-Object)
    $expectedStaged = @('PackageManifest.json','SmartM365-EndpointMigration-Install.ps1','SmartM365-EndpointMigration-Run-IntuneUpgrade.ps1','SmartM365-EndpointMigration-SetupMediaIntegrity.ps1','SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1')
    Test-Step 'builder stages exactly the new package file names' { ($stagedNames -join ',') -eq ($expectedStaged -join ',') }
    Test-Step 'builder embeds the NEW endpoint script unchanged' { (Get-FileHash -LiteralPath (Join-Path $stagedSource 'SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1')).Hash -eq (Get-FileHash -LiteralPath $endpoint).Hash }
    Test-Step 'builder stages media under SetupMedia\<MediaId>-<Language> and validates it' { (Test-SmartM365SetupMediaIntegrity -MediaRoot (Join-Path $stagedSource 'SetupMedia\Win11-fr-FR')).Files -eq 4 }
    $stagedManifest = $null
    try { $stagedManifest = Get-Content -LiteralPath (Join-Path $stagedSource 'PackageManifest.json') -Raw | ConvertFrom-Json } catch { }
    Test-Step 'package manifest records identity, cache folder and endpoint version' { $null -ne $stagedManifest -and $stagedManifest.PackageId -eq $expectedPackageId -and $stagedManifest.SetupCacheFolder -eq 'Win11-fr-FR' -and $stagedManifest.PackageMode -eq 'WithMedia' -and $stagedManifest.EndpointVersion -eq '1.0.0' -and $stagedManifest.PackageVersion -eq '1.0.0' }
    $outputDir = Join-Path $fakeIntune "Output\$expectedPackageId"
    $generatedDetect = Join-Path $outputDir 'SmartM365-EndpointMigration-Detect.ps1'
    $generatedDetectText = if (Test-Path -LiteralPath $generatedDetect) { Get-Content -LiteralPath $generatedDetect -Raw } else { '' }
    Test-Step 'generated detection script replaces both placeholders' { $generatedDetectText.Contains("`$packageId = '$expectedPackageId'") -and $generatedDetectText.Contains("`$packageVersion = '1.0.0'") -and -not $generatedDetectText.Contains('__PACKAGE_') }
    Test-Step 'generated detection script parses and reads the new registry root' { (Get-ScriptAst -Path $generatedDetect).Errors.Count -eq 0 -and $generatedDetectText.Contains($newRegistryRoot) }
    $commandsText = Get-Content -LiteralPath (Join-Path $outputDir 'Intune-App-Commands.txt') -Raw -ErrorAction SilentlyContinue
    Test-Step 'Intune command file uses the new installer name' { [string]$commandsText -match '-File \.\\SmartM365-EndpointMigration-Install\.ps1\r?\n' -and [string]$commandsText -match '-File \.\\SmartM365-EndpointMigration-Install\.ps1 -Uninstall' }
    Test-Step 'builder passes the new installer as IntuneWinAppUtil setup file' { $builderText.Contains('& $IntuneWinAppUtilPath -c $prepSource -s $installerFileName -o $prepOutput -q') -and (Get-AssignedStringValue -Path $builder -VariableName 'installerFileName') -eq 'SmartM365-EndpointMigration-Install.ps1' }
    Test-Step 'builder validates source, staged and prep media' { ([regex]::Matches($builderText, 'Test-SmartM365SetupMediaIntegrity\s+-MediaRoot').Count) -eq 3 }

    $buildAllResult = $null
    try { $buildAllResult = @(& (Join-Path $fakeIntune 'SmartM365-EndpointMigration-Build-AllWindows11IntunePackages.ps1') -SkipIntuneWinBuild -Force -ContentPrepRoot $prepRoot 6> $null) }
    catch { Write-TestResult -Condition $false -Message 'batch builder runs with -SkipIntuneWinBuild' -Detail $_.Exception.Message }
    Test-Step 'batch builder builds the FR-fr package with the new id and removes staging' { $null -ne $buildAllResult -and $buildAllResult.Count -eq 1 -and $buildAllResult[0].PackageId -eq $expectedPackageId -and -not (Test-Path -LiteralPath (Join-Path $fakeIntune "Build\$expectedPackageId")) }

    # ---------- Publisher metadata parsing (no Graph) ----------
    foreach ($functionName in @('ConvertTo-Base64Utf8','New-EndpointRequirementRule','Read-CompanionPackageMetadata','Get-LanguageFromPackageId','Set-DetectScriptPackageMetadata')) { Import-ScriptFunction -Path $publisher -Name $functionName }
    $fakeIntuneWin = Join-Path $outputDir "$expectedPackageId.intunewin"
    $companion = Read-CompanionPackageMetadata -Path $fakeIntuneWin
    Test-Step 'publisher reads package id, version, language and cache folder from generated files' { $companion.PackageId -eq $expectedPackageId -and $companion.PackageVersion -eq '1.0.0' -and $companion.Language -eq 'fr-FR' -and $companion.SetupCacheFolder -eq 'Win11-fr-FR' -and $companion.PackageMode -eq 'WithMedia' }
    Test-Step 'publisher reads the new display name from the package manifest' { $companion.DisplayName -eq 'Smart Endpoint Migration Toolkit - Windows 11 Upgrade (fr-FR)' }
    Test-Step 'publisher derives the language from the new package id' { (Get-LanguageFromPackageId -Value $expectedPackageId) -eq 'fr-FR' -and (Get-LanguageFromPackageId -Value 'SmartM365-SEMT-Windows11Upgrade-Win11-de-DE') -eq 'de-DE' }
    $patchedDetect = Set-DetectScriptPackageMetadata -ScriptContent $generatedDetectText -PackageId $expectedPackageId -PackageVersion '1.0.7'
    Test-Step 'publisher rewrites the detection package version' { $patchedDetect.Contains("`$packageVersion = '1.0.7'") -and $patchedDetect.Contains($newRegistryRoot) }
    Test-Step 'publisher default install and uninstall commands use the new installer' {
        (Get-ParameterDefault -Path $publisher -ParameterName 'InstallCommandLine') -eq 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-EndpointMigration-Install.ps1' -and
        (Get-ParameterDefault -Path $publisher -ParameterName 'UninstallCommandLine') -eq 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-EndpointMigration-Install.ps1 -Uninstall'
    }
    Test-Step 'publisher refuses package ids without the SEMT prefix' { $publisherText -match "\`$script:SemtPackageIdPrefix = 'SmartM365-SEMT-Windows11Upgrade-'" -and $publisherText -match 'if \(-not \$PackageId\.StartsWith\(\$script:SemtPackageIdPrefix' }
    Test-Step 'batch publisher skips packages without the SEMT prefix' { (Get-Content -LiteralPath $publisherAll -Raw) -match "\`$packageIdPrefix = 'SmartM365-SEMT-Windows11Upgrade-'" }
    $publisherCode = (Get-CodeWithoutComments -Path $publisher) + ' ' + (Get-CodeWithoutComments -Path $publisherAll)
    Test-Step 'publishers never call Intune assignment endpoints' { $publisherCode -notmatch '(?i)/assign|assignments' }
    Test-Step 'publisher detection rule runs in 64-bit PowerShell' { $publisherText -match "win32LobAppPowerShellScriptDetection'\s+enforceSignatureCheck = \`$false\s+runAs32Bit = \`$false" }

    $requirementRule = New-EndpointRequirementRule -Language 'fr-FR'
    $requirementScript = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String([string]$requirementRule.scriptContent))
    Test-Step 'language requirement reads InstallLanguage with InstalledUICulture fallback and no SystemLocale' { ($requirementScript -match 'InstallLanguage') -and ($requirementScript -match 'InstalledUICulture') -and ($requirementScript -notmatch 'Get-WinSystemLocale') -and -not $requirementRule.runAs32Bit }
    $mockLanguage = "function Get-ItemProperty {`r`n    [CmdletBinding()]`r`n    param([string]`$LiteralPath)`r`n    [pscustomobject]@{ InstallLanguage = '040C' }`r`n}`r`n"
    $runspace = [powershell]::Create()
    try {
        $output = @($runspace.AddScript($mockLanguage + $requirementScript).Invoke())
        Test-Step 'language requirement accepts InstallLanguage 040C as fr-FR' { -not $runspace.HadErrors -and ($output -join ',') -eq 'OK' }
    }
    finally { $runspace.Dispose() }
    $cacheRule = New-EndpointRequirementRule -Language 'fr-FR' -SetupCacheFolder 'Win11-fr-FR' -RequireSetupCache
    $cacheScript = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String([string]$cacheRule.scriptContent))
    Test-Step 'cache-only requirement checks the new data root cache and the legacy cache' { $cacheScript.Contains("'$newDataRoot\SetupMedia'") -and $cacheScript.Contains("'C:\ProgramData\SmartM365\Windows11UpgradeToolkit\SetupMedia'") }
    foreach ($requirementCase in @(
        @{ Name = 'no reusable cache'; Present = ''; Expected = 'MISSING_REUSABLE_SETUP_CACHE' },
        @{ Name = 'cache of this product'; Present = "$newDataRoot\SetupMedia\Win11-fr-FR"; Expected = 'OK' },
        @{ Name = 'legacy cache only'; Present = 'C:\ProgramData\SmartM365\Windows11UpgradeToolkit\SetupMedia\Win11-fr-FR'; Expected = 'OK' }
    )) {
        $mockCache = $mockLanguage + ("function Test-Path {{ param([string]`$LiteralPath, [string]`$PathType) if ([string]::IsNullOrEmpty('{0}')) {{ return `$false }}; return `$LiteralPath.StartsWith('{0}\', [System.StringComparison]::OrdinalIgnoreCase) }}`r`n" -f $requirementCase.Present)
        $runspace = [powershell]::Create()
        try {
            $output = @($runspace.AddScript($mockCache + $cacheScript).Invoke())
            Write-TestResult -Condition (-not $runspace.HadErrors -and ($output -join ',') -eq $requirementCase.Expected) -Message ("cache-only requirement ({0}): {1}" -f $requirementCase.Name,$requirementCase.Expected) -Detail ($output -join ',')
        }
        finally { $runspace.Dispose() }
    }

    # ---------- Root CMD, docs and template ----------
    $cmdText = Get-Content -LiteralPath $rootCmd -Raw
    Test-Step 'root CMD uses setlocal DisableDelayedExpansion, pushd/popd and exit /b' { ($cmdText -match '(?im)^setlocal DisableDelayedExpansion') -and ($cmdText -match '(?im)^pushd ') -and ($cmdText -match '(?im)^popd') -and ($cmdText -match '(?im)^exit /b %SEMT_EXIT%') }
    Test-Step 'root CMD targets the new manifest generator and campaign SetupSource' { $cmdText.Contains('%SEMT_ROOT_DIR%Campaigns\Windows11Upgrade\Scripts\SmartM365-EndpointMigration-New-SetupMediaManifest.ps1') -and $cmdText.Contains('%SEMT_ROOT_DIR%Campaigns\Windows11Upgrade\SetupSource') -and (Test-Path -LiteralPath $manifestGenerator) }
    Test-Step 'root CMD uses CRLF line endings' { -not ([regex]::IsMatch($cmdText, '(?<!\r)\n')) }
    $readmeText = Get-Content -LiteralPath $intuneReadme -Raw
    Test-Step 'README documents the new identities' { $readmeText.Contains($newDataRoot) -and $readmeText.Contains('HKLM\' + $newRegistryRoot) -and $readmeText.Contains($newTaskName) -and $readmeText.Contains($newPackagePrefix) -and $readmeText.Contains('Smart Endpoint Migration Toolkit - Windows 11 Upgrade (<Language>)') }
    Test-Step 'README documents the unassign-then-assign switch order and in-place legacy media reuse' { ($readmeText -match 'remove the assignment of the old') -and ($readmeText -match 'reused in place') -and ($readmeText.IndexOf('remove the assignment of the old') -lt $readmeText.IndexOf('Assign the new')) }
    Test-Step 'README states the no-Graph-assignment policy' { $readmeText -match 'never create,\s+change or remove Intune assignments' }
    Test-Step 'gates README requires Modify on the UNC gate root for target computer accounts' { (Get-Content -LiteralPath $gatesReadme -Raw) -match 'need Modify on the UNC gate root' }
    $mapRows = @(Import-Csv -LiteralPath $sourceMapTemplate)
    Test-Step 'setup source map template keeps its columns and UNC sources' { $mapRows.Count -ge 1 -and @($mapRows | Where-Object { -not ([string]$_.SetupSourcePath).StartsWith('\\') }).Count -eq 0 -and ($mapRows[0].PSObject.Properties.Name -join ',') -eq 'ScopeType,ScopeValue,SetupSourcePath,Priority' }
}
catch {
    Write-TestResult -Condition $false -Message 'test harness completed without an unexpected error' -Detail ($_.Exception.Message + ' @ ' + $_.InvocationInfo.PositionMessage)
}
finally {
    foreach ($name in @('Set-Registry64String','Write-InstallLog','Write-RunnerLog','Set-PackageDetectionState','Invoke-IntuneLogMaintenance','Set-PackageInstalledForWindows11','Get-EndpointAttemptOutcome','ConvertTo-Base64Utf8','New-EndpointRequirementRule','Read-CompanionPackageMetadata','Get-LanguageFromPackageId','Set-DetectScriptPackageMetadata','ConvertTo-LongLiteralPath')) {
        Remove-Item -Path ("Function:\global:{0}" -f $name) -Force -ErrorAction SilentlyContinue
    }
    Remove-Variable -Name SyntheticRegistryWrites -Scope Global -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $script:Passed,$script:Failed)
if ($script:Failed -gt 0) { exit 1 }
exit 0
