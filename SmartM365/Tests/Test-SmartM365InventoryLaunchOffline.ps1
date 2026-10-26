#Requires -Version 5.1
<#
.SYNOPSIS
Validate offline the SmartInventory launch engine and its catalog against the legacy .cmd launchers.
.VERSION
1.1.0
.NOTES
Runs under Windows PowerShell 5.1 and PowerShell 7. Nothing is executed outside a fictional fixture.
Checks:
- catalog validity (scripts, runtimes, execution, includes, tokens, parameters checked by AST);
- parity of the 186 legacy launchers (Launchers\Cloud, OnPremises, Orchestrator, Jobs and
  LaunchersByOrchestrator), frozen in SmartInventory-LegacyLaunchers.psd1 before their removal,
  with their catalog entry (engine, script, arguments, working directory, tenant, cache and
  include, confirmation);
- the old names are one-line aliases of SmartInventory.cmd, which starts Windows PowerShell 5.1;
- tenant and execution-mode resolution order, Auto zone detection;
- real runs on a fictional fixture in %TEMP%: Direct without copy, LocalCache copy, missing
  Include (missing source, failed copy, absent after copy), output inside the cache refused,
  exit codes transmitted (5.1 and 7), extra arguments, the mailbox identity helper loads under 5.1.
#>
[CmdletBinding()]
param()
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$smartRoot = Split-Path $PSScriptRoot -Parent
$inventoryRoot = Join-Path $smartRoot 'SmartInventory'
$orchestratorRoot = Join-Path $inventoryRoot 'Orchestrator'
$moduleFile = Join-Path $orchestratorRoot 'SmartM365.Inventory.Launch.psm1'
$launchModule = Import-Module -Name $moduleFile -Force -PassThru -ErrorAction Stop
$script:checks = 0

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
    $script:checks++
}

function Assert-Throws {
    param([scriptblock]$Script, [string]$Pattern, [string]$Message)
    $thrown = $null
    try { & $Script } catch { $thrown = $_.Exception.Message }
    if ($null -eq $thrown) { throw "$Message (no error)" }
    if ($Pattern -and $thrown -notmatch $Pattern) { throw "$Message (unexpected error: $thrown)" }
    $script:checks++
}

# ---------------------------------------------------------------- frozen reference of the legacy launchers
function ConvertTo-NormalizedArgumentSet {
    <# Named-parameter view: "-name=value" groups, sorted (the order of named parameters does not matter). #>
    param([AllowEmptyCollection()][string[]]$Arguments)
    $groups = New-Object System.Collections.Generic.List[string]
    $current = $null
    foreach ($argument in @($Arguments)) {
        if ($argument -match '^-[A-Za-z]') {
            if ($null -ne $current) { $groups.Add($current) }
            $current = $argument.ToLowerInvariant()
            continue
        }
        $value = $argument
        try { if ([IO.Path]::IsPathRooted($value)) { $value = [IO.Path]::GetFullPath($value) } } catch { $value = $argument }
        $current = '{0}={1}' -f $current, $value.ToLowerInvariant()
    }
    if ($null -ne $current) { $groups.Add($current) }
    return ((@($groups) | Sort-Object) -join ' ')
}

function Get-PlanCommandView {
    <# Arguments of a plan after -File <script>, without -Tenant, with tokens resolved. #>
    param($Plan, [string]$InventoryRoot)
    $command = Get-SmartM365InventoryCommand -Plan $Plan -EnginePath 'engine.exe' -ScriptPath $Plan.ScriptPath -SmartInventoryRoot $InventoryRoot
    $all = @($command.Arguments)
    $index = [array]::IndexOf($all, '-File')
    $arguments = New-Object System.Collections.Generic.List[string]
    for ($i = $index + 2; $i -lt $all.Count; $i++) {
        if ($all[$i] -eq '-Tenant') { $i++; continue }
        $arguments.Add($all[$i])
    }
    [pscustomobject]@{ Arguments = $arguments.ToArray(); WorkingDirectory = $command.WorkingDirectory }
}

# Documented, verified differences between a legacy launcher and its catalog entry.
$tolerances = @{
    # The job adds -Connect like the orchestrator; in this script -Connect only writes a log line (checked below).
    'Launchers\Cloud\Start-SmartM365-AzureADConnect-SyncHealth-Inventory.cmd' = 'Connect'
    # The generated job launcher had no pushd; the script does not depend on the current directory (checked below).
    'Launchers\Jobs\Start-WorkplaceEvidence-Prepare.cmd' = 'WorkingDirectory'
    # The launcher preferred pwsh when installed; the catalog entry runs Windows PowerShell 5.1 (decision of the migration).
    'Launchers\Orchestrator\Start-SmartM365-Inventory-OrchestratorScheduledTask-Installer.cmd' = 'Runtime'
}
$syncHealthScript = [IO.File]::ReadAllText((Join-Path $inventoryRoot 'M365Inventory\SyncHealth\SmartM365-AzureADConnect-SyncHealth-Inventory.ps1'))
Assert-True (([regex]::Matches($syncHealthScript, '\$Connect\b')).Count -eq 2 -and $syncHealthScript -match 'Connect switch specified; Graph connection will be established by this script') 'SyncHealth tolerance is stale: -Connect does more than log.'
$prepareScript = [IO.File]::ReadAllText((Join-Path $inventoryRoot 'PreparedEvidence\SmartM365-WorkplaceEvidence-Prepare.ps1'))
Assert-True ($prepareScript -notmatch 'Get-Location|\$PWD\b|CurrentDirectory') 'WorkplaceEvidence-Prepare tolerance is stale: the script reads the current directory.'
$installerScript = [IO.File]::ReadAllText((Join-Path $orchestratorRoot 'Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1'))
Assert-True ($installerScript -match 'Windows PowerShell 5\.1 or PowerShell 7') 'Scheduled task installer tolerance is stale: the installer no longer supports Windows PowerShell 5.1.'

function Assert-LauncherParity {
    param([Parameter(Mandatory)]$Record, [Parameter(Mandatory)]$Plan)
    $label = [string]$Record.Launcher
    $tolerance = ''
    if ($tolerances.ContainsKey($label)) { $tolerance = $tolerances[$label] }
    $view = Get-PlanCommandView -Plan $Plan -InventoryRoot $inventoryRoot
    if ($tolerance -eq 'Runtime') { Assert-True ($Record.Runtime -eq 'PowerShell7OrWindowsPowerShell' -and $Plan.Runtime -eq 'WindowsPowerShell') "$label : runtime tolerance" }
    else { Assert-True ($Record.Runtime -eq $Plan.Runtime) "$label : runtime $($Record.Runtime) <> $($Plan.Runtime)" }
    $expectedScript = [IO.Path]::GetFullPath((Join-Path $inventoryRoot $Record.Script))
    Assert-True ([IO.Path]::GetFullPath($Plan.ScriptPath) -ieq $expectedScript) "$label : script $expectedScript <> $($Plan.ScriptPath)"
    $expectedArguments = @(@($Record.Arguments) | ForEach-Object { ([string]$_).Replace('<SmartInventory>\', $inventoryRoot + '\') })
    $actualArguments = @($view.Arguments)
    if ($Record.Action -eq 'Request') { $actualArguments = @($Plan.Arguments) }
    if ($tolerance -eq 'Connect') { $actualArguments = @($actualArguments | Where-Object { $_ -ne '-Connect' }) }
    $expectedSet = ConvertTo-NormalizedArgumentSet -Arguments $expectedArguments
    $actualSet = ConvertTo-NormalizedArgumentSet -Arguments $actualArguments
    Assert-True ($expectedSet -eq $actualSet) "$label : arguments [$expectedSet] <> [$actualSet]"
    $expectedDirectory = ''
    if ($Record.WorkingDirectory) { $expectedDirectory = [IO.Path]::GetFullPath((Join-Path $inventoryRoot $Record.WorkingDirectory)) }
    if ($tolerance -ne 'WorkingDirectory') { Assert-True ($expectedDirectory -ieq $view.WorkingDirectory) "$label : working directory '$expectedDirectory' <> '$($view.WorkingDirectory)'" }
    $actualTenant = ''
    if ($Plan.PassTenant) { $actualTenant = $Plan.Tenant }
    Assert-True ($Record.Tenant -eq $actualTenant) "$label : tenant '$($Record.Tenant)' <> '$actualTenant'"
    $expectedMode = 'Direct'
    if ($Record.CacheName) { $expectedMode = 'LocalCache' }
    Assert-True ($Plan.Mode -eq $expectedMode) "$label : mode $expectedMode <> $($Plan.Mode)"
    if ($Record.CacheName) { Assert-True ($Plan.CacheName -eq $Record.CacheName) "$label : cache $($Record.CacheName) <> $($Plan.CacheName)" }
    Assert-True ((@($Record.Include) -join '|') -eq (@($Plan.Include) -join '|')) "$label : include [$(@($Record.Include) -join '|')] <> [$(@($Plan.Include) -join '|')]"
    Assert-True ([bool]$Record.RequiresConfirmation -eq [bool]$Plan.RequiresConfirmation) "$label : confirmation $($Record.RequiresConfirmation) <> $($Plan.RequiresConfirmation)"
}

# ---------------------------------------------------------------- real catalog
$previousTenant = $env:SMARTM365_ORCHESTRATOR_TENANT
try {
    $env:SMARTM365_ORCHESTRATOR_TENANT = 'prod'
    $catalog = Get-SmartM365InventoryCatalog -SmartInventoryRoot $inventoryRoot
    $catalogErrors = @(Test-SmartM365InventoryCatalog -Catalog $catalog)
    Assert-True ($catalogErrors.Count -eq 0) ("Catalog errors:`n" + ($catalogErrors -join "`n"))
    $template = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($catalog.TemplatePath))
    Assert-True (@($template.Launches).Count -gt 0 -and @($template.Launches | Where-Object { $_.Manual -ne $true }).Count -eq 0) 'Every Launches entry must declare "Manual": true.'
    Assert-True (@($template.Jobs | Where-Object { $_.PSObject.Properties['Manual'] }).Count -eq 0) 'Manual entries must stay out of the Jobs array.'

    $reference = @((Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'SmartInventory-LegacyLaunchers.psd1')).Launchers)
    Assert-True ($reference.Count -eq 186) "The frozen reference must describe the 186 legacy launchers ($($reference.Count))."

    # Every hand-written legacy launcher is mapped once in the catalog, to the entry recorded in the reference.
    $legacyMap = @{}
    foreach ($entry in $catalog.Entries) {
        $legacyProperty = $entry.Definition.PSObject.Properties['LegacyLaunchers']
        if ($null -eq $legacyProperty) { continue }
        foreach ($launcher in @($legacyProperty.Value)) { $legacyMap[$launcher.ToLowerInvariant()] = $entry.Name }
    }
    $handWritten = @($reference | Where-Object { $_.Launcher -match '^Launchers\\(Cloud|OnPremises|Orchestrator)\\' })
    Assert-True ($handWritten.Count -eq 77 -and $legacyMap.Count -eq 77) "Catalog LegacyLaunchers must map the 77 hand-written launchers ($($legacyMap.Count))."
    foreach ($record in $reference) {
        if ($record.Launcher -match '^Launchers\\(Cloud|OnPremises|Orchestrator)\\') {
            Assert-True ($legacyMap[$record.Launcher.ToLowerInvariant()] -eq $record.Entry) "$($record.Launcher): mapped to '$($legacyMap[$record.Launcher.ToLowerInvariant()])', reference '$($record.Entry)'."
        }
        $plan = Resolve-SmartM365InventoryLaunch -Catalog $catalog -Name $record.Entry -Action $record.Action -DefaultOnPremExecutionMode 'LocalCache'
        if ($record.Launcher -like 'Launchers\Jobs\*') { Assert-True ($plan.Kind -eq 'Job') "$($record.Launcher): $($record.Entry) is not a job." }
        Assert-LauncherParity -Record $record -Plan $plan
    }

    # Old names stay one version as one-line aliases of SmartInventory.cmd.
    $entryPoint = Join-Path $inventoryRoot 'SmartInventory.cmd'
    $entryPointText = [IO.File]::ReadAllText($entryPoint)
    Assert-True ($entryPointText -match '"%POWERSHELL5%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Orchestrator\\Invoke-SmartM365InventoryJob\.ps1" %\*' -and $entryPointText -match 'WindowsPowerShell\\v1\.0\\powershell\.exe' -and $entryPointText -notmatch '(?i)pwsh') 'SmartInventory.cmd must start Windows PowerShell 5.1 with the engine.'
    foreach ($record in $reference) {
        $aliasPath = Join-Path $inventoryRoot $record.Launcher
        Assert-True (Test-Path -LiteralPath $aliasPath -PathType Leaf) "Alias missing: $($record.Launcher)"
        $lines = @([IO.File]::ReadAllLines($aliasPath) | Where-Object { $_.Trim() })
        $depth = ($record.Launcher.Split('\').Count - 1)
        $prefix = '@call "%~dp0' + ('..\' * $depth) + 'SmartInventory.cmd"'
        $arguments = if ($record.Launcher -like '*\Request-SmartM365-OrchestratorJob.cmd') { '-Request' } elseif ($record.Action -eq 'Request') { $record.Entry + ' -Request' } else { $record.Entry }
        Assert-True ($lines.Count -eq 1 -and $lines[0] -ceq ($prefix + ' ' + $arguments + ' %*')) "$($record.Launcher): alias must be '$prefix $arguments %*', found '$($lines -join ' | ')'."
    }
    $aliasFiles = @(Get-ChildItem -LiteralPath (Join-Path $inventoryRoot 'Launchers'), (Join-Path $inventoryRoot 'LaunchersByOrchestrator') -Filter '*.cmd' -File -Recurse)
    Assert-True ($aliasFiles.Count -eq $reference.Count) "Only the aliases of the legacy names may remain ($($aliasFiles.Count) .cmd files)."

    $requestPlan = Resolve-SmartM365InventoryLaunch -Catalog $catalog -Name 'AD-Inventory' -Action Request -ExtraArguments @('-IncludeDependencies', '-NoWait')
    $requestCommand = Get-SmartM365InventoryCommand -Plan $requestPlan -EnginePath 'pwsh.exe' -ScriptPath $requestPlan.ScriptPath -SmartInventoryRoot $inventoryRoot
    Assert-True (((@($requestCommand.Arguments) | Select-Object -Skip 5) -join ' ') -eq '-Tenant prod -Job AD-Inventory -Collect -IncludeDependencies -NoWait') ('Request extra arguments must be passed once: ' + (@($requestCommand.Arguments) -join ' '))
    Assert-Throws { Resolve-SmartM365InventoryLaunch -Catalog $catalog -Name 'Exchange2016-ProxyAddresses-Write' -Action Request } 'manual launch' 'A manual entry must not be requested through the orchestrators.'

    # Legacy launchers that ignored SMARTM365_ORCHESTRATOR_TENANT (hard-coded prod) now follow it.
    $env:SMARTM365_ORCHESTRATOR_TENANT = 'demo'
    $write = Resolve-SmartM365InventoryLaunch -Catalog $catalog -Name 'Exchange2016-ProxyAddresses-Write'
    Assert-True ($write.Tenant -eq 'demo' -and $write.Interactive -and $write.RequiresConfirmation -and $write.Mode -eq 'LocalCache') 'ProxyAddresses Write: tenant, confirmation or cache.'
    $onlyAd = Resolve-SmartM365InventoryLaunch -Catalog $catalog -Name 'Exchange2016-Local-Mailboxes-OnlyADPermissions'
    Assert-True ($onlyAd.Tenant -eq 'demo') 'OnlyADPermissions must follow the orchestrator tenant.'
    foreach ($installer in @('Exchange-HealthChecker-ScheduledTask-Installer', 'Orchestrator-ScheduledTask-Installer')) {
        $installerPlan = Resolve-SmartM365InventoryLaunch -Catalog $catalog -Name $installer
        Assert-True ($installerPlan.Kind -eq 'Manual' -and $installerPlan.Runtime -eq 'WindowsPowerShell' -and $installerPlan.Mode -eq 'Direct' -and $installerPlan.Interactive) "$installer must be a manual Windows PowerShell entry in Direct mode."
    }
}
finally {
    $env:SMARTM365_ORCHESTRATOR_TENANT = $previousTenant
}

# ---------------------------------------------------------------- fictional fixture
$fixtureBase = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365LaunchTest-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
$fixtureRoot = Join-Path $fixtureBase 'SmartM365'
$fixtureInventory = Join-Path $fixtureRoot 'SmartInventory'
$workspace = Join-Path $fixtureBase 'Workspace'
$markerFolder = Join-Path $fixtureBase 'Markers'
$cacheRoot = Join-Path $fixtureBase 'Cache'

function Write-FixtureFile {
    param([string]$Path, [string]$Content)
    $folder = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    [IO.File]::WriteAllText($Path, $Content, (New-Object System.Text.UTF8Encoding($false)))
}

function Set-FixtureGlobalConfig {
    param([string]$WorkspaceRootPath)
    $global = [ordered]@{ DefaultTenant = 'demo'; SmartM365RootPath = '.'; WorkspaceRootPath = $WorkspaceRootPath; OnPremExecutionMode = 'LocalCache' }
    Write-FixtureFile -Path (Join-Path $fixtureRoot 'Config\SmartM365.global.local.json') -Content ($global | ConvertTo-Json)
}

function Read-Marker {
    param([string]$Name)
    $path = Join-Path $markerFolder ($Name + '.json')
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    return (ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($path)))
}

try {
    foreach ($file in @('Config\SmartM365-TenantContext.ps1', 'Config\SmartM365.global.local.json.template', 'Config\SmartM365-JsonTransport.policy.psd1', 'Config\Tenants\tenant.local.json.template', 'Modules\SmartM365.Core\SmartM365.JsonTransport.psd1', 'Modules\SmartM365.Core\SmartM365.JsonTransport.psm1')) {
        $destination = Join-Path $fixtureRoot $file
        New-Item -ItemType Directory -Path (Split-Path $destination -Parent) -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $smartRoot $file) -Destination $destination
    }
    New-Item -ItemType Directory -Path (Join-Path $fixtureInventory 'Orchestrator'), $markerFolder, $workspace -Force | Out-Null
    Copy-Item -LiteralPath $moduleFile, (Join-Path $orchestratorRoot 'Invoke-SmartM365InventoryJob.ps1') -Destination (Join-Path $fixtureInventory 'Orchestrator')
    Set-FixtureGlobalConfig -WorkspaceRootPath $workspace
    $tenantProfile = [ordered]@{ ProfileKey = 'demo'; OrganizationKey = 'contoso'; EnvironmentKey = 'lab'; TenantKey = 'contoso-lab' }
    Write-FixtureFile -Path (Join-Path $fixtureRoot 'Config\Tenants\demo.local.json') -Content ($tenantProfile | ConvertTo-Json)

    $probe = @'
param([string]$Tenant = 'none', [int]$ExitCode = 0, [string]$Marker = 'probe', [string]$OutputPath = '', [string]$Value = '', [switch]$Flag)
$markerFolder = $env:SMARTM365_TEST_MARKERS
$data = [ordered]@{ ScriptRoot = $PSScriptRoot; Tenant = $Tenant; Edition = $PSVersionTable.PSEdition; Major = $PSVersionTable.PSVersion.Major; Value = $Value; Flag = [bool]$Flag; CurrentDirectory = [Environment]::CurrentDirectory }
[IO.File]::WriteAllText((Join-Path $markerFolder ($Marker + '.json')), ($data | ConvertTo-Json))
exit $ExitCode
'@
    Write-FixtureFile -Path (Join-Path $fixtureInventory 'Fixture\SmartM365-Fixture-Probe.ps1') -Content $probe
    Write-FixtureFile -Path (Join-Path $fixtureInventory 'Common\SmartM365.Fixture.Include.ps1') -Content '# fixture include'
    $identityHelper = Join-Path $inventoryRoot 'Common\SmartM365.IdentityDateEvidence.ps1'
    Copy-Item -LiteralPath $identityHelper -Destination (Join-Path $fixtureInventory 'Common\SmartM365.IdentityDateEvidence.ps1')
    Write-FixtureFile -Path (Join-Path $fixtureInventory 'Fixture\prices.csv') -Content 'Sku,Price'
    $env:SMARTM365_TEST_MARKERS = $markerFolder

    $fixtureCatalog = [ordered]@{
        Jobs = @(
            [ordered]@{ Name = 'Probe-Direct'; ScriptPath = 'Fixture\SmartM365-Fixture-Probe.ps1'; Arguments = '-Marker direct -ExitCode 7'; PowerShellEdition = 'PowerShell7'; Schedule = [ordered]@{ Type = 'Daily'; Times = @('01:00') } },
            [ordered]@{ Name = 'Probe-Desktop'; ScriptPath = 'Fixture\SmartM365-Fixture-Probe.ps1'; Arguments = '-Marker desktop -ExitCode 5'; Runtime = 'WindowsPowerShell' },
            [ordered]@{ Name = 'Probe-Cache'; ScriptPath = 'Fixture\SmartM365-Fixture-Probe.ps1'; LaunchArguments = '-Flag'; Arguments = '-Marker cache -ExitCode 3'; Runtime = 'WindowsPowerShell'; Execution = [ordered]@{ CacheName = 'Fixture\Probe'; Include = @('SmartInventory\Common\SmartM365.Fixture.Include.ps1', 'SmartInventory\Common\SmartM365.IdentityDateEvidence.ps1') } },
            [ordered]@{ Name = 'Probe-MissingInclude'; ScriptPath = 'Fixture\SmartM365-Fixture-Probe.ps1'; Arguments = '-Marker missing'; Runtime = 'WindowsPowerShell'; Execution = [ordered]@{ CacheName = 'Fixture\Missing'; Include = @('SmartInventory\Common\SmartM365.Absent.ps1') } },
            [ordered]@{ Name = 'Probe-JobMode'; ScriptPath = 'Fixture\SmartM365-Fixture-Probe.ps1'; Arguments = '-Marker jobmode'; Execution = [ordered]@{ Mode = 'Direct'; CacheName = 'Fixture\JobMode'; Include = @() } }
        )
        Launches = @(
            [ordered]@{ Name = 'Probe-Manual'; Manual = $true; Description = 'Fixture manual entry.'; ScriptPath = 'Fixture\SmartM365-Fixture-Probe.ps1'; Arguments = '-Marker manual -Value {{PriceList}}'; WorkingDirectory = 'Fixture'; Tokens = [ordered]@{ PriceList = [ordered]@{ Value = '{{ScriptDirectory}}\prices.csv'; MustExist = $true } } },
            [ordered]@{ Name = 'Probe-Write'; Manual = $true; ScriptPath = 'Fixture\SmartM365-Fixture-Probe.ps1'; Arguments = '-Marker write'; RequiresConfirmation = $true; Interactive = $true }
        )
    }
    Write-FixtureFile -Path (Join-Path $fixtureInventory 'Orchestrator\Orchestrator-Jobs.json.template') -Content ($fixtureCatalog | ConvertTo-Json -Depth 10)
    $fixture = Get-SmartM365InventoryCatalog -SmartInventoryRoot $fixtureInventory
    $fixtureErrors = @(Test-SmartM365InventoryCatalog -Catalog $fixture)
    Assert-True ($fixtureErrors.Count -eq 1 -and $fixtureErrors[0] -match 'Probe-MissingInclude : Include not found') ('Fixture catalog errors: ' + ($fixtureErrors -join '; '))
    $cacheCandidates = @((Join-Path $cacheRoot 'Unwritable'), (Join-Path $cacheRoot 'Probe'))
    # A file at the first candidate makes it unusable, as an unwritable folder would.
    Write-FixtureFile -Path (Join-Path $cacheRoot 'Unwritable') -Content 'not a folder'

    # Catalog validation: unknown parameter, missing script, undefined token (checked by AST, nothing runs).
    $broken = [ordered]@{ Jobs = @([ordered]@{ Name = 'Broken'; ScriptPath = 'Fixture\SmartM365-Fixture-Probe.ps1'; Arguments = '-NoSuchParameter -Value {{Undefined}}' }, [ordered]@{ Name = 'Gone'; ScriptPath = 'Fixture\Absent.ps1' }) }
    $brokenRoot = Join-Path $fixtureBase 'Broken\SmartInventory'
    Write-FixtureFile -Path (Join-Path $brokenRoot 'Orchestrator\Orchestrator-Jobs.json.template') -Content ($broken | ConvertTo-Json -Depth 10)
    Copy-Item -LiteralPath (Join-Path $fixtureInventory 'Fixture') -Destination (Join-Path $brokenRoot 'Fixture') -Recurse
    $brokenErrors = @(Test-SmartM365InventoryCatalog -Catalog (Get-SmartM365InventoryCatalog -SmartInventoryRoot $brokenRoot)) -join "`n"
    Assert-True ($brokenErrors -match 'NoSuchParameter is not declared') 'AST check must report an unknown parameter.'
    Assert-True ($brokenErrors -match 'Gone : script not found') 'Catalog check must report a missing script.'
    Assert-True ($brokenErrors -match '\{\{Undefined\}\} is not defined') 'Catalog check must report an undefined token.'

    # Tenant: parameter, then SMARTM365_ORCHESTRATOR_TENANT, then DefaultTenant.
    $env:SMARTM365_ORCHESTRATOR_TENANT = ''
    Assert-True ((Resolve-SmartM365InventoryTenant -SmartM365Root $fixtureRoot).Value -eq 'demo') 'Tenant must fall back to DefaultTenant.'
    $env:SMARTM365_ORCHESTRATOR_TENANT = 'fromorchestrator'
    Assert-True ((Resolve-SmartM365InventoryTenant -SmartM365Root $fixtureRoot).Source -eq 'SMARTM365_ORCHESTRATOR_TENANT') 'Tenant must follow SMARTM365_ORCHESTRATOR_TENANT.'
    Assert-True ((Resolve-SmartM365InventoryTenant -Tenant 'explicit' -SmartM365Root $fixtureRoot).Value -eq 'explicit') 'Tenant parameter must win.'
    $env:SMARTM365_ORCHESTRATOR_TENANT = ''

    # Execution mode: parameter, then job (runtime manifest, then template), then OnPremExecutionMode.
    $cacheEntry = Find-SmartM365InventoryCatalogEntry -Catalog $fixture -Name 'Probe-Cache'
    Assert-True ((Resolve-SmartM365InventoryExecutionMode -Entry $cacheEntry -Catalog $fixture).Mode -eq 'LocalCache') 'Global OnPremExecutionMode must apply.'
    Assert-True ((Resolve-SmartM365InventoryExecutionMode -Entry $cacheEntry -Catalog $fixture -ExecutionMode Direct).Mode -eq 'Direct') 'Parameter must win.'
    Assert-True ((Resolve-SmartM365InventoryExecutionMode -Entry (Find-SmartM365InventoryCatalogEntry -Catalog $fixture -Name 'Probe-JobMode') -Catalog $fixture).Source -eq 'Job') 'Job Execution.Mode must win over the global mode.'
    $overlay = [ordered]@{ Jobs = @([ordered]@{ Name = 'Probe-Cache'; Arguments = '-Marker overlay'; Execution = [ordered]@{ Mode = 'Direct' } }) }
    $overlayPath = Join-Path $fixtureBase 'Orchestrator-Jobs.json'
    Write-FixtureFile -Path $overlayPath -Content ($overlay | ConvertTo-Json -Depth 10)
    $overlaid = Resolve-SmartM365InventoryLaunch -Catalog (Get-SmartM365InventoryCatalog -SmartInventoryRoot $fixtureInventory -OverlayPath $overlayPath) -Name 'Probe-Cache'
    Assert-True ($overlaid.Mode -eq 'Direct' -and (@($overlaid.Arguments) -join ' ') -eq '-Flag -Marker overlay') 'The runtime manifest must override Arguments and Execution.Mode, LaunchArguments stay.'
    Assert-Throws { Resolve-SmartM365InventoryLaunch -Catalog $fixture -Name 'Probe-Direct' -ExecutionMode LocalCache } 'only run Direct' 'An entry without Execution cannot run from the cache.'
    foreach ($case in @(@('Intranet', 'Direct'), @('Trusted', 'Direct'), @('MyComputer', 'Direct'), @('Internet', 'LocalCache'), @('Restricted', 'LocalCache'))) {
        $zone = $case[0]
        & $launchModule { param($value) $script:PathZoneResolver = [scriptblock]::Create("'$value'") } $zone
        $auto = Resolve-SmartM365InventoryExecutionMode -Entry $cacheEntry -Catalog $fixture -ExecutionMode Auto
        Assert-True ($auto.Mode -eq $case[1] -and $auto.Zone -eq $zone) "Auto with zone $zone must give $($case[1])."
    }
    & $launchModule { $script:PathZoneResolver = $null }
    Assert-True ((Get-SmartM365InventoryPathZone -Path $fixtureInventory) -eq 'MyComputer') 'A local folder is in the MyComputer zone.'

    # Direct: the script runs in place, nothing is copied, the exit code is transmitted (pwsh 7).
    $plan = Resolve-SmartM365InventoryLaunch -Catalog $fixture -Name 'Probe-Direct' -ExtraArguments @('-Value', 'two words', '-Flag')
    $exitCode = Invoke-SmartM365InventoryLaunch -Plan $plan -Catalog $fixture -CacheBaseCandidates $cacheCandidates
    $marker = Read-Marker -Name 'direct'
    Assert-True ($exitCode -eq 7) "Direct exit code 7 expected, got $exitCode."
    Assert-True ($null -ne $marker -and $marker.ScriptRoot -ieq (Join-Path $fixtureInventory 'Fixture') -and $marker.Edition -eq 'Core') 'Direct must run the script in place with pwsh 7.'
    Assert-True ($marker.Tenant -eq 'demo' -and $marker.Value -eq 'two words' -and $marker.Flag) 'Tenant and extra arguments must be passed.'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $cacheRoot 'Probe'))) 'Direct must not create a cache.'

    # Windows PowerShell 5.1 runtime and its exit code.
    $exitCode = Invoke-SmartM365InventoryLaunch -Plan (Resolve-SmartM365InventoryLaunch -Catalog $fixture -Name 'Probe-Desktop') -Catalog $fixture
    $marker = Read-Marker -Name 'desktop'
    Assert-True ($exitCode -eq 5 -and $marker.Edition -eq 'Desktop' -and $marker.Major -eq 5) 'WindowsPowerShell runtime must run powershell.exe 5.1 and return its exit code.'

    # Manual entry: working directory, token resolved and checked.
    $exitCode = Invoke-SmartM365InventoryLaunch -Plan (Resolve-SmartM365InventoryLaunch -Catalog $fixture -Name 'Probe-Manual') -Catalog $fixture
    $marker = Read-Marker -Name 'manual'
    Assert-True ($exitCode -eq 0 -and $marker.CurrentDirectory.TrimEnd('\') -ieq (Join-Path $fixtureInventory 'Fixture') -and $marker.Value -ieq (Join-Path $fixtureInventory 'Fixture\prices.csv')) 'WorkingDirectory and tokens must be applied.'
    Remove-Item -LiteralPath (Join-Path $fixtureInventory 'Fixture\prices.csv')
    Assert-Throws { Invoke-SmartM365InventoryLaunch -Plan (Resolve-SmartM365InventoryLaunch -Catalog $fixture -Name 'Probe-Manual') -Catalog $fixture } 'path not found' 'A missing MustExist token must stop the launch.'

    # Confirmation and interactive entries never run unattended (the test console input is redirected or not interactive).
    $inputRedirected = $true
    try { $inputRedirected = [Console]::IsInputRedirected } catch { $inputRedirected = $true }
    if ($inputRedirected -or -not [Environment]::UserInteractive) {
        Assert-Throws { Invoke-SmartM365InventoryLaunch -Plan (Resolve-SmartM365InventoryLaunch -Catalog $fixture -Name 'Probe-Write') -Catalog $fixture } 'interactive console' 'A confirmation entry must refuse to run unattended.'
        Assert-True ($null -eq (Read-Marker -Name 'write')) 'The confirmation entry must not have run.'
    }

    # LocalCache: same folders, order, /MIR, log, include, unblock; script runs from the cache.
    $plan = Resolve-SmartM365InventoryLaunch -Catalog $fixture -Name 'Probe-Cache'
    $exitCode = Invoke-SmartM365InventoryLaunch -Plan $plan -Catalog $fixture -CacheBaseCandidates $cacheCandidates
    $marker = Read-Marker -Name 'cache'
    $cacheBase = Join-Path $cacheRoot 'Probe'
    Assert-True ($exitCode -eq 3) "LocalCache exit code 3 expected, got $exitCode."
    Assert-True ($marker.ScriptRoot -ieq (Join-Path $cacheBase 'SmartM365\SmartInventory\Fixture') -and $marker.Edition -eq 'Desktop' -and $marker.Flag) 'LocalCache must run the cached copy with LaunchArguments.'
    foreach ($relative in @('SmartM365\Config\SmartM365-TenantContext.ps1', 'SmartM365\Modules\SmartM365.Core\SmartM365.JsonTransport.psm1', 'SmartM365\SmartInventory\Common\SmartM365.Fixture.Include.ps1', 'PrepareCache.log')) {
        Assert-True (Test-Path -LiteralPath (Join-Path $cacheBase $relative)) "LocalCache must create $relative."
    }
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $cacheBase 'SmartM365\SmartInventory\Orchestrator'))) 'LocalCache must copy only the script folder of SmartInventory.'
    $cachedHelper = Join-Path $cacheBase 'SmartM365\SmartInventory\Common\SmartM365.IdentityDateEvidence.ps1'
    Assert-True ((Get-FileHash -LiteralPath $cachedHelper).Hash -eq (Get-FileHash -LiteralPath $identityHelper).Hash) 'The cached identity helper must be byte-identical.'
    $powershell5 = Get-SmartM365InventoryEnginePath -Runtime WindowsPowerShell
    $loaded = & $powershell5 -NoProfile -ExecutionPolicy Bypass -Command ". '$cachedHelper'; if (-not (Get-Command Read-IdentityDate -ErrorAction SilentlyContinue)) { exit 7 }; 'HELPER_LOADED'"
    Assert-True ($LASTEXITCODE -eq 0 -and @($loaded) -contains 'HELPER_LOADED') 'The cached identity helper must load under Windows PowerShell 5.1.'
    Write-FixtureFile -Path (Join-Path $cacheBase 'SmartM365\SmartInventory\Fixture\stale.txt') -Content 'stale'
    [void](Invoke-SmartM365InventoryLaunch -Plan $plan -Catalog $fixture -CacheBaseCandidates $cacheCandidates)
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $cacheBase 'SmartM365\SmartInventory\Fixture\stale.txt'))) 'The cache must be emptied before each run.'

    # Missing Include: the launch fails before the script.
    Assert-Throws { Invoke-SmartM365InventoryLaunch -Plan (Resolve-SmartM365InventoryLaunch -Catalog $fixture -Name 'Probe-MissingInclude') -Catalog $fixture -CacheBaseCandidates @((Join-Path $cacheRoot 'Missing')) } 'Include source missing' 'A missing Include must stop the launch.'
    Assert-True ($null -eq (Read-Marker -Name 'missing')) 'The script must not run when an Include is missing.'
    # A failed Include copy (robocopy 8) and a success without the file both stop the launch.
    & $launchModule {
        $script:OriginalRobocopy = ${function:Invoke-SmartM365InventoryRobocopy}
        function script:Invoke-SmartM365InventoryRobocopy {
            param([string]$Source, [string]$Destination, [string[]]$Options, [string]$LogOption)
            if (@($Options) -notcontains '/MIR') { return $script:IncludeRobocopyResult }
            return (& $script:OriginalRobocopy -Source $Source -Destination $Destination -Options $Options -LogOption $LogOption)
        }
    }
    try {
        foreach ($case in @(@(8, 'Include copy failed'), @(0, 'Include absent after copy'))) {
            & $launchModule { param($code) $script:IncludeRobocopyResult = $code } $case[0]
            Remove-Item -LiteralPath (Join-Path $markerFolder 'cache.json') -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath (Join-Path $cacheRoot 'Include') -Recurse -Force -ErrorAction SilentlyContinue
            Assert-Throws { Invoke-SmartM365InventoryLaunch -Plan $plan -Catalog $fixture -CacheBaseCandidates @((Join-Path $cacheRoot 'Include')) } $case[1] "Include robocopy result $($case[0]) must stop the launch."
            Assert-True ($null -eq (Read-Marker -Name 'cache')) "The script must not run after: $($case[1])."
        }
    }
    finally { & $launchModule { ${function:script:Invoke-SmartM365InventoryRobocopy} = $script:OriginalRobocopy } }

    # Outputs inside the cache are refused (workspace left at the SmartM365 root, then a path argument).
    Set-FixtureGlobalConfig -WorkspaceRootPath '{{SmartM365RootPath}}'
    Remove-Item -LiteralPath (Join-Path $markerFolder 'cache.json') -ErrorAction SilentlyContinue
    Assert-Throws { Invoke-SmartM365InventoryLaunch -Plan $plan -Catalog $fixture -CacheBaseCandidates $cacheCandidates } 'inside the disposable launcher cache' 'A data root inside the cache must be refused.'
    Assert-True ($null -eq (Read-Marker -Name 'cache')) 'The script must not run when its output resolves inside the cache.'
    Set-FixtureGlobalConfig -WorkspaceRootPath $workspace
    $pathPlan = Resolve-SmartM365InventoryLaunch -Catalog $fixture -Name 'Probe-Cache' -ExtraArguments @('-OutputPath', (Join-Path $cacheBase 'SmartM365\out'))
    Assert-Throws { Invoke-SmartM365InventoryLaunch -Plan $pathPlan -Catalog $fixture -CacheBaseCandidates $cacheCandidates } 'argument' 'A path argument inside the cache must be refused.'

    # Entry script under Windows PowerShell 5.1: -WhatIf, -List and the exit code of the job.
    $engineScript = Join-Path $fixtureInventory 'Orchestrator\Invoke-SmartM365InventoryJob.ps1'
    $powershell5 = Get-SmartM365InventoryEnginePath -Runtime WindowsPowerShell
    $whatIf = & $powershell5 -NoProfile -ExecutionPolicy Bypass -File $engineScript 'Probe-Cache' -WhatIf 2>&1 | Out-String
    Assert-True ($LASTEXITCODE -eq 0 -and $whatIf -match 'Execution mode: LocalCache' -and $whatIf -match 'SmartM365LauncherCache\\Fixture\\Probe' -and $whatIf -match '-Flag') "-WhatIf must show the resolved command: $whatIf"
    Assert-True ($null -eq (Read-Marker -Name 'cache')) '-WhatIf must not run the script.'
    $list = & $powershell5 -NoProfile -ExecutionPolicy Bypass -File $engineScript -List 2>&1 | Out-String
    Assert-True ($LASTEXITCODE -eq 0 -and $list -match 'Probe-Manual\s+Manual') '-List must show the manual entries.'
    $defaultList = & $powershell5 -NoProfile -ExecutionPolicy Bypass -File $engineScript 2>&1 | Out-String
    Assert-True ($LASTEXITCODE -eq 0 -and $defaultList -match 'Probe-Manual\s+Manual') 'Without arguments the engine must list the catalog.'
    & $powershell5 -NoProfile -ExecutionPolicy Bypass -File $engineScript 'Probe-Direct' -Tenant 'cli' -Value 'cli value' *> $null
    $marker = Read-Marker -Name 'direct'
    Assert-True ($LASTEXITCODE -eq 7 -and $marker.Tenant -eq 'cli' -and $marker.Value -eq 'cli value' -and $marker.Edition -eq 'Core') "The entry script must relaunch pwsh 7 and return the job exit code (got $LASTEXITCODE)."
    & $powershell5 -NoProfile -ExecutionPolicy Bypass -File $engineScript 'No-Such-Job' *> $null
    Assert-True ($LASTEXITCODE -eq 1) 'An unknown entry must return 1.'
}
finally {
    $env:SMARTM365_ORCHESTRATOR_TENANT = $previousTenant
    Remove-Item Env:\SMARTM365_TEST_MARKERS -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $fixtureBase -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Output ("SmartInventory launch engine offline test passed: {0} checks ({1} legacy launchers and aliases; PowerShell {2})." -f $script:checks, $reference.Count, $PSVersionTable.PSVersion)
