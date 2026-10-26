#Requires -Version 5.1
<#
.SYNOPSIS
Validate offline the SmartInventory launch engine and its catalog against the legacy .cmd launchers.
.VERSION
1.4.0
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
  exit codes transmitted (5.1 and 7), extra arguments, the mailbox identity helper loads under 5.1;
- menu of SmartInventory.cmd (domains, advanced entries on request, choice by number or name, run,
  show the command, quit) and HideInGui, Advanced and Group validation;
- cache hardening: explicit ACL, junction and untrusted owner refused, code integrity checked
  before the launch, orchestrator candidates without the shared Windows\Temp and C:\Temp.
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

    # Menu: every entry once, jobs first, then manual launches, then the orchestrator entries.
    $menu = @(Get-SmartM365InventoryMenu -Catalog $catalog)
    $domains = @(Get-SmartM365InventoryMenuDomains -Items $menu)
    $allDomains = @(Get-SmartM365InventoryMenuDomains -Items $menu -IncludeAdvanced)
    $advanced = @($menu | Where-Object Advanced | ForEach-Object Name)
    Assert-True ($menu.Count -eq @($catalog.Entries).Count -and @($menu | Where-Object { $_.Domain -eq 'Other' }).Count -eq 0) 'Every catalog entry must belong to a menu domain.'
    Assert-True ($advanced.Count -eq 9 -and $advanced -contains 'WorkplaceEvidence-ConvertMetadata' -and $advanced -contains 'Orchestrator-ScheduledTask-Installer' -and $advanced -notcontains 'M365-Usage-Inventory-D180') 'Nine advanced entries: migration, repair, diagnostics and setup tools.'
    Assert-True ((($domains | Measure-Object -Property Count -Sum).Sum) -eq ($menu.Count - 9) -and (($allDomains | Measure-Object -Property Count -Sum).Sum) -eq $menu.Count) 'Advanced entries are listed only on request.'
    Assert-True (@($domains.Domain) -join '|' -eq 'Active Directory|Exchange Online|Exchange 2016 (on-premises)|Intune|Microsoft 365|SharePoint (on-premises)|Workplace Evidence|SmartFinOps|Orchestrator') ('Unexpected menu domains: ' + (@($domains.Domain) -join '|'))
    $adItems = @($domains[0].Items)
    Assert-True ($adItems[0].Kind -eq 'Job' -and $adItems[-1].Kind -eq 'Manual') 'Jobs come before manual entries in a domain.'
    Assert-True ((Resolve-SmartM365InventoryMenuChoice -Choices $domains -Items $menu -Choice ' 2 ').Domain -eq 'Exchange Online' -and (Resolve-SmartM365InventoryMenuChoice -Choices $domains -Items $menu -Choice 'workplaceevidence-convertmetadata').Name -eq 'WorkplaceEvidence-ConvertMetadata' -and $null -eq (Resolve-SmartM365InventoryMenuChoice -Choices $domains -Items $menu -Choice '99')) 'A choice resolves by number, or by name even for an advanced entry.'
    $hiddenInGui = @($catalog.Entries | Where-Object { $_.Definition.PSObject.Properties['HideInGui'] -and $_.Definition.HideInGui } | ForEach-Object { $_.Name })
    Assert-True ($hiddenInGui.Count -eq 8 -and $hiddenInGui -contains 'Orchestrator-GUI' -and $hiddenInGui -contains 'Exchange-HealthChecker-ScheduledTask-Installer') 'The orchestrator entries and the installers are hidden in the GUI.'

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
            [ordered]@{ Name = 'Probe-Manual'; Manual = $true; Group = 'Fixture'; Advanced = $true; Description = 'Fixture manual entry.'; ScriptPath = 'Fixture\SmartM365-Fixture-Probe.ps1'; Arguments = '-Marker manual -Value {{PriceList}}'; WorkingDirectory = 'Fixture'; Tokens = [ordered]@{ PriceList = [ordered]@{ Value = '{{ScriptDirectory}}\prices.csv'; MustExist = $true } } },
            [ordered]@{ Name = 'Probe-Write'; Manual = $true; Group = 'Fixture'; ScriptPath = 'Fixture\SmartM365-Fixture-Probe.ps1'; Arguments = '-Marker write'; RequiresConfirmation = $true; Interactive = $true }
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
    $broken = [ordered]@{ Jobs = @([ordered]@{ Name = 'Broken'; ScriptPath = 'Fixture\SmartM365-Fixture-Probe.ps1'; Arguments = '-NoSuchParameter -Value {{Undefined}}'; HideInGui = $true }, [ordered]@{ Name = 'Gone'; ScriptPath = 'Fixture\Absent.ps1' }) }
    $brokenRoot = Join-Path $fixtureBase 'Broken\SmartInventory'
    Write-FixtureFile -Path (Join-Path $brokenRoot 'Orchestrator\Orchestrator-Jobs.json.template') -Content ($broken | ConvertTo-Json -Depth 10)
    Copy-Item -LiteralPath (Join-Path $fixtureInventory 'Fixture') -Destination (Join-Path $brokenRoot 'Fixture') -Recurse
    $brokenErrors = @(Test-SmartM365InventoryCatalog -Catalog (Get-SmartM365InventoryCatalog -SmartInventoryRoot $brokenRoot)) -join "`n"
    Assert-True ($brokenErrors -match 'NoSuchParameter is not declared') 'AST check must report an unknown parameter.'
    Assert-True ($brokenErrors -match 'Gone : script not found') 'Catalog check must report a missing script.'
    Assert-True ($brokenErrors -match '\{\{Undefined\}\} is not defined') 'Catalog check must report an undefined token.'
    Assert-True ($brokenErrors -match 'Broken : HideInGui must be') 'HideInGui is reserved for manual entries.'

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

    # Hardened cache: explicit ACL limited to SYSTEM, Administrators and the launching account.
    $currentSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $cacheAcl = Get-Acl -LiteralPath $cacheBase
    $aclSids = @($cacheAcl.Access | ForEach-Object { $_.IdentityReference.Translate([Security.Principal.SecurityIdentifier]).Value } | Sort-Object -Unique)
    Assert-True ($cacheAcl.AreAccessRulesProtected -and @($aclSids | Where-Object { $_ -notin @('S-1-5-18', 'S-1-5-32-544', $currentSid) }).Count -eq 0 -and $aclSids -contains $currentSid) "The cache folder ACL must be explicit (found: $($aclSids -join ', '))."
    # A junction and a folder owned by another account are refused; the next safe candidate is used.
    $junctionTarget = Join-Path $fixtureBase 'JunctionTarget'
    New-Item -ItemType Directory -Path $junctionTarget | Out-Null
    $junction = Join-Path $cacheRoot 'Junction'
    & $env:ComSpec /d /c mklink /J "$junction" "$junctionTarget" | Out-Null
    $untrusted = Join-Path $cacheRoot 'Untrusted'
    New-Item -ItemType Directory -Path $untrusted | Out-Null
    & $launchModule { $script:OwnerResolver = { param($path) if ($path -like '*\Untrusted') { 'S-1-5-32-545' } else { (Get-Acl -LiteralPath $path).GetOwner([Security.Principal.SecurityIdentifier]).Value } } }
    try {
        Assert-True ((Test-SmartM365InventoryCacheLocation -Path $junction) -match 'junction') 'A junction must be refused as cache folder.'
        Assert-True ((Test-SmartM365InventoryCacheLocation -Path (Join-Path $untrusted 'Child')) -match 'owned by S-1-5-32-545') 'A folder under a parent owned by another account must be refused.'
        Remove-Item -LiteralPath (Join-Path $markerFolder 'cache.json') -ErrorAction SilentlyContinue
        $exitCode = Invoke-SmartM365InventoryLaunch -Plan $plan -Catalog $fixture -CacheBaseCandidates @($junction, $untrusted, (Join-Path $cacheRoot 'Safe'))
        $marker = Read-Marker -Name 'cache'
        Assert-True ($exitCode -eq 3 -and $marker.ScriptRoot -like (Join-Path $cacheRoot 'Safe\SmartM365\SmartInventory\Fixture')) 'Refused candidates must be skipped for the next safe one.'
        Assert-True (@(Get-ChildItem -LiteralPath $junctionTarget -Force).Count -eq 0) 'Nothing may be written through a junction.'
        Assert-Throws { Invoke-SmartM365InventoryLaunch -Plan $plan -Catalog $fixture -CacheBaseCandidates @($junction, $untrusted) } 'safe local cache' 'Without a safe candidate the launch must stop.'
    }
    finally {
        & $launchModule { $script:OwnerResolver = $null }
        & $env:ComSpec /d /c rmdir "$junction" | Out-Null
    }
    # Code integrity: a cached code file that differs from its source, or an extra one, stops the launch.
    $sourceFolder = Join-Path $fixtureInventory 'Fixture'
    $cachedFolder = Join-Path $cacheBase 'SmartM365\SmartInventory\Fixture'
    Assert-SmartM365InventoryCacheIntegrity -Source $sourceFolder -Destination $cachedFolder
    [IO.File]::AppendAllText((Join-Path $cachedFolder 'SmartM365-Fixture-Probe.ps1'), '# tampered')
    Assert-Throws { Assert-SmartM365InventoryCacheIntegrity -Source $sourceFolder -Destination $cachedFolder } 'differs from its source' 'A modified cached script must be detected.'
    Write-FixtureFile -Path (Join-Path $cachedFolder 'Extra.ps1') -Content '# extra'
    Assert-Throws { Assert-SmartM365InventoryCacheIntegrity -Source $sourceFolder -Destination $cachedFolder } 'unexpected code file' 'An extra cached script must be detected.'
    # The orchestrator never uses the shared Windows\Temp and C:\Temp folders.
    $orchestratedCandidates = @(Get-SmartM365InventoryCacheCandidates -CacheName 'Fixture\Probe' -Orchestrated)
    Assert-True ($orchestratedCandidates[0] -like '*\SmartM365\LauncherCache\Fixture\Probe' -and @($orchestratedCandidates | Where-Object { $_ -like 'C:\Temp\*' -or $_ -like (Join-Path $env:SystemRoot 'Temp\*') }).Count -eq 0) 'Orchestrator cache candidates must exclude Windows\Temp and C:\Temp.'
    Assert-True (@(Get-SmartM365InventoryCacheCandidates -CacheName 'Fixture\Probe').Count -eq (@($orchestratedCandidates).Count + 2)) 'Manual runs keep the four legacy candidates.'

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
    # Menu driven through the input: show the command, go back, choose by name, run now; quitting runs nothing.
    Remove-Item -LiteralPath (Join-Path $markerFolder 'direct.json') -ErrorAction SilentlyContinue
    # Domain 2 (Other: the fixture jobs, after the Fixture group), entry 3, show the command, back, then the entry by name, run.
    $menuOutput = "2`n3`nW`nB`nprobe-direct`nR`n" | & $powershell5 -NoProfile -ExecutionPolicy Bypass -File $engineScript -Menu -Tenant 'menu' 2>&1 | Out-String
    $marker = Read-Marker -Name 'direct'
    Assert-True ($LASTEXITCODE -eq 7 -and $null -ne $marker -and $marker.Tenant -eq 'menu') "The menu must run the chosen entry and return its exit code (got $LASTEXITCODE)."
    Assert-True ($menuOutput -match 'Domains' -and $menuOutput -match '\s1\s+Fixture \(1\)' -and $menuOutput -match '\s3\s+Probe-Direct' -and $menuOutput -match 'Command: ' -and $menuOutput -notmatch 'Probe-Manual') "The menu must list the domains and entries, hide advanced entries and show the command: $menuOutput"
    Remove-Item -LiteralPath (Join-Path $markerFolder 'direct.json') -ErrorAction SilentlyContinue
    $quitOutput = "A`nQ`n" | & $powershell5 -NoProfile -ExecutionPolicy Bypass -File $engineScript -Menu 2>&1 | Out-String
    Assert-True ($quitOutput -match 'advanced entries shown' -and $quitOutput -match 'Fixture \(2\)') 'A must show the advanced entries.'
    Assert-True ($LASTEXITCODE -eq 0 -and $null -eq (Read-Marker -Name 'direct')) 'Quitting the menu must run nothing.'
    & $powershell5 -NoProfile -ExecutionPolicy Bypass -File $engineScript 'No-Such-Job' *> $null
    Assert-True ($LASTEXITCODE -eq 1) 'An unknown entry must return 1.'
}
finally {
    $env:SMARTM365_ORCHESTRATOR_TENANT = $previousTenant
    Remove-Item Env:\SMARTM365_TEST_MARKERS -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $fixtureBase -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Output ("SmartInventory launch engine offline test passed: {0} checks ({1} legacy launchers and aliases; PowerShell {2})." -f $script:checks, $reference.Count, $PSVersionTable.PSVersion)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAwW2q88515ds8j
# A/7Se5POSMTkLdan6jKcNBVDgzjrcKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJNz5FfPzVOvsJXsHw9yYsxo7BsoSInhf+gTu4z/UzJDMA0GCSqG
# SIb3DQEBAQUABIIBgF3mSeKtgItl2ikl3Py9SCMwE4DnF0L7/HxYOD3Q0BiW5ZpF
# f2hxDfrialxKkGS84vl0VcUX27WRc09JmCVsRI1bDit7+kh3fjdAdmRbHmOmmugZ
# 8S0/tkcDADodq6iCC3UfD0Xo8cXEVgZTCLVyixC6ln2EQYZjmKJ6YY5TrCtblIQY
# H9SEg/7vXHqNwMtMetr7whdN5FOaxLvpNpVdpPMNy7aIBzwbeFTmdfUT2S/YyDeY
# J1zBiN9twxlaZ9BaxKYH86cVqRxG9Pwnq9pOJ2Vb5zmhLCfbB0ek0Wg/mL/1NLVf
# FWVAxwtuLT//mlJDbJi4qFz1kMWIzLwXZlkibs75woRhNw3vCHtCH6abYkVQOmDG
# F6Ljl+7X6tnns8dgFTUA/lKZbzEZPs91fx1T2KAWV1Xgso/IgPR6GMlw4zqwYCcM
# xZZKYgIPpAtchRMcFQVFB3IIYL3iEEx+2TcsiwKZzrbA45LB5fVFDDSWByLSNwqV
# /mRcphrZ/ivaXf2MwKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjYxMTI2
# NTNaMC8GCSqGSIb3DQEJBDEiBCBCZ9lopA4Y9XAmni4QzmqkP2SSLvrYekFr7S40
# XryosjANBgkqhkiG9w0BAQEFAASCAgCkvavWAyGvJZl6ErL79P3x9/+p9SKPdBCQ
# b2smrbocA72aoLudv67LFnOk+ylpWwt5M7U4C6mIIOjHib9XCnGiE8eUHiAiBLYb
# 60WWlGwE1NgkHtW5WXCOH6As05HCwdvwMR4oFQauVjfXQ4Hpc3cRqJYNTcnugcFE
# 0nR+jQwnbQE4LQLXYUh+IHYmaY+EamLVVoCikKZZtTBFeGOysuJ/uQC5RvX1oaiR
# hWLkWxd4NQ/3jOpqJaSmn0WQCVabj+IZCTbmARf5XkTqSJbDxwH4aTl5tizth4I7
# 0LUUpXNdVlEFtGdWRzWtCXOvBQgfsrWsyfuSpd65YaF7N1+Y5vI4UcjP3HkqQwbJ
# SC0SLvHgLxs3h8xFBMwJGaDnacfKC6zrIEOVpaCIapDJrjdrILuI0xFK+KdcPPzN
# JT/8p2356D0ka3B7HhVpHY9Jl3xjJprlIEXBR2V04N72pX/DOmcRDwN8Aq4+jhkH
# HZ8DfewFvFwW5OZKocfLYOY1Zzn3PL/ZuFv2sMzqWtk5cSzNkvJhVXY7uTt/ZJ6r
# TGCQpzPf4dUpOGqmyAtMhJp6DLsxRH/CDuI/W6DwOPv1IbJZicOTYnCy715bglLD
# SA01ZSP0aJrXvfAY7L5+IlXYJpIV+FPvZOVcUTcF1oiVpZQls1Ce90U4NWtM4Emm
# qq6O54GtSA==
# SIG # End signature block
