<#
.SYNOPSIS
Runs or requests a SmartInventory job, or a manual launch, from the jobs catalog.

.DESCRIPTION
Single entry point that replaces the SmartInventory .cmd launchers; SmartInventory.cmd starts it
with Windows PowerShell 5.1. The catalog is
Orchestrator-Jobs.json.template (Jobs and Launches); see SmartM365.Inventory.Launch.psm1.

-Name <Job> [-Now] runs the entry now on this server with its runtime (pwsh 7 or Windows
PowerShell 5.1) and execution mode (Direct or LocalCache), outside the orchestrator.
-Name <Job>[,<Job>...] -Request submits the jobs through SmartM365-Inventory-Pipeline.ps1: the
resident orchestrators run them on the elected server with their dependencies and locks.
-List (or no argument) shows the catalog. -WhatIf shows the resolved command without running anything.
Any other argument is passed to the launched script (or to the pipeline with -Request).

Tenant: -Tenant, then SMARTM365_ORCHESTRATOR_TENANT, then DefaultTenant of
SmartM365.global.local.json. Execution mode: -ExecutionMode, then Execution.Mode of the job,
then OnPremExecutionMode of SmartM365.global.local.json (LocalCache by default).

The exit code of the launched script is returned. Preparation errors return 1.
-List and -WhatIf show the WorkplaceCloudHub banners; a launched script shows its own.

.EXAMPLE
.\Invoke-SmartM365InventoryJob.ps1 -List

.EXAMPLE
.\Invoke-SmartM365InventoryJob.ps1 Exchange2016-Local-Mailboxes-Inventory -WhatIf

.EXAMPLE
.\Invoke-SmartM365InventoryJob.ps1 AD-Inventory -Request -IncludeDependencies

.VERSION
1.1.0
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Position = 0)][string[]]$Name = @(),
    [switch]$List,
    [switch]$Now,
    [switch]$Request,
    [string]$Tenant = '',
    [ValidateSet('', 'Auto', 'Direct', 'LocalCache')][string]$ExecutionMode = '',
    [string]$JobsManifestPath = '',
    [Parameter(ValueFromRemainingArguments)][string[]]$ExtraArguments = @()
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.1.0'

function Write-Host {
    param([Parameter(Position = 0)][AllowEmptyString()][string]$Object = '', [ConsoleColor]$ForegroundColor = 'Gray')
    foreach ($line in ($Object -split "\r?\n")) {
        Microsoft.PowerShell.Utility\Write-Host ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $line) -ForegroundColor $ForegroundColor
    }
}

function Resolve-InventoryJobsManifestPath {
    <# Runtime jobs manifest that overrides job Arguments: the central manifest, else the local one. #>
    param([Parameter(Mandatory)]$Catalog, [Parameter(Mandatory)][string]$TenantValue)
    $orchestratorFolder = Join-Path $Catalog.SmartInventoryRoot 'Orchestrator'
    $localConfigPath = Get-SmartM365JsonReadPath -Path (Join-Path $orchestratorFolder 'SmartM365-Inventory-Orchestrator.local.json') -Optional
    if (-not $localConfigPath) { $localConfigPath = Join-Path $orchestratorFolder 'SmartM365-Inventory-Orchestrator.local.json.template' }
    $localConfig = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($localConfigPath, [Text.Encoding]::UTF8))
    $centralValue = $localConfig.PSObject.Properties['CentralConfigurationEnabled']
    $centralEnabled = ($null -eq $centralValue) -or ([string]$centralValue.Value -in @('True', 'true', '1'))
    if (-not $centralEnabled) { return (Get-SmartM365JsonReadPath -Path (Join-Path $orchestratorFolder 'Orchestrator-Jobs.json') -Optional) }

    $dataFolder = '{{DataAllRootPath}}\Orchestrator'
    $dataFolderValue = $localConfig.PSObject.Properties['OrchestratorDataFolderPath']
    if ($null -ne $dataFolderValue -and -not [string]::IsNullOrWhiteSpace([string]$dataFolderValue.Value) -and [string]$dataFolderValue.Value -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) { $dataFolder = [string]$dataFolderValue.Value }
    if ($dataFolder -match '\{\{') {
        $roots = Get-SmartM365InventoryOutputRoots -SmartM365Root $Catalog.SmartM365Root -Tenant $TenantValue
        foreach ($property in $roots.PSObject.Properties) { $dataFolder = $dataFolder.Replace('{{' + $property.Name + '}}', [string]$property.Value) }
    }
    if ($dataFolder -match '\{\{') { throw "OrchestratorDataFolderPath is not resolved: $dataFolder" }
    $sharedDataFolder = $dataFolder
    if ((Split-Path -Path $sharedDataFolder -Leaf) -eq $env:COMPUTERNAME) { $sharedDataFolder = Split-Path -Path $sharedDataFolder -Parent }
    return (Get-SmartM365JsonReadPath -Path (Join-Path $sharedDataFolder 'Config\Orchestrator-Jobs.json') -Optional)
}

function Wait-InventoryLauncherKey {
    # Same as the pause of the legacy .cmd launchers: only in a user console, never unattended.
    $inputRedirected = $true
    try { $inputRedirected = [Console]::IsInputRedirected } catch { $inputRedirected = $true }
    if ([Environment]::UserInteractive -and -not $inputRedirected -and -not $env:SMARTM365_ORCHESTRATOR_TENANT) {
        [void](Read-Host -Prompt 'Press Enter to close')
    }
}

$exitCode = 0
try {
    . (Join-Path $PSScriptRoot '..\..\Config\SmartM365-TenantContext.ps1')
    $launchModuleFile = Join-Path $PSScriptRoot 'SmartM365.Inventory.Launch.psm1'
    Import-Module -Name $launchModuleFile -Force -ErrorAction Stop

    $catalog = Get-SmartM365InventoryCatalog -SmartInventoryRoot (Split-Path -Path $PSScriptRoot -Parent)
    if ($List -or @($Name).Count -eq 0) {
        Write-SmartM365StartupBanner
        Write-Host ('SmartInventory catalog: {0}' -f $catalog.TemplatePath)
        foreach ($entry in $catalog.Entries) {
            $definition = $entry.Definition
            $runtime = if ($definition.PSObject.Properties['Runtime'] -and $definition.Runtime) { [string]$definition.Runtime } elseif ($definition.PSObject.Properties['PowerShellEdition'] -and $definition.PowerShellEdition) { [string]$definition.PowerShellEdition } else { 'PowerShell7' }
            $execution = if ($definition.PSObject.Properties['Execution'] -and $null -ne $definition.Execution) { 'Execution:' + [string]$definition.Execution.CacheName } else { 'Direct' }
            $description = if ($definition.PSObject.Properties['Description']) { [string]$definition.Description } else { [string]$definition.ScriptPath }
            Write-Host ('{0,-58} {1,-6} {2,-17} {3,-42} {4}' -f $entry.Name, $entry.Kind, $runtime, $execution, $description)
        }
        Write-SmartM365CompletionBanner -Status Success -ScriptName 'Invoke-SmartM365InventoryJob'
        exit 0
    }
    if ($Now -and $Request) { throw 'Choose either -Now or -Request.' }
    $action = if ($Request) { 'Request' } else { 'Now' }

    if (-not [string]::IsNullOrWhiteSpace($JobsManifestPath) -or $action -eq 'Now') {
        $overlayPath = $JobsManifestPath
        if ([string]::IsNullOrWhiteSpace($overlayPath)) {
            try {
                $tenantForOverlay = (Resolve-SmartM365InventoryTenant -Tenant $Tenant -SmartM365Root $catalog.SmartM365Root).Value
                $overlayPath = Resolve-InventoryJobsManifestPath -Catalog $catalog -TenantValue $tenantForOverlay
            }
            catch {
                Write-Host ('[WARN] Runtime jobs manifest not resolved; template Arguments are used. {0}' -f $_.Exception.Message) -ForegroundColor Yellow
                $overlayPath = ''
            }
        }
        if ($overlayPath) { $catalog = Get-SmartM365InventoryCatalog -SmartInventoryRoot $catalog.SmartInventoryRoot -OverlayPath $overlayPath }
    }

    $plan = Resolve-SmartM365InventoryLaunch -Catalog $catalog -Name $Name -Action $action -Tenant $Tenant -ExecutionMode $ExecutionMode -ExtraArguments $ExtraArguments
    if ($WhatIfPreference) {
        Write-SmartM365StartupBanner
        Write-Host (Format-SmartM365InventoryPlan -Plan $plan -Catalog $catalog)
        Write-Host ('Catalog: {0}; runtime manifest: {1}' -f $catalog.TemplatePath, $(if ($catalog.OverlayPath) { $catalog.OverlayPath } else { 'none' }))
        Write-SmartM365CompletionBanner -Status Success -ScriptName 'Invoke-SmartM365InventoryJob'
        exit 0
    }
    $exitCode = Invoke-SmartM365InventoryLaunch -Plan $plan -Catalog $catalog
    if ($exitCode -ne 0) { Write-Host ('{0} ended with exit code {1}.' -f $plan.Name, $exitCode) -ForegroundColor Yellow }
}
catch {
    Write-Host ('[ERROR] {0}' -f $_.Exception.Message) -ForegroundColor Red
    Wait-InventoryLauncherKey
    $exitCode = 1
}
exit $exitCode
