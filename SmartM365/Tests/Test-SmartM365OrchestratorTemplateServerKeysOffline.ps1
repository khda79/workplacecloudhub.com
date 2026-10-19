#requires -Version 7.0
<#
.SYNOPSIS
Checks that the committed orchestrator local template never injects example server names into a runtime configuration.

.DESCRIPTION
Template enrichment adds missing keys recursively inside nested objects. A runtime configuration that only lists
the real servers is synchronised with the committed template in a temporary folder; no server-keyed entry may be
added. A control template carrying example server keys proves that the check detects the regression.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'

$smartM365Root = Split-Path -Path $PSScriptRoot -Parent
$orchestratorRoot = Join-Path -Path $smartM365Root -ChildPath 'SmartInventory\Orchestrator'
$templatePath = Join-Path -Path $orchestratorRoot -ChildPath 'SmartM365-Inventory-Orchestrator.local.json.template'
. (Join-Path -Path $smartM365Root -ChildPath 'Config\SmartM365-TenantContext.ps1')

$script:passed = 0
function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
    $script:passed++
}

$serverKeyedObjects = @('MaxConcurrencyByServer', 'ElectionWeightsByServer', 'ServerJobPolicies')
$serverKeyedLists = @('AllowedServers', 'ExpectedOrchestratorServers')
$realServers = @('REAL-SRV-A', 'REAL-SRV-B', 'REAL-EXCH-C')

$template = Get-Content -LiteralPath $templatePath -Raw | ConvertFrom-Json -ErrorAction Stop
foreach ($name in $serverKeyedObjects) {
    $property = $template.PSObject.Properties[$name]
    Assert-True -Condition ($null -ne $property -and $property.Value -is [pscustomobject] -and @($property.Value.PSObject.Properties).Count -eq 0) -Message "Template setting $name must be an empty object."
}
foreach ($name in $serverKeyedLists) {
    $property = $template.PSObject.Properties[$name]
    Assert-True -Condition ($null -ne $property -and @($property.Value).Count -eq 0) -Message "Template setting $name must be an empty list."
}

function New-RuntimeConfiguration {
    param([Parameter(Mandatory)][string]$Path)
    $runtime = [ordered]@{
        AllowedServers = $realServers
        ExpectedOrchestratorServers = $realServers
        MaxConcurrency = 6
        MaxConcurrencyByServer = [ordered]@{ 'REAL-SRV-A' = 6; 'REAL-SRV-B' = 9; 'REAL-EXCH-C' = 6 }
        ElectionWeightsByServer = [ordered]@{ 'REAL-SRV-A' = 1.0; 'REAL-SRV-B' = 1.1; 'REAL-EXCH-C' = 1.0 }
        ServerJobPolicies = [ordered]@{ 'REAL-EXCH-C' = [ordered]@{ OnlyJobsRequiring = @('ExchangeOnPrem') } }
    }
    [IO.File]::WriteAllText($Path, ($runtime | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
}

function Get-ServerKeys {
    param([Parameter(Mandatory)]$Config)
    $keys = foreach ($name in $serverKeyedObjects) {
        foreach ($entry in $Config.$name.PSObject.Properties) { '{0}.{1}' -f $name, $entry.Name }
    }
    $keys += foreach ($name in $serverKeyedLists) { foreach ($value in @($Config.$name)) { '{0}[{1}]' -f $name, $value } }
    return @($keys | Sort-Object)
}

$temporaryRoot = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ('SmartM365-OrchestratorTemplateServerKeys-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $temporaryRoot
try {
    $runtimePath = Join-Path -Path $temporaryRoot -ChildPath 'SmartM365-Inventory-Orchestrator.local.json.txt'
    New-RuntimeConfiguration -Path $runtimePath
    $before = Get-Content -LiteralPath $runtimePath -Raw | ConvertFrom-Json
    $expectedKeys = Get-ServerKeys -Config $before

    $synced = Sync-SmartM365JsonConfigWithTemplate -Config $before -Path $runtimePath -TemplatePath $templatePath 6>$null
    $reloaded = Get-Content -LiteralPath $runtimePath -Raw | ConvertFrom-Json
    Assert-True -Condition (((Get-ServerKeys -Config $synced) -join '|') -ceq ($expectedKeys -join '|')) -Message 'Template enrichment added server-keyed entries in memory.'
    Assert-True -Condition (((Get-ServerKeys -Config $reloaded) -join '|') -ceq ($expectedKeys -join '|')) -Message 'Template enrichment added server-keyed entries to the runtime file.'
    Assert-True -Condition ($null -ne $reloaded.PSObject.Properties['MaxLifetimeHours']) -Message 'Template enrichment did not run: a missing ordinary key was not added.'

    $secondPass = Sync-SmartM365JsonConfigWithTemplate -Config $reloaded -Path $runtimePath -TemplatePath $templatePath 6>$null
    Assert-True -Condition (((Get-ServerKeys -Config $secondPass) -join '|') -ceq ($expectedKeys -join '|')) -Message 'A second start added server-keyed entries.'

    # Control: a template with example server keys must be detected by the same check.
    $controlTemplatePath = Join-Path -Path $temporaryRoot -ChildPath 'control.local.json.template'
    $controlTemplate = Get-Content -LiteralPath $templatePath -Raw | ConvertFrom-Json
    $controlTemplate.MaxConcurrencyByServer = [pscustomobject]@{ 'EXAMPLE-SRV' = 6 }
    [IO.File]::WriteAllText($controlTemplatePath, ($controlTemplate | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
    $controlRuntimePath = Join-Path -Path $temporaryRoot -ChildPath 'control.local.json.txt'
    New-RuntimeConfiguration -Path $controlRuntimePath
    $control = Sync-SmartM365JsonConfigWithTemplate -Config (Get-Content -LiteralPath $controlRuntimePath -Raw | ConvertFrom-Json) -Path $controlRuntimePath -TemplatePath $controlTemplatePath 6>$null
    Assert-True -Condition ($null -ne $control.MaxConcurrencyByServer.PSObject.Properties['EXAMPLE-SRV']) -Message 'Control template did not reproduce the example-key injection.'

    Write-Host ("SmartM365 orchestrator template server-key tests passed ({0} checks)." -f $script:passed)
}
finally {
    Remove-Item -LiteralPath $temporaryRoot -Recurse -Force -ErrorAction SilentlyContinue
}
