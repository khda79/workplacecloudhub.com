<#
.SYNOPSIS
    Regenerates the configuration templates from the option catalog.

.DESCRIPTION
    Writes Config\SmartEndpointMigrationToolkit.local.json.template (every option, inherit marker)
    and Campaigns\<Campaign>\Templates\Wave\Wave.local.json.template (options applicable to the
    campaign). Tests fail when the committed templates differ from this generator.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param([switch]$Check)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$packageRoot = Split-Path -Path $PSScriptRoot -Parent
Import-Module (Join-Path -Path $packageRoot -ChildPath 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
$catalog = Import-SemtOptionCatalog -PackageRoot $packageRoot

function ConvertTo-TemplateText {
    param([Parameter(Mandatory = $true)]$Object)
    return (($Object | ConvertTo-Json -Depth 6) -replace "`r`n","`n") + "`n"
}

$targets = New-Object System.Collections.Generic.List[object]
$globalOptions = [ordered]@{}
foreach ($definition in @($catalog.Options)) { $globalOptions[[string]$definition.Name] = '__INHERIT__' }
$targets.Add([pscustomobject]@{
    Path = Join-Path $packageRoot 'Config\SmartEndpointMigrationToolkit.local.json.template'
    Text = ConvertTo-TemplateText -Object ([ordered]@{
        SchemaVersion = 1
        Product = 'SmartEndpointMigrationToolkit'
        Description = 'Global options of this technician installation. __INHERIT__ keeps the campaign default of Config\SmartEndpointMigrationToolkit.options.json.txt. A Wave, a SEMT_* environment variable, a GUI choice or an explicit CLI option overrides it.'
        Options = $globalOptions
    })
})
foreach ($campaign in @('HybridJoin','Windows11Upgrade')) {
    $waveOptions = [ordered]@{}
    foreach ($definition in @(Get-SemtCampaignOptionDefinitions -Catalog $catalog -Campaign $campaign)) { $waveOptions[[string]$definition.Name] = '__INHERIT__' }
    $targets.Add([pscustomobject]@{
        Path = Join-Path $packageRoot ("Campaigns\{0}\Templates\Wave\Wave.local.json.template" -f $campaign)
        Text = ConvertTo-TemplateText -Object ([ordered]@{
            SchemaVersion = 1
            Product = 'SmartEndpointMigrationToolkit'
            Campaign = $campaign
            Description = ''
            Source = 'Manual'
            CreatedUtc = ''
            Options = $waveOptions
        })
    })
}

$different = @()
foreach ($target in $targets) {
    $current = if (Test-Path -LiteralPath $target.Path -PathType Leaf) { ([System.IO.File]::ReadAllText($target.Path) -replace "`r`n","`n") } else { '' }
    # Compared as JSON content: Windows PowerShell 5.1 and PowerShell 7 indent ConvertTo-Json
    # differently, and a template must not be reported out of date for that reason alone.
    $same = $false
    if ($current) {
        try { $same = (($current | ConvertFrom-Json | ConvertTo-Json -Depth 10 -Compress) -ceq ($target.Text | ConvertFrom-Json | ConvertTo-Json -Depth 10 -Compress)) } catch { $same = $false }
    }
    if (-not $same) {
        $different += $target.Path
        if (-not $Check) {
            New-Item -ItemType Directory -Path (Split-Path -Path $target.Path -Parent) -Force | Out-Null
            [System.IO.File]::WriteAllText($target.Path, ($target.Text -replace "`n","`r`n"), (New-Object System.Text.UTF8Encoding($false)))
        }
    }
}
if ($Check -and $different.Count -gt 0) { throw ("Configuration templates are out of date: {0}" -f ($different -join ', ')) }
[pscustomobject]@{ Checked = $targets.Count; Updated = $(if ($Check) { 0 } else { $different.Count }); Different = $different }
