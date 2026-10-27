# Version 1.0.0
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$root=Split-Path $PSScriptRoot -Parent
$catalog=Get-Content (Join-Path $root 'SmartInventory/Orchestrator/Orchestrator-Jobs.json.template') -Raw | ConvertFrom-Json
$retired=@('Exchange-HybridIdentity-Issues','Intune-Windows11-Readiness-Issues')
foreach($job in $catalog.Jobs){
    if($job.Name -in $retired){throw 'Retired job is still active.'}
    if($job.PSObject.Properties['DependsOn'] -and @($job.DependsOn | Where-Object {$_ -in $retired}).Count){throw 'Retired dependency is still active.'}
}
$cleanup=Join-Path $root 'SmartInventory/Maintenance/Invoke-SmartM365-ServerDataCleanup.ps1'
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($cleanup,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Cleanup parser failed.'}
$function=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-TimestampedCsvDecisions'},$true)
. ([scriptblock]::Create($function.Extent.Text))
$tempRoot=Join-Path ([IO.Path]::GetTempPath()) ('legacy-preservation-'+[guid]::NewGuid().ToString('N'))
$dataLast=Join-Path $tempRoot 'DATA-LAST';$dataAll=Join-Path $tempRoot 'DATA-ALL'
$null=New-Item -ItemType Directory -Path $dataLast,$dataAll
try {
    foreach($prefix in @('Exchange_HybridIdentity_Issues','Exchange_HybridIdentity_Issues_Summary','Intune_Windows11_Readiness_Issues','Intune_Windows11_Readiness_Issues_Summary')) {
        '"TenantKey"' | Set-Content (Join-Path $dataLast ($prefix+'.csv'))
        '"TenantKey"' | Set-Content (Join-Path $dataAll ($prefix+'_20200101_010000.csv'))
    }
    $before=@(Get-ChildItem $dataAll -File | Get-FileHash | Select-Object Path,Hash)
    $decisions=@(Get-TimestampedCsvDecisions)
    if($decisions.Count){throw 'Retired CSV was eligible for timestamped cleanup.'}
    foreach($entry in $before){if((Get-FileHash $entry.Path).Hash -ne $entry.Hash){throw 'History changed.'}}
    'PASS: retired jobs/dependencies absent; all four legacy CSV families preserved.'
} finally {
    $resolved=[IO.Path]::GetFullPath($tempRoot)
    if(-not $resolved.StartsWith([IO.Path]::GetTempPath(),[StringComparison]::OrdinalIgnoreCase)){throw 'Unsafe test cleanup.'}
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
