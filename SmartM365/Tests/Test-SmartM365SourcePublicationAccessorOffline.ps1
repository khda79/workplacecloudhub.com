#Requires -Version 7.0
<#
.SYNOPSIS
Synthetic checks of the read-only current source publication accessor and run-state notes. No tenant, API, mail or synchronized data access.
.VERSION
1.0.0
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic Core identity globals are saved and restored in finally.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.90' -Force -PassThru
$producerScript = Get-ChildItem (Join-Path $root 'SmartInventory') -Recurse -File -Filter 'SmartM365-LicensePricing-Inventory.ps1' | Select-Object -First 1
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-PublicationAccessor-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporary | Out-Null
$saved = @{}
foreach ($name in @('SmartM365TenantKey','SmartM365OrganizationKey','SmartM365EnvironmentKey','SmartM365TenantId','csvGeneratedPaths')) {
  $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
  $saved[$name] = @{ Exists=($null -ne $variable); Value=$(if ($variable) { $variable.Value } else { $null }) }
}
$global:SmartM365TenantKey = 'synthetic'; $global:SmartM365OrganizationKey = 'contoso'
$global:SmartM365EnvironmentKey = 'test'; $global:SmartM365TenantId = 'synthetic-tenant'
$global:csvGeneratedPaths = $null
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
try {
  foreach ($name in @('Register-SmartM365SourceCsv','Get-SmartM365CurrentSourcePublication','Add-SmartM365SourceRunNote')) {
    Check ([bool](Get-Command -Name $name -Module $core.Name -ErrorAction SilentlyContinue)) "SmartM365.Core does not export $name."
  }
  Check ($null -eq (Get-SmartM365CurrentSourcePublication)) 'No publication is reported before a receipt starts.'
  Start-SmartM365SourceReceipt -ScriptPath $producerScript.FullName -SourceRootPath $temporary
  $runPath = Join-Path $temporary 'SmartInventory_SmartM365-LicensePricing-Inventory.run.json.txt'
  $running = Get-Content -LiteralPath $runPath -Raw | ConvertFrom-Json
  $empty = Get-SmartM365CurrentSourcePublication
  Check ($empty.RunId -eq $running.RunId -and $empty.Producer -eq 'SmartM365-LicensePricing-Inventory.ps1' -and $empty.Hashes.Count -eq 0) 'Accessor reports the active run before publication.'
  $csv = Join-Path $temporary 'M365_License_Prices.csv'
  [IO.File]::WriteAllText($csv, "TenantKey,SkuPartNumber`r`nsynthetic,SPE_E3`r`n", [Text.UTF8Encoding]::new($false))
  Register-SmartM365SourceCsv -Path $csv
  $publication = Get-SmartM365CurrentSourcePublication
  $fullPath = [IO.Path]::GetFullPath($csv)
  Check ($publication.Hashes[$fullPath] -eq (Get-FileHash -LiteralPath $csv -Algorithm SHA256).Hash) 'Accessor returns the hash recorded at publication.'
  Check ($publication.Hashes.ContainsKey($fullPath.ToUpperInvariant())) 'Accessor hashes are path case-insensitive.'
  $publication.Hashes[$fullPath] = 'CHANGED'
  $publication.Hashes['X:\other.csv'] = 'ADDED'
  $again = Get-SmartM365CurrentSourcePublication
  Check ($again.Hashes[$fullPath] -ne 'CHANGED' -and -not $again.Hashes.ContainsKey('X:\other.csv')) 'Accessor returns a copy, not the receipt state.'
  Add-SmartM365SourceRunNote -Message 'Synthetic output not republished in this run.'
  $noted = Get-Content -LiteralPath $runPath -Raw | ConvertFrom-Json
  Check ($noted.Status -eq 'Publishing' -and @($noted.Notes) -contains 'Synthetic output not republished in this run.' -and $noted.RunId -eq $running.RunId) 'Run note is written immediately to the run state.'
  $receiptPath = Complete-SmartM365SourceReceipt -Status 'Completed'
  $receipt = Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
  $final = Get-Content -LiteralPath $runPath -Raw | ConvertFrom-Json
  Check ($receipt.Status -eq 'Completed' -and $receipt.RunId -eq $publication.RunId) 'Completed receipt keeps the accessor run identity.'
  Check ($final.Status -eq 'Completed' -and @($final.Notes).Count -eq 1) 'Final run state retains the run note.'
  Check ($null -eq $receipt.PSObject.Properties['Notes']) 'Run notes stay in the run state, not in the completion proof.'
  Check ($null -eq (Get-SmartM365CurrentSourcePublication)) 'No publication is reported after completion.'
  Add-SmartM365SourceRunNote -Message 'ignored'
  Check (@((Get-Content -LiteralPath $runPath -Raw | ConvertFrom-Json).Notes).Count -eq 1) 'A note without an active receipt is ignored.'
}
finally {
  foreach ($name in $saved.Keys) {
    if ($saved[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $saved[$name].Value }
    else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
  }
  if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}
"PASS: $script:checks source publication accessor checks; PowerShell $($PSVersionTable.PSVersion)."
