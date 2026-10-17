#Requires -Version 7.0
<#
.SYNOPSIS
Synthetic checks of the LicensePricing public baseline, private override merge, legacy replacement mode,
Windows 10 ESU price validation and optional receipt registration. No tenant, API, mail or synchronized data access.
.VERSION
1.0.0
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic Core identity globals are saved and restored in finally.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$licensingFolder = Join-Path $root 'SmartInventory/M365Inventory/Licensing'
$scriptPath = Join-Path $licensingFolder 'SmartM365-LicensePricing-Inventory.ps1'
$baselinePath = Join-Path $licensingFolder 'M365_License_Prices_Public-France.csv'
# Real production modules: a missing export or helper fails here as it would in the collector.
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.90' -Force -PassThru
Import-Module (Join-Path $root 'SmartInventory/Common/SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.11' -Force
$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'LicensePricing script does not parse.' }
$names = @('ConvertTo-LicensePriceRows','Merge-LicensePriceRows','ConvertTo-WindowsEsuPriceRows','Resolve-LicensePricingRows')
$definitions = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $names -contains $n.Name }, $true))
if ($definitions.Count -ne $names.Count) { throw 'LicensePricing functions are missing.' }
foreach ($definition in $definitions) { Invoke-Expression $definition.Extent.Text }
# Read by the extracted collector functions (SourceScriptVersion).
$ScriptVersion = '1.1.0'
if ($ast.Extent.Text -notmatch "\`$ScriptVersion='1\.1\.0'" -or $ast.Extent.Text -notmatch '(?m)^\.VERSION\s*\r?\n1\.1\.0') { throw 'Header and runtime versions must both be 1.1.0.' }

$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-LicensePricing-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporary | Out-Null
$saved = @{}
foreach ($name in @('SmartM365TenantKey','SmartM365OrganizationKey','SmartM365EnvironmentKey','SmartM365TenantId','csvGeneratedPaths')) {
  $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
  $saved[$name] = @{ Exists=($null -ne $variable); Value=$(if ($variable) { $variable.Value } else { $null }) }
}
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
function Assert-Throws { param([scriptblock]$Action, [string]$Like, [string]$Message)
  $threw = $false; $text = ''
  try { & $Action | Out-Null } catch { $threw = $true; $text = $_.Exception.Message }
  Check ($threw -and $text -like $Like) "$Message (threw=$threw; message='$text')"
}
function Write-Fixture { param([string]$Name, [string[]]$Lines) $path = Join-Path $temporary $Name; [IO.File]::WriteAllText($path, (($Lines -join "`r`n") + "`r`n"), [Text.UTF8Encoding]::new($false)); return $path }
$priceHeader = 'SkuPartNumber,Currency,UnitPriceMonthly,BillingCadence,EffectiveFrom,EffectiveTo,PriceSource,PricingBasis,ContractReference,IsEstimated'
$esuHeader = 'ProgramKey,CoverageYear,Currency,UnitPricePerDeviceYear,EffectiveFrom,EffectiveTo,IsCumulative,PriceSource,PricingBasis,ContractReference,IsEstimated'
try {
  # 1. Public baseline: existing prices kept, eleven verified additions, all EUR estimates.
  $baseline = @(Import-Csv -LiteralPath $baselinePath)
  Check ($baseline.Count -eq 34) "Public baseline must have 34 rows, got $($baseline.Count)."
  Check (@($baseline | Where-Object { $_.Currency -ne 'EUR' -or $_.IsEstimated -ne 'True' -or $_.BillingCadence -ne 'Monthly' }).Count -eq 0) 'Public baseline rows are EUR monthly estimates.'
  $expected = [ordered]@{
    'SPE_E3'='37.78'; 'SPE_F1'='8.66'; 'Microsoft_365_Copilot'='26.00'
    'SPE_E5'='58.13'; 'M365_F1'='2.60'; 'M365_F1_COMM'='2.60'; 'MCOEV'='8.70'; 'M365_Copilot'='26.00'; 'EXCHANGEARCHIVE_ADDON'='2.60'
    'O365_w/o_Teams_Bundle_M5'='50.71'; 'Microsoft_365_F3_EEA_(no_Teams)'='7.73'; 'Microsoft_365_F1_EEA_(no_Teams)'='2.16'
    'Microsoft_Teams_EEA_New'='7.41'; 'Microsoft_Teams_Premium'='8.70'
  }
  foreach ($sku in $expected.Keys) { Check (@($baseline | Where-Object { $_.SkuPartNumber -ceq $sku -and $_.UnitPriceMonthly -eq $expected[$sku] }).Count -eq 1) "Baseline price for $sku." }
  Check (@($baseline | Where-Object EffectiveFrom -eq '2026-10-17').Count -eq 11) 'Eleven rows were observed on 2026-10-17.'
  Check (@($baseline | Where-Object SkuPartNumber -ceq 'Microsoft_Teams_Premium')[0].PricingBasis -like '*introductory price not applied*') 'Teams Premium introductory price is explicitly not applied.'
  foreach ($absent in @('EXCHANGEARCHIVE','SHAREPOINTSTORAGE','Microsoft_Intune_Suite','PROJECTPREMIUM','POWERAPPS_PER_APP_NEW')) { Check (@($baseline | Where-Object SkuPartNumber -ceq $absent).Count -eq 0) "$absent stays Not valued." }

  # 2. Public baseline alone.
  $public = Resolve-LicensePricingRows -SourcePath $baselinePath -PublicBaselinePath $baselinePath -RunId r -CollectedAtUtc c
  Check ($public.Mode -eq 'PublicBaseline' -and $public.Rows.Count -eq 34 -and $public.Messages[0].Level -eq 'INFO') 'Public baseline mode.'

  $baselineCopy = Join-Path $temporary 'baseline-copy.csv'; Copy-Item -LiteralPath $baselinePath -Destination $baselineCopy
  $copied = Resolve-LicensePricingRows -SourcePath $baselineCopy -PublicBaselinePath $baselinePath -RunId r -CollectedAtUtc c
  Check ($copied.Mode -eq 'PublicBaseline' -and $copied.Rows.Count -eq 34) 'A byte-identical baseline reached through another path is not legacy mode.'

  # 3. Private override replaces every public period of the same SKU and currency only.
  $override = Write-Fixture 'override.csv' @($priceHeader,
    'SPE_E3,EUR,30.00,Monthly,2026-01-01,2026-12-31,Contoso agreement,Negotiated,CONTOSO-EA-1,False',
    'SPE_E3,EUR,31.00,Monthly,2027-01-01,,Contoso agreement,Negotiated,CONTOSO-EA-1,False',
    'SPE_F1,USD,9.00,Monthly,2026-01-01,,Contoso agreement,Negotiated,CONTOSO-EA-1,False',
    'CONTOSO_ADDON,EUR,4.00,Monthly,2026-01-01,,Contoso agreement,Negotiated,CONTOSO-EA-1,False')
  $merged = Resolve-LicensePricingRows -SourcePath $baselinePath -OverridePath $override -PublicBaselinePath $baselinePath -RunId r -CollectedAtUtc c
  Check ($merged.Mode -eq 'PublicBaselineWithOverride' -and $merged.Rows.Count -eq 37) "Merged rows: 34 - 1 + 4 = 37, got $($merged.Rows.Count)."
  Check (@($merged.Rows | Where-Object { $_.SkuPartNumber -eq 'SPE_E3' -and $_.Currency -eq 'EUR' }).Count -eq 2 -and @($merged.Rows | Where-Object { $_.SkuPartNumber -eq 'SPE_E3' -and $_.IsEstimated -eq 'True' }).Count -eq 0) 'Private SPE_E3 EUR replaces all public SPE_E3 EUR periods.'
  Check (@($merged.Rows | Where-Object { $_.SkuPartNumber -eq 'SPE_F1' -and $_.Currency -eq 'EUR' -and $_.UnitPriceMonthly -eq 8.66 }).Count -eq 1) 'A private price in another currency does not replace the public EUR price.'
  Check (@($merged.Rows | Where-Object SkuPartNumber -eq 'CONTOSO_ADDON').Count -eq 1) 'Private-only SKU is added.'
  Check ($merged.Messages[0].Text -like '*1 public row(s) replaced*') 'Merge counts are logged.'
  $lowerCase = Write-Fixture 'override-case.csv' @($priceHeader, 'spe_e5,EUR,50.00,Monthly,2026-01-01,,Contoso agreement,Negotiated,CONTOSO-EA-1,False')
  $caseMerge = Resolve-LicensePricingRows -SourcePath $baselinePath -OverridePath $lowerCase -PublicBaselinePath $baselinePath -RunId r -CollectedAtUtc c
  Check (@($caseMerge.Rows | Where-Object { $_.SkuPartNumber -ieq 'SPE_E5' }).Count -eq 1) 'SKU matching is case-insensitive.'

  # 4. Legacy: PricingSourceCsvPath pointing to a private file keeps whole replacement with a warning.
  $legacy = Resolve-LicensePricingRows -SourcePath $override -PublicBaselinePath $baselinePath -RunId r -CollectedAtUtc c
  Check ($legacy.Mode -eq 'LegacyReplacement' -and $legacy.Rows.Count -eq 4) 'Legacy private source replaces the whole baseline.'
  Check ($legacy.Messages[0].Level -eq 'WARNING' -and $legacy.Messages[0].Text -like '*Migrate private prices to PricingOverrideCsvPath*') 'Legacy mode recommends the migration.'
  Assert-Throws { Resolve-LicensePricingRows -SourcePath $override -OverridePath $override -PublicBaselinePath $baselinePath -RunId r -CollectedAtUtc c } '*PricingOverrideCsvPath is also set*' 'Legacy source and override together are refused.'
  Assert-Throws { Resolve-LicensePricingRows -SourcePath $baselinePath -OverridePath (Join-Path $temporary 'missing.csv') -PublicBaselinePath $baselinePath -RunId r -CollectedAtUtc c } '*must reference an existing private CSV*' 'Missing override is refused.'
  $emptyOverride = Write-Fixture 'override-empty.csv' @($priceHeader)
  Assert-Throws { Resolve-LicensePricingRows -SourcePath $baselinePath -OverridePath $emptyOverride -PublicBaselinePath $baselinePath -RunId r -CollectedAtUtc c } '*contains no rows*' 'Empty override is refused.'
  $duplicateOverride = Write-Fixture 'override-duplicate.csv' @($priceHeader, 'SPE_E3,EUR,30.00,Monthly,2026-01-01,,A,B,C,False', 'SPE_E3,EUR,30.50,Monthly,2026-01-01,,A,B,C,False')
  Assert-Throws { Resolve-LicensePricingRows -SourcePath $baselinePath -OverridePath $duplicateOverride -PublicBaselinePath $baselinePath -RunId r -CollectedAtUtc c } '*Duplicate license price period*' 'Duplicate override periods are refused.'

  # 5. Windows 10 ESU: private EUR source only, explicit cumulative years, no conversion.
  $esuLines = @($esuHeader,
    'WIN10_ESU,1,EUR,55.00,2025-11-01,2026-10-31,True,Contoso agreement,Per device per ESU year,CONTOSO-ESU-1,False',
    'WIN10_ESU,2,EUR,110.00,2026-11-01,2027-10-31,True,Contoso agreement,Per device per ESU year,CONTOSO-ESU-1,False',
    'WIN10_ESU,3,EUR,220.00,2027-11-01,2028-10-31,True,Contoso agreement,Per device per ESU year,CONTOSO-ESU-1,False')
  $esu = @(ConvertTo-WindowsEsuPriceRows @(Import-Csv -LiteralPath (Write-Fixture 'esu.csv' $esuLines)) r c)
  Check ($esu.Count -eq 3 -and $esu[1].UnitPricePerDeviceYear -eq 110 -and $esu[1].IsCumulative -eq 'True' -and $esu[0].SourceScriptVersion -eq '1.1.0') 'Valid ESU years are normalized.'
  $invalid = [ordered]@{
    'USD price' = @($esuHeader, 'WIN10_ESU,1,USD,61.00,2025-11-01,2026-10-31,True,A,B,C,True')
    'Unknown program' = @($esuHeader, 'WIN11_ESU,1,EUR,1.00,2025-11-01,2026-10-31,True,A,B,C,True')
    'Year 4' = @($esuHeader, 'WIN10_ESU,4,EUR,1.00,2028-11-01,2029-10-31,True,A,B,C,True')
    'Negative price' = @($esuHeader, 'WIN10_ESU,1,EUR,-1,2025-11-01,2026-10-31,True,A,B,C,True')
    'Missing end date' = @($esuHeader, 'WIN10_ESU,1,EUR,1.00,2025-11-01,,True,A,B,C,True')
    'Reversed period' = @($esuHeader, 'WIN10_ESU,1,EUR,1.00,2026-11-01,2025-10-31,True,A,B,C,True')
    'Overlap' = @($esuHeader, 'WIN10_ESU,1,EUR,1.00,2025-11-01,2026-12-31,True,A,B,C,True', 'WIN10_ESU,2,EUR,2.00,2026-11-01,2027-10-31,True,A,B,C,True')
    'Duplicate year' = @($esuHeader, 'WIN10_ESU,1,EUR,1.00,2025-11-01,2026-10-31,True,A,B,C,True', 'WIN10_ESU,1,EUR,2.00,2026-11-01,2027-10-31,True,A,B,C,True')
    'Years out of order' = @($esuHeader, 'WIN10_ESU,1,EUR,1.00,2026-11-01,2027-10-31,True,A,B,C,True', 'WIN10_ESU,2,EUR,2.00,2025-11-01,2026-10-31,True,A,B,C,True')
    'Cumulative flag' = @($esuHeader, 'WIN10_ESU,1,EUR,1.00,2025-11-01,2026-10-31,Yes,A,B,C,True')
  }
  foreach ($case in $invalid.Keys) {
    $path = Write-Fixture ("esu-$($case -replace '\W','')" + '.csv') $invalid[$case]
    Assert-Throws { ConvertTo-WindowsEsuPriceRows @(Import-Csv -LiteralPath $path) r c } '*' "ESU case is refused: $case"
  }
  Assert-Throws { ConvertTo-WindowsEsuPriceRows @(Import-Csv -LiteralPath (Write-Fixture 'esu-usd.csv' $invalid['USD price'])) r c } '*must be in EUR*' 'USD ESU prices are refused without conversion.'

  # 6. Receipt registration with the real Core module: the ESU dataset is an optional output.
  $global:SmartM365TenantKey = 'synthetic'; $global:SmartM365OrganizationKey = 'contoso'; $global:SmartM365EnvironmentKey = 'test'; $global:SmartM365TenantId = 'synthetic-tenant'
  $latest = Join-Path $temporary 'DATA-LAST'; New-Item -ItemType Directory -Path $latest | Out-Null
  foreach ($withEsu in @($true, $false)) {
    $global:csvGeneratedPaths = $null
    Start-SmartM365SourceReceipt -ScriptPath $scriptPath -SourceRootPath $latest
    $pricesPath = Join-Path $latest 'M365_License_Prices.csv'
    [IO.File]::WriteAllText($pricesPath, "TenantKey,SkuPartNumber`r`nsynthetic,SPE_E3`r`n", [Text.UTF8Encoding]::new($false))
    Register-SmartM365SourceCsv -Path $pricesPath
    if ($withEsu) {
      $esuPath = Join-Path $latest 'Windows_ESU_Prices.csv'
      [IO.File]::WriteAllText($esuPath, "TenantKey,ProgramKey`r`nsynthetic,WIN10_ESU`r`n", [Text.UTF8Encoding]::new($false))
      Register-SmartM365SourceCsv -Path $esuPath
    }
    $receipt = Get-Content -LiteralPath (Complete-SmartM365SourceReceipt -Status 'Completed') -Raw | ConvertFrom-Json
    $optional = @($receipt.OptionalFiles | Where-Object File -eq 'Windows_ESU_Prices.csv')
    Check ($receipt.Status -eq 'Completed' -and @($receipt.RequiredFiles) -contains 'M365_License_Prices.csv' -and $optional.Count -eq 1) 'Receipt declares the ESU dataset as optional.'
    Check ([bool]$optional[0].Published -eq $withEsu -and (@($receipt.Files | Where-Object File -eq 'Windows_ESU_Prices.csv').Count -eq [int]$withEsu)) "ESU publication is recorded only when published (published=$withEsu)."
  }

  # 7. Template exposes the private keys; the collector creates and completes its runtime JSON.
  $template = Get-Content -LiteralPath (Join-Path $licensingFolder 'SmartM365-LicensePricing-Inventory.local.json.template') -Raw | ConvertFrom-Json
  Check ($template.PSObject.Properties['PricingOverrideCsvPath'] -and $template.PricingOverrideCsvPath -eq '' -and $template.PSObject.Properties['EsuPricingCsvPath'] -and $template.EsuPricingCsvPath -eq '') 'Template exposes empty private override and ESU keys.'
  Check ($ast.Extent.Text -match "Read-SmartM365JsonConfig -Path \(Join-Path \`$PSScriptRoot 'SmartM365-LicensePricing-Inventory\.local\.json'\) -Required") 'Runtime JSON is created and completed from the template.'
  Check ([bool](Get-Command -Name Register-SmartM365SourceCsv -Module $core.Name -ErrorAction SilentlyContinue)) 'Core exports the registration used by the receipt.'
}
finally {
  foreach ($name in $saved.Keys) {
    if ($saved[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $saved[$name].Value }
    else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
  }
  if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}
"PASS: $script:checks license pricing merge, ESU and receipt checks; PowerShell $($PSVersionTable.PSVersion)."
