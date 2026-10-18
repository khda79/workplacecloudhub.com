#Requires -Version 7.0
<#
.SYNOPSIS
Synthetic checks of the LicensePricing public baseline, private override merge, legacy replacement mode,
Windows 10 ESU price validation and optional receipt registration. No tenant, API, mail or synchronized data access.
.VERSION
1.1.0
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
$ScriptVersion = '1.1.1'
if ($ast.Extent.Text -notmatch "\`$ScriptVersion='1\.1\.1'" -or $ast.Extent.Text -notmatch '(?m)^\.VERSION\s*\r?\n1\.1\.1') { throw 'Header and runtime versions must both be 1.1.1.' }

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
  Check ($esu.Count -eq 3 -and $esu[1].UnitPricePerDeviceYear -eq 110 -and $esu[1].IsCumulative -eq 'True' -and $esu[0].SourceScriptVersion -eq '1.1.1') 'Valid ESU years are normalized.'
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

  # 7. Invariant prices whatever the host culture (fr-FR exported a double as "37,78").
  $savedCulture = [Globalization.CultureInfo]::CurrentCulture
  try {
    [Globalization.CultureInfo]::CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')
    Check ((1.5).ToString() -eq '1,5') 'The fr-FR culture is active for the following checks.'
    $frRows = @((Resolve-LicensePricingRows -SourcePath $baselinePath -OverridePath $override -PublicBaselinePath $baselinePath -RunId r -CollectedAtUtc c).Rows)
    $frCsv = @($frRows | ConvertTo-Csv -NoTypeInformation | ConvertFrom-Csv)
    Check (@($frCsv | Where-Object { $_.SkuPartNumber -ceq 'SPE_F1' -and $_.Currency -eq 'EUR' })[0].UnitPriceMonthly -ceq '8.66') 'fr-FR: public price written as 8.66.'
    Check (@($frCsv | Where-Object { $_.SkuPartNumber -ceq 'SPE_E3' -and $_.EffectiveFrom -eq '2026-01-01' })[0].UnitPriceMonthly -ceq '30') 'fr-FR: private price written with the invariant format.'
    Check (@($frCsv | Where-Object { $_.UnitPriceMonthly -match ',' }).Count -eq 0) 'fr-FR: no decimal comma in UnitPriceMonthly.'
    $frEsu = @(@(ConvertTo-WindowsEsuPriceRows @(Import-Csv -LiteralPath (Write-Fixture 'esu-fr.csv' @($esuHeader, 'WIN10_ESU,1,EUR,55.25,2025-11-01,2026-10-31,True,A,B,C,False'))) r c) | ConvertTo-Csv -NoTypeInformation | ConvertFrom-Csv)
    Check ($frEsu[0].UnitPricePerDeviceYear -ceq '55.25') 'fr-FR: ESU price written as 55.25.'
    $frFile = Join-Path $temporary 'prices-fr.csv'
    $frRows | Export-Csv -LiteralPath $frFile -NoTypeInformation -Encoding utf8NoBOM
    Check (-not ((Get-Content -LiteralPath $frFile -Raw) -match '"\d+,\d+"')) 'fr-FR: exported CSV has no decimal comma value.'
  }
  finally { [Globalization.CultureInfo]::CurrentCulture = $savedCulture }
  foreach ($bad in @('"30,00"', '"1,234.50"', '1 234.5')) {
    $path = Write-Fixture ('override-bad-' + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.csv') @($priceHeader, "SPE_E3,EUR,$bad,Monthly,2026-01-01,,A,B,C,False")
    Assert-Throws { Resolve-LicensePricingRows -SourcePath $baselinePath -OverridePath $path -PublicBaselinePath $baselinePath -RunId r -CollectedAtUtc c } '*Invalid UnitPriceMonthly*' "Price $bad is refused instead of being misread."
  }
  Assert-Throws { ConvertTo-WindowsEsuPriceRows @(Import-Csv -LiteralPath (Write-Fixture 'esu-comma.csv' @($esuHeader, 'WIN10_ESU,1,EUR,"55,25",2025-11-01,2026-10-31,True,A,B,C,False'))) r c } '*Invalid UnitPricePerDeviceYear*' 'ESU price with a decimal comma is refused.'

  # 8. Template exposes the private keys; the collector creates and completes its runtime JSON.
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBHynQjhvny9z0X
# usm+77eCdyhWBNnBhFrCmmyHLmqBA6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKlxZ2lJsbvv6idJ3QQiDDVBm5n6ZEB1MMZW22Coigk/MA0GCSqG
# SIb3DQEBAQUABIIBgBeC/pPHmF1jL54+JLBl5gcxHEw0CpPduZzdVYifOAZuGX4K
# x9SRyspkBYEqDn1b2Cd3KIWwqlTbgKlQEy5hwIbNlcMrHFuf0mjXhyRWsTNovuqM
# gV7d3T/cNl1jGxVkpchmW1H5nbN5ye+N2uK/BNhgWahkGiSPPMndTcmSfdH7Nlhk
# BEXE2VRb3Dr5YNDWqCupb3w7ZhkHnKnU1t9TQRiCwyQUI9d+e5i2u8XSa1T/EUAk
# ti05K5LiMpq5intMsPu+ursZO/UpE2guZrKNudNpsQS1eY4yGNgZk7IIxTR3cYEr
# aFk26eBEs2pXMyFwjcGF/y7IcMIhnj2QI43trdDT8Mn/CdcksVkOqwZwB9E4lZas
# CHGO7zs5fEofiCcmkuNrvrvRgIUiRs5IuE4k9K4meDFnmG/vjMOV68H/n0oNUub7
# Hd1K325KdqgqTFeo6ll2/hQwtMFnp0DAisIiYwUVuJtr8u2jfVQ3WRHmx21oFIV1
# iIjXaYm2dJlgnZaDT6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgwMTU1
# MDRaMC8GCSqGSIb3DQEJBDEiBCB8USdfUemray6pUkYJf4DdayshNlmas5KwPVnz
# NJFVjTANBgkqhkiG9w0BAQEFAASCAgA+l+uUPQdXTLrMe3DHxAG4yLnXPdigrDEF
# xEGsC1rvFbnwolissZi26xs7mxCmodteRF47lKeJyyXTanypTg1p7FRGKSjSrGsG
# GhXGBikaA5ogsHU5Qas4L768flJzKwNgFOoeIWnmjMBXiwVgZ2bIf6GIiafMLGbF
# rxjGa1fz8YgQh/d4KWX052BFHC6MYSvnkTvQHxFl2bOAVOLarUasJkWfpoM9qmF0
# nddJzOYFSSij548f/EtB+d8JArSnGR+Zuv5Ad2qgW0ArMZSNI+ByricGgkS/8Zqi
# eJ7pYAeuF8Js1YPF9ZOoIE/A2e9Ih9LXkZt+p+CSNZgtMv2l6mY4eFK3KcBk8IK0
# Pyr5QVyB6+OQt4nm1Wq+sGKBn3Jd+bvk3DgZbTBptDUqxo+z2oIo85nnzz+EYCZK
# 5DHEG2ktgilLQXiyURicDvIaEvJWqvYVks7P5vys5ESmH3O2aK2q41ejfEv0ZX8j
# YoNia3xKEiZJknUj8iPPk7URqb4q+2hrMoJQlwJvvlQVyrRb6MDZ/gjWQ4UIvcPD
# yCIZeVZSgH1WIJHXKQe2pQzGYkjWB+eaxdXfk4HuoFMJwZuqUt7ylb1K2eJPgWgj
# U5gU3LIC4zrS/LzzxUkzi1IlMtk9CZNcCEnaMqux8mEfX8qFtoeP3GW1WZN9g4pR
# xt/T4jBKsQ==
# SIG # End signature block
