#Requires -Version 7.0
<#
.SYNOPSIS
Validates and normalizes the governed Microsoft 365 license-pricing CSV and optional Windows 10 ESU prices.
The public France baseline is merged with an optional private override (PricingOverrideCsvPath): a private
SKU+Currency replaces every public period of that SKU+Currency. A PricingSourceCsvPath pointing to another
file keeps the former whole-replacement behaviour with a migration warning. Windows 10 ESU prices come only
from a private EUR source (EsuPricingCsvPath) and are published as Windows_ESU_Prices.csv.
.VERSION
1.1.5
.REQUIREMENTS
No Microsoft Graph permission. PricingSourceCsvPath must reference the public baseline or a private CSV.
#>
[CmdletBinding()]
param([string]$Tenant='test',[string]$PricingSourceCsvPath,[string]$PricingOverrideCsvPath,[string]$EsuPricingCsvPath,[switch]$ValidateOnly,[switch]$SelfTest,[switch]$EnableConfiguredExternalActions,[string]$OutputPath,[string]$LatestCsvFolderPath)
# Local lifecycle only; no collection or publication is performed by this guard.
$smartM365LifecycleRoot=$PSScriptRoot
while ($smartM365LifecycleRoot -and -not (Test-Path -LiteralPath (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1') -PathType Leaf)) {
    $smartM365LifecycleParent=Split-Path $smartM365LifecycleRoot -Parent
    if($smartM365LifecycleParent -eq $smartM365LifecycleRoot){break}
    $smartM365LifecycleRoot=$smartM365LifecycleParent
}
if(-not $smartM365LifecycleRoot){throw 'SmartM365 lifecycle support is missing; update the complete runtime.'}
. (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1')
$smartM365RunGuard=Start-SmartM365InventoryRun -ScriptPath $PSCommandPath -RequireReturn
try {
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop';$ScriptVersion='1.1.5';$script:Runtime=$null;$script:CompletionStatus='Success';$script:CompletionError=$null

function ConvertTo-LicensePriceRows{
    param([object[]]$InputRows,[string]$RunId,[string]$CollectedAtUtc)
    $seen=@{};$rows=[System.Collections.Generic.List[object]]::new()
    foreach($row in $InputRows){
        $sku=[string](Get-SmartM365EvidenceProperty $row @('SkuPartNumber'));$currency=[string](Get-SmartM365EvidenceProperty $row @('Currency'));$priceText=[string](Get-SmartM365EvidenceProperty $row @('UnitPriceMonthly'))
        if([string]::IsNullOrWhiteSpace($sku)-or $currency-notmatch'^[A-Z]{3}$'){throw "Invalid SKU or ISO currency in license pricing source: '$sku' / '$currency'."}
        $price=0.0;if($priceText.Trim()-notmatch'^\d+(\.\d+)?$'-or -not[double]::TryParse($priceText,[System.Globalization.NumberStyles]::Number,[System.Globalization.CultureInfo]::InvariantCulture,[ref]$price)-or $price-lt 0){throw "Invalid UnitPriceMonthly for SKU '$sku'. Use an invariant non-negative number (decimal point, no group separator)."}
        $from=[string](Get-SmartM365EvidenceProperty $row @('EffectiveFrom'));$to=[string](Get-SmartM365EvidenceProperty $row @('EffectiveTo'));$key="$sku|$currency|$from|$to";if($seen.ContainsKey($key)){throw "Duplicate license price period: $key"};$seen[$key]=$true
        if($from-and $to-and [datetime]$from-gt[datetime]$to){throw "EffectiveFrom is after EffectiveTo for SKU '$sku'."}
        $rows.Add([pscustomobject][ordered]@{RunId=$RunId;CollectedAtUtc=$CollectedAtUtc;CollectionStatus='GovernedInput';IsPartialInventory=$false;SourceScriptVersion=$ScriptVersion;SkuPartNumber=$sku;Currency=$currency;UnitPriceMonthly=$price.ToString([System.Globalization.CultureInfo]::InvariantCulture);BillingCadence=[string](Get-SmartM365EvidenceProperty $row @('BillingCadence') 'Monthly');EffectiveFrom=$from;EffectiveTo=$to;PriceSource=[string](Get-SmartM365EvidenceProperty $row @('PriceSource'));PricingBasis=[string](Get-SmartM365EvidenceProperty $row @('PricingBasis'));ContractReference=[string](Get-SmartM365EvidenceProperty $row @('ContractReference'));IsEstimated=[string](Get-SmartM365EvidenceProperty $row @('IsEstimated') 'False')})
    }
    return @($rows)
}

# A private SKU+Currency replaces every public period of that SKU+Currency; private-only SKUs are added.
function Merge-LicensePriceRows{
    param([AllowEmptyCollection()][object[]]$PublicRows,[AllowEmptyCollection()][object[]]$OverrideRows)
    $overridden=[System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach($row in $OverrideRows){[void]$overridden.Add(('{0}|{1}' -f $row.SkuPartNumber,$row.Currency))}
    $kept=@($PublicRows|Where-Object{-not $overridden.Contains(('{0}|{1}' -f $_.SkuPartNumber,$_.Currency))})
    return [pscustomobject]@{Rows=@($kept+@($OverrideRows));ReplacedPublicRows=($PublicRows.Count-$kept.Count);OverrideKeys=$overridden.Count}
}

function ConvertTo-WindowsEsuPriceRows{
    param([object[]]$InputRows,[string]$RunId,[string]$CollectedAtUtc)
    $rows=[System.Collections.Generic.List[object]]::new();$years=@{};$periods=@{}
    $culture=[System.Globalization.CultureInfo]::InvariantCulture
    foreach($row in $InputRows){
        $program=[string](Get-SmartM365EvidenceProperty $row @('ProgramKey'));$currency=[string](Get-SmartM365EvidenceProperty $row @('Currency'))
        if($program-cne'WIN10_ESU'){throw "Unsupported ESU program key: '$program'. Expected WIN10_ESU."}
        # No currency conversion: only governed EUR prices are accepted.
        if($currency-cne'EUR'){throw "Windows ESU prices must be in EUR; got '$currency' for $program."}
        $year=0;if(-not[int]::TryParse([string](Get-SmartM365EvidenceProperty $row @('CoverageYear')),[ref]$year)-or $year-lt 1-or $year-gt 3){throw "CoverageYear must be 1, 2 or 3 for $program."}
        $priceText=[string](Get-SmartM365EvidenceProperty $row @('UnitPricePerDeviceYear'));$price=0.0;if($priceText.Trim()-notmatch'^\d+(\.\d+)?$'-or -not[double]::TryParse($priceText,[System.Globalization.NumberStyles]::Number,$culture,[ref]$price)-or $price-lt 0){throw "Invalid UnitPricePerDeviceYear for $program year $year. Use an invariant non-negative number (decimal point, no group separator)."}
        $fromText=[string](Get-SmartM365EvidenceProperty $row @('EffectiveFrom'));$toText=[string](Get-SmartM365EvidenceProperty $row @('EffectiveTo'))
        $from=[datetime]::MinValue;$to=[datetime]::MinValue
        if(-not[datetime]::TryParseExact($fromText,'yyyy-MM-dd',$culture,[System.Globalization.DateTimeStyles]::None,[ref]$from)-or
           -not[datetime]::TryParseExact($toText,'yyyy-MM-dd',$culture,[System.Globalization.DateTimeStyles]::None,[ref]$to)-or $from-gt $to){throw "ESU year $year needs EffectiveFrom <= EffectiveTo as yyyy-MM-dd."}
        $cumulative=[string](Get-SmartM365EvidenceProperty $row @('IsCumulative'));$estimated=[string](Get-SmartM365EvidenceProperty $row @('IsEstimated'))
        if($cumulative-notin@('True','False')-or $estimated-notin@('True','False')){throw "IsCumulative and IsEstimated must be True or False for ESU year $year."}
        $key="$program|$currency";if(-not$years.ContainsKey($key)){$years[$key]=@{};$periods[$key]=[System.Collections.Generic.List[object]]::new()}
        if($years[$key].ContainsKey($year)){throw "Duplicate ESU coverage year $year for $key."};$years[$key][$year]=$from
        foreach($other in $periods[$key]){if($from-le $other.To-and $other.From-le $to){throw "ESU coverage periods overlap for $key."}}
        $periods[$key].Add([pscustomobject]@{From=$from;To=$to})
        $rows.Add([pscustomobject][ordered]@{RunId=$RunId;CollectedAtUtc=$CollectedAtUtc;CollectionStatus='GovernedInput';IsPartialInventory=$false;SourceScriptVersion=$ScriptVersion;ProgramKey=$program;CoverageYear=$year;Currency=$currency;UnitPricePerDeviceYear=$price.ToString([System.Globalization.CultureInfo]::InvariantCulture);EffectiveFrom=$fromText;EffectiveTo=$toText;IsCumulative=$cumulative;PriceSource=[string](Get-SmartM365EvidenceProperty $row @('PriceSource'));PricingBasis=[string](Get-SmartM365EvidenceProperty $row @('PricingBasis'));ContractReference=[string](Get-SmartM365EvidenceProperty $row @('ContractReference'));IsEstimated=$estimated})
    }
    foreach($key in $years.Keys){
        $ordered=@($years[$key].Keys|Sort-Object)
        for($i=1;$i-lt $ordered.Count;$i++){if($years[$key][$ordered[$i]]-le $years[$key][$ordered[$i-1]]){throw "ESU coverage years must follow each other in time for $key."}}
    }
    return @($rows)
}

# PricingSourceCsvPath keeps its meaning; a private file there keeps the former whole-replacement behaviour.
function Resolve-LicensePricingRows{
    param([Parameter(Mandatory)][string]$SourcePath,[string]$OverridePath,[Parameter(Mandatory)][string]$PublicBaselinePath,[string]$RunId,[string]$CollectedAtUtc)
    if(-not(Test-Path -LiteralPath $SourcePath -PathType Leaf)){throw 'PricingSourceCsvPath must reference an existing CSV.'}
    # Same file, or a byte-identical copy reached through another path (mapped drive, UNC), is the public baseline.
    $legacy=[System.IO.Path]::GetFullPath($SourcePath)-ine[System.IO.Path]::GetFullPath($PublicBaselinePath)-and
        (Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash-cne(Get-FileHash -LiteralPath $PublicBaselinePath -Algorithm SHA256).Hash
    $hasOverride=-not[string]::IsNullOrWhiteSpace($OverridePath)
    if($legacy-and $hasOverride){throw 'PricingSourceCsvPath points to a private file while PricingOverrideCsvPath is also set. Point PricingSourceCsvPath back to the public baseline and keep private prices in PricingOverrideCsvPath.'}
    $sourceRows=@(ConvertTo-LicensePriceRows @(Import-Csv -LiteralPath $SourcePath) $RunId $CollectedAtUtc)
    $messages=[System.Collections.Generic.List[object]]::new()
    if($legacy){
        $messages.Add([pscustomobject]@{Level='WARNING';Text=("PricingSourceCsvPath points to a file other than the public baseline; it replaces the whole public baseline (legacy behaviour, {0} row(s)). Migrate private prices to PricingOverrideCsvPath and point PricingSourceCsvPath back to the public baseline." -f $sourceRows.Count)})
        return [pscustomobject]@{Mode='LegacyReplacement';Rows=$sourceRows;Messages=$messages.ToArray()}
    }
    if(-not $hasOverride){
        $messages.Add([pscustomobject]@{Level='INFO';Text=("Public license price baseline used without private override: {0} row(s)." -f $sourceRows.Count)})
        return [pscustomobject]@{Mode='PublicBaseline';Rows=$sourceRows;Messages=$messages.ToArray()}
    }
    if(-not(Test-Path -LiteralPath $OverridePath -PathType Leaf)){throw 'PricingOverrideCsvPath must reference an existing private CSV.'}
    $overrideRows=@(ConvertTo-LicensePriceRows @(Import-Csv -LiteralPath $OverridePath) $RunId $CollectedAtUtc)
    if($overrideRows.Count-eq 0){throw 'The private license price override contains no rows.'}
    $merged=Merge-LicensePriceRows -PublicRows $sourceRows -OverrideRows $overrideRows
    $messages.Add([pscustomobject]@{Level='INFO';Text=("Public baseline {0} row(s) merged with private override {1} row(s) for {2} SKU/currency key(s); {3} public row(s) replaced." -f $sourceRows.Count,$overrideRows.Count,$merged.OverrideKeys,$merged.ReplacedPublicRows)})
    return [pscustomobject]@{Mode='PublicBaselineWithOverride';Rows=$merged.Rows;Messages=$messages.ToArray()}
}
if($SelfTest){Import-Module (Join-Path $PSScriptRoot '..\..\Common\SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.22' -Force;$sample=[pscustomobject]@{SkuPartNumber='SPE_E5';Currency='EUR';UnitPriceMonthly='42.50';EffectiveFrom='2026-01-01';EffectiveTo='2026-12-31';PriceSource='Contract'};$test=ConvertTo-LicensePriceRows @($sample) r n;if($test.Count-ne 1-or $test[0].UnitPriceMonthly-ne 42.5){throw 'License pricing self-test failed.'};'PASS: License pricing offline contract';$smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return}
try{
    $tenantContextPath=Join-Path (Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent) 'Config\SmartM365-TenantContext.ps1';. $tenantContextPath
    $effectiveConfig=Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
    # Create the runtime JSON from its template when absent and add missing template keys without overwriting values.
    Read-SmartM365JsonConfig -Path (Join-Path $PSScriptRoot 'SmartM365-LicensePricing-Inventory.local.json') -Required | Out-Null
    Import-Module (Join-Path $PSScriptRoot '..\..\..\Modules\SmartM365.Core\SmartM365.Core.psd1') -MinimumVersion '1.0.92' -Global -ErrorAction Stop
    Import-Module (Join-Path $PSScriptRoot '..\..\Common\SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.22' -Force
    $script:Runtime=Initialize-SmartM365EvidenceRuntime -ScriptPath $PSCommandPath -EffectiveConfig $effectiveConfig -DefaultOutputRelativePath 'M365\Licensing\Pricing' -OutputPath $OutputPath -LatestCsvFolderPath $LatestCsvFolderPath -ValidateOnly:$ValidateOnly -EnableConfiguredExternalActions:$EnableConfiguredExternalActions
    Start-SmartM365SourceReceipt -ScriptPath $PSCommandPath -SourceRootPath $script:Runtime.LatestCsvFolderPath -ReadOnly:($ValidateOnly -or $SelfTest)
    Enable-SmartM365DailyCaptureRetention -OutputRoot $script:Runtime.OutputPath
    if([string]::IsNullOrWhiteSpace($PricingSourceCsvPath)){$PricingSourceCsvPath=[string](Get-SmartM365EvidenceProperty $script:Runtime.Config @('PricingSourceCsvPath'))}
    if([string]::IsNullOrWhiteSpace($PricingSourceCsvPath)){throw 'PricingSourceCsvPath must reference an existing CSV.'}
    if([string]::IsNullOrWhiteSpace($PricingOverrideCsvPath)){$PricingOverrideCsvPath=[string](Get-SmartM365EvidenceProperty $script:Runtime.Config @('PricingOverrideCsvPath'))}
    if([string]::IsNullOrWhiteSpace($EsuPricingCsvPath)){$EsuPricingCsvPath=[string](Get-SmartM365EvidenceProperty $script:Runtime.Config @('EsuPricingCsvPath'))}
    $resolved=Resolve-LicensePricingRows -SourcePath $PricingSourceCsvPath -OverridePath $PricingOverrideCsvPath -PublicBaselinePath (Join-Path $PSScriptRoot 'M365_License_Prices_Public-France.csv') -RunId $script:Runtime.RunId -CollectedAtUtc $script:Runtime.CollectedAtUtc
    foreach($message in $resolved.Messages){Write-SmartM365EvidenceLog -Message $message.Text -Level $message.Level}
    $rows=@($resolved.Rows)
    if($rows.Count-eq 0){throw 'The license pricing source contains no rows.'}
    $esuRows=@()
    if([string]::IsNullOrWhiteSpace($EsuPricingCsvPath)){Write-SmartM365EvidenceLog -Message 'Windows 10 ESU prices are not configured (EsuPricingCsvPath); ESU remains Not valued and Windows_ESU_Prices.csv is not published.' -Level INFO}
    else{
        if(-not(Test-Path -LiteralPath $EsuPricingCsvPath -PathType Leaf)){throw 'EsuPricingCsvPath must reference an existing private CSV.'}
        $esuRows=@(ConvertTo-WindowsEsuPriceRows @(Import-Csv -LiteralPath $EsuPricingCsvPath) $script:Runtime.RunId $script:Runtime.CollectedAtUtc)
        if($esuRows.Count-eq 0){throw 'The Windows ESU price source contains no rows.'}
    }
    if($ValidateOnly){Write-SmartM365EvidenceLog -Message ("License pricing source valid ({0}); {1} governed row(s); {2} Windows ESU row(s). No CSV published." -f $resolved.Mode,$rows.Count,$esuRows.Count) -Level SUCCESS;$smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return}
    Export-SmartM365EvidenceDataset $script:Runtime 'M365_License_Prices' $rows @('RunId','CollectedAtUtc','CollectionStatus','IsPartialInventory','SourceScriptVersion','SkuPartNumber','Currency','UnitPriceMonthly','BillingCadence','EffectiveFrom','EffectiveTo','PriceSource','PricingBasis','ContractReference','IsEstimated')|Out-Null
    if($esuRows.Count-gt 0){Export-SmartM365EvidenceDataset $script:Runtime 'Windows_ESU_Prices' $esuRows @('RunId','CollectedAtUtc','CollectionStatus','IsPartialInventory','SourceScriptVersion','ProgramKey','CoverageYear','Currency','UnitPricePerDeviceYear','EffectiveFrom','EffectiveTo','IsCumulative','PriceSource','PricingBasis','ContractReference','IsEstimated')|Out-Null}

    $smartM365RunGuard.ReachedTerminal=$true
}catch{
    if ($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]) { throw }
    $smartM365RunGuard.ErrorRecord=$_
$script:CompletionStatus='Failed';$script:CompletionError=$_;throw}finally{
    $smartM365RunGuard.InCleanup=$true
if($script:Runtime){Complete-SmartM365EvidenceRuntime -Status $script:CompletionStatus -ErrorRecord $script:CompletionError -FailureStage 'LicensePricingInventory'}}
    $smartM365RunGuard.ReachedTerminal=$true
    $smartM365RunGuard.ReturnedNormally=$true
} catch {
    if ($_.Exception -isnot [Management.Automation.PipelineStoppedException] -and $_.Exception -isnot [OperationCanceledException]) { $smartM365RunGuard.ErrorRecord=$_ }
    throw
} finally {
    Complete-SmartM365InventoryRun -State $smartM365RunGuard
    if($smartM365RunGuard.ExitCode -eq 130){exit 130}
}
if($smartM365RunGuard.ExitCode -ne 0){exit $smartM365RunGuard.ExitCode}


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBO6UjZ+fPgBiiZ
# GAEuvKOlQZQIcfQbHXUoAHkdTuBjoqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEID1HXs5UiDKmjqesMtPpXGkJTFGryFZek68Bf7myA3VJMA0GCSqG
# SIb3DQEBAQUABIIBgHgsVEKo0v3LUQaKC0aiq4P7KDzGhWAoRQFpbb1i09HSR7kE
# Z1cpvsXOPXu7lROTa0VLKG6pn0dnHs6paLWuYIyBZpKhza+tPQFAojtm3k/X+U0v
# tHnnKXaUpZ0y8W2QOxp7bLtkgIsSz4gBHx2beleG6zsS6neJioOQ3DBp09nqP0wR
# sD+JtuCSJcEyk0RYsjSie4q9g7XaVM0dCNwV2IYrt6T7SA6noS8BlV/8tRgKM43X
# nAXGyMCagtLTskPBilqypp6Q7NFzPLfNmFHOAuA9nVSmI+sh8+VDfviBsVV8ot+c
# rznNxxj9YujhKAIM0jdTH3PsPiJndq7TYzT3PYS5pbqa/tsBbSiWmxHZg+hddCL7
# gcsJ8L3nX/FFaPqV2a/HxpD5cqD44UJvmelVSdl20x6WUTVKVVKkzjO9XfjkZfj/
# 8f5//tRySS9cCOcckRdh8aOstezK8PAUYulkS9YhOl1JNYkdeDtKSs47cqFqo/97
# +p/CkiZvlkzzJYADnKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzExNDA3
# MjRaMC8GCSqGSIb3DQEJBDEiBCAByjE7Li79uOOuzgXitFWKlYuhufADPibq4QGP
# i1/zVzANBgkqhkiG9w0BAQEFAASCAgA330zSOtyt70B1nFPlr51AWFVnJd0BNVjA
# IwPy55R92V3Pj60FwNGG2pwJ8O23ajXMgYhbkB/g0H3ZaakocogFC2v2fTNSXROw
# pHsuRdIuhCxH0VWUZoBLIVf6TOjvfcrUp/6cWVFCz/ZRAf/BuT/c0RvgkimC8gsO
# IhqxwA10wex/cz2KG1TBimPloQdfcgWDMeoO3CX/S20saHclqkKXuLZDMYQkx2L6
# TgAybN5Ot0aQaP/gQbbi6v7hvhYO+NJjMS6N9VYWJdtfeKChB7aFDJgLwQDX1WhX
# MZAQbqz+drCHH0l9ly8wqUDoTLYCpJVlvS2qzADDOHugwsqz9oxu07ku6Ri39UFb
# P25dbr+iDUJ4uC/MOxK5ISX9MRA/fPFK6sK7esv1rQIYFSIrDxBen+2KB5ug6ip6
# YdtEmsy9ztMPKlBc5XDaVrxoFJ43CN747ofqnTd3CYx6TEC1Qn7ZPy1CpUlrk5dn
# RfkmVXNUnSf5piS1AduDswQAyRzo4Cg/Lis2YIsx2/EyPAy+AtpDj93MYWyI+jl7
# qlpRyHxiUi8wB14fqM7N794wTjuP/6PhRBNMpw0O8TiEF4cUnUX3g8mBkayeto3r
# D2ZSPJRdJLewZbTnhzANa43NIZvHDPBzVKKafsaGiSqODshjagAcHR8Xo1KlqvRM
# Hbnp7OzrCg==
# SIG # End signature block
