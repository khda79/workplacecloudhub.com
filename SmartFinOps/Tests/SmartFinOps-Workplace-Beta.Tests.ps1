[CmdletBinding()]
param([string]$EvidenceRoot = (Join-Path $env:TEMP ('SmartFinOps-beta-tests-' + [guid]::NewGuid().ToString('N'))))
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2
Add-Type -AssemblyName Microsoft.VisualBasic
$project = Split-Path $PSScriptRoot -Parent
$script:RunId = 'synthetic-beta-test'
$script:SmartFinOpsExcludedFileNamePatterns = @('MAXITEMS')
$script:SmartFinOpsMaxSourceAgeHours = 72
$script:SmartM365LatestCsvFolderPath = $EvidenceRoot
New-Item $EvidenceRoot -ItemType Directory -Force | Out-Null
$tokens=$null; $errors=$null
$ast=[System.Management.Automation.Language.Parser]::ParseFile((Join-Path $project 'SmartFinOps-Workplace-Analyze.ps1'),[ref]$tokens,[ref]$errors)
foreach($name in @('Get-RowPropertyValue','ConvertTo-DateTimeOrNull','ConvertTo-DecimalOrZero','ConvertTo-BoolOrNull','Get-MonthlySkuPrice')) {
    $fn=$ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$false) | Where-Object Name -eq $name | Select-Object -First 1
    . ([scriptblock]::Create($fn.Extent.Text))
}
function Write-SmartFinOpsLog { param($Message,$Level) }
. (Join-Path $project 'Config/SmartFinOps-Workplace-DataContract.ps1')
. (Join-Path $project 'Config/SmartFinOps-Workplace-LicenseDecision.ps1')
. (Join-Path $project 'Config/SmartFinOps-Workplace-ExchangeDecision.ps1')
. (Join-Path $project 'Config/SmartFinOps-Workplace-ValueModel.ps1')
$script:results=[System.Collections.Generic.List[object]]::new()
function Assert-Equal($Name,$Actual,$Expected) {
    if($Actual -ne $Expected) { throw "$Name : expected [$Expected], got [$Actual]" }
    $script:results.Add([pscustomobject]@{Test=$Name;Status='Passed'})
}
function Read-Fixture($Name,$Rows,[char]$Delimiter=',') {
    $Rows | Export-Csv (Join-Path $EvidenceRoot 'fixture.csv') -NoTypeInformation -Encoding UTF8 -Delimiter $Delimiter
    $script:quality=[System.Collections.Generic.List[object]]::new()
    return @(Import-SmartFinOpsSourceCsv -SourceName $Name -FileNames @('fixture.csv') -RequiredColumns @('UserPrincipalName') -DataQualityRows $script:quality)
}
$fresh=(Get-Date).AddHours(-2).ToString('s'); $stale=(Get-Date).AddDays(-10).ToString('s')
$data=@(Read-Fixture 'Fixture' @([pscustomobject]@{UserPrincipalName='one@example.invalid';'Report Refresh Date'=$fresh},[pscustomobject]@{UserPrincipalName='two@example.invalid';'Report Refresh Date'=$stale}) ';')
Assert-Equal 'Semicolon import rows' $data.Count 2
Assert-Equal 'Oldest refresh across all rows' $quality[0].FreshnessStatus 'Stale'
$data=@(Read-Fixture 'Fixture' @([pscustomobject]@{UserPrincipalName='one@example.invalid';'Report Refresh Date'='invalid'}))
Assert-Equal 'Malformed refresh is unknown despite new file' $quality[0].FreshnessStatus 'Unknown'
$data=@(Read-Fixture 'Fixture' @([pscustomobject]@{UserPrincipalName='one@example.invalid';'Report Refresh Date'=(Get-Date).AddDays(1).ToString('s')}))
Assert-Equal 'Future refresh is unknown' $quality[0].FreshnessStatus 'Unknown'
$data=@(Read-Fixture 'Fixture' @([pscustomobject]@{WrongColumn='x'}))
Assert-Equal 'Invalid schema excluded' $data.Count 0
Assert-Equal 'Invalid schema reported' $quality[0].Status 'Invalid schema'
$data=@(Read-Fixture 'M365 active users' @([pscustomobject]@{UserPrincipalName=' ONE@example.invalid '},[pscustomobject]@{UserPrincipalName='one@example.invalid'}))
Assert-Equal 'Duplicate normalized directory identity excluded' $data.Count 0
Assert-Equal 'MAXITEMS excluded' (Test-SmartFinOpsExcludedCsv 'Example_MAXITEMS.csv') $true
$quality=[System.Collections.Generic.List[object]]::new()
$data=@(Import-SmartFinOpsSourceCsv -SourceName 'Absent' -FileNames @('absent.csv') -DataQualityRows $quality)
Assert-Equal 'Missing CSV returns empty collection' $data.Count 0
Assert-Equal 'Missing CSV reported' $quality[0].Status 'Missing'
Assert-Equal 'Missing sources prevent financial decision' (Test-SmartFinOpsDecisionSources @()) $false
Set-Content (Join-Path $EvidenceRoot 'empty.csv') 'UserPrincipalName' -Encoding UTF8
$quality=[System.Collections.Generic.List[object]]::new()
$data=@(Import-SmartFinOpsSourceCsv -SourceName 'Empty' -FileNames @('empty.csv') -RequiredColumns @('UserPrincipalName') -DataQualityRows $quality)
Assert-Equal 'Header-only CSV reported' $quality[0].Status 'Empty'
Assert-Equal 'Missing numeric is unknown' ($null -eq (ConvertTo-SmartFinOpsNumberOrNull '')) $true
Assert-Equal 'Malformed numeric is unknown' ($null -eq (ConvertTo-SmartFinOpsNumberOrNull 'oops')) $true
Assert-Equal 'Negative numeric is unknown' ($null -eq (ConvertTo-SmartFinOpsNumberOrNull '-1')) $true
Assert-Equal 'Comma decimal never interpreted as thousands' ($null -eq (ConvertTo-SmartFinOpsNumberOrNull '1,5')) $true
Assert-Equal 'Invariant decimal' (ConvertTo-SmartFinOpsNumberOrNull '1.5') 1.5
$price=[pscustomobject]@{MonthlyUnitPriceBySkuPartNumber=[pscustomobject]@{SPE_E3=[decimal]31.20;SPE_F1=[decimal]6.45}}
$license=[pscustomobject]@{UserPrincipalName='one@example.invalid';SkuPartNumber='SPE_E3';Source='Direct'}
$user=[pscustomobject]@{UserPrincipalName='one@example.invalid';AccountEnabled=$false}
$ad=[pscustomobject]@{UserPrincipalName='one@example.invalid';Enabled=$false;M365LicenseTargetPersona='None';AccountType='Named Account'}
$map=@{}; $ev=Get-SmartFinOpsEvidenceRecord $map 'one@example.invalid'; $ev.HasUserActivityEvidence=$true
$args=@{LicenseRows=@($license,$license);M365Users=@($user);ADUsers=@($ad);EvidenceMap=$map;M365RecentCutoff=(Get-Date).AddDays(-90);TechnicalRecentCutoff=(Get-Date).AddDays(-90);M365EvidenceAsOfDate=(Get-Date);TechnicalEvidenceAsOfDate=(Get-Date);PriceModel=$price;DecisionSourcesHealthy=$true}
$decisions=@(New-SmartFinOpsUserLicenseDecisionRows @args)
Assert-Equal 'Repeated assignment creates one decision' $decisions.Count 1
Assert-Equal 'Complete inactive disabled user' $decisions[0].DecisionClass 'Recommended'
Assert-Equal 'E3 monthly removal arithmetic' $decisions[0].IndicativeMonthlyDifferenceEUR 31.20
$values=@(New-SmartFinOpsValueOpportunityRows -SummaryRows @() -LicenseRows @($license,$license) -UserDecisionRows $decisions -LicenseCapacityRows @() -PriceModel $price)
Assert-Equal 'Duplicate assignments not double valued' (Get-SmartFinOpsDecimalSum $values 'MonthlyValueEUR') 31.20
Assert-Equal 'Annual arithmetic' (Get-SmartFinOpsDecimalSum $values 'AnnualValueEUR') 374.40
$args.DecisionSourcesHealthy=$false
$decisions=@(New-SmartFinOpsUserLicenseDecisionRows @args)
Assert-Equal 'Unhealthy sources downgrade recommendation' $decisions[0].DecisionClass 'Review'
Assert-Equal 'Unhealthy sources unquantified' ($null -eq $decisions[0].IndicativeMonthlyDifferenceEUR) $true
$args.DecisionSourcesHealthy=$true; $user.AccountEnabled=$true
$decisions=@(New-SmartFinOpsUserLicenseDecisionRows @args)
Assert-Equal 'Contradictory directory state unquantified' ($null -eq $decisions[0].IndicativeMonthlyDifferenceEUR) $true
$ad.Enabled=$true; $ad.M365LicenseTargetPersona='F3'; [void]$ev.RecentM365Services.Add('Teams')
$decisions=@(New-SmartFinOpsUserLicenseDecisionRows @args)
Assert-Equal 'Missing storage blocks F3 value' ($null -eq $decisions[0].IndicativeMonthlyDifferenceEUR) $true
$ev.HasMailboxStorageEvidence=$true; $ev.HasOneDriveStorageEvidence=$true; $ev.MailboxStorageBytes=2GB; $ev.OneDriveStorageBytes=2GB
$decisions=@(New-SmartFinOpsUserLicenseDecisionRows @args)
Assert-Equal 'F3 exact 2GB conditional only' $decisions[0].DecisionClass 'Conditional'
Assert-Equal 'E3 minus F3' $decisions[0].IndicativeMonthlyDifferenceEUR 24.75
$ev.MailboxStorageBytes=2GB+1
$decisions=@(New-SmartFinOpsUserLicenseDecisionRows @args)
Assert-Equal 'F3 above 2GB blocked' $decisions[0].RecommendedLicense 'Keep M365 E3 - F3 technical blocker'
$ev.MailboxStorageBytes=0; $ev.HasDesktopAppsActivation=$true
$decisions=@(New-SmartFinOpsUserLicenseDecisionRows @args)
Assert-Equal 'Desktop activation blocks F3' $decisions[0].RecommendedLicense 'Keep M365 E3 - F3 technical blocker'
$ev.HasDesktopAppsActivation=$false; $ad.AccountType='Service Account'
$decisions=@(New-SmartFinOpsUserLicenseDecisionRows @args)
Assert-Equal 'Special account separated' $decisions[0].RecommendedLicense 'Separate review - special account'
$capacity=[pscustomobject]@{SkuPartNumber='SPE_E3';SkuDisplayName='E3';AvailableUnits=2;EvidenceFresh=$true}
$values=@(New-SmartFinOpsValueOpportunityRows -SummaryRows @() -LicenseRows @() -UserDecisionRows @() -LicenseCapacityRows @($capacity) -PriceModel $price)
Assert-Equal 'Reuse separately classified' $values[0].ValuePillar 'Cost avoidance'
Assert-Equal 'Capacity multiplication' $values[0].MonthlyValueEUR 62.40
$capacity.EvidenceFresh=$false
$values=@(New-SmartFinOpsValueOpportunityRows -SummaryRows @() -LicenseRows @() -UserDecisionRows @() -LicenseCapacityRows @($capacity) -PriceModel $price)
Assert-Equal 'Stale capacity unquantified' $values.Count 0
$mailbox=[pscustomobject]@{UserPrincipalName='one@example.invalid';PrimarySmtpAddress='one@example.invalid';RecipientTypeDetails='UserMailbox';AccountEnabled=$false;TotalItemSizeGB='44.99';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False';FullAccess='delegate@example.invalid';SendAs='';GrantSendOnBehalfTo=''}
$user.AccountEnabled=$false; $ad.Enabled=$false; $ad.AccountType='Named Account';$ad.M365LicenseTargetPersona='None'
$decisions=@(New-SmartFinOpsUserLicenseDecisionRows @args)
$exchangeArgs=@{MailboxRows=@($mailbox);MailboxStatsRows=@();MailboxArchiveRows=@();MailboxPermissionRows=@();MailboxUsageRows=@();M365Users=@($user);ADUsers=@($ad);UserDecisionRows=$decisions}
$exchange=@(New-SmartFinOpsSharedMailboxConversionRows @exchangeArgs)
Assert-Equal 'Mailbox below 45GB' $exchange[0].Status 'Strong candidate'
$mailbox.TotalItemSizeGB='45';$exchange=@(New-SmartFinOpsSharedMailboxConversionRows @exchangeArgs)
Assert-Equal 'Mailbox 45GB capacity review' $exchange[0].Status 'Capacity review'
$mailbox.TotalItemSizeGB='50';$exchange=@(New-SmartFinOpsSharedMailboxConversionRows @exchangeArgs)
Assert-Equal 'Mailbox 50GB excluded' $exchange[0].Status 'Excluded'
$mailbox.TotalItemSizeGB='10';$mailbox.LitigationHoldEnabled='';$exchange=@(New-SmartFinOpsSharedMailboxConversionRows @exchangeArgs)
Assert-Equal 'Unknown hold excludes conversion' $exchange[0].Status 'Excluded'
$mailbox.LitigationHoldEnabled='True';$exchange=@(New-SmartFinOpsSharedMailboxConversionRows @exchangeArgs)
Assert-Equal 'Known hold excludes conversion' $exchange[0].Status 'Excluded'
$summary=@([pscustomobject]@{Metric='Strong shared mailbox conversion candidates';Value=1})
$values=@(New-SmartFinOpsValueOpportunityRows -SummaryRows $summary -LicenseRows @() -UserDecisionRows @() -LicenseCapacityRows @() -PriceModel $price)
Assert-Equal 'Mailbox conversion never adds license value' (Get-SmartFinOpsDecimalSum $values 'MonthlyValueEUR') 0
$script:SmartFinOpsPriceBaselinePath=Join-Path $project 'Config/SmartFinOps-Workplace-FrancePriceBaseline.json'
$ScriptLocalConfig=[pscustomobject]@{}
$script:priceOverride=$null
function Get-SmartFinOpsScriptConfigValue {param($Config,$Name,$DefaultValue) return $script:priceOverride}
$model=Get-PriceModel
Assert-Equal 'Baseline price stays numeric' $model.MonthlyUnitPriceBySkuPartNumber.SPE_E3 31.20
foreach($invalid in @('31,20','-1','oops',$null)) {
 $script:priceOverride=[pscustomobject]@{Currency='EUR';MonthlyUnitPriceBySkuPartNumber=[pscustomobject]@{SPE_E3=$invalid}}
 $thrown=$false;try {$null=Get-PriceModel} catch {$thrown=$true}
 Assert-Equal "Invalid contract price rejected [$invalid]" $thrown $true
}
$script:priceOverride=[pscustomobject]@{Currency='USD';MonthlyUnitPriceBySkuPartNumber=[pscustomobject]@{SPE_E3=10}}
$thrown=$false;try {$null=Get-PriceModel} catch {$thrown=$true}
Assert-Equal 'Non-EUR contract rejected' $thrown $true
$script:priceOverride=[pscustomobject]@{Currency='EUR';MonthlyUnitPriceBySkuPartNumber=[pscustomobject]@{SPE_E3=0}}
Assert-Equal 'Explicit zero contract price accepted' (Get-PriceModel).MonthlyUnitPriceBySkuPartNumber.SPE_E3 0
Assert-Equal 'Ambiguous mailbox number rejected' ($null -eq (ConvertTo-SmartFinOpsExchangeDecimalOrNull '4,5')) $true
$script:results | ConvertTo-Json | Set-Content (Join-Path $EvidenceRoot 'test-results.json') -Encoding UTF8
Write-Output ("Passed {0} synthetic assertions. Evidence: {1}" -f $results.Count,$EvidenceRoot)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDtXkb2xambIgoM
# ef5fQ+AwZE77VB8ObCzmDhL7NkqcmaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINyBRkENK8whOmA+tZI87dW0Qgz5UN+/xrG7/N/1P5AJMA0GCSqG
# SIb3DQEBAQUABIIBgDN0hwCVL8fvB8ciiEBQtKEDQppDPp1OCWPTsFYMRksFOWUR
# c23LZveI1x3kxMnQvBTArA2+OU3Gu5SFNNXvTy+mYo9HuZCYQErHR9cV75fGRXZ+
# qMRW32oVcc+jzslopiQwC1KE3KcFe0ggUbTRx2n3P/+vmbeAg9SUyYQ6FS7FZE3J
# Map2UUMmnKk6VHfOhY7uCoK1FJUFscKdZ6wEoHiXkUJqLWR0vu4r2/Hl/ZOKzz3N
# vA7uS2vEjn+5+ZSaa+xO+T73pWh1+2qJs0z9Y0SGXJCQgow1XHVUNrFJRPF2u1HK
# uGhqb40dN7dnxpbsXke/P6/DJhDS9jPvrJZE+miPa2/2Y1OnaYHmC4nnkXpzDM2x
# mwolvnfBcRhMrTPdZs+8MBFSzriWOZY/HrvuAotVRHTzhql+K8uApAoD8R4fTYXR
# 42K8G07nF6SKsYRXRVPuXC7/KBSlTOuwGHLe2UupebcktrhD05l0YWGV2gaTWNf+
# pTauc1TtAd2AfBDzH6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MDkwODA4
# MzBaMC8GCSqGSIb3DQEJBDEiBCDiF/739hmKYAScTXqRri3DKJDfEarprP5/vJ46
# HLfh4DANBgkqhkiG9w0BAQEFAASCAgCyFepOs4QE0m4G28K3R4B1nvKcaBAglgL0
# N1dE/Yp8zFTuqFbKfXRiwBh0u/lL/gVdSNIq4P1nG8LW8Nxc/sNTq3ZY8IDEWKZw
# mb0j5P25nyIrKpuf2zulYRO26X+Ln0UfBufHez/N/YZRbQt8JCQ5VtJ9JqhF0jkG
# dgNgCrMp4Yh5jnz/a2T0R2oMI6XEgjafrW2XQYjsiPm5DM9RszXYWtWQmBxqYLu8
# oxXuIFJikVnMAWlziegtPHM8qtBQJ7qz5IBnkc9izYm0o2znRqCWB8CoBcxSVldu
# EFo6c1WmjwSwQizQAgndhEr0YThrFysbR1FRSaiTfMuyuY9obR0KtQc9ajPJi/qU
# aPQQ7+3sQacS8CsG4vQJ65Gkv2XqqbHpxfZL7oljImY08s1Fb2HYPWv4DXCz/C1/
# izf1TMolyPvpKvjgIzRzEhAAQ4vu+ZsFKz+Xsyayg7oNzaU5yfa7YRxk9GQ/kkPX
# A5FAlvkC/hDWI/H8KrrJM+RZ6etTR/a8YsRpZ2/ioGNTjTXoTL8yB/B8tUsa7Uk+
# CwValzKwKX00ZiNW0JiIN4Av1pMbDz9r4iUe8TYujJXxOpSyLD+a2fvycuivbIKu
# ZO5YwoSzR3KUmq89ICLdyy0K+g15zQQldW4kkFT+bLQa5YiGkYIaGbQZOMAX3CR7
# FdWkoTlV/g==
# SIG # End signature block
