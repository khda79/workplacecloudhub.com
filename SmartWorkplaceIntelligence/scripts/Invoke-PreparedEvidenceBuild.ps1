# Version 1.2.0: license decisions come from the SmartInventory Licences report; New-LicensingEvidence is retired.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$DataRoot,
    [Parameter(Mandatory)][string]$OutputRoot,
    [ValidateSet('All','Estate','Workforce','Devices','Applications','IntuneOperations','Collaboration','Content','Mailbox','DailyTrends','Security','Backup','Trust','Executive')]
    [string]$Only = 'All',
    [switch]$SkipWorkforceHistory,
    [switch]$ReuseStagedWorkforce,
    [string]$AccountClassificationConfigPath,
    [ValidatePattern('^[A-Fa-f0-9]{64}$')][string]$AccountClassificationExpectedSHA256,
    [Nullable[long]]$AccountClassificationExpectedBytes,
    [string]$AdCountryRulesPath,
    [string]$EndpointProtectionPolicyPath,
    [Nullable[double]]$DeviceSystemMaxObservationAgeHours,
    [string]$DeviceSystemOffsetlessGraphTimeZoneId,
    [string]$TenantKey,
    [string]$IntuneSourcePolicyPath,
    [string]$NativeEvidenceAuditPath,
    [string]$PythonCommand='python',
    [string]$DailyTrendsQualificationPath,
    [double]$DailyTrendsPartialThresholdPercent = 1,
    [int]$DailyTrendsPersistentAbsenceObservations = 7,
    # Optional reference clock passed to every generator; omitted means each generator uses the current time.
    [Nullable[datetimeoffset]]$AsOfUtc
)

# Offline preparation only. No collection, upload, Desktop control or promotion.
# Each child process releases its memory before the next domain starts.
$ErrorActionPreference = 'Stop'
$DataRoot = (Resolve-Path -LiteralPath $DataRoot).ProviderPath
$last = Join-Path $DataRoot 'DATA-LAST'
if (-not (Test-Path -LiteralPath $last -PathType Container)) { throw 'DATA-LAST is required.' }
$OutputRoot = [IO.Path]::GetFullPath($OutputRoot)
if ($OutputRoot.TrimEnd('\') -eq $DataRoot.TrimEnd('\') -or $OutputRoot.TrimEnd('\') -eq $last.TrimEnd('\')) {
    throw 'Outputs must use a dedicated staging directory, never the raw source root.'
}
New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
$shell = (Get-Process -Id $PID).Path
if ($PSVersionTable.PSVersion.Major -lt 7) { throw 'PowerShell 7 is required.' }
function OutPath([string]$Name) { Join-Path $OutputRoot ($Name + '.csv') }
function Run([string]$Domain, [string]$Script, [object[]]$Arguments) {
    if ($Only -ne 'All' -and $Only -ne $Domain) { return }
    $started = [datetime]::UtcNow
    Write-Host ('[{0:O}] Starting {1}' -f $started, $Domain)
    & $shell -NoProfile -File (Join-Path $PSScriptRoot $Script) @Arguments
    if ($LASTEXITCODE -ne 0) { throw "Preparation failed: $Domain (exit $LASTEXITCODE). Nothing was published." }
    Write-Host ('[{0:O}] Completed {1}: {2:n1} seconds' -f [datetime]::UtcNow, $Domain, ([datetime]::UtcNow - $started).TotalSeconds)
}
$clockArgs=@();$clockDateArgs=@();$clockUtcArgs=@()
if ($null -ne $AsOfUtc) {
    $reference=$AsOfUtc.ToUniversalTime()
    $clockArgs=@('-AsOfUtc',$reference.ToString('o',[Globalization.CultureInfo]::InvariantCulture))
    $clockDateArgs=@('-AsOfDate',$reference.UtcDateTime.ToString('yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture))
    $clockUtcArgs=@('-AsOfUtc',$reference.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ss.fffffff',[Globalization.CultureInfo]::InvariantCulture))
}
$countryArgs=@();if($AdCountryRulesPath){$countryArgs=@('-AdCountryRulesPath',$AdCountryRulesPath)}
$classificationStart = $null
if (($AccountClassificationExpectedSHA256 -or $null -ne $AccountClassificationExpectedBytes) -and -not $AccountClassificationConfigPath) {
    throw 'An expected account classification capture requires its explicit file path.'
}
if ($Only -in @('All','Workforce') -and -not $ReuseStagedWorkforce -and $AccountClassificationConfigPath) {
    Import-Module (Join-Path $PSScriptRoot 'PreparedClassification.psm1')
    $classificationCheck = @{Path=$AccountClassificationConfigPath;Stage='BuildStart'}
    if ($AccountClassificationExpectedSHA256) { $classificationCheck.ExpectedSHA256=$AccountClassificationExpectedSHA256 }
    if ($null -ne $AccountClassificationExpectedBytes) { $classificationCheck.ExpectedBytes=$AccountClassificationExpectedBytes }
    $classificationStart = Test-PreparedClassificationCapture @classificationCheck
}
Run 'Estate' 'New-EnterpriseEstateEvidence.ps1' (@('-DataRoot',$DataRoot,'-OutputRoot',$OutputRoot)+$countryArgs)
if ($Only -eq 'All') {
    if (-not $NativeEvidenceAuditPath) { throw 'Unified preparation requires a receipt-qualified native capture audit.' }
    Import-Module (Join-Path $PSScriptRoot 'NativeEvidence.psm1') -Force
    Invoke-NativeEvidenceStage -LastRoot $last -TenantKey $TenantKey -AuditPath $NativeEvidenceAuditPath -OutputRoot $OutputRoot -PythonCommand $PythonCommand
}
$deviceArgs=@('-DataRoot',$DataRoot,'-OutputPath',(OutPath 'DeviceInventoryEvidence'),'-SignalsOutputPath',(OutPath 'DeviceRiskEvidenceSummary'),'-DirectorySummaryOutputPath',(OutPath 'DeviceDirectorySummary'),'-SystemEvidenceOutputPath',(OutPath 'DeviceSystemEvidence'))
if ($null -ne $DeviceSystemMaxObservationAgeHours) { $deviceArgs+=@('-DeviceSystemMaxObservationAgeHours',$DeviceSystemMaxObservationAgeHours) }
if ($DeviceSystemOffsetlessGraphTimeZoneId) { $deviceArgs+=@('-DeviceSystemOffsetlessGraphTimeZoneId',$DeviceSystemOffsetlessGraphTimeZoneId) }
$deviceArgs+=$clockArgs
# Current workforce may use the current staged console evidence, never a future snapshot in history.
Run 'Devices' 'New-DeviceInventoryEvidence.ps1' $deviceArgs
$wf = @('-DataRoot',$DataRoot,'-UserOutputPath',(OutPath 'UserInventoryEvidence'),'-HistoryOutputPath',(OutPath 'UserActivityHistoryEvidence'),'-TrendOutputPath',(OutPath 'WorkforceTrendEvidence'),'-IdentityOutputPath',(OutPath 'IdentityReconciliationEvidence'),'-SignalsOutputPath',(OutPath 'WorkforceOperationalSignals'))
if ($Only -eq 'All') { $wf+=@('-SystemEvidencePath',(OutPath 'DeviceSystemEvidence')) }
if ($AccountClassificationConfigPath) { $wf += @('-AccountClassificationConfigPath',$AccountClassificationConfigPath) }
if ($SkipWorkforceHistory) { $wf += '-SkipHistory' }
$wf+=$clockArgs
if ($ReuseStagedWorkforce) {
    foreach ($name in @('UserInventoryEvidence','IdentityReconciliationEvidence','WorkforceOperationalSignals')) {
        if (-not (Test-Path -LiteralPath (OutPath $name))) { throw "Missing staged workforce output: $name" }
    }
} else {
    if ($null -ne $classificationStart) {
        $null = Test-PreparedClassificationCapture -Path $AccountClassificationConfigPath -Stage 'BeforeWorkforce' -ExpectedSHA256 $classificationStart.SHA256 -ExpectedBytes $classificationStart.Bytes
    }
    Run 'Workforce' 'New-WorkforceIdentityEvidence.ps1' $wf
}
if ($EndpointProtectionPolicyPath) {
    Run 'Devices' 'New-EndpointProtectionEvidence.ps1' (@('-DataRoot',$DataRoot,'-InventoryEvidencePath',(OutPath 'DeviceInventoryEvidence'),'-SystemEvidencePath',(OutPath 'DeviceSystemEvidence'),'-PolicyPath',$EndpointProtectionPolicyPath,'-OutputRoot',$OutputRoot)+$clockArgs)
}
Run 'Devices' 'New-WindowsMigrationEvidence.ps1' (@('-DataRoot',$DataRoot,'-PreparedInventoryPath',(OutPath 'DeviceInventoryEvidence'),'-OutputRoot',$OutputRoot)+$countryArgs)
Run 'Devices' 'New-DeviceLifecycleExperienceTrendEvidence.ps1' @('-DataRoot',$DataRoot,'-WindowsOutputPath',(OutPath 'WindowsLifecycleTrendEvidence'),'-EndpointOutputPath',(OutPath 'EndpointExperienceTrendEvidence'))
Run 'Applications' 'New-ApplicationInventoryEvidence.ps1' @('-DataRoot',$DataRoot,'-InventoryOutputPath',(OutPath 'ApplicationInventoryEvidence'),'-TrendOutputPath',(OutPath 'ApplicationTrendEvidence'))
Run 'Collaboration' 'New-CollaborationEvidence.ps1' (@('-DataRoot',$DataRoot,'-AdoptionOutputPath',(OutPath 'CollaborationAdoptionEvidence'),'-ObjectOutputPath',(OutPath 'CollaborationObjectEvidence'),'-TrendOutputPath',(OutPath 'CollaborationTrendEvidence'))+$clockArgs)
Run 'Content' 'New-ContentStorageEvidence.ps1' (@('-DataRoot',$DataRoot,'-ObjectOutputPath',(OutPath 'ContentStorageEvidence'),'-TrendOutputPath',(OutPath 'ContentStorageTrendEvidence'))+$clockArgs)
Run 'Mailbox' 'New-MailboxEvidence.ps1' @('-DataRoot',$last,'-UserEvidencePath',(OutPath 'UserInventoryEvidence'),'-OutputPath',(OutPath 'MailboxEvidence'))
$dailyTrendArgs=@('-DataRoot',$DataRoot,'-OutputRoot',$OutputRoot,'-PythonCommand',$PythonCommand,'-PartialThresholdPercent',[string]$DailyTrendsPartialThresholdPercent,'-PersistentAbsenceObservations',[string]$DailyTrendsPersistentAbsenceObservations)
if($DailyTrendsQualificationPath){$dailyTrendArgs+=@('-QualificationPath',$DailyTrendsQualificationPath)}
Run -Domain 'DailyTrends' -Script 'New-DailyTrendsEvidence.ps1' -Arguments $dailyTrendArgs
Run 'Security' 'New-SecurityEvidence.ps1' (@('-DataRoot',$last,'-ExtendedEvidenceRoot',$last,'-UserEvidencePath',(OutPath 'UserInventoryEvidence'),'-DeviceEvidencePath',(OutPath 'DeviceInventoryEvidence'),'-OutputPath',(OutPath 'SecurityControlEvidence'))+$clockUtcArgs)
Run 'Backup' 'New-BackupResilienceEvidence.ps1' (@('-DataRoot',$last,'-OutputRoot',$OutputRoot)+$clockDateArgs)
$intuneArgs=@('-DataRoot',$DataRoot,'-OutputRoot',$OutputRoot)
if($TenantKey){$intuneArgs+=@('-TenantKey',$TenantKey)}
if($IntuneSourcePolicyPath){$intuneArgs+=@('-SourcePolicyPath',$IntuneSourcePolicyPath)}
$intuneArgs+=$clockArgs
Run 'IntuneOperations' 'New-IntuneOperationalEvidence.ps1' $intuneArgs
$trustArgs=@('-DataRoot',$DataRoot,'-OutputPath',(OutPath 'DataTrustEvidence'))
if($Only -eq 'All'){$trustArgs+=@('-IntuneOperationalAuditPath',(Join-Path $OutputRoot 'IntuneOperationalAudit.json.txt'))}
$trustArgs+=$clockDateArgs+$clockUtcArgs
Run 'Trust' 'New-DataTrustEvidence.ps1' $trustArgs
Run 'Executive' 'New-ExecutiveTrendEvidence.ps1' @('-DataRoot',$DataRoot,'-OutputPath',(OutPath 'ExecutiveKPITrends'))
Write-Host 'Preparation finished. Outputs are staging only; schema/data validation and promotion are separate steps.'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCvwnbEQpqMbYf7
# IO1I6z90GwH9H6Ubp1AceCZ19TyqlaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINmmvX76iNiFSFC4oEhO/+vAUxWtO3qaEeFvfEuXThmtMA0GCSqG
# SIb3DQEBAQUABIIBgKnE8XA5LOz3q6g+FD3w5hPY8KrZ9Nh+F0ahpYKMV9GgHFpc
# DPbAJvdl0omI93pIjaaQBFMoeyHh+YpKR+Q+thcrcx0v3JiCpzAbEfZ1YRCvZVyh
# Ar37+yKkRy6cfz2CuTckF8xDWPBr81lI2T4I1g2zl1+/W8PafYCmnQ9d3kV1VjmT
# KnjKKpHADbt+IQMVW+ispboWKZQOHYaAO7nLrN9EvksxvI7SoLjGCFZaz5CrKdJM
# MTuBNL3VSphzJNDiwTvw8EhHSCR+RQ+t0T7uZaOb8wqYWgtkB5l1/N3E24afO2VA
# ln9k8yVurJp6TdM42WZxcxw1q6YUrEmfz6KC8SXYIacAA0KDjQ2LlEgHn2QebSZI
# 5bJIpwBsjT+r7U+JrEW9Zqp4TolNHTeLjCLmepoda65vUcldU4BLYYrck4mbh+Zg
# +XDmFQUNALusWwzM4wkp8FIqAuWRexAGTs2J3B0zuRwPcdpPjfu7oHIFdehMrIO5
# LbVWZ4zkqVT2vrOduaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTcyMzE1
# NTVaMC8GCSqGSIb3DQEJBDEiBCAZ6i8d+wJ6MF1cIqzUHZG9Q9cef6jE9fKU6DxL
# UeKdVjANBgkqhkiG9w0BAQEFAASCAgCDJwkLeCTpj0eKIJrGHWtT1ABtsh7qbEe9
# w/DVTluFx8PSg+JMEIHkMg8imOQQH3JKD+ttcZvRCbHtdIkQZxV0qJi4wM+EoVQ0
# GHeR3Pk/gR6cKXTk5hlRdqlPhn7oKTMglJ344AStZ/zFIhrS0fxH94fnh1sf8e5u
# oKcFfkyrZvclbK914SeqYZxm7QYoCJvbj+KntGmrHmyFzwmzgWh7lr6EXhLilWux
# 4gPqWGRmYYQN6wRpExaEtXmekAH3ummwDvviwndd3B4iJD1rrf7oHagJ+TNZx8Dz
# 1htqQIvSwq/wRr57onpMu6dwk2c9QXmeNrUXkLv9tPlJZ58SW2pb4ALMHuz9km0P
# iCvbVImf2tcS2FcwgPBbxrck+lT0/mU54zoWEvnwPYg4FcmHwmPgcBp8QKBolG1y
# MoH90e0tMjkOmIC5Z4WCX0E8StJJwkNt1RGruazF3oecHhYtEMuqwI1TsMwvZyVc
# tt4qAOuWWEF6dZMqwMriwmnRYeTqi7ktrzfodWZDQt1Y+NT1EU5G+lgKpsOeP7T8
# Jr9gTq573am9ID3qPA2L5SFcXLq0dJFDXpfdq/aoK6WUJllaJUlrzi1PlsriF0Jw
# BoEH1LOVC5r9ftLQRkLWiToO9UPVrFEd5p7UhLY/7KJN58BZkOVZQIBgVP1mBaS9
# S0didENnPQ==
# SIG # End signature block
