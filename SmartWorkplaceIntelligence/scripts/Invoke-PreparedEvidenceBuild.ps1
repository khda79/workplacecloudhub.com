# Version 1.1.0
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$DataRoot,
    [Parameter(Mandatory)][string]$OutputRoot,
    [ValidateSet('All','Estate','Workforce','Devices','Applications','IntuneOperations','Collaboration','Content','Mailbox','DailyTrends','Licensing','Security','Backup','Trust','Executive')]
    [string]$Only = 'All',
    [switch]$SkipWorkforceHistory,
    [switch]$ReuseStagedWorkforce,
    [string]$AccountClassificationConfigPath,
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
} else { Run 'Workforce' 'New-WorkforceIdentityEvidence.ps1' $wf }
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
Run 'Licensing' 'New-LicensingEvidence.ps1' (@('-DataRoot',$last,'-UserEvidencePath',(OutPath 'UserInventoryEvidence'),'-MailboxEvidencePath',(OutPath 'MailboxEvidence'),'-OutputPath',(OutPath 'LicenseEvidence'),'-OptimizationOutputPath',(OutPath 'LicenseOptimizationEvidence'))+$clockArgs)
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCA8GiaYC8op+Ux
# jSZ5sEfc6BYe8K8m3LhwelsDPSG/z6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAmSJq5OGqSLRrO/h50aje0
# PWrPRZsegm11z+P8ki0W2zANBgkqhkiG9w0BAQEFAASCAYAaKr2L0KwxLa4UwJUB
# 73f69kbYZw0mO5WrBejTvam7m2Svsm4QII13q3Tsg5n6Xw8nCBkjWTezV75DkdLW
# Tf2KBdd/DfG9fcR89rwKuJDTh89EnM8Dy99rabRz/Gfspz3b6AHiKeLbLn5xqVvh
# MVjuNiEDjQzy7Ett3VSCrPHWIGk3mnfk24CIvERPbVafT1puCp5/5l7fMR0YLBS0
# n0UeknAjpYP1Djv7tO+sBGiaJmik8IDENUKDTSlNXOfPE7ffaRan5J73cRrrOpRW
# +LP8SBPupM19kakp/LBmGCIuw4Qy9NgCa8eMmM5zlAGwbVYY95d44PtLJJGCu26D
# f1b4cpp+7qIi+sRT+rMq/g1cR7QQoe91rCrhdVURiHCCANdJr60zzaFzIOFm7rEH
# XzymCVJ3LyCFjC71hIYOSS/Hpw4Rej+kpunXWJLc0Ub5oa681zFUoW7RRrrmvLQR
# kbQvE49W8UCkS+8UrcKgty4W67k2OJHZLBzjmz4drW4KFAQ=
# SIG # End signature block
