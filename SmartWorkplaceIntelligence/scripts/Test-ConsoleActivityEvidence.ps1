#Requires -Version 7.0
# Synthetic-only tests: no collector, production preparation, Graph or Power BI.
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'ConsoleActivityEvidence.psm1') -Force
$script:checks=0
function Assert-Test([bool]$Condition,[string]$Message){$script:checks++;if(-not $Condition){throw $Message}}
$asOf=[datetimeoffset]'2026-10-13T12:00:00Z'
$sid='S-1-5-21-1-2-3-1001'
$ad=[pscustomobject]@{TenantKey='synthetic';ObjectSID=$sid;ObjectGUID='ad-user';UserPrincipalName='person@example.invalid';LastLogonDate='2026-04-01'}
function New-ConsoleRow([string]$Id='device-one') {
    [pscustomobject]@{TenantKey='synthetic';'Device Source ID'=$Id;'Local Data Source'='IntuneRemediation';'Local Data Freshness'='Fresh';Freshness='Within policy';'Inventory Match'='Matched';'Prepared Inventory Match'='Matched';'Observation UTC'='2026-10-13T11:00:00Z';'Last Console Logon Status'='Observed';'Last Console Logon User'='EXAMPLE\person';'Last Console Logon SID'=$sid;'Last Console Logon UTC'='2026-10-11T09:00:00Z'}
}
function Build([object[]]$Rows=@((New-ConsoleRow)),[object[]]$AdRows=@($ad),[datetimeoffset]$At=$asOf){New-ConsoleActivityIndex -SystemRows $Rows -AdRows $AdRows -AsOfUtc $At}
$index=Build
$activity=Get-ConsoleActivityForAdUser -Index $index -AdRow $ad
Assert-Test ($activity.MatchMethod -eq 'SID' -and $activity.DeviceCount -eq 1) 'Unique SID attribution failed.'
Assert-Test ($activity.Utc -eq [datetimeoffset]'2026-10-11T09:00:00Z') 'Qualified console time changed.'
$other=New-ConsoleRow 'device-two';$other.'Last Console Logon UTC'='2026-10-12T09:00:00Z'
$activity=Get-ConsoleActivityForAdUser -Index (Build -Rows @((New-ConsoleRow),$other)) -AdRow $ad
Assert-Test ($activity.DeviceCount -eq 2 -and $activity.Utc.Day -eq 12) 'Latest activity across devices was not selected.'
foreach($state in @('NotCollected','Stale','Error','Unknown','Pending','NotApplicable','Failed','ScriptError','InvalidOutput','InvalidTimestamp','FutureObservation')) {
    $row=New-ConsoleRow;$row.'Local Data Freshness'=$state
    Assert-Test ((Build -Rows @($row)).ByAdSid.Count -eq 0) 'Unavailable local collection became activity.'
}
foreach($state in @('NotCollected','NotFoundInScan','Error','Unknown')) {
    $row=New-ConsoleRow;$row.'Last Console Logon Status'=$state
    Assert-Test ((Build -Rows @($row)).ByAdSid.Count -eq 0) 'Unobserved console state became activity.'
}
foreach($state in @('Stale','Future observation','Unqualified recurring time','Missing')) {
    $row=New-ConsoleRow;$row.Freshness=$state
    Assert-Test ((Build -Rows @($row)).ByAdSid.Count -eq 0) 'Unqualified observation became activity.'
}
foreach($field in @('Inventory Match','Prepared Inventory Match')) {
    $row=New-ConsoleRow;$row.$field='Unmatched'
    Assert-Test ((Build -Rows @($row)).ByAdSid.Count -eq 0) 'Unmatched device became an attributed user.'
}
foreach($badSid in @('','not-a-sid','S-1-5-18','S-1-12-1-1-2-3-4','S-1-5-21-5-6-7-1001')) {
    $row=New-ConsoleRow;$row.'Last Console Logon SID'=$badSid
    Assert-Test ((Build -Rows @($row)).ByAdSid.Count -eq 0) 'Missing, local, cloud or unmatched SID inferred a user.'
}
Assert-Test ((Build -AdRows @($ad,$ad)).ByAdSid.Count -eq 0) 'Ambiguous AD SID was accepted.'
Assert-Test ((Build -Rows @((New-ConsoleRow),(New-ConsoleRow))).ByAdSid.Count -eq 0) 'Duplicate device evidence was accepted.'
$row=New-ConsoleRow;$row.TenantKey='other'
Assert-Test ((Build -Rows @($row)).ByAdSid.Count -eq 0) 'Cross-tenant SID attribution occurred.'
$otherAd=$ad.PSObject.Copy();$otherAd.TenantKey='other'
$index=Build -Rows @((New-ConsoleRow),$row) -AdRows @($ad,$otherAd)
Assert-Test ($index.ByAdSid.Count -eq 2) 'Equal SIDs in isolated tenants were conflated.'
foreach($time in @('','bad','2026-10-11T09:00:00','2026-10-13T11:30:00Z','2026-10-14T09:00:00Z')) {
    $row=New-ConsoleRow;$row.'Last Console Logon UTC'=$time
    Assert-Test ((Build -Rows @($row)).ByAdSid.Count -eq 0) 'Unqualified or future console time became activity.'
}
$row=New-ConsoleRow;$row.'Last Console Logon UTC'='2026-10-11T11:00:00+02:00'
Assert-Test ((Get-ConsoleActivityForAdUser -Index (Build -Rows @($row)) -AdRow $ad).Utc -eq [datetimeoffset]'2026-10-11T09:00:00Z') 'Explicit console offset was not normalized.'
$row=New-ConsoleRow;$row.'Observation UTC'='2026-10-14T09:00:00Z'
Assert-Test ((Build -Rows @($row)).ByAdSid.Count -eq 0) 'Future collection was accepted.'
Assert-Test ((Build -At ([datetimeoffset]'2026-10-01T00:00:00Z')).ByAdSid.Count -eq 0) 'A future snapshot leaked into history.'
Assert-Test ((Build -Rows @()).ByAdSid.Count -eq 0) 'Empty console evidence invented activity.'

# Load only the real pure classification functions. Never execute the workforce generator.
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'New-WorkforceIdentityEvidence.ps1'),[ref]$tokens,[ref]$errors)
Assert-Test ($errors.Count -eq 0) 'Workforce generator did not parse.'
foreach($name in @('Get-NormalizedKey','Get-PropertyValue','Convert-ToDateTimeOrNull','Convert-AdLogonDateTimeOrNull','Get-ActivityClassification')) {
    $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    Assert-Test ($null -ne $node) 'Classification helper missing.'
    . ([scriptblock]::Create($node.Extent.Text))
}
$cloud=[pscustomobject]@{LastActivityDate='2026-07-01';HasAnyM365Activity='true'}
$console=Get-ConsoleActivityForAdUser -Index (Build) -AdRow $ad
$current=Get-ActivityClassification -ActivityRow $cloud -AdRow $ad -AsOfDate $asOf.UtcDateTime.Date -ConsoleActivity $console -IncludeConsole
Assert-Test ($current.ActivityState -eq 'Active <=30D' -and $current.DaysSinceLastActivity -eq 2) 'Recent console activity did not prevent false inactivity.'
Assert-Test ($current.ActivitySource -eq 'Active Directory + Microsoft 365 + Local console') 'Activity provenance lost the console source.'
Assert-Test ($current.M365LastActivityDate.Date -eq [datetime]'2026-07-01') 'Console activity changed the M365 usage date.'
foreach($days in @(30,31,90,91)) {
    $dated=[pscustomobject]@{Utc=$asOf.AddDays(-$days)}
    $classified=Get-ActivityClassification -AsOfDate $asOf.UtcDateTime.Date -ConsoleActivity $dated -IncludeConsole
    $expected=if($days -le 30){'Active <=30D'}elseif($days -le 90){'Dormant 31-90D'}else{'Stale >90D'}
    Assert-Test ($classified.ActivityState -eq $expected) 'The existing inactivity thresholds changed.'
}
$emptyCloud=[pscustomobject]@{LastActivityDate='';HasAnyM365Activity='false'}
$unknown=Get-ActivityClassification -ActivityRow $emptyCloud -AsOfDate $asOf.UtcDateTime.Date -IncludeConsole
Assert-Test ($unknown.WorkforceStatus -eq 'Unknown' -and $unknown.EvidenceStatus -eq 'Unknown') 'No attributable console evidence became a never-used workforce account.'
$history=Get-ActivityClassification -ActivityRow $cloud -AdRow $ad -AsOfDate ([datetime]'2026-10-01') -ConsoleActivity $console
Assert-Test ($null -eq $history.ConsoleLastActivityDate -and $history.ActivitySource -eq 'Active Directory + Microsoft 365') 'The historical classification consumed the current console snapshot.'
Assert-Test ((Get-ActivityClassification -ActivityRow $emptyCloud -AsOfDate $asOf.UtcDateTime.Date).ActivityState -eq 'Observed never used') 'Historical legacy classification was silently rewritten.'
$pipeline=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'Invoke-PreparedEvidenceBuild.ps1'))
Assert-Test ($pipeline.IndexOf("Add-Task 'DeviceInventory' 'Devices' 'New-DeviceInventoryEvidence.ps1'") -lt $pipeline.IndexOf("Add-Task 'Workforce' 'Workforce' 'New-WorkforceIdentityEvidence.ps1'") -and $pipeline.Contains("Add-Task 'Workforce' 'Workforce' 'New-WorkforceIdentityEvidence.ps1' `$wf -DependsOn 'DeviceInventory'")) 'Workforce runs before current staged device evidence.'
Assert-Test ($pipeline.Contains("'-SystemEvidencePath',(OutPath 'DeviceSystemEvidence')")) 'The prepared console source is not wired to workforce.'
$historyCall=@($ast.FindAll({param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$weekClassification'},$true))
Assert-Test ($historyCall.Count -eq 1 -and $historyCall[0].Extent.Text -notmatch 'ConsoleActivity|IncludeConsole') 'Historical weeks receive today console evidence.'
$metadataTokens=$null;$metadataErrors=$null
$metadataAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1'),[ref]$metadataTokens,[ref]$metadataErrors)
Assert-Test ($metadataErrors.Count -eq 0) 'Preparation pipeline did not parse.'
$metadataNode=$metadataAst.Find({param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$provenance.WorkforceActivity'},$true)
Assert-Test ($null -ne $metadataNode) 'Activity-definition provenance is missing.'
$provenance=@{}
. ([scriptblock]::Create($metadataNode.Extent.Text))
Assert-Test ($provenance.WorkforceActivity.DefinitionVersion -eq 'AD-M365-Console-v1' -and -not $provenance.WorkforceActivity.CurrentConsoleInjectedIntoHistory -and $provenance.WorkforceActivity.HistoricalSources.Count -eq 2) 'Current/history definition boundary is not recorded.'
[pscustomobject]@{Passed=$true;Checks=$script:checks;Sources='Synthetic only';Collection=$false;Preparation=$false;Transfer=$false;Refresh=$false} | ConvertTo-Json

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD0EcLIN97Gi9GJ
# M4NyQQWw0qRWDXUlCi24ehy5h8LjhKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAHQKRICg6eKijLZqDvoKea1UMx3JVKK5HjWSIb4ucKuMA0GCSqG
# SIb3DQEBAQUABIIBgGv0RqWIgEVKhV3Ul1Nq/eltCip9l9b7G817yl9l4Ck8QGDI
# +qlOLThOB+XjE9NA9P+N2p75/IBLT5top7zGwGnUNPUmONS9SERSUBvLvcVm3bE7
# +TqxUR8ykZQH5In9GMMNSUyWBD0UuGP8Kdbr6nXDmqZotpgXZ4lotsgIfokMx/TA
# wAoARsuYMQCv8/WpfykKgTK7lg7OdZRjxwWBpKy4BKqNfRwYCtgZroPVa/HdGutm
# L4N0jdH10r1lcvOM+Al5LMWDtz/310xRgWUek0uX3rYc1yQWos6RVXaaMJbKjyeo
# mhdMJuqgc6JKo7FFnT8rWb5ZetOQog6NaAb+g0x7+riKPWL+TTE9wfnB5hxTFF1j
# atf2ZSgcVLwzWRglfwCrzhzABIqSiXCjFc34H/JC80LkywMza89qXPBAHOuA4TI1
# 3jcBw2t2B+SjBRvAncayD7uIU1SUXVfaiylUAnt/dwcKo0d8HzmNetHP8ujj2gdW
# 34hXMVFekNQWqC1N1qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTMxNzMx
# MzlaMC8GCSqGSIb3DQEJBDEiBCCqtm74qUVrxljX3AGbypr7T7+PB+BXpU4BRzAx
# XLk8dTANBgkqhkiG9w0BAQEFAASCAgCJwmMZ/Z7r9mVL4Nb31Z9hfrWY7sRflres
# jbINv63LLYBT7edntvpyDRBlwi2QzbYFib2fryCUVocf5JyjeJ2OQFDB7CgTLaf7
# O4LYTOOGHbknTnnpSNo5FFUnmDpvwXcppMz3V1BapJjXouiNGCrk7Uwb1MevfNdW
# Dos0vHtuJ/UB4NAbAe4H+w63VQvznGNrkr3A47ggLwu4W6l0TDJbEhmaceLsVvcy
# 4yI3ZN+RVuDFfhDwV1445DiZBo+xznqNAYEpmS5V0D9l0Z1oo9Kac1AiJ9i2zj4I
# CK4u7EX/e31mDTipWZmj20Gxh5oz4+LBEXvGmUThmZqm4c1c1IrKPUBHlY5miNpu
# zhj7zw/hFJ1Y1M1Cuk6TRRZPM2d5/J3kXZmJx5736zRxqe8p+7bM9VVvRXPzZK8+
# o39mJHS6voPG5Xrrix/YUee8R6pAApziNUU9JtgT0NPJf0Dqc6hsAAra6rlE7hAc
# KRycfGsJl4kmsrcJduEylWTb2Kirl+bVCKw1ZlcmWHylvwvtpvt3btLK0y2tzvxW
# 6Y2y0Qryk9C4JKAZxT9kou1nwUcLmcm3Nuh2e6rwvamUahOOCsCq3iWmuL2gQB0R
# yzeTkx0IszhDkkYzTEMA9X16hbtc/GdRLTY6fcHrLplBQipAESEhBm3DLF3M0pqH
# pTEnf4RjIw==
# SIG # End signature block
