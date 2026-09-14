<#
.SYNOPSIS
Normalizes SharePoint and Teams collaboration data for Power BI.

.VERSION
1.1.0
#>
[CmdletBinding()]
param(
    [Alias('ProfileKey')][string]$Tenant='default',
    [string]$OrganizationKey,[string]$EnvironmentKey,[string]$TenantKey,[string]$TenantId,
    [string]$DataRootPath,[string]$DataAllRootPath,[string]$LatestOutputRootPath,[string]$LogRootPath,
    [string]$GlobalConfigPath,[string]$TenantConfigPath,[switch]$NoConfigWrite,[switch]$ValidateOnly
)
$ScriptVersion='1.1.0';$ErrorActionPreference='Stop';Set-StrictMode -Version 2.0
function Get-HeaderStatus{param([string]$Path,[string[]]$Columns)if(-not(Test-Path -LiteralPath $Path -PathType Leaf)){return 'Missing'};$line=Get-Content -LiteralPath $Path -TotalCount 1;$actual=if([string]::IsNullOrWhiteSpace($line)){@()}else{@($line.Split(',')|ForEach-Object{$_.Trim().Trim('"')})};if(($actual -join [char]31) -ceq ($Columns -join [char]31)){return 'Valid'};return 'Incompatible'}
function Get-KeyText{param([AllowNull()]$Value)if($null -eq $Value){return ''};return ([string]$Value).Trim().ToLowerInvariant()}
function Get-ActivityState{param([string]$LastActivityDate,[string]$CollectedDateTime)if([string]::IsNullOrWhiteSpace($LastActivityDate)){return 'Unknown'};$last=[datetimeoffset]::Parse($LastActivityDate);$collected=[datetimeoffset]::Parse($CollectedDateTime);if(($collected.UtcDateTime-$last.UtcDateTime).TotalDays -le 90){return 'Active (90d)'};return 'Inactive (>90d)'}
$scriptRoot=Split-Path -Parent $MyInvocation.MyCommand.Path;$projectRoot=Split-Path -Parent(Split-Path -Parent $scriptRoot)
$module=Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1';$rawPath=Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json';$curatedPath=Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json';Import-Module $module -Force
$bound=@{};foreach($key in $PSBoundParameters.Keys){$bound[$key]=$PSBoundParameters[$key]};$context=Resolve-SmartWorkplaceCMDBContext -BoundParameters $bound -GlobalConfigPath $GlobalConfigPath -TenantConfigPath $TenantConfigPath -NoConfigWrite:($ValidateOnly-or$NoConfigWrite);$paths=$context.Paths
$raw=Get-SmartWorkplaceCMDBTableContract $rawPath;$curated=Get-SmartWorkplaceCMDBTableContract $curatedPath
$maps=@(
 @{Raw='M365_SharePointSites.csv';Curated='DimSharePointSite.csv'},@{Raw='M365_Teams.csv';Curated='DimTeam.csv'},@{Raw='M365_TeamMembers.csv';Curated='FactTeamMember.csv'}
)
$defs=@();foreach($map in $maps){$r=@($raw.tables|Where-Object { $_.name -eq $map.Raw });$c=@($curated.tables|Where-Object { $_.name -eq $map.Curated });if($r.Count -ne 1 -or $c.Count -ne 1){throw "Collaboration contract definition missing or duplicated: $($map.Raw) / $($map.Curated)"};$input=[IO.Path]::GetFullPath((Join-Path $paths.LatestOutputRootPath (Join-Path ([string]$r[0].area) $map.Raw)));$output=[IO.Path]::GetFullPath((Join-Path $paths.LatestOutputRootPath (Join-Path ([string]$c[0].area) $map.Curated)));$status=Get-HeaderStatus $input @($r[0].columns|ForEach-Object{[string]$_});if($status -eq 'Incompatible'){throw "Collaboration CSV contract is incompatible: $input"};$defs+=@{Map=$map;Raw=$r[0];Target=$c[0];Input=$input;Output=$output;Status=$status}}
if($ValidateOnly){foreach($def in $defs){if($def.Status-eq'Valid'){Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $def.Input -Paths $paths|Out-Null}};[pscustomobject]@{Status='Valid';ScriptVersion=$ScriptVersion;DatasetCount=$defs.Count;RawContractVersion=[string]$raw.contractVersion;CuratedContractVersion=[string]$curated.contractVersion}|Format-List;return}
$identity=@{TenantKey=$paths.TenantKey;OrganizationKey=$paths.OrganizationKey;EnvironmentKey=$paths.EnvironmentKey;TenantId=$paths.TenantId};$published=@()
foreach($def in $defs){if($def.Status-eq'Missing'){throw "Required raw collaboration inventory is missing: $($def.Input)"};$rows=@(Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $def.Input -Paths $paths);$targetRows=@($rows|ForEach-Object{$row=$_;switch($def.Map.Raw){
 'M365_SharePointSites.csv'{$key=Get-KeyText $row.SiteId;[pscustomobject][ordered]@{TenantSiteKey=('{0}|site|{1}'-f$paths.TenantKey,$key);SiteId=[string]$row.SiteId;SiteUrl=[string]$row.SiteUrl;SiteName=[string]$row.SiteName;OwnerPrincipalName=[string]$row.OwnerPrincipalName;LastActivityDate=[string]$row.LastActivityDate;ActivityState=Get-ActivityState $row.LastActivityDate $row.SourceCollectedDateTime;StorageUsedBytes=[string]$row.StorageUsedBytes;StorageAllocatedBytes=[string]$row.StorageAllocatedBytes;RootWebTemplate=[string]$row.RootWebTemplate;IsDeleted=[string]$row.IsDeleted;SourceCollectedDateTime=[string]$row.SourceCollectedDateTime}}
 'M365_Teams.csv'{$key=Get-KeyText $row.TeamId;[pscustomobject][ordered]@{TenantTeamKey=('{0}|team|{1}'-f$paths.TenantKey,$key);TeamId=[string]$row.TeamId;DisplayName=[string]$row.DisplayName;Visibility=[string]$row.Visibility;CreatedDateTime=[string]$row.CreatedDateTime;LastActivityDate=[string]$row.LastActivityDate;ActivityState=Get-ActivityState $row.LastActivityDate $row.SourceCollectedDateTime;OwnerCount=[string]$row.OwnerCount;MemberCount=[string]$row.MemberCount;GuestCount=[string]$row.GuestCount;UnresolvedMemberCount=[string]$row.UnresolvedMemberCount;MembershipCoverageStatus=[string]$row.MembershipCoverageStatus;IsArchived=[string]$row.IsArchived;SourceCollectedDateTime=[string]$row.SourceCollectedDateTime}}
 'M365_TeamMembers.csv'{$team=Get-KeyText $row.TeamId;$user=Get-KeyText $row.UserId;[pscustomobject][ordered]@{TenantTeamMemberKey=('{0}|team-member|{1}|{2}'-f$paths.TenantKey,$team,$user);TenantTeamKey=('{0}|team|{1}'-f$paths.TenantKey,$team);TeamId=[string]$row.TeamId;TenantUserKey=('{0}|user|{1}'-f$paths.TenantKey,$user);UserId=[string]$row.UserId;UserPrincipalName=[string]$row.UserPrincipalName;UserType=[string]$row.UserType;Role=[string]$row.Role;SourceCollectedDateTime=[string]$row.SourceCollectedDateTime}}
 }});Export-SmartWorkplaceCMDBCsv -InputObject $targetRows -Path $def.Output -Columns @($def.Target.columns|ForEach-Object{[string]$_}) @identity;if((Get-HeaderStatus $def.Output @($def.Target.columns|ForEach-Object{[string]$_}))-ne'Valid'){throw "Curated collaboration output failed validation: $($def.Output)"};$published+=[pscustomobject]@{Table=$def.Map.Curated;Count=$targetRows.Count;Path=$def.Output}}
Write-Information ("SmartWorkplaceCMDB collaboration normalization completed. Tables={0}; Rows={1}."-f$published.Count,(($published|Measure-Object Count -Sum).Sum)) -InformationAction Continue
[pscustomobject]@{Status='Completed';ScriptVersion=$ScriptVersion;Published=$published;RawContractVersion=[string]$raw.contractVersion;CuratedContractVersion=[string]$curated.contractVersion}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCmjyftrVLbhbvh
# GLktovqZVrMctKGJs10NC8f6JF9vBqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFPdp8BKllCnTYj6HXP8A+Yih+NI+cO+/NsFTmtjXYo9MA0GCSqG
# SIb3DQEBAQUABIIBgKq5Iw1Q473xhceogZi1DqxCTCC9J8z5J6yD54tQ413zAUKf
# SXfjGcIUQr7bOvkj3GB/HayCbN2UsWg85TCbcW+0bbPf6Ybzsd1WG+SR7m0ZRA37
# 5xriRQru67xWsfUGRINAfIdWH917RSUkfa/vkURbqKgJGFgeE3q3yikhuMhz7WZr
# BXT6ywTrKwxER7kgPqycGGeGWiBvI8y9Oq87Axc7KMwuV5xdnNyNFvepwfQ7uaj6
# 4wfaqTjuxweKnKlnEh73MVgWOzB16zJH2e4g80o/X97qNFMJTnqLFj5/+lb3ItBI
# FlV7cG9qOQrbsFUXicXw7YrygSZs3g9mkXq4Tg5WsAVRcC6b/HM+TviKZg8gZttH
# 3dO/yS0lc1k7Ag/giucO4KGxZAVH7NEhMGczPj4h6e8Wb45ZUHuFhRIR0CbWMYw8
# zOBazchiAVSb6MZwY0ejqiKq5f3sJw7I+tg2CmSG5rGEkFzlnwnmgRKFE+yBcIah
# Jy4luc179EWIjPXNkqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTQwNDI0
# NTVaMC8GCSqGSIb3DQEJBDEiBCB8+C2LX0qI1xNrZH+3b4v84llQXFU2AVjjPAQO
# GG62lTANBgkqhkiG9w0BAQEFAASCAgAW8pwrPOWAEUdPMrYXK7n/cH9LXb4Jtixg
# IXVH0iAzdLmSGqY/T9UmI9IEcSBASWE9fM38egM0P4LE9hLi/8KtAeT34llkRgXX
# m2dcyaSNMKnVXHabeM3T3ouWVP9dsc5/98lgqE0s0qLS70/edE27hEqes8fU9np2
# xDBvsziW3HMliRGKG+ZC5E5kv923POI6keUmoMo6JxV84DzcOXHLakbTcqS1ajE9
# BvehoH1jRONhmSVOHfKLHLq4p1e5Yg6hxh672LkPEQ1XVkiBy9/B4hJX3ZSh1HMt
# ENn/flCbA8K/IPT0h5iOG/bvSgP655wY+24YL1IedhWMP3iIalHL1SdaU/XKm5Tf
# 4CtdZLzEGK7HSKwCx2AsuzXKdRs5bFG7d5eKy+P50lh8bYsW2KcLmiW2oXbtbYb/
# EnSc1Y3WOPOi8aZXSh9Bxf/T0BgSRdivyVHUbCJdkKXhpXz2hYza7ePfaqlbBu4M
# 5L2LRgAyRwGcrlayZ6sB5YLfb5FCcRo2tD8n/hZp7P3A77uLVFXmlfOUQ1U/1IXT
# XgB3FwqwdIhzanAWOiPz5vtZoSYsp/+ObPwH4OiutWcW1gM5WRI7K0xFQuKCeyTv
# ygvKuKe13ExeLzJFINNX+m8OS9W55fJWk3wOVlsRBcuZ+biqlOzdeCQvnZ4yVQEI
# SUd0atEhTA==
# SIG # End signature block
