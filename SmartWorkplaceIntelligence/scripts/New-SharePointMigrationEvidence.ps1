# Offline migration planning projection. No collection, refresh or transfer.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$DataRoot,
    [Parameter(Mandatory)][string]$OutputRoot,
    [Parameter(Mandatory)][string]$TenantKey,
    [switch]$Enabled,
    [string]$WorkbookFileName='SmartWorkplaceIntelligence-Migration-to-SPO.xlsx',
    [string]$BindingsFileName='',
    [string]$PreviousHistoryPath='',
    [string]$TargetTenantUrl='',
    [string]$SourcePolicyPath='',
    [ValidateRange(1,3650)][int]$ReviewMaxAgeDays=7,
    [ValidateRange(1,8760)][int]$MaxSourceAgeHours=168,
    [string]$ReportingTimeZoneId='UTC',
    [datetimeoffset]$AsOfUtc=[datetimeoffset]::UtcNow,
    [string]$PythonCommand='python'
)
$ErrorActionPreference='Stop'
$python=Get-Command $PythonCommand -CommandType Application -ErrorAction Stop | Select-Object -First 1
$pythonVersion=& $python.Source -B -c 'import sys; print(".".join(map(str,sys.version_info[:3])))'
if($LASTEXITCODE -ne 0 -or [version]$pythonVersion -lt [version]'3.10'){throw 'Python 3.10+ is required for migration planning evidence.'}
$zone=[TimeZoneInfo]::FindSystemTimeZoneById($ReportingTimeZoneId)
$reportingDate=[TimeZoneInfo]::ConvertTime($AsOfUtc,$zone).ToString('yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)
$arguments=@('-B',(Join-Path $PSScriptRoot 'workplace_spo_migration.py'),'--data-root',$DataRoot,'--output-root',$OutputRoot,'--tenant',$TenantKey,
    '--workbook',$WorkbookFileName,'--review-days',[string]$ReviewMaxAgeDays,'--max-age-hours',[string]$MaxSourceAgeHours,
    '--as-of',$AsOfUtc.ToUniversalTime().ToString('o'),'--reporting-date',$reportingDate)
if($Enabled){$arguments+='--enabled'}
if($BindingsFileName){$arguments+=@('--bindings',$BindingsFileName)}
if($PreviousHistoryPath){$arguments+=@('--previous-history',$PreviousHistoryPath)}
if($TargetTenantUrl){$arguments+=@('--target-tenant-url',$TargetTenantUrl)}
if($SourcePolicyPath){$arguments+=@('--source-policy',$SourcePolicyPath)}
& $python.Source @arguments
if($LASTEXITCODE -ne 0){throw 'SharePoint migration evidence projection failed. No source workbook was changed.'}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAUco8NFQtRUGxS
# Yr8I48Ql56UsHL0gw018DAIXa+yCQKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBO5idl78wBxLnAUyXRY9GV
# XNjYLPkRsMwXnRa1C/wUuDANBgkqhkiG9w0BAQEFAASCAYAQR+Rz2+qTv/SILRnd
# StJd+ciy/+Q5SeHGhwkhy+ZD8J913TO+kjrPyiZbdQzWACJJ1EOxqyngkxt3MOp0
# 36eftC5UjcNiUBHSSJDp+7YkyWWETFfh9mmkJQXBLNuR5LafPo+RDTbDHWNFswvO
# hSAPuDk8Q/VXvB0je9mdkZwd2ossNp381eeGaYj5ArlRclUVa06mlQswQNqWDN8w
# 8VwQg85GExsnQtOdu2xfeyAY/H9szkXgySDQIOBgDKLF7VRLDYdEQ6qR2ujadsmF
# nlqiQACfVi9sLQ3uiqlb22B4eLRVkQVSFoKMbun60sCtoeldnGN3xDAq+zfqcfqP
# TmWPS+hUmakQUeJUtwz8vPKV/+xf+pop0wE8A3sMhWYKRIpmEgrjEeERcQka7Frt
# 4ou31UmJXvyoVUi4AG5aLi1waHkvgq4hpNnnk+0UIDScyvkXjccsTIxYLz1DRHGJ
# 7rFVmMarBf6spUXISOjVoZtx/Zv1VwmiQLWVFO63+lasmDA=
# SIG # End signature block
