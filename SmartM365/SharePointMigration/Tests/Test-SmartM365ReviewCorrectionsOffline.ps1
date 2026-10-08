<#
.SYNOPSIS
    Offline checks for scan ordering, site-scoped memberships and connection/token handling.
.VERSION
    1.0.0
#>
#Requires -Version 5.1
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2
$root=Join-Path $PSScriptRoot '..'
. (Join-Path $root 'Scripts/Launchers/SmartM365-SharePointMigration-LauncherCommon.ps1')
function Import-TestFunction {
    param([string]$Path,[string[]]$Names)
    $tokens=$null; $errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseInput([IO.File]::ReadAllText($Path),[ref]$tokens,[ref]$errors)
    if($errors.Count) { throw ($errors | ForEach-Object Message) }
    foreach($name in $Names) {
        $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
        if(-not $node) { throw "Missing function: $name" }
        $body=$node.Body.Extent.Text
        Set-Item "function:global:$name" ([scriptblock]::Create($body.Substring(1,$body.Length-2)))
    }
}
function Assert-True { param([bool]$Condition,[string]$Message) if(-not $Condition) { throw $Message } }
$script:Messages=New-Object 'Collections.Generic.List[string]'
$script:Warnings=New-Object 'Collections.Generic.List[string]'
$script:Delays=New-Object 'Collections.Generic.List[int]'
function Write-ConsoleMessage { param($Message,$ForegroundColor) $script:Messages.Add($Message) }
function Write-ConsoleWarning { param($Message) $script:Warnings.Add($Message) }
function Write-SPOReadRetry { param($Message) }
function Start-Sleep { param([int]$Seconds) $script:Delays.Add($Seconds) }
$temp=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ReviewTest-'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $temp)
try {
    Import-TestFunction (Join-Path $root 'SmartM365-SharePointMigration-Summary.ps1') @('Get-SmartM365LatestPortfolioScan')
    Import-TestFunction (Join-Path $root 'SmartM365-SharePointMigration-GUI.ps1') @('Get-LatestCsvFile','Get-CsvFileItems')
    Import-TestFunction (Join-Path $root 'Scripts/Launchers/Generic/SmartM365-SharePointMigration-Launcher.ps1') @('Get-LatestCsv')
    $old=Join-Path $temp 'Inventory-20261007-010000.csv'; $new=Join-Path $temp 'Inventory-20261008-010000.csv'
    foreach($p in @($old,$new)) { [IO.File]::WriteAllText($p,"Name;Value`r`nfile;1`r`n") }
    (Get-Item $old).LastWriteTime=(Get-Date).AddDays(1)
    Assert-True ((Get-LatestCsv $temp 'Inventory-*.csv') -eq $new) 'Launcher selected copy time rather than collection time.'
    Assert-True ((Get-LatestCsvFile $temp 'Inventory-*.csv').FullName -eq $new -and (Get-SmartM365LatestPortfolioScan $temp 'Inventory-*.csv').File.FullName -eq $new) 'GUI and launcher scan ordering disagree.'
    @{SchemaVersion=1; InventoryFile=[IO.Path]::GetFileName($old); Rows=1; Sha256=(Get-FileHash $old).Hash; CompletedAtUtc=[datetime]::UtcNow.AddMinutes(1).ToString('o')} | ConvertTo-Json | Set-Content "$old.manifest.json.txt" -Encoding utf8
    Assert-True ((Get-LatestCsv $temp 'Inventory-*.csv') -eq $old) 'Receipt completion date did not take precedence over the filename.'
    $display=@(Get-CsvFileItems $temp 'Inventory-*.csv')[0].Display
    Assert-True ($display.StartsWith((Get-SmartM365LatestPortfolioScan $temp 'Inventory-*.csv').Date.ToString('yyyy-MM-dd HH:mm'))) 'GUI item label uses copy time.'
    [IO.File]::WriteAllText("$old.manifest.json.txt",'invalid receipt')
    Assert-True ((Get-LatestCsv $temp 'Inventory-*.csv') -eq $new) 'Invalid receipt remained eligible.'
    [IO.File]::WriteAllText((Join-Path $temp 'Inventory-20261008-010000-Errors.csv'),'Scope;Message')
    $failed=$false; try { Get-LatestCsv $temp 'Inventory-*.csv' | Out-Null } catch { $failed=$true }
    Assert-True ($failed -and @(Get-CsvFileItems $temp 'Inventory-*.csv').Count -eq 0) 'An incomplete scan remained eligible.'

    $permission=Join-Path $root 'Scripts/Inventory/SmartM365-SharePointTarget-PermissionInventory.ps1'
    Import-TestFunction $permission @('Get-PrincipalMembershipInfo','Get-EmptyPrincipalMembershipInfo','Join-PrincipalMemberValues','Invoke-SPORead','Get-SPOExceptionDetails','Write-InventoryError','Connect-ToSPOWeb','ConvertFrom-Base64Url','Write-PnPTokenSummary')
    $script:ErrorPath=Join-Path $temp 'membership-errors.csv'; $script:ErrorCsvCreated=$false
    $script:SharePointGroupMembershipCache=@{}; $script:SPOPermissionConnection='fixture'; $script:GroupCalls=0; $script:GroupMode='OK'
    function Get-PnPGroupMember {
        [CmdletBinding()] param($Group,$Connection)
        $script:GroupCalls++
        if($script:GroupMode -eq 'Denied') { throw 'Access denied.' }
        if($script:GroupMode -eq 'Transient' -and $script:GroupCalls -lt 3) { throw 'Connection reset.' }
        [pscustomobject]@{LoginName=$script:GroupUser; Title=$script:GroupUser; PrincipalType='User'}
    }
    $group=[pscustomobject]@{Id=7;Title='Members';LoginName='Members'}
    $script:GroupUser='first@example.test'
    $first=Get-PrincipalMembershipInfo $group 'SharePointGroup' 'https://example.test/sites/first'
    $script:GroupUser='second@example.test'
    $second=Get-PrincipalMembershipInfo $group 'SharePointGroup' 'https://example.test/sites/second'
    Assert-True ($first.PrincipalMemberLoginNames -ne $second.PrincipalMemberLoginNames -and $script:GroupCalls -eq 2) 'Target membership cache crosses site collections.'
    [void](Get-PrincipalMembershipInfo $group 'SharePointGroup' 'https://example.test/sites/first')
    Assert-True ($script:GroupCalls -eq 2) 'Successful membership was not cached.'
    $script:GroupMode='Denied'; $script:GroupCalls=0
    $info=Get-PrincipalMembershipInfo $group 'SharePointGroup' 'https://example.test/sites/denied'
    Assert-True ($script:ErrorCsvCreated -and $info.PrincipalMemberLookupStatus -like 'Failed:*' -and $script:GroupCalls -eq 1) 'Membership denial was silently accepted or retried.'
    $script:GroupMode='OK'; [void](Get-PrincipalMembershipInfo $group 'SharePointGroup' 'https://example.test/sites/denied')
    Assert-True ($script:GroupCalls -eq 2) 'Failed memberships were cached.'
    Remove-Item $script:ErrorPath; $script:ErrorCsvCreated=$false; $script:GroupMode='Transient'; $script:GroupCalls=0; $script:Delays.Clear()
    $info=Get-PrincipalMembershipInfo $group 'SharePointGroup' 'https://example.test/sites/recover'
    Assert-True ($info.PrincipalMemberLookupStatus -eq 'OK' -and $script:GroupCalls -eq 3 -and -not $script:ErrorCsvCreated -and ($script:Delays -join ',') -eq '5,15') 'Transient membership read did not recover without losing completeness.'

    $script:Token='header.'+[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('{"azp":"fixture-app","roles":["Sites.Read.All"]}')).TrimEnd('=').Replace('+','-').Replace('/','_')+'.signature'
    function Get-PnPAccessToken { [CmdletBinding()] param($ResourceTypeName,$Connection) return $script:Token }
    $script:Messages.Clear(); $script:Warnings.Clear()
    Write-PnPTokenSummary -Connection ([pscustomobject]@{})
    Assert-True ($script:Messages[0] -match 'fixture-app.*Sites.Read.All' -and $script:Warnings.Count -eq 0 -and ($script:Messages -join '') -notlike "*$script:Token*") 'Optional token claims failed strict mode or exposed a token.'
    $script:Token=ConvertTo-SecureString -String $script:Token -AsPlainText -Force
    Write-PnPTokenSummary -Connection ([pscustomobject]@{})
    Assert-True ($script:Messages[1] -match 'fixture-app.*Sites.Read.All') 'Secure token return was not handled.'
    function Get-PnPAccessToken { [CmdletBinding()] param($ResourceTypeName,$Connection) throw 'Optional resource unavailable.' }
    Write-PnPTokenSummary -Connection ([pscustomobject]@{})
    Assert-True ($script:Warnings.Count -eq 0 -and $script:Messages[2] -like '*unavailable*') 'Optional token diagnostics produced warning noise.'

    $Tenant='fixture'; $TenantId=''; $ClientId='fixture'; $Thumbprint=''; $DeviceLogin=$false; $Interactive=$true; $ForceAuthentication=$false
    $script:SPOPermissionConnections=@{}; $script:Connections=0
    $script:TokenSummaryContexts=New-Object 'Collections.Generic.HashSet[string]'
    function Connect-PnPOnline { [CmdletBinding()] param($Url,$Tenant,$ClientId,[switch]$Interactive,[switch]$ReturnConnection,[switch]$PersistLogin) $script:Connections++; return [pscustomobject]@{Url=$Url; Number=$script:Connections} }
    function Write-SPOConnectionIdentity { param($Connection,$Url) $script:SPOConnectedAccount='fixture@example.test' }
    Connect-ToSPOWeb 'https://example.test/sites/first'
    $initial=$script:SPOPermissionConnection
    Connect-ToSPOWeb 'https://example.test/sites/second'
    Connect-ToSPOWeb 'https://example.test/sites/first/'
    Assert-True ($script:Connections -eq 2 -and $script:SPOPermissionConnection -eq $initial) 'Permission connection was not reused or crossed web URLs.'
    Import-TestFunction (Join-Path $root 'Scripts/Inventory/SmartM365-SharePointTarget-FileInventory.ps1') @('Connect-SPOInventory')
    function Write-Info { param($Color,$Message) }
    $UseEnvironmentVariables=$false; $ManagedIdentity=$false; $CertificatePath=''; $CertificatePassword=$null; $PersistLogin=$false
    $script:ForceAuthenticationAlreadyUsed=$false; $script:UsePersistedLoginForRun=$false; $script:PersistedLoginCleared=$false; $script:SPOFileConnections=@{}; $script:Connections=0
    $initial=Connect-SPOInventory 'https://example.test/sites/first'
    [void](Connect-SPOInventory 'https://example.test/sites/second')
    $again=Connect-SPOInventory 'https://example.test/sites/first/'
    Assert-True ($script:Connections -eq 2 -and $again -eq $initial) 'File connection was not reused or crossed web URLs.'

    # A small native-type fixture exercises the on-premises SPGroup branch without a farm.
    Add-Type -TypeDefinition @"
namespace Microsoft.SharePoint {
 public class SPGroup {
  public int ID = 7; public string LoginName = "Members"; public object[] Data; public string Fault; public int Calls;
  public object[] Users { get { Calls++; if(Fault == "Denied") throw new System.UnauthorizedAccessException("Access denied."); if(Fault == "Transient" && Calls < 3) throw new System.TimeoutException("Request timed out."); return Data; } }
 }
}
"@
    Import-TestFunction (Join-Path $root 'Scripts/Inventory/SmartM365-SharePointSource-PermissionInventory.ps1') @('Get-PrincipalMembershipInfo')
    $script:SharePointGroupMembershipCache=@{}
    $a=New-Object Microsoft.SharePoint.SPGroup; $c=New-Object Microsoft.SharePoint.SPGroup
    $a.Data=@([pscustomobject]@{LoginName='first';Name='First';IsDomainGroup=$false})
    $c.Data=@([pscustomobject]@{LoginName='second';Name='Second';IsDomainGroup=$false})
    $first=Get-PrincipalMembershipInfo $a 'https://example.test/sites/first'
    $second=Get-PrincipalMembershipInfo $c 'https://example.test/sites/second'
    Assert-True ($first.PrincipalMemberLoginNames -eq 'first' -and $second.PrincipalMemberLoginNames -eq 'second') 'Source membership cache crosses site collections.'
    $a.Fault='Denied'; $script:ErrorCsvCreated=$false
    $info=Get-PrincipalMembershipInfo $a 'https://example.test/sites/denied'
    Assert-True ($script:ErrorCsvCreated -and $info.PrincipalMemberLookupStatus -like 'Failed:*') 'Source membership denial was silently accepted.'
    Remove-Item $script:ErrorPath; $script:ErrorCsvCreated=$false; $a.Fault='Transient'; $a.Calls=0; $script:Delays.Clear()
    $info=Get-PrincipalMembershipInfo $a 'https://example.test/sites/recover'
    Assert-True ($info.PrincipalMemberLookupStatus -eq 'OK' -and $a.Calls -eq 3 -and -not $script:ErrorCsvCreated -and ($script:Delays -join ',') -eq '5,15') 'Source transient membership read did not recover.'
    'PASS: consistent scan selection, site-scoped caches, blocking membership failures, token diagnostics and exact-web connection reuse.'
}
finally {
    $resolved=[IO.Path]::GetFullPath($temp)
    if($resolved -notlike ([IO.Path]::GetTempPath().TrimEnd('\')+'\SmartM365-ReviewTest-*')) { throw 'Unexpected cleanup path.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBigFAfxysDM5PG
# /14K3QdKXH8Ov4YZMKNMhq3Tn5mro6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJYdEKtNWax4Wd/Dm/4to11EBGkS2CtVkVaGSoYPtM1TMA0GCSqG
# SIb3DQEBAQUABIIBgHBVbHHavYuprzEp3JolD+rA0AkbdTdUXVQ0Y1zaf1ox7Jit
# qa2ejAQdMhBJP2RGHCkzTA+dVI1DTLaUzO+fkO079FCm3Gr+3j2/DeQH3T+DLroq
# WmMoEeP/Jg2RtJahSYioAO9tqWgJ1ykdzTWWKDkmV31qC/TbiBfblRu3XKoHWTMf
# rXnvW8+wmy5dwR5R13wmcsVHD08xNLMkOvbSmLeENqWGZRhpwyAiMIMeRHAaz4ta
# bHe32eOJz4HJP7HtnVU9uae/PTYtXXcOAl+Ollc1xQchB1wWaFKIzIp7igk+k0YL
# SheDueSPmWPgX2KanQsTPIsJVyRkbY1ZTeAkMD9insuxOlIYmh+m736sT+EueR+J
# izuVx1S5uyRSlnVRPLDehE5dTb0wIscIW6BpQjFxbwEuBUdRqejg5bAt9BSDggXw
# cjDAuCar/KNO2mkuogNLZTXGsAehCT3+yVq2rDbj26cfrUWQ8lZdTocbmzAfrcJc
# Duc3xaFGwmVxqN9K5KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDgxNDU4
# MjVaMC8GCSqGSIb3DQEJBDEiBCACGpZODp37bEpw1+qTa6cJy1wb9f85pOiTp9re
# 4xBA9jANBgkqhkiG9w0BAQEFAASCAgB3Gk3vucdOyGd9+O/UxZD9bp5bGgY8/wIn
# +RS/AWsKyqd0TBr4j5QiD1YBZ6pxVtj80FfAcheEXcvMBiILzj3YDLQxtD/hXDa7
# IIyhDW3FegKS5XDSmsESJ6TtvesJxyYhL4OV2RZKJYlMWuoJq9uO6X4j+/872PBM
# XLfr+kYprJDrd8fTzl5IHehViOYRYhQEb+EnduuiDVQmTSy59gr9DXLz5QIgYpfj
# OwKLcK8GGvUTEdHkvYg87M762PcewAXtxHQ5pP6PR/wurLqhEE6THUGgB2dS/HfZ
# QPoDt2iBNk6AMWPanZM/LQd29eLwEGtZ+dKBrP0QFdtchijHtC4FV11jL9xMKTNt
# SFXWWXrCb5AAz1d7vNOq1+PBbmbfkF1jogtNj0XpdQn3Ixh6OdUT9jMFw711R1ih
# sZ87tWbG3Y4tSVrz/F6DQsmYTcIvl5ahXwLf3tR9gFNY3PHMUCaXvrVZVFOY+zdQ
# yXjTakfBy4foD7ZMLbgWuMBgUM1IugW1bfrHIscKru+8TXWixZ8ci5WlfuKFyBOo
# WSQ0OCray50D0LBzAFQsahSxlVK1DfvwoS/JGc5uE9cOdgG+y4AnO9FB4eZtaPzY
# WynlYpkrgKN3EXyRKGl/39IxefZT6FOn6ZS5uSxPG7kbV1FJzQCziamRvewK6+pR
# egYKDsM3Jg==
# SIG # End signature block
