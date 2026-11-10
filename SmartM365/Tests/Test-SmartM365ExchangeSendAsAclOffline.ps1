#Requires -Version 5.1
<# .SYNOPSIS Offline tests for the local mailbox SendAs DACL reader and its Get-ADPermission fallback. .VERSION 1.0.0 #>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$smartRoot = Split-Path $PSScriptRoot -Parent
$checks = 0
function Check([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message }; $script:checks++ }
function Wire($Value) { [Management.Automation.PSSerializer]::Deserialize([Management.Automation.PSSerializer]::Serialize($Value, 10)) }

$collector = Join-Path $smartRoot 'SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($collector, [ref]$tokens, [ref]$errors)
Check ($errors.Count -eq 0) 'Collector does not parse'
foreach ($name in @('Get-SmartM365SendAsFromSecurityDescriptor', 'Initialize-SmartM365SendAsAclContext', 'Resolve-SmartM365SendAsSid', 'Get-SmartM365MailboxSendAsUsers')) {
    $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    Check ($null -ne $node) ('SendAs dependency missing: ' + $name)
    . ([scriptblock]::Create($node.Extent.Text))
}
$blockNode = $ast.Find({ param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Extent.Text -match '^If \(\$IncludeADPermission\)' }, $true)
Check ($null -ne $blockNode) 'IncludeADPermission block missing'
Check ($blockNode.Extent.Text -match 'Get-SmartM365MailboxSendAsUsers') 'IncludeADPermission block does not use the DACL reader'

$sendAs = 'ab721a54-1e2f-11d0-9819-00aa0040529b'
$receiveAs = 'ab721a56-1e2f-11d0-9819-00aa0040529b'
$sendTo = 'ab721a55-1e2f-11d0-9819-00aa0040529b'
$names = @{
    'S-1-5-10' = 'NT AUTHORITY\SELF'
    'S-1-5-21-1-2-3-1001' = 'CONTOSO\alice'
    'S-1-5-21-1-2-3-1002' = 'CONTOSO\bob'
    'S-1-5-21-1-2-3-1003' = 'CONTOSO\carol'
    'S-1-5-21-1-2-3-1004' = 'CONTOSO\dave'
    'S-1-5-21-1-2-3-1005' = 'CONTOSO\erin'
}
$resolver = { param($sid) if ($names.ContainsKey($sid)) { $names[$sid] } else { $sid } }.GetNewClosure()
function New-Sd([string]$Sddl) { $sd = New-Object System.DirectoryServices.ActiveDirectorySecurity; $sd.SetSecurityDescriptorSddlForm($Sddl); return $sd }

# 1. Pure DACL filter.
$sddl = 'D:' +
    "(OA;;CR;$sendAs;;S-1-5-10)" +
    "(OA;;CR;$sendAs;;S-1-5-21-1-2-3-1001)" +
    "(OA;;CR;$sendAs;;S-1-5-21-1-2-3-1001)" +
    "(OA;ID;CR;$sendAs;;S-1-5-21-1-2-3-1002)" +
    "(OD;;CR;$sendAs;;S-1-5-21-1-2-3-1003)" +
    "(OA;;CR;$receiveAs;;S-1-5-21-1-2-3-1004)" +
    "(OA;;RP;$sendAs;;S-1-5-21-1-2-3-1005)" +
    "(OA;;CR;$sendAs;;S-1-5-21-9-9-9-9999)" +
    "(OA;;CR;$sendTo;;S-1-5-21-1-2-3-1005)" +
    '(A;;CR;;;S-1-5-21-1-2-3-1004)' +
    '(A;;GA;;;BA)'
$rights = @{ $sendAs = 'Send-As'; $sendTo = 'Send-To' }
$r = Get-SmartM365SendAsFromSecurityDescriptor -SecurityDescriptor (New-Sd $sddl) -SendRightGuids $rights -ResolveSid $resolver
Check (($r.Users -join ';') -ceq 'CONTOSO\alice;CONTOSO\carol;CONTOSO\erin') ("Filter users: " + ($r.Users -join ';'))
Check ($r.MatchedAceCount -eq 6 -and $r.DenyMatchedAceCount -eq 1 -and $r.AllExtendedRightsAceCount -eq 1) 'Filter counters'
Check ($r.TotalAceCount -eq 11 -and $r.ExplicitAceCount -eq 10) 'ACE counters'
$e = Get-SmartM365SendAsFromSecurityDescriptor -SecurityDescriptor (New-Sd 'D:(A;;GA;;;BA)') -SendRightGuids $rights -ResolveSid $resolver
Check (@($e.Users).Count -eq 0) 'Empty DACL returned users'

# 2. AD reader with mocked ActiveDirectory cmdlets.
$script:adCalls = New-Object 'System.Collections.Generic.List[string]'
$script:failAd = $false
$script:objectSddl = $sddl
function WriteLog { param([string]$Message, [string]$Level = 'INFO') $script:logLines += ,("[$Level] $Message") }
function Get-ADRootDSE { [CmdletBinding()] param() [pscustomobject]@{ configurationNamingContext = 'CN=Configuration,DC=example,DC=invalid' } }
function Get-ADObject {
    [CmdletBinding()]
    param($Identity, $Server, $SearchBase, $LDAPFilter, $Properties)
    if ($SearchBase) {
        $script:adCalls.Add('rights')
        return @(
            [pscustomobject]@{ Name = 'Send-As'; rightsGuid = $sendAs },
            [pscustomobject]@{ Name = 'Send-To'; rightsGuid = $sendTo },
            [pscustomobject]@{ Name = 'Receive-As'; rightsGuid = $receiveAs }
        )
    }
    $script:adCalls.Add("object|$Server|$Identity")
    if ($script:failAd) { throw 'Synthetic DC unavailable' }
    [pscustomobject]@{ DistinguishedName = $Identity; nTSecurityDescriptor = (New-Sd $script:objectSddl) }
}
$script:logLines = @()
$script:SendAsRightGuids = $null
Initialize-SmartM365SendAsAclContext
Check (@($script:logLines | Where-Object { $_ -match 'SendAs ACL reader ready: 2 extended rights' }).Count -eq 1) 'Reader readiness not logged'
# Synthetic SIDs cannot be translated by LSA: pre-seed the reader cache.
foreach ($k in $names.Keys) { $script:SendAsSidCache[$k] = $names[$k] }
$script:SendAsSidCache['S-1-5-21-9-9-9-9999'] = 'S-1-5-21-9-9-9-9999'
$dn ='CN=Shared One,OU=Users,DC=fr,DC=example,DC=invalid'
$users = @(Get-SmartM365MailboxSendAsUsers -DistinguishedName $dn)
Check (($users -join ';') -ceq 'CONTOSO\alice;CONTOSO\carol;CONTOSO\erin') 'Reader users'
Check ($script:adCalls[0] -eq 'rights' -and $script:adCalls[1] -eq "object|fr.example.invalid|$dn") ('AD calls: ' + ($script:adCalls -join ' / '))
Check ($script:SendAsRightGuids.Count -eq 2) 'Rights map should keep Send-As and Send-To only'
$null = Get-SmartM365MailboxSendAsUsers -DistinguishedName $dn
Check (@($script:adCalls | Where-Object { $_ -eq 'rights' }).Count -eq 1) 'Rights map loaded more than once'
# A cached SID that the reader resolved must be reused.
$script:SendAsSidCache = @{ 'S-1-5-21-1-2-3-1001' = 'CACHED\alice' }
$script:objectSddl = "D:(OA;;CR;$sendAs;;S-1-5-21-1-2-3-1001)"
Check ((@(Get-SmartM365MailboxSendAsUsers -DistinguishedName $dn) -join ';') -eq 'CACHED\alice') 'SID cache ignored'
$threw = $false
try { $null = Get-SmartM365MailboxSendAsUsers -DistinguishedName 'CN=NoDomain' } catch { $threw = $true }
Check $threw 'DN without domain component must fail'

# 3. Actual IncludeADPermission block: DACL path, fallback path, error path.
function Write-LogMailboxesProcessing { param([string]$Message) $script:logLines += ,$Message }
function Write-Host { param([Parameter(ValueFromRemainingArguments = $true)]$Object) }
$script:adPermissionMode = 'Rows'
function Get-ADPermission {
    [CmdletBinding()]
    param($Identity)
    if ($script:adPermissionMode -eq 'Throw') { throw 'Synthetic Get-ADPermission failure' }
    Wire @(
        [pscustomobject]@{ User = 'NT AUTHORITY\SELF'; ExtendedRights = @('Send-As'); IsInherited = $false },
        [pscustomobject]@{ User = 'CONTOSO\alice'; ExtendedRights = @('Send-As'); IsInherited = $false },
        [pscustomobject]@{ User = 'CONTOSO\alice'; ExtendedRights = @('Send-As'); IsInherited = $false },
        [pscustomobject]@{ User = 'CONTOSO\bob'; ExtendedRights = @('Send-As'); IsInherited = $true },
        [pscustomobject]@{ User = 'S-1-5-21-9-9-9-9999'; ExtendedRights = @('Send-As'); IsInherited = $false },
        [pscustomobject]@{ User = 'CONTOSO\dave'; ExtendedRights = @('Receive-As'); IsInherited = $false },
        [pscustomobject]@{ User = 'CONTOSO\carol'; ExtendedRights = @('Send-As'); IsInherited = $false }
    )
}
function Invoke-Block {
    $script:userObj = [pscustomobject]@{ Name = 'Shared One' }
    $userObj = $script:userObj
    $Mbx = Wire ([pscustomobject]@{ DistinguishedName = $dn; DisplayName = 'Shared One' })
    $IncludeADPermission = $true
    . ([scriptblock]::Create($blockNode.Extent.Text))
    return $userObj
}
$script:SendAsSidCache = @{}
foreach ($k in $names.Keys) { $script:SendAsSidCache[$k] = $names[$k] }
$script:SendAsSidCache['S-1-5-21-9-9-9-9999'] = 'S-1-5-21-9-9-9-9999'
$script:SendAsAclReadCount = 0; $script:SendAsAclFallbackCount = 0
$script:failAd = $false

$script:objectSddl = $sddl
$o = Invoke-Block
Check ($o.SendAs -ceq 'CONTOSO\alice;CONTOSO\carol;CONTOSO\erin' -and $o.SendAsCount -eq 3) ("DACL path: $($o.SendAs) / $($o.SendAsCount)")
$script:objectSddl = "D:(OA;;CR;$sendAs;;S-1-5-21-1-2-3-1001)(A;;GA;;;BA)"
$o = Invoke-Block
Check ($o.SendAs -ceq 'CONTOSO\alice' -and $o.SendAsCount -eq 1) 'Single SendAs count'
$script:objectSddl = 'D:(A;;GA;;;BA)'
$o = Invoke-Block
Check ($o.SendAs -ceq '' -and $o.SendAsCount -eq 0) 'Empty SendAs'
Check ($script:SendAsAclReadCount -eq 3 -and $script:SendAsAclFallbackCount -eq 0) 'ACL read counters'

$script:failAd = $true
$script:logLines = @()
$o = Invoke-Block
Check ($o.SendAs -ceq 'CONTOSO\alice;CONTOSO\carol' -and $o.SendAsCount -eq 2) ("Fallback path: $($o.SendAs) / $($o.SendAsCount)")
Check ($script:SendAsAclFallbackCount -eq 1) 'Fallback counter'
Check (@($script:logLines | Where-Object { $_ -match 'Falling back to Get-ADPermission' }).Count -ge 1) 'Fallback not logged'

$script:adPermissionMode = 'Throw'
$o = Invoke-Block
Check ($o.SendAs -eq 'Error retrieving SendAs' -and $o.SendAsCount -eq 'Error') 'Double failure must publish Error'

Write-Output ('PASS: {0} SendAs ACL checks; {1}.' -f $checks, $PSVersionTable.PSVersion)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCb/3cALE5iaJKP
# dViUQ/DG80Jth/e2E83qemmRypSOF6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILEvX7f1TxH3XtD7Zp7Hpp8iUfVWDPQjOklt7dobjyLAMA0GCSqG
# SIb3DQEBAQUABIIBgAJLdcn+ZC94oCgO+ykfznRulgNzeVk0NFHrEj4fLOVsapRY
# QoaY/mLlu8RPDnqVtJc/4BOpxebdxiptvFfzf1+Pievldt+sjtBdlWbVQbUkNjQH
# CXtwal8aK0eK4SIalXUwFkImItvjMZNq1OiHsHDC2dTKrA7kBo4K96DTjxAEpV7R
# WDrVbm3w5fGkzMmClkSJ4BGKP1ogxl+00gLxr9lKf8wWiZu/M35rfUyL8t6WSxnS
# 2baQAaQJk2JJBTEYcZjM8WPfZ7taFjBgMZRkt2Z4oNqfE+jTfnGLZzYA+zOePYHJ
# EksLebebqFkcC+m1DfesjAd5Sl2rwaKvoZUb78iqEcB/LYl/PtQ/sFhjjRcUWGEb
# y/GI/TJfjsYbLJ+/eDWC43p1279au+yZ0paFt9aKOB8YqRRBxuETihnleH5Wi/4h
# QmvYyxqiohYxpeXab7UI4fs21BRkuwGtbl7Pr/APXltpdr0uxqYf/1e/e5icXjN+
# fpkpLiHZKoQ9fc2CX6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTAxODI2
# NTRaMC8GCSqGSIb3DQEJBDEiBCCAB/Z8JS+Tq+1V/5tZWqTXuA+7lJmFGGJeGavS
# aqXWqDANBgkqhkiG9w0BAQEFAASCAgA0FJWwM55FlyusdJCMLdxqfURvQ9gdz5oc
# CZ5WsR+oUgInhc4xhHvqZmSQCROdFMFO58qSfbD8gDpgDEkDtl4BSrpJ571vYWH3
# EQp6oFdnXx9pVep33yupsy2EqB6vWOTROb8Pt1y5isMFEGO/xMbrTH1jkOp0+5cv
# iKx7J8SCN9pFhPvFGNU4ps6nqJpw55r9v/QYFAggML1+2DQe45VbgpiVzLEDRx/s
# liP+xHH8tAUVWN+u4HRG5mKcox+8QDE6rfsHtbL955IlJpsq14CejbLX736i7aLe
# tIFQExhf4yWOyeMzwI3Az64qycPyYWf8dIeWjJbrBFlzcvX8YdjgEH+HxlFo5wfT
# Ie7r+QjsXXWGkUM3dRfTKRe880YFsWVrRIIjj5J+syEHWpevBZmMtbv1EO9bcTCr
# WmgP/nY7rIRbLRqfr+9eMS8z+LWZT4NtQvhFpU5I90Adx6P7NmemgGR1Sb3UYyO7
# ++DBsec7PRh//7HVVAoiqld1Ubcc5X33AN4S3T3YJ/zEeebdckJjyeB+Smq0gFqs
# ZBmlKgzbxCaiXswMcMaqajfWqBzXGAwD8QQql4ELeklORQqCK1ZjyXntfTuto4sK
# ruEpkwC4asJgM+TzIMY/NXJzSI1OunxEKqIGwjOUNr3yimk6AWhfX4i1//WRiTLK
# bJqqqgzzxg==
# SIG # End signature block
