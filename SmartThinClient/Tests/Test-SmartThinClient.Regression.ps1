param([string]$SourcePath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'SmartThinClient-Shell.ps1'))
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($SourcePath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
# Load function declarations only. Never execute the application entry point.
foreach ($fn in $ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.FunctionDefinitionAst] }) {
    . ([scriptblock]::Create($fn.Extent.Text))
}
$script:ScriptRoot = Split-Path $SourcePath -Parent
$script:LogPath = $null; $script:CliMode = $false
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('ThinClient-tests-' + [guid]::NewGuid())
New-Item -ItemType Directory $testRoot | Out-Null
$script:OutputRoot = $testRoot
$script:AppName = 'Smart ThinClient Shell'; $script:RunId = 'synthetic'
$script:RollbackRoot = $testRoot
$script:EvidenceJsonPath = Join-Path $testRoot 'evidence.json'; $script:EvidenceCsvPath = Join-Path $testRoot 'evidence.csv'
$Profile = 'Citrix'; $TargetUserMode = 'Auto'; $TargetUserName = ''; $DedicatedUserPassword = $null
$script:passed = 0; $script:failed = 0
function Check($Name, [scriptblock]$Body) {
    try { & $Body; $script:passed++; Write-Host "PASS $Name" }
    catch { $script:failed++; Write-Host "FAIL $Name : $_" }
}
function Assert($Condition, $Message) { if (-not $Condition) { throw $Message } }
function Reject([scriptblock]$Body) { $caught = $false; try { & $Body } catch { $caught = $true }; Assert $caught 'Expected rejection' }
# Every system-write primitive used below is replaced; unanticipated mutations fail.
function Set-StringRegistryValue { throw 'Unexpected registry write' }
function Set-DwordRegistryValue { throw 'Unexpected registry write' }
function Restore-RegistryValue { throw 'Unexpected registry restore' }
function New-LocalUser { throw 'Unexpected user creation' }
function Set-AssignedAccess { throw 'Unexpected kiosk write' }
function Clear-AssignedAccess { throw 'Synthetic restore failure' }
function Invoke-CimMethod { throw 'Unexpected CIM call' }
function Export-Evidence { param($Evidence) }
function Test-IsAdministrator { return $true }
function Get-RegistryValueSnapshot { param($Path,$Name) [pscustomobject]@{Path=$Path;Name=$Name;ValueExists=$false;ValueKind='';Value=$null} }
$audit = [pscustomobject]@{EffectiveProfile='Citrix';CitrixWorkspacePath=$PSHOME;AvdClassicClientPath='';BrowserPath='';TargetUserName='TEST-USER';TargetUserSid='S-1-5-21-1-2-3-1001'}
Check 'Native configuration survives profile templates' {
    $merged = Merge-ProfileConfig -Config ([ordered]@{PreferredAccessMode='Native';UseWebShell=$false}) -EffectiveProfile Citrix
    Assert ($merged.PreferredAccessMode -eq 'Native' -and -not $merged.UseWebShell) 'Profile overrode explicit global configuration'
}
Check 'None restrictions write nothing' { Apply-ShellLimitations -Config ([ordered]@{EnableShellLimitations=$true;ShellRestrictionLevel='None'}) }
Check 'AutoLaunchMode None writes nothing' { Apply-AutoLaunch -Config ([ordered]@{AutoLaunchMode='None'}) -CommandLine 'synthetic' }
Check 'Rollback excludes untouched Winlogon values' {
    $state = New-RollbackState -Config ([ordered]@{AutoLaunchMode='RunKey'}) -Audit $audit -EffectiveProfile Citrix
    Assert (@($state.Registry | Where-Object {$_.Path -match 'Winlogon'}).Count -eq 0) 'Rollback includes unrelated system values'
}
Check 'Launch uses isolated file' {
    function Invoke-ThinClientAudit { param($Config,$RequestedProfile,$RequestedAction) return $audit }
    function Start-Process { param($FilePath,$ArgumentList,$WindowStyle) }
    $installed = New-LauncherScript -Config ([ordered]@{}) -EffectiveProfile Citrix -Audit $audit
    $before = (Get-FileHash $installed).Hash
    $result = Invoke-ThinClientLaunchOnly -Config ([ordered]@{WebShellTitle='Different launch-only title'})
    Assert ($result.LauncherPath -ne $installed) 'Launch reused installed file'
    Assert ((Get-FileHash $installed).Hash -eq $before) 'Launch changed installed file'
}
Check 'Native process omits empty argument list and reports errors' {
    $launcher = New-LauncherScript -Config ([ordered]@{}) -EffectiveProfile Citrix -Audit $audit
    $a = [Management.Automation.Language.Parser]::ParseFile($launcher,[ref]$tokens,[ref]$errors)
    Assert ($errors.Count -eq 0) 'Generated script parse error'
    $f = $a.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Start-WorkspaceProcess'},$true)
    . ([scriptblock]::Create($f.Extent.Text))
    function Start-Process { param($FilePath,[ValidateNotNullOrEmpty()][string[]]$ArgumentList) if ($PSBoundParameters.ContainsKey('ArgumentList')) { throw 'Empty args passed' } }
    Assert (Start-WorkspaceProcess -FilePath $PSHOME -Arguments '') 'No native success result'
}
Check 'Invalid access mode rejected before launcher generation' {
    Reject { New-LauncherScript -Config ([ordered]@{PreferredAccessMode="Web'; throw 'injected"}) -EffectiveProfile Citrix -Audit $audit }
}
Check 'All four generated profiles parse with quoted configuration' {
    foreach ($provider in @('Citrix','AVD','WebOnly','Hybrid')) {
        $launcher = New-LauncherScript -Config ([ordered]@{WebShellTitle="Example's title";CitrixWebUrl='https://example.invalid/?a=1&b=2'}) -EffectiveProfile $provider -Audit $audit
        [void][Management.Automation.Language.Parser]::ParseFile($launcher,[ref]$tokens,[ref]$errors)
        Assert ($errors.Count -eq 0) "Parse error in $provider"
    }
}
Check 'JSON string false cannot enable Apply' { Reject { Assert-ConfigValues -Config ([ordered]@{AllowApply='false'}) } }
Check 'CIM error code is not success' { Reject { Assert-CimSuccess -Response ([pscustomobject]@{ReturnValue=5}) } }
Check 'CIM success accepted' { Assert-CimSuccess -Response ([pscustomobject]@{ReturnValue=0}) }
Check 'Rollback captures existing launcher bytes' {
    $path = Join-Path $testRoot 'Launcher\SmartThinClient-LaunchWorkspace.ps1'
    [IO.File]::WriteAllBytes($path,[byte[]]@(1,2,3,4))
    $state = New-RollbackState -Config ([ordered]@{AutoLaunchMode='None'}) -Audit $audit -EffectiveProfile Citrix
    Assert ($state.LauncherExisted -and $state.LauncherContent -eq 'AQIDBA==') 'Previous launcher missing'
    Assert (@($state.Registry).Count -eq 0) 'No-write apply must not snapshot unrelated registry'
}
Check 'Foreign rollback rejected before any write' {
    $state = New-RollbackState -Config ([ordered]@{}) -Audit $audit -EffectiveProfile Citrix
    $state.ComputerName = 'OTHER-SYNTHETIC-COMPUTER'
    Reject { Assert-RollbackState -State ([pscustomobject]$state) }
}
Check 'Legacy rollback rejected before any write' { Reject { Assert-RollbackState -State ([pscustomobject]@{AppName=$script:AppName}) } }
Check 'Malformed late snapshot rejected before any restore' {
    $state = New-RollbackState -Config ([ordered]@{}) -Audit $audit -EffectiveProfile Citrix
    $state.Registry += [pscustomobject]@{Path='HKLM:\Unrelated';Name='Bad'}
    $RollbackPath = Join-Path $testRoot 'malformed.json'
    [pscustomobject]$state | ConvertTo-Json -Depth 12 | Set-Content $RollbackPath
    $ConfirmRestore = 'RESTORE WINDOWS SHELL'
    Reject { Invoke-ThinClientRestore -Config ([ordered]@{AllowRestore=$true}) }
}
Check 'Restore preserves original launcher bytes' {
    $state = New-RollbackState -Config ([ordered]@{AutoLaunchMode='None'}) -Audit $audit -EffectiveProfile Citrix
    $RollbackPath = Join-Path $testRoot 'valid.json'
    [pscustomobject]$state | ConvertTo-Json -Depth 12 | Set-Content $RollbackPath
    [IO.File]::WriteAllText($state.LauncherPath,'Changed')
    $ConfirmRestore = 'RESTORE WINDOWS SHELL'
    $result = Invoke-ThinClientRestore -Config ([ordered]@{AllowRestore=$true})
    Assert ($result.Status -eq 'Restored') 'Missing restored result'
    Assert ([Convert]::ToBase64String([IO.File]::ReadAllBytes($state.LauncherPath)) -eq $state.LauncherContent) 'Original bytes not restored'
}
Check 'Assigned Access restore failure propagates' {
    $state = New-RollbackState -Config ([ordered]@{AutoLaunchMode='None';EnableAssignedAccess=$true}) -Audit $audit -EffectiveProfile Citrix
    $RollbackPath = Join-Path $testRoot 'kiosk-failure.json'
    [pscustomobject]$state | ConvertTo-Json -Depth 12 | Set-Content $RollbackPath
    $ConfirmRestore = 'RESTORE WINDOWS SHELL'
    Reject { Invoke-ThinClientRestore -Config ([ordered]@{AllowRestore=$true}) }
}
Check 'Missing confirmation rejects apply' { $ConfirmApply=''; Reject { Assert-ApplyAllowed -Config ([ordered]@{AllowApply=$true}) } }
Check 'Dedicated user missing password rejected before mutation' {
    $target = [pscustomobject]@{Mode='DedicatedUser';LocalUserExists=$false}
    Reject { Assert-ApplyPreflight -Config ([ordered]@{CreateDedicatedLocalUser=$true}) -Audit $audit -TargetUser $target }
}
Check 'Conflicting kiosk modes rejected' {
    Reject { Assert-ApplyPreflight -Config ([ordered]@{EnableAssignedAccess=$true;EnableShellLauncher=$true}) -Audit $audit -TargetUser $null }
}
Check 'Existing Assigned Access is refused before writes' {
    function Get-LocalGroupMember { param($SID) }
    function Get-AssignedAccess { [pscustomobject]@{UserName='OTHER-USER'} }
    $target=[pscustomobject]@{Mode='ExistingUser';LocalUserExists=$true;Enabled=$true;Sid=$audit.TargetUserSid}
    $kioskAudit=[pscustomobject]@{AssignedAccessCmdletAvailable=$true}
    Reject { Assert-ApplyPreflight -Config ([ordered]@{EnableAssignedAccess=$true;AssignedAccessAppUserModelId='SYNTHETIC!App'}) -Audit $kioskAudit -TargetUser $target }
}
Check 'Empty Assigned Access preflight passes without writes' {
    function Get-LocalGroupMember { param($SID) }
    function Get-AssignedAccess { }
    $target=[pscustomobject]@{Mode='ExistingUser';LocalUserExists=$true;Enabled=$true;Sid=$audit.TargetUserSid}
    Assert-ApplyPreflight -Config ([ordered]@{EnableAssignedAccess=$true;AssignedAccessAppUserModelId='SYNTHETIC!App'}) -Audit ([pscustomobject]@{AssignedAccessCmdletAvailable=$true}) -TargetUser $target
}
Check 'Disabled Shell Launcher is refused' {
    function Get-LocalGroupMember { param($SID) }
    function Invoke-CimMethod { param($Namespace,$ClassName,$MethodName) [pscustomobject]@{ReturnValue=0;Enabled=$false} }
    $target=[pscustomobject]@{Mode='ExistingUser';LocalUserExists=$true;Enabled=$true;Sid=$audit.TargetUserSid}
    Reject { Assert-ApplyPreflight -Config ([ordered]@{EnableShellLauncher=$true}) -Audit ([pscustomobject]@{ShellLauncherClassAvailable=$true}) -TargetUser $target }
}
Check 'Existing custom shell is refused' {
    function Get-LocalGroupMember { param($SID) }
    function Invoke-CimMethod { param($Namespace,$ClassName,$MethodName) [pscustomobject]@{ReturnValue=0;Enabled=$true} }
    function Get-CimInstance { param($Namespace,$ClassName) [pscustomobject]@{Sid=$audit.TargetUserSid} }
    $target=[pscustomobject]@{Mode='ExistingUser';LocalUserExists=$true;Enabled=$true;Sid=$audit.TargetUserSid}
    Reject { Assert-ApplyPreflight -Config ([ordered]@{EnableShellLauncher=$true}) -Audit ([pscustomobject]@{ShellLauncherClassAvailable=$true}) -TargetUser $target }
}
Check 'Apply saves rollback before account mutation and respects None' {
    $script:sequence = @()
    function Invoke-ThinClientAudit { param($Config,$RequestedProfile,$RequestedAction) $audit }
    function Get-TargetUserState { param($Config) [pscustomobject]@{Mode='ExistingUser';LocalUserExists=$true} }
    function Save-RollbackState { param($State) $script:sequence += 'save'; 'synthetic-rollback.json' }
    function New-DedicatedUserIfNeeded { param($Config,$TargetUser) $script:sequence += 'account'; $false }
    $ConfirmApply='APPLY SMARTTHINCLIENT'
    $result=Invoke-ThinClientApply -Config ([ordered]@{AllowApply=$true;AutoLaunchMode='None'})
    Assert ($result.Status -eq 'Applied') 'Missing Applied result'
    Assert (($script:sequence -join ',') -eq 'save,account,save') 'Account attempted before rollback'
}
Write-Host "RESULT passed=$script:passed failed=$script:failed"
if ($script:failed) { exit 1 }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAJ+baajRLjAzuD
# ktVlljx17I702jRGCtx5Lwqj1IPB+KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIL9iyONJNd+YT5ivfbR02RSkXx3b4SochwYfqQqPhI+tMA0GCSqG
# SIb3DQEBAQUABIIBgH1S+H+aTWxoF4hoDXzZlaBhqdAUTM6tYy04ERNb7IT6e6o2
# luPDtGVHsKpNUQ5jGftbbqVbpWI03Iuo469QKCtBOYLEEAqOJuU6HK/VcMAXNHCE
# 3iSfShpaDWh5OH/grrXVBcGxQAn5hX6pVQ0Sfg5TosxMp0Kd1bJpV/3VVoUMiTya
# OyvFLnJNDqaHMKqo6WB4npmKlIER3QeF08oJpKrZEwYsHcJ0jvq1vhRbCZjT6w7o
# VDAZCZs1xKdAg7HMOeXlrNJrE4nPWaFRuPi7M7LXnl7WebKBWJDLgQLg1LVP3+we
# E3ml4Xp3RvFiEJXwz3Oh1V9WMwRoAaXxr/+cwSlRjOe5T3BS7sRD+RnT887eCOYj
# ftvTKhG6WW96lZ/II9tZpCiJ3ZRHYk8BfnG91TvOeWgkJtIwwS8Ct/fLglnsjr+z
# MyksZ7oRoducyumg/g3nVCb1hEnlDsaY4rISRCm28o+xV0z7IepSKRrnPL2IuR6I
# qq1mfeIp6qCsD72fl6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MDkwNjM5
# MjFaMC8GCSqGSIb3DQEJBDEiBCA2xIT3jWNVjYKkpQ7q1wGng4KC0nBEb8xykicc
# r6gREDANBgkqhkiG9w0BAQEFAASCAgCDCksy5xEfHhqDSonpfJY5NQgrh/hh2m5b
# GRDRZy/xY9L0/Zoue0rZhrdXOA1+iUeHk8bjyQQvQuNSdp1nzNW0lvjeXtMNnC5X
# S8em3f79Nc6btyXrof16uf9zjPfgMBsonYGfZmER5xmv4bQngh554tjUKBJwrHjb
# JPalYShoU0GKtXYODPBb8wMVMmU/0oFOqr4SV+p4Hd1MydSec0Nd+bI/F8+7FX4Z
# CqyNm9c87OUx/RwOpojhtgGHW5SIjSbeLepOtfjbeR8u1x/PJv869r1s2qUNuX99
# CnLybk64R5yBQKg59CbWKJJSvgP1ZFbSbr8K/Ib4SLFkjvBSodEVFDB1U1yaOT6n
# MWBig8Fqa8gVwe8fjIgznSJ0+VtWBnVchFYjPpJ1IS7ajoBuKYBmoSxOhXcwSsLl
# LFZ18SpRoIz82l4b6L3UGC9hcOjs6kCg160rv/7/B2n/RheelH0W9v5DNOUGllWe
# Wii/cAvsbiAUO+3h1ETIXT4zyAF2x9bBOMu9SzXUhd310CyDYfbsTUEsPtQJ7p72
# iGZjoLaF57aJy/rfxX2bvM0/saFxTuxBtPTEdipUy089V1umu+yKOw1IQ8Puzj8b
# 7SOnfUXu/n/dZpCpix9cURxpeQ+1tCHkvL5mqKMYT2VhngCcy4/1PV8nNd95kVZG
# 1YfKVbh+DQ==
# SIG # End signature block
