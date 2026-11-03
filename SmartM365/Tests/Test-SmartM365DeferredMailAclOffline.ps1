#Requires -Version 5.1
<#
.SYNOPSIS
Checks queue ACL admission with synthetic ACLs; no filesystem permission changes.
.VERSION
1.0.0
#>
$ErrorActionPreference='Stop'
$path=Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365-DeferredMail.ps1'
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
if($errors.Count){throw ($errors.Message -join ';')}
$definition=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Test-SmartM365DeferredMailDirectory'},$true)
$runtime=New-Module -ScriptBlock ([scriptblock]::Create($definition.Extent.Text))
try {
 & $runtime {
  $script:parent=[pscustomobject]@{FullName='parent';Attributes=[IO.FileAttributes]::Directory;Parent=$null}
  $script:directory=[pscustomobject]@{FullName='queue';PSIsContainer=$true;Attributes=[IO.FileAttributes]::Directory;Parent=$script:parent}
  $script:acl=New-Object Security.AccessControl.DirectorySecurity
  $script:acl.SetOwner((New-Object Security.Principal.SecurityIdentifier('S-1-5-32-544')))
  $script:acl.SetAccessRuleProtection($true,$false)
  $script:acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule((New-Object Security.Principal.SecurityIdentifier('S-1-5-18')),'FullControl','Allow')))
  $script:children=@()
  function script:Get-Item {param($LiteralPath,[switch]$Force,$ErrorAction) $script:directory}
  function script:Get-Acl {param($LiteralPath,$ErrorAction) if($LiteralPath -eq 'parent'){throw 'Parent ACL must not be an admission requirement'};if($LiteralPath -eq 'queue'){$script:acl}else{$script:childAcl}}
  function script:Get-ChildItem {param($LiteralPath,[switch]$Force,$ErrorAction) $script:children}
  if(-not (Test-SmartM365DeferredMailDirectory queue @())){throw 'Protected queue refused'}
  $script:acl.SetAccessRuleProtection($false,$true)
  if(Test-SmartM365DeferredMailDirectory queue @()){throw 'Inherited queue accepted'}
  $script:acl.SetAccessRuleProtection($true,$false)
  $rule=New-Object Security.AccessControl.FileSystemAccessRule((New-Object Security.Principal.SecurityIdentifier('S-1-1-0')),'Read','Allow')
  $script:acl.AddAccessRule($rule)
  if(Test-SmartM365DeferredMailDirectory queue @()){throw 'Unauthorized queue reader accepted'}
  $script:acl.RemoveAccessRuleSpecific($rule)
  $script:parent.Attributes=[IO.FileAttributes]::ReparsePoint
  if(Test-SmartM365DeferredMailDirectory queue @()){throw 'Linked ancestor accepted'}
  $script:parent.Attributes=[IO.FileAttributes]::Directory
  $script:childAcl=New-Object Security.AccessControl.FileSecurity
  $script:childAcl.SetOwner((New-Object Security.Principal.SecurityIdentifier('S-1-5-32-544')))
  $script:childAcl.AddAccessRule($rule)
  $script:children=@([pscustomobject]@{FullName='payload';PSIsContainer=$false;Attributes=[IO.FileAttributes]::Normal})
  if(Test-SmartM365DeferredMailDirectory queue @()){throw 'Unauthorized payload reader accepted'}
 }
 Write-Host 'PASS: 5 queue ACL checks; no permission changes.'
} finally {Remove-Module $runtime -Force}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA6g77DNYPtHCi0
# 2dmBNSZN1qaFOM7yTxN8KFEvhacBvqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDULCDvppeb+lAU/bc++ySR
# f+6x4OsTwONhxb7YMtGEozANBgkqhkiG9w0BAQEFAASCAYA55DKEuuSRuQnAEnQp
# ShYUqjqTQR1LtubCsDRn+XhR77W5rDr4kaaM/uPAv0a3f85sgvNUPd/nOoPObbkY
# Gq4k0br89tCXisRh0JCowZxDyajRzNyhEjKK+Goo9OX8uoR+D0EMFsfuJCFO6IIY
# mcAZhcIYDuP1e2fiTF0pJm+YnugdSMf4H1R9nXWK2IgRPVCtbscwnx0HMxQE1wyk
# d3xYw7WMEFBTKeW41ryTDQxS/BJBiNNeA+cKAIjYOtGrUlD10oIXmdAfqQjdYHVF
# czX4fF1Eu6AoIlB0cE5g4uxPaC+tVRriAUEFlhI2wNhRxBRiYgwpHzI/Vj9Mj2ji
# c/C2D6pYEc1kmQbO8QThjCvHXbJGx9rk3CXTiHbUZdK6g9fpqOyrQ+nIKQvPxTaH
# DoCb5WicjCDgivIDAuPy42ykeTZbSE6rL+ii0/1/+KLDG5ma+jNZSb7z/HuCjit6
# n4yBmydCvwLK0WqDwySb5UyCjgQ8fAy2asFv1dL6EL65Js0=
# SIG # End signature block
