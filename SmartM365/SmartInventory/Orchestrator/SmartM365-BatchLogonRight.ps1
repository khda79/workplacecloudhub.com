# Grant only the batch-logon right for one resolved account. Loading this file makes no changes.
function Initialize-SiBatchLogonNative {
    if ('SmartM365Setup.BatchLogonNative' -as [type]) {return}
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Principal;
using Microsoft.Win32.SafeHandles;
namespace SmartM365Setup {
 public static class BatchLogonNative {
  [StructLayout(LayoutKind.Sequential)] struct OBJECT_ATTRIBUTES {public int Length; public IntPtr RootDirectory,ObjectName; public uint Attributes; public IntPtr SecurityDescriptor,SecurityQualityOfService;}
  [StructLayout(LayoutKind.Sequential)] struct UNICODE_STRING {public ushort Length,MaximumLength; public IntPtr Buffer;}
  [DllImport("advapi32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern bool LogonUserW(string user,string domain,IntPtr password,int type,int provider,out SafeAccessTokenHandle token);
  [DllImport("advapi32.dll")] static extern uint LsaOpenPolicy(IntPtr system,ref OBJECT_ATTRIBUTES attributes,uint access,out IntPtr policy);
  [DllImport("advapi32.dll")] static extern uint LsaEnumerateAccountRights(IntPtr policy,IntPtr sid,out IntPtr rights,out uint count);
  [DllImport("advapi32.dll")] static extern uint LsaAddAccountRights(IntPtr policy,IntPtr sid,UNICODE_STRING[] rights,uint count);
  [DllImport("advapi32.dll")] static extern uint LsaFreeMemory(IntPtr memory);
  [DllImport("advapi32.dll")] static extern uint LsaClose(IntPtr policy);
  [DllImport("advapi32.dll")] static extern uint LsaNtStatusToWinError(uint status);
  static void Check(uint status) {if(status!=0) throw new Win32Exception((int)LsaNtStatusToWinError(status));}
  static IntPtr SidPointer(string text) {var sid=new SecurityIdentifier(text);var bytes=new byte[sid.BinaryLength];sid.GetBinaryForm(bytes,0);var p=Marshal.AllocHGlobal(bytes.Length);Marshal.Copy(bytes,0,p,bytes.Length);return p;}
  public static int Probe(string user,string domain,SecureString password,string expectedSid) {
   return ProbeType(user,domain,password,expectedSid,4);
  }
  public static int ProbeInteractive(string user,string domain,SecureString password,string expectedSid) {
   return ProbeType(user,domain,password,expectedSid,2);
  }
  static int ProbeType(string user,string domain,SecureString password,string expectedSid,int logonType) {
   IntPtr secret=Marshal.SecureStringToGlobalAllocUnicode(password);
   try {SafeAccessTokenHandle token;if(!LogonUserW(user,String.IsNullOrEmpty(domain)?null:domain,secret,logonType,0,out token)) return Marshal.GetLastWin32Error();
    using(token) using(var identity=new WindowsIdentity(token.DangerousGetHandle())) {if(identity.User.Value!=expectedSid) throw new InvalidOperationException("Authenticated account SID differs from the reviewed account.");return 0;}
   } finally {Marshal.ZeroFreeGlobalAllocUnicode(secret);}
  }
  public static string[] Read(string sidText) {
   var attrs=new OBJECT_ATTRIBUTES();attrs.Length=Marshal.SizeOf(typeof(OBJECT_ATTRIBUTES));IntPtr policy=IntPtr.Zero,sid=IntPtr.Zero,rights=IntPtr.Zero;
   try {Check(LsaOpenPolicy(IntPtr.Zero,ref attrs,0x800,out policy));sid=SidPointer(sidText);uint count;uint status=LsaEnumerateAccountRights(policy,sid,out rights,out count);
    if(status!=0) {int error=(int)LsaNtStatusToWinError(status);if(error==2) return new string[0];throw new Win32Exception(error);}
    var result=new List<string>();int size=Marshal.SizeOf(typeof(UNICODE_STRING));for(int i=0;i<count;i++){var item=(UNICODE_STRING)Marshal.PtrToStructure(IntPtr.Add(rights,i*size),typeof(UNICODE_STRING));result.Add(Marshal.PtrToStringUni(item.Buffer,item.Length/2));}return result.ToArray();
   } finally {if(rights!=IntPtr.Zero)LsaFreeMemory(rights);if(sid!=IntPtr.Zero)Marshal.FreeHGlobal(sid);if(policy!=IntPtr.Zero)LsaClose(policy);}
  }
  public static void AddBatchRight(string sidText) {
   AddRight(sidText,"SeBatchLogonRight");
  }
  public static void AddInteractiveRight(string sidText) {
   AddRight(sidText,"SeInteractiveLogonRight");
  }
  static void AddRight(string sidText,string right) {
   var attrs=new OBJECT_ATTRIBUTES();attrs.Length=Marshal.SizeOf(typeof(OBJECT_ATTRIBUTES));IntPtr policy=IntPtr.Zero,sid=IntPtr.Zero,name=IntPtr.Zero;
   try {Check(LsaOpenPolicy(IntPtr.Zero,ref attrs,0x810,out policy));sid=SidPointer(sidText);name=Marshal.StringToHGlobalUni(right);
    var item=new UNICODE_STRING{Length=(ushort)(right.Length*2),MaximumLength=(ushort)((right.Length+1)*2),Buffer=name};Check(LsaAddAccountRights(policy,sid,new UNICODE_STRING[]{item},1));
   } finally {if(name!=IntPtr.Zero)Marshal.FreeHGlobal(name);if(sid!=IntPtr.Zero)Marshal.FreeHGlobal(sid);if(policy!=IntPtr.Zero)LsaClose(policy);}
  }
 }
}

'@
}

function Ensure-SiBatchLogonRight {
    param([Parameter(Mandatory)][Management.Automation.PSCredential]$Credential,[Parameter(Mandatory)][string]$ExpectedSid,[scriptblock]$Probe,[scriptblock]$ReadRights,[scriptblock]$AddRight)
    if ($ExpectedSid -notmatch '^S-1-(?:\d+-)+\d+$' -or $ExpectedSid -eq 'S-1-5-18') {throw 'Invalid or forbidden service account SID.'}
    $user=$Credential.UserName;$domain=$null
    if ($user -match '^([^\\]+)\\([^\\]+)$') {$domain=$Matches[1];$user=$Matches[2]}
    elseif ($user -notmatch '^[^@\\\s]+@[^@\\\s]+$') {throw 'A domain account is required.'}
    if (-not $Probe -or -not $ReadRights -or -not $AddRight) {Initialize-SiBatchLogonNative}
    if (-not $Probe) {$Probe={param($u,$d,$p,$sid) [SmartM365Setup.BatchLogonNative]::Probe($u,$d,$p,$sid)}}
    if (-not $ReadRights) {$ReadRights={param($sid) [SmartM365Setup.BatchLogonNative]::Read($sid)}}
    if (-not $AddRight) {$AddRight={param($sid) [SmartM365Setup.BatchLogonNative]::AddBatchRight($sid)}}
    $code=[int](& $Probe $user $domain $Credential.Password $ExpectedSid)
    if ($code -eq 0) {return [pscustomobject]@{State='AlreadyEffective';Added=$false}}
    if ($code -ne 1385) {throw "Service account batch authentication failed (Windows error $code); no right was added."}
    $rights=@(& $ReadRights $ExpectedSid)
    if ('SeDenyBatchLogonRight' -in $rights) {throw 'The service account is directly denied batch logon. Resolve the policy before continuing.'}
    if ('SeBatchLogonRight' -in $rights) {throw 'Batch logon is denied despite an existing direct grant. Check group denies and effective policy.'}
    & $AddRight $ExpectedSid
    try {$rights=@(& $ReadRights $ExpectedSid)} catch {throw 'The batch right may have been added but its state could not be verified. Inspect local policy before retrying.'}
    if ('SeBatchLogonRight' -notin $rights) {throw 'The batch logon grant could not be verified; inspect local policy before retrying.'}
    try {$code=[int](& $Probe $user $domain $Credential.Password $ExpectedSid)} catch {throw 'The batch right was added, but the account logon could not be verified. Inspect local policy before retrying.'}
    if ($code -ne 0) {throw "The batch right was added, but batch logon still failed (Windows error $code). Check effective GPO and group denies."}
    [pscustomobject]@{State='GrantedAndVerified';Added=$true}
}

function Ensure-SiInteractiveLogonRight {
    param([Parameter(Mandatory)][Management.Automation.PSCredential]$Credential,[Parameter(Mandatory)][string]$ExpectedSid,[scriptblock]$Probe,[scriptblock]$ReadRights,[scriptblock]$AddRight)
    if ($ExpectedSid -notmatch '^S-1-(?:\d+-)+\d+$' -or $ExpectedSid -eq 'S-1-5-18') {throw 'Invalid or forbidden service account SID.'}
    $user=$Credential.UserName;$domain=$null
    if ($user -match '^([^\\]+)\\([^\\]+)$') {$domain=$Matches[1];$user=$Matches[2]}
    elseif ($user -notmatch '^[^@\\\s]+@[^@\\\s]+$') {throw 'A domain account is required.'}
    if (-not $Probe -or -not $ReadRights -or -not $AddRight) {Initialize-SiBatchLogonNative}
    if (-not $Probe) {$Probe={param($u,$d,$p,$sid) [SmartM365Setup.BatchLogonNative]::ProbeInteractive($u,$d,$p,$sid)}}
    if (-not $ReadRights) {$ReadRights={param($sid) [SmartM365Setup.BatchLogonNative]::Read($sid)}}
    if (-not $AddRight) {$AddRight={param($sid) [SmartM365Setup.BatchLogonNative]::AddInteractiveRight($sid)}}
    $code=[int](& $Probe $user $domain $Credential.Password $ExpectedSid)
    if ($code -eq 0) {return [pscustomobject]@{State='AlreadyEffective';Added=$false}}
    if ($code -ne 1385) {throw "Service account interactive authentication failed (Windows error $code); no right was added."}
    $rights=@(& $ReadRights $ExpectedSid)
    if ('SeDenyInteractiveLogonRight' -in $rights) {throw 'The service account is directly denied interactive logon. Resolve the policy before continuing.'}
    if ('SeInteractiveLogonRight' -in $rights) {throw 'Interactive logon is denied despite an existing direct grant. Check group denies and effective policy.'}
    & $AddRight $ExpectedSid
    try {$rights=@(& $ReadRights $ExpectedSid)} catch {throw 'The interactive right may have been added but its state could not be verified. Inspect local policy before retrying.'}
    if ('SeInteractiveLogonRight' -notin $rights) {throw 'The interactive logon grant could not be verified; inspect local policy before retrying.'}
    try {$code=[int](& $Probe $user $domain $Credential.Password $ExpectedSid)} catch {throw 'The interactive right was added, but the account logon could not be verified. Inspect local policy before retrying.'}
    if ($code -ne 0) {throw "The interactive right was added, but interactive logon still failed (Windows error $code). Check effective GPO and group denies."}
    [pscustomobject]@{State='GrantedAndVerified';Added=$true}
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAeshrmK5OOXS1Y
# vi+M9E57fKB89ax+8ydvkcUAaZLJ6aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPA+xiAKYBhKReQyueUDf9W/1s6RIh1iEg2V+FNYTLPSMA0GCSqG
# SIb3DQEBAQUABIIBgIj6DnIZ0ij7QjNjn00ZrTS9ASFFfoqPwHZ9cQM9plfovqDa
# m3Q5qPrfjY5bLwMePUN8VzEuF/KBEPWirdtbaItVDt5NX9m8+rEi/EnugMac8IOq
# laXaIGw61e9tudhdUa5nEqkDug/FRwnYsoeflmYULd5x9+yjsLVqZgLgfwWNdD0I
# xiz5r9ArX6xmwnlN/d+22Au+L7j3uAzr3Gokh5pew7pKeQWFhiAm4Z64CpK2qoIP
# DPz2wXLotwnJ0Nom96w2IpxmKs0remuW42V0ZoBou726ZNdr0iClLRuABc9l95hK
# OksoMBRazV/XxOZm5FjTDWpPRK+8zS8nAapXjrLkieqGbUkgvB+99qjvj0DtSwtz
# rkD3OdnJOGX5fAu7PvK958MOO2k2JUHXMylmmUxOhpJzJbU0sxKpUd/kHnqUWzDI
# NMLMb4c2B8CwB0JoUZHmefFVp8zjqNj6B9rfNsZxfmxbVt5LHO5+Ef6+kT/sWbQ7
# etvUgBvxm0NSq+zcs6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTMx
# MTFaMC8GCSqGSIb3DQEJBDEiBCBwUyQ8/e0AZhP8boZo9ATvZPUMdIzubzzVcUUz
# y/rF+jANBgkqhkiG9w0BAQEFAASCAgBZhScOTHZ6pKGzMChkoU9XAeEhNh1Af5WE
# fRlsgNiwtiQQJH1K8kljWIX16DcZAVKpXDYRdxAVGd3vq89KR9fbb/fBypecyRtM
# 15MNEAEj41ODsxdWhgsysj8vuNzRLXboSft9U9LfLy+0SpTi9JjZ2q6qXu0MveKU
# trrqKwhF5MB7hsjSysxh7xSCVuuaE2swoUq6CHa1abhkgQyaClu9TwOsVcmRQyU1
# 9cpTtOEIPg6+HmJFhWKeR0u0kjxsV1iSr+UVCNzLJlHiDdk75CPL+z5bh2d1eD6k
# Xr/1osCLCGo3PLH6ZZnau973pRR0jIVzZiTFsMeKppYMydUUdfyN70JsQ3to2wOk
# 0J/nwS8sB46IGhcMRRTXnfJDVEuTPgnHKo4srFW7eaMErWg/08hExdu3/xpm4uek
# lf1PA4i1MQjLnf+Rtkgz4NvhPVUW0MCQeHcM7ub7GMvgEdJ7HrlXI5lv+l+Zfp2A
# mL/rSsl/W2pgDbkWXtWjgbSlqQAmW+io4iiITuaUmJI6X6zj9IJYbnmIO0s5iwsJ
# OBXCu2rXjbS9tRXQurIUYYrqnYIUne8z9AQbPFIC2rdXR4h9jRL0OcItxoZmbM4l
# Pe51WBJ2glYlwjGEW97nucGX42hSuo0s8L0uw/sfnnBK/WdHSEHY3qi3WPajMijo
# 3PZ7CFeiGQ==
# SIG # End signature block
