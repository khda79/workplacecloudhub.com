# Service-user certificate import helpers. No action runs when this file is loaded.
function Get-SiServiceCertificatePfx {
    param([Parameter(Mandatory)][string]$Root,[Parameter(Mandatory)][string]$Thumbprint)
    if ($Thumbprint -notmatch '^[0-9A-Fa-f]{40}$' -or $Thumbprint -match '^0+$') {throw 'The application certificate thumbprint is invalid.'}
    $folder=Join-Path $Root 'Config/Certificats'
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) {throw 'The application certificate export folder is unavailable.'}
    $item=Get-Item -LiteralPath $folder -Force -ErrorAction Stop
    if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {throw 'The certificate export folder must not be a reparse point.'}
    $files=@(Get-ChildItem -LiteralPath $folder -File -ErrorAction Stop | Where-Object { $_.Name -imatch ('^SmartM365-App-'+[regex]::Escape($Thumbprint)+'-[0-9a-f]{32}\.pfx$') -and -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) } | Sort-Object Name)
    if (-not $files.Count) {throw 'No exported PFX matches the configured application certificate.'}
    $files
}

# Keep this validation separate from process creation so it can be exercised offline.
function Test-SiServiceCertificateImportResult {
    param([Parameter(Mandatory)][string]$ExpectedSid,[Parameter(Mandatory)][string]$ExpectedThumbprint,[Parameter(Mandatory)]$Result)
    if ($ExpectedSid -notmatch '^S-1-(?:\d+-)+\d+$' -or $ExpectedSid -eq 'S-1-5-18') {throw 'The service account SID is invalid.'}
    if ($ExpectedThumbprint -notmatch '^[0-9A-Fa-f]{40}$' -or $ExpectedThumbprint -match '^0+$') {throw 'The certificate thumbprint is invalid.'}
    if ($Result.State -notin @('Imported','AlreadyPresent') -or [string]$Result.Sid -cne $ExpectedSid -or [string]$Result.Thumbprint -ine $ExpectedThumbprint -or $Result.HasPrivateKey -ne $true) {throw 'The certificate import could not be verified in the selected service account profile.'}
    return $true
}

function Initialize-SiServiceCertificateNative {
    if ('SmartM365Setup.ServiceCertificateNative' -as [type]) {return}
    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Principal;
using System.Text;
using Microsoft.Win32.SafeHandles;
namespace SmartM365Setup {
  public static class ServiceCertificateNative {
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
    private struct STARTUPINFO {
      public int cb; public string reserved,desktop,title; public int x,y,xsize,ysize,xcount,ycount,fill,flags;
      public short show,reserved2; public IntPtr reservedPtr,stdin,stdout,stderr;
    }
    [StructLayout(LayoutKind.Sequential)] private struct PROCESS_INFORMATION {public IntPtr process,thread;public int pid,tid;}
    [DllImport("advapi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    private static extern bool LogonUserW(string user,string domain,IntPtr password,int logonType,int provider,out SafeAccessTokenHandle token);
    [DllImport("advapi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    private static extern bool CreateProcessWithLogonW(string user,string domain,IntPtr password,int logonFlags,string app,StringBuilder command,int flags,IntPtr env,string directory,ref STARTUPINFO startup,out PROCESS_INFORMATION process);
    [DllImport("kernel32.dll", SetLastError=true)] private static extern bool CloseHandle(IntPtr handle);
    [DllImport("kernel32.dll", SetLastError=true)] private static extern bool GetNamedPipeClientProcessId(SafePipeHandle pipe,out uint pid);
    public static int ClientPid(SafePipeHandle pipe) {uint id;if(!GetNamedPipeClientProcessId(pipe,out id)) throw new Win32Exception(Marshal.GetLastWin32Error());return checked((int)id);}
    public static int Start(string user,string domain,SecureString password,string expectedSid,string executable,string encoded,string directory) {
      IntPtr secret=Marshal.SecureStringToGlobalAllocUnicode(password);
      try {
        SafeAccessTokenHandle token;
        if (!LogonUserW(user,String.IsNullOrEmpty(domain)?null:domain,secret,2,0,out token)) {
          int error=Marshal.GetLastWin32Error();
          throw new Win32Exception(error,"Service account interactive logon failed (Windows error "+error+"). Check the account's local logon rights.");
        }
        using(token) {
          using(var identity=new WindowsIdentity(token.DangerousGetHandle())) {if(identity.User.Value!=expectedSid) throw new InvalidOperationException("Authenticated service SID differs from the selected account.");}
          var si=new STARTUPINFO();si.cb=Marshal.SizeOf(typeof(STARTUPINFO));
          var command=new StringBuilder("\""+executable+"\" -NoProfile -NonInteractive -EncodedCommand "+encoded);
          if(command.Length >= 1024) throw new InvalidOperationException("Certificate worker launch command exceeds the Windows limit.");
          PROCESS_INFORMATION pi;
          if(!CreateProcessWithLogonW(user,String.IsNullOrEmpty(domain)?null:domain,secret,1,executable,command,0x00000010,IntPtr.Zero,directory,ref si,out pi)) {
            int error=Marshal.GetLastWin32Error();
            throw new Win32Exception(error,"Could not launch the service certificate worker with the user profile (Windows error "+error+": "+new Win32Exception(error).Message+")");
          }
          try {return pi.pid;} finally {CloseHandle(pi.thread);CloseHandle(pi.process);}
        }
      } finally {Marshal.ZeroFreeGlobalAllocUnicode(secret);}
    }
  }
}
'@
}

function Invoke-SiServiceCertificateImport {
    param(
        [Parameter(Mandatory)][string]$PfxPath,
        [Parameter(Mandatory)][Security.SecureString]$PfxPassword,
        [Parameter(Mandatory)][Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory)][string]$ExpectedSid,
        [Parameter(Mandatory)][string]$Thumbprint,
        [Parameter(Mandatory)][string]$PowerShell,
        [scriptblock]$LaunchWorker
    )
    if ($ExpectedSid -notmatch '^S-1-(?:\d+-)+\d+$' -or $ExpectedSid -eq 'S-1-5-18') {throw 'Invalid service account SID.'}
    if ($Thumbprint -notmatch '^[0-9A-Fa-f]{40}$' -or $Thumbprint -match '^0+$') {throw 'Invalid application certificate thumbprint.'}
    if (-not (Test-Path -LiteralPath $PfxPath -PathType Leaf) -or -not (Test-Path -LiteralPath $PowerShell -PathType Leaf)) {throw 'PFX file or PowerShell executable not found.'}
    $file=Get-Item -LiteralPath $PfxPath -Force -ErrorAction Stop
    if (($file.Attributes -band [IO.FileAttributes]::ReparsePoint) -or $file.Length -gt 4194304 -or $file.Length -eq 0) {throw 'PFX file is unsafe or has an invalid size.'}
    $user=$Credential.UserName;$domain=$null
    if ($user -match '^([^\\]+)\\([^\\]+)$') {$domain=$Matches[1];$user=$Matches[2]}
    elseif ($user -notmatch '^[^@\\\s]+@[^@\\\s]+$') {throw 'Use a domain service account.'}
    Initialize-SiServiceCertificateNative
    $pipeName='SmartM365-Cert-'+[guid]::NewGuid().ToString('N')
    $security=[IO.Pipes.PipeSecurity]::new();$security.SetAccessRuleProtection($true,$false)
    foreach ($sid in @([Security.Principal.SecurityIdentifier]::new($ExpectedSid),[Security.Principal.WindowsIdentity]::GetCurrent().User,[Security.Principal.SecurityIdentifier]::new('S-1-5-18'))) {
        $security.AddAccessRule([IO.Pipes.PipeAccessRule]::new($sid,[IO.Pipes.PipeAccessRights]::FullControl,[Security.AccessControl.AccessControlType]::Allow))
    }
    $pipe=[IO.Pipes.NamedPipeServerStreamAcl]::Create($pipeName,[IO.Pipes.PipeDirection]::InOut,1,[IO.Pipes.PipeTransmissionMode]::Byte,[IO.Pipes.PipeOptions]::Asynchronous,4096,4096,$security,[IO.HandleInheritability]::None,[IO.Pipes.PipeAccessRights]0)
    $worker=@'
$ErrorActionPreference='Stop'
$reader=[IO.StreamReader]::new($pipe,[Text.Encoding]::UTF8,$false,4096,$true)
$writer=[IO.StreamWriter]::new($pipe,[Text.UTF8Encoding]::new($false),4096,$true);$writer.AutoFlush=$true
$bytes=$null;$password=$null;$cert=$null;$store=$null
try {
    $sid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $writer.WriteLine('READY|'+$sid)
    $data=$reader.ReadLine() | ConvertFrom-Json -ErrorAction Stop
    if ($sid -cne [string]$data.Sid -or [string]$data.Thumbprint -notmatch '^[0-9A-Fa-f]{40}$') {throw 'Identity mismatch.'}
    $bytes=[Convert]::FromBase64String([string]$data.Pfx)
    $password=ConvertTo-SecureString -String ([string]$data.Password) -AsPlainText -Force
    $store=[Security.Cryptography.X509Certificates.X509Store]::new('My',[Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    $store.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $existing=@($store.Certificates | Where-Object {$_.Thumbprint -ieq [string]$data.Thumbprint})
    if ($existing.Count) {
        if (@($existing | Where-Object HasPrivateKey).Count -eq 0) {throw 'A certificate without its private key is already present.'}
        $state='AlreadyPresent'
    } else {
        $ephemeral=[Security.Cryptography.X509Certificates.X509Certificate2]::new($bytes,$password,[Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet)
        try {if ($ephemeral.Thumbprint -ine [string]$data.Thumbprint -or -not $ephemeral.HasPrivateKey) {throw 'PFX certificate does not match.'}} finally {$ephemeral.Dispose()}
        $flags=[Security.Cryptography.X509Certificates.X509KeyStorageFlags]::UserKeySet -bor [Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet
        $cert=[Security.Cryptography.X509Certificates.X509Certificate2]::new($bytes,$password,$flags)
        $store.Add($cert);$state='Imported'
    }
    $verified=@($store.Certificates | Where-Object {$_.Thumbprint -ieq [string]$data.Thumbprint -and $_.HasPrivateKey})
    if (-not $verified.Count) {throw 'The private key was not found in the user store.'}
    $writer.WriteLine((@{State=$state;Sid=$sid;Thumbprint=[string]$data.Thumbprint;HasPrivateKey=$true} | ConvertTo-Json -Compress))
} catch {$writer.WriteLine('{"State":"Failed"}')}
finally {if($store){$store.Close()};if($cert){$cert.Dispose()};if($password){$password.Dispose()};if($bytes){[Array]::Clear($bytes,0,$bytes.Length)};$reader.Dispose();$writer.Dispose();$pipe.Dispose()}
'@
    $workerEncoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($worker))
    $bootstrap='$p=[IO.Pipes.NamedPipeClientStream]::new(''.'',''__PIPE__'',3);$p.Connect(30000);$r=[IO.StreamReader]::new($p,[Text.Encoding]::UTF8,0,128,1);$s=$r.ReadLine();$r.Dispose();$pipe=$p;iex([Text.Encoding]::Unicode.GetString([Convert]::FromBase64String($s)))'
    $bootstrap=$bootstrap.Replace('__PIPE__',$pipeName)
    $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($bootstrap))
    $process=$null;$reader=$null;$writer=$null;$bytes=$null;$bstr=[IntPtr]::Zero;$plain=$null;$sent=$false
    try {
        $pidValue=if ($LaunchWorker) {& $LaunchWorker $PowerShell $encoded $pipeName $ExpectedSid} else {[SmartM365Setup.ServiceCertificateNative]::Start($user,$domain,$Credential.Password,$ExpectedSid,$PowerShell,$encoded,[Environment]::GetFolderPath('System'))}
        $process=Get-Process -Id $pidValue -ErrorAction Stop
        $wait=$pipe.WaitForConnectionAsync()
        if (-not $wait.Wait(30000)) {throw 'The service certificate worker did not connect.'}
        if ([SmartM365Setup.ServiceCertificateNative]::ClientPid($pipe.SafePipeHandle) -ne $pidValue) {throw 'The certificate worker identity could not be confirmed.'}
        $reader=[IO.StreamReader]::new($pipe,[Text.Encoding]::UTF8,$false,4096,$true)
        $writer=[IO.StreamWriter]::new($pipe,[Text.UTF8Encoding]::new($false),4096,$true);$writer.AutoFlush=$true
        $writer.WriteLine($workerEncoded)
        $handshake=$reader.ReadLineAsync();if (-not $handshake.Wait(10000) -or $handshake.Result -cne ('READY|'+$ExpectedSid)) {throw 'The service certificate worker identity could not be confirmed.'}
        $bytes=[IO.File]::ReadAllBytes($PfxPath)
        if ($bytes.Length -ne $file.Length -or $bytes.Length -gt 4194304) {throw 'The PFX file changed while being read.'}
        $bstr=[Runtime.InteropServices.Marshal]::SecureStringToBSTR($PfxPassword);$plain=[Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        $payload=@{Sid=$ExpectedSid;Thumbprint=$Thumbprint;Pfx=[Convert]::ToBase64String($bytes);Password=$plain} | ConvertTo-Json -Compress
        $writer.WriteLine($payload);$sent=$true;$payload=$null;$plain=$null
        $response=$reader.ReadLineAsync();if (-not $response.Wait(120000)) {throw 'Certificate import result is uncertain; inspect the service account store before retrying.'}
        $result=$response.Result | ConvertFrom-Json -ErrorAction Stop
        $null=Test-SiServiceCertificateImportResult -ExpectedSid $ExpectedSid -ExpectedThumbprint $Thumbprint -Result $result
        return $result
    } catch {
        if ($sent) {throw 'Certificate import failed or is uncertain. Inspect the service account certificate store before retrying.'}
        throw
    } finally {
        if ($bstr -ne [IntPtr]::Zero) {[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)}
        if ($bytes) {[Array]::Clear($bytes,0,$bytes.Length)}
        if ($reader) {$reader.Dispose()};if ($writer) {$writer.Dispose()};$pipe.Dispose();if ($process) {$process.Dispose()}
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCE+iVAzwGlo5M/
# QgCAf0hsW5l+yPDOc1zZ/ZHB9lM6wKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPp+DjeTh8HNkrswPiKSzSqLQpqmb9OgZWyHL08OrlJ4MA0GCSqG
# SIb3DQEBAQUABIIBgC2ldZC/wDBZAi3CDXnszxyYH+/fuJtocjKwRPvneBBngjji
# bs9tqsE/bCPKOKR4dV0Qj1YjJ8FMikN9BtEtyM111gjfyJdGVDTtdaBLeZBPWEPs
# d8uNABvJywvr+Iu/eXzWrSltrm/c6xk6FmlCZj6x2J75WtnIe7fbZD54QkiXOwv0
# /yttkIQa6EAY+Epio1cRirUgl2VotgFAYpWJ+3f/jc2NtiH9heUJpkChJ8PZm6Xp
# gFlOmR4r1pBGKl0dqHziGcI4zPlRQGlRk1tCN/XTaKkf53dMpLSs7fQSnmdZOO5x
# WapNG3p6i6UOzugwA2z2JqzUw2IVYeBBY+s0+TkAoE0V/B465IvvmxKXfxJCMpjo
# nYE4T3aqTchI7N3Uzqt+DVxhaWsUXnSM5Pm+4r1yISAGRYMhtUoJODkUOLzkxfP8
# gS7z6CZ5sK/0B+VPKl+DVYIWc9WbuBKMI7bqCSFdq4VQOasvv6jn01wv9vaqeQvY
# 1OwSzyAKdPTAwEp3bKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTMx
# MTJaMC8GCSqGSIb3DQEJBDEiBCBXzmb3lXGong3rjIUnYNczUiMxJPC1TehuXYp4
# acIx8DANBgkqhkiG9w0BAQEFAASCAgAzEx7/NZCGN4fInGy4VJDTCJxGP5J2stRI
# P4pSN4IUkF3l6hi8io9x/WfEL5gh+rEEbtIH3a8H3gImDjGbTtQnY8ucupAboJii
# o59J1AKo7DrGE7d7kbTg/BhTs8mN1+dpST4NSumIPSxnOL1N94b6PTAX51tGYCle
# OMx2V88Mq9XuYu0lxBN84btUbKZkfFnDajuQpGBrUEorGxXJPqrsYOqXlqnVNQaP
# AmoczMy4zpyVShiwT1ac8XpRia9PClcGyfabbuciEOYVEVBTuIsE6UFFGi8fFZRy
# EXXoYpxmJhMQUSWsGSWzrZjfjqV42gR+fHsISUKDtHZdeDQHSkyVGlMduUcRo4fZ
# vhsYveRi8pD1JXD123GsbtICE4VPagbA0qwpMJnpyRPOxW2Z/AInROJJZiaB1kGu
# /jbL6AAwL4xQ9awWk5/6QPg33mEftCsK4NOLTQNORT+PlynfRYdd43xAjUPyU4K5
# AlV6J1nCpg/Yr1LjkFvLE0ulMNZTIZibkyV0mRj8vKqZz7eDc31MgbJJNlnUpHJ4
# Gpp0VE1AndJYQJ3Pr0+4lDMX7/dWe0OwQEP1R+V14+jpg/jBbYplPQd4YVRYpc6y
# kBbOWFQjkvIou6rrt+AadLOGUoqVs1C0aN9SRgYhm1/SUt/dbUJLHQpuWFIUsMO8
# f3bvSzQc8Q==
# SIG # End signature block
