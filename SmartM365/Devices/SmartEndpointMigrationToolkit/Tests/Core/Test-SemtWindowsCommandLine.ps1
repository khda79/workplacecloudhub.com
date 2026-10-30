#requires -Version 5.1
<#
.SYNOPSIS
Offline round-trip tests for native Windows command-line argument quoting.
.VERSION
1.0.0
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
Import-Module (Join-Path $root 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class SemtCommandLineTest {
  [DllImport("shell32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
  public static extern IntPtr CommandLineToArgvW(string line, out int count);
  [DllImport("kernel32.dll")] public static extern IntPtr LocalFree(IntPtr memory);
  public static string[] Parse(string line) {
    int count; IntPtr ptr = CommandLineToArgvW(line, out count);
    if (ptr == IntPtr.Zero) throw new System.ComponentModel.Win32Exception();
    try {
      var result = new string[count];
      for (int i=0; i<count; i++) result[i] = Marshal.PtrToStringUni(Marshal.ReadIntPtr(ptr, i*IntPtr.Size));
      return result;
    } finally { LocalFree(ptr); }
  }
}
'@
$cases = @('', 'plain', 'C:\Program Files\WCH\script.ps1', 'espace accentué', 'a"b', 'a\"b', 'C:\with space\', 'many\\"quotes\\')
$line = 'test.exe ' + (($cases | ForEach-Object { ConvertTo-SemtWindowsCommandLineArgument -Value $_ }) -join ' ')
$parsed = [SemtCommandLineTest]::Parse($line)
if ($parsed.Count -ne ($cases.Count + 1)) { throw "Argument count mismatch: $($parsed.Count)" }
for ($i = 0; $i -lt $cases.Count; $i++) {
    if ($parsed[$i + 1] -cne $cases[$i]) { throw "Round-trip failed at argument $i" }
}
Write-Host ("Summary: Passed={0}; Failed=0" -f $cases.Count)

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD3sjk7zk/IaZss
# k8yR2a2nbl127VXQJqNntYX2JhTqiaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBWlxPFJ901dnnuJjIp+JMa
# COIVbyryYfI/zquHa0DwHTANBgkqhkiG9w0BAQEFAASCAYCDuVkrjXYFmHZ+61nK
# NU+tVHTxeyX0o1mMQePahhTPM6198GQoQmuJcYywmbjxytTR0o3qKVLSqkD0tJ1o
# li382E6/qdh6CCWJI9gILFOMYsxTXnO3r1AX8WaNH+vsxoaO6/KKdFcMNbJ7I8MN
# qw0taVSvLIJn020rV47HZkBuXgE+IPnSTZT3JbJY6Is1CDrBDZ/QBu3m+TdIER9r
# PlwZR9RvUVUCZNB3f4CS39NlUWwWUperlcRnh/ZLPlCs0qCEoyW2KXHYHCWOBUFY
# GOxyc3bE7g9B9ha0kCyIT9M1n6hydUDK4EmWX9fZtDNpIfQyWvwUfePhH9BTXUG0
# HXKqhuSY33iGMlcbW6JbpJ3ugdHiio74io7MATWWloGa5tfcviFYSpBA8QAiZrcf
# FHpdhclR6rQ8O2cgSDszRI/sqanhmHObONQCyX5fTjgA4Z1KipmGje/DZ5+FePy9
# 4ngSuoCHJFM0TUQNvCgppfcduXRK5I77hpm5TsqrTb175hg=
# SIG # End signature block
