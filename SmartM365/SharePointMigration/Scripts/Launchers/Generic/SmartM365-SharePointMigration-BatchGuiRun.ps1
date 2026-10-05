<#
.SYNOPSIS
    Run an existing CMD batch launcher in a dedicated dashboard console.
.VERSION
    1.0.0
.DESCRIPTION
    Internal launcher. The JSON request is data, not PowerShell code. Records the
    actual child exit code independently of whether the console stays open.
#>
#requires -Version 7.4
[CmdletBinding()]
param([Parameter(Mandatory)][string]$RequestPath)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-BatchGui.ps1')

function Save-BatchReceipt {
    param($Receipt)
    $temporary=$RequestPath + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    [IO.File]::WriteAllText($temporary, ($Receipt | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false))
    # A GUI reader or sync client can briefly hold the previous receipt open.
    for($attempt=0;$attempt -lt 5;$attempt++) {
        try { [IO.File]::Move($temporary, $RequestPath, $true); return }
        catch [IO.IOException] { if($attempt -eq 4){throw}; Start-Sleep -Milliseconds 100 }
    }
}

$request=$null
$process=$null
try {
    $request=Get-Content -LiteralPath $RequestPath -Raw | ConvertFrom-Json -AsHashtable
    $arguments=@{Kind=[string]$request.Kind;Root=[string]$request.Root;Mode=[string]$request.Mode;AuthMode=[string]$request.AuthMode;Names=@($request.Names);PlanOnly=[bool]$request.PlanOnly;BatchId=[string]$request.BatchId}
    $invocation=Get-SmartM365BatchCommand @arguments
    if(-not (Test-Path -LiteralPath $invocation.Launcher -PathType Leaf)){throw "Batch launcher not found: $($invocation.Launcher)"}
    $request.Status='Running';$request.Error='';Save-BatchReceipt $request
    Write-Host ('[{0}] {1} batch: {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$request.Kind,$invocation.Command)
    $info=[Diagnostics.ProcessStartInfo]::new($env:ComSpec,$invocation.CmdArguments)
    $info.UseShellExecute=$false;$info.WorkingDirectory=$env:TEMP
    $process=[Diagnostics.Process]::Start($info)
    $process.WaitForExit()
    $request.ExitCode=$process.ExitCode
    $request.Status=if($process.ExitCode -ne 0){'Failed'}elseif($request.PlanOnly){'Previewed'}else{'Succeeded'}
} catch {
    if($request){$request.Status='Failed';$request.ExitCode=1;$request.Error=$_.Exception.Message}
    foreach($line in ($_.Exception.Message -split '\r?\n')){Write-Host ('[{0}] Batch launch failed: {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$line) -ForegroundColor Red}
} finally {
    if($request){$request.FinishedUtc=[DateTimeOffset]::UtcNow.ToString('o');Save-BatchReceipt $request;Write-Host ('[{0}] Batch status: {1}; exit code: {2}; receipt: {3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$request.Status,$request.ExitCode,$RequestPath)}
    if($process){$process.Dispose()}
}
if($request -and $request.ExitCode -ne 0){exit $request.ExitCode}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC8buIGk3iIGCQW
# JgR8Cs2Z0TjTE7kI9IndtWFyv+3IjaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCFhAKYAuFkp5rhtzgfjqP1
# Vm0/+1rfmGQGKf46zaFP5DANBgkqhkiG9w0BAQEFAASCAYBWBq4+4L0Yg6UGekYb
# 8qY+4sRh5Jp4Ddws+nsil3Cp6oZbmovh4Ahp0pbTBGGU6B/DG9P0oBJm4IulL51N
# 6byMPApplOJJOb6BA+5RRIbDObZQ43UA1LG864DJppBS4zQwdJqcgjKS/Olu7V3c
# cwKE7YGDr0vd5EPebEKNGxAiapNCGDHrSK893YTCjG7be8lyc6PHfwcKcYFrGzux
# mWUteFCLGyFv99oou+gvd8LR/tSRT3NrtqBGY1dGD16c1lB08jxUfL2jDoEY83+h
# 50VgKFfdFSCFx2nkn5yPuCav2OV/vlNncW0XWFsU/ZDivnHicKs02FR1ffDP+NCQ
# itWYhdS/pXbTZvumrQRjtzWDuhh7pSNqv26ToXm4FYnUtGXygL8b9vucIOBNLjzV
# 8RDIdLqnvsEAvt7aztRLbYAtmnjvg5WVy7jVAILqX2ahVP5t0CWAaz4LX6p7yIp6
# kMDb4fqtKRmHwhbdKdsgvNrx5hwH5hjNb0KyJ1+v5Y7ieIc=
# SIG # End signature block
