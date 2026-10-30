<#
.SYNOPSIS
    Offline gpupdate timeout tests using simulated processes.
.VERSION
    1.0.0
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$hj = Get-Content -LiteralPath (Join-Path $root 'Campaigns\HybridJoin\Scripts\SmartM365-Invoke-EndpointMigration-HybridJoin.ps1') -Raw
$w11 = Get-Content -LiteralPath (Join-Path $root 'Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1') -Raw
$passed = 0
$failed = 0
function Assert-Test([string]$Name, [bool]$Condition) {
    if ($Condition) { $script:passed++; Write-Host "PASS: $Name" }
    else { $script:failed++; Write-Host "FAIL: $Name" }
}
function Import-Function([string]$Source, [string]$Name) {
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseInput($Source,[ref]$tokens,[ref]$errors)
    if ($errors.Count) { throw ('Parse error in endpoint: ' + $errors[0].Message) }
    $f = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name },$true)
    if (-not $f) { throw "Missing function $Name" }
    return $f.Extent.Text
}
. ([scriptblock]::Create((Import-Function $hj 'Invoke-ProcessWithTimeout')))
. ([scriptblock]::Create((Import-Function $w11 'Invoke-SemtGpUpdateWithTimeout')))
function Start-Process {
    param($FilePath,$ArgumentList,[switch]$PassThru,$WindowStyle,$RedirectStandardOutput,$RedirectStandardError)
    Set-Content -LiteralPath $RedirectStandardOutput -Value 'gpupdate output'
    Set-Content -LiteralPath $RedirectStandardError -Value 'diagnostic'
    $script:lastCall = @{ FilePath=$FilePath; Arguments=@($ArgumentList) }
    $p = [pscustomobject]@{ ExitCode=7; Killed=$false; WaitCount=0 }
    $p | Add-Member ScriptMethod WaitForExit { param($ms) $this.WaitCount++; if ($this.WaitCount -eq 1) { return -not $script:simulateTimeout }; return $true }
    $p | Add-Member ScriptMethod Kill { if ($script:denyKill) { throw 'termination denied' }; $this.Killed=$true }
    $p | Add-Member ScriptMethod Refresh { }
    $script:lastProcess = $p
    return $p
}
$temp = Join-Path $env:TEMP ('SEMT-test-gpupdate-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp | Out-Null
try {
    foreach ($campaign in @('HybridJoin','Windows11')) {
        foreach ($timeout in @($false,$true)) {
            $script:simulateTimeout = $timeout
            $script:denyKill = $false
            $outputFile = Join-Path $temp ($campaign + $timeout + '.txt')
            $result = if ($campaign -eq 'HybridJoin') {
                Invoke-ProcessWithTimeout -FilePath 'gpupdate.exe' -Arguments @('/target:computer','/force') -TimeoutSeconds 1 -OutputFile $outputFile
            } else { Invoke-SemtGpUpdateWithTimeout -TimeoutSeconds 1 -OutputFile $outputFile }
            Assert-Test "$campaign timeout=$timeout preserves output and arguments" ($result.TimedOut -eq $timeout -and (Get-Content -LiteralPath $outputFile -Raw) -match 'gpupdate output' -and $script:lastCall.FilePath -eq 'gpupdate.exe' -and ($script:lastCall.Arguments -join ' ') -eq '/target:computer /force')
            Assert-Test "$campaign timeout=$timeout handles process" ($(if ($timeout) { $script:lastProcess.Killed -and $result.ExitCode -eq '' } else { -not $script:lastProcess.Killed -and $result.ExitCode -eq 7 }))
        }
        $script:simulateTimeout = $true
        $script:denyKill = $true
        $outputFile = Join-Path $temp ($campaign + '-denied.txt')
        $result = if ($campaign -eq 'HybridJoin') {
            Invoke-ProcessWithTimeout -FilePath 'gpupdate.exe' -Arguments @('/target:computer','/force') -TimeoutSeconds 1 -OutputFile $outputFile
        } else { Invoke-SemtGpUpdateWithTimeout -TimeoutSeconds 1 -OutputFile $outputFile }
        Assert-Test "$campaign reports a failed termination" ($result.TimedOut -and $result.StopDetail -match 'termination denied')
    }
    # Execute the real initial policy-refresh branch with a fake process. The next
    # statement represents continuation toward the identity and MDM checks.
    $start = $hj.IndexOf('    if ($AuditOnly) {', $hj.IndexOf('    $domainPreflightComplete = $true'))
    $end = $hj.IndexOf('    if ($domainPreflightComplete) {', $start)
    if ($start -lt 0 -or $end -lt 0) { throw 'Initial policy-refresh branch not found.' }
    $policyBranch = [scriptblock]::Create($hj.Substring($start, $end - $start))
    function Write-RunLog { param($Message) }
    function Invoke-ProcessWithTimeout {
        param($FilePath,$Arguments,$TimeoutSeconds,$OutputFile)
        $script:gpCalls++
        return [pscustomobject]@{ ExitCode=0; Output=''; TimedOut=$false; StopDetail='' }
    }
    $OutputDir = $temp; $ComputerName = 'PC'; $RunId = 'test'; $GpUpdateTimeoutSeconds = 180
    $GpUpdateWaitSeconds = 0; $gpUpdateTimedOut = $false; $gpUpdateAttempted = $false
    $script:gpCalls = 0; $AuditOnly = $true; $continued = $false
    . $policyBranch
    $continued = $true
    Assert-Test 'Hybrid Join audit skips gpupdate and continues to identity checks' ($script:gpCalls -eq 0 -and -not $gpUpdateAttempted -and $continued)
    $AuditOnly = $false; $gpUpdateAttempted = $false
    . $policyBranch
    Assert-Test 'Hybrid Join repair still runs gpupdate' ($script:gpCalls -eq 1 -and $gpUpdateAttempted)
    Assert-Test 'Hybrid Join final gpupdate also excludes audit' ($hj -match '(?m)^\s*if \(-not \$AuditOnly -and \$finalGpUpdateReasons\.Count -gt 0\)')
    Assert-Test 'Hybrid Join stops with distinct initial and final timeout statuses' ($hj -match "status = 'GPUPDATE_TIMEOUT'" -and $hj -match "status = 'FINAL_GPUPDATE_TIMEOUT'" -and $hj -match 'INVESTIGATE_GROUP_POLICY_TIMEOUT' -and $hj -notmatch 'gpupdate\.exe /target:computer /force')
    Assert-Test 'Windows 11 policy repair returns timeout status' ($w11 -match "status = 'GPUPDATE_TIMEOUT'" -and $w11 -match 'PolicyRepairGpUpdateTimeout' -and $w11 -notmatch 'gpupdate\.exe /target:computer /force')
} finally { Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue }
Write-Host "Summary: Passed=$passed; Failed=$failed"
if ($failed) { exit 1 }

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAWpC3Pqy4ZkWLx
# tH6FOm6zkoF5jO6MLCB/vX55mClCVqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCA/iSn9P6Qsk67+xjTHtkhq
# X9wL3B/fAhDHISYk1zsKfjANBgkqhkiG9w0BAQEFAASCAYCJHhMif9jQ4Uk/nYeG
# d3I3KxSQH2BDkWkIRmPo+YcJKdcKACiyoXzICFIdmpd4otlBBQDncqVFC4+Ts1a2
# YejbzUYwafxpbrgtRcGq3gqHWDJCTP1SxxexU5qw6u26uleeszvERToQ1xsSeJtV
# FtzCnSApKnfYOhkgaHY1DAAgAbldhFMQxU6Wjb5uboBwW+M7vMhLHSd5EXdWFYwF
# /64JWNFqVOLomLBDMqUjh9mNpS66TGrydYXK7ETmAzzFajuiHnP9HAJXKpIohTyu
# CHpxy9muvEA93/bjVprd55d7X8wklGEGFy7ikmw2XJmghqx6cdOra1zaQcahWPtk
# JP9LtzWjvyWmqmyuQJsl+xokpdZohvIWntwxV6dkVmj7mmitmQxNea9SIiFk2Q4r
# seJw10k1uBvhe4AC0zDtIsXuKvdeGryAo6aQIqNyp4lLql8ZZf8TAlOHO5J0gl2b
# Ex9dx4DWzJaYNM6FbazsZIag/8NxfNu8cuRZfxiaO6CxWEc=
# SIG # End signature block
