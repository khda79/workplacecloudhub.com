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
        return [pscustomobject]@{ ExitCode=0; Output=''; TimedOut=$script:mockTimeout; StopDetail='' }
    }
    $OutputDir = $temp; $ComputerName = 'PC'; $RunId = 'test'; $GpUpdateTimeoutSeconds = 180
    $GpUpdateWaitSeconds = 0; $gpUpdateTimedOut = $false; $gpUpdateAttempted = $false
    $script:gpCalls = 0; $script:mockTimeout = $false; $AuditOnly = $true; $continued = $false
    . $policyBranch
    $continued = $true
    Assert-Test 'Hybrid Join audit skips gpupdate and continues to identity checks' ($script:gpCalls -eq 0 -and -not $gpUpdateAttempted -and $continued)
    $AuditOnly = $false; $gpUpdateAttempted = $false
    . $policyBranch
    Assert-Test 'Hybrid Join repair still runs gpupdate' ($script:gpCalls -eq 1 -and $gpUpdateAttempted)
    Assert-Test 'Hybrid Join final gpupdate also excludes audit' ($hj -match '(?m)^\s*if \(-not \$AuditOnly -and -not \$gpUpdateTimedOut -and \$finalGpUpdateReasons\.Count -gt 0\)')
    $script:mockTimeout = $true; $gpUpdateTimedOut = $false; $status = 'SUCCESS'; $ExitCode = 0
    . $policyBranch
    Assert-Test 'Hybrid Join continues after initial gpupdate timeout without changing enrollment status' ($script:gpCalls -eq 2 -and $gpUpdateTimedOut -and $status -eq 'SUCCESS' -and $ExitCode -eq 0)
    $startFinal = $hj.IndexOf('    $finalGpUpdateReasons = New-Object')
    $endFinal = $hj.IndexOf('    # Enrollment events describe observations', $startFinal)
    if ($startFinal -lt 0 -or $endFinal -lt 0) { throw 'Final policy-refresh branch not found.' }
    $finalBranch = [scriptblock]::Create($hj.Substring($startFinal, $endFinal - $startFinal))
    $leaveAttempted = $true; $nonIntuneMdmRemovalAttempted = $false; $gpUpdateMdmPolicyWarning = $false
    $intuneEnrolled = $false; $rebootAttempted = $false; $status = 'INTUNE_PENDING'; $ExitCode = 3
    $finalGpUpdateTimedOut = $false; $AuditOnly = $false; $gpUpdateTimedOut = $false
    . $finalBranch
    Assert-Test 'Final gpupdate timeout preserves prior enrollment status' ($finalGpUpdateTimedOut -and $status -eq 'INTUNE_PENDING' -and $ExitCode -eq 3)
    $priorCalls = $script:gpCalls; $gpUpdateTimedOut = $true
    . $finalBranch
    Assert-Test 'No second gpupdate after an initial timeout' ($script:gpCalls -eq $priorCalls)
    Assert-Test 'Hybrid Join records both timeouts and preserves the final enrollment status' ($hj -match 'final gpupdate exceeded' -and $hj -notmatch "status = 'FINAL_GPUPDATE_TIMEOUT'" -and $hj -match 'GpUpdateTimedOut' -and $hj -match 'FinalGpUpdateTimedOut')
    # Exercise the Windows 11 decision before the upgrade gate with simulated policy repair.
    $startW11 = $w11.IndexOf('    # Repair the existing policy gate before choosing the upgrade path.')
    $endW11 = $w11.IndexOf("    if (`$os.MajorFamily -eq 'Windows11') {", $startW11)
    if ($startW11 -lt 0 -or $endW11 -lt 0) { throw 'Windows 11 policy preflight not found.' }
    $w11Branch = [scriptblock]::Create($w11.Substring($startW11, $endW11 - $startW11))
    function Write-SmartLog { param($Message,$Level) }
    function Repair-WindowsUpdatePolicyBlockers { $script:repairCalls++; [pscustomobject]@{TimedOut=$true;StopDetail='';OutputFile='test.txt'} }
    function Get-WindowsUpdatePolicySummary { [pscustomobject]@{HasLegacyBlocker=$script:blockerRemains;Issues='test'} }
    $os=[pscustomobject]@{MajorFamily='Windows10'}; $hardwareReadinessBlocksUpgrade=$false
    $DirectSetupUpgrade=$false; $intune=[pscustomobject]@{IsIntuneEnrolled=$true}
    $indicators=[pscustomobject]@{ActionableBlocking=$false}; $freeGb=50; $MinimumFreeDiskGB=10
    $pendingReboot=$false; $policy=[pscustomobject]@{HasLegacyBlocker=$true;Issues='test'}
    $AllowPolicyRepair=$true; $AuditOnly=$false; $policyGpUpdateTimedOut=$false
    $policyRepairCompletedNormally=$false; $script:repairCalls=0; $script:blockerRemains=$false
    . $w11Branch
    Assert-Test 'Windows 11 rechecks policy and can continue after gpupdate timeout' ($script:repairCalls -eq 1 -and $policyGpUpdateTimedOut -and -not $policy.HasLegacyBlocker)
    $policy=[pscustomobject]@{HasLegacyBlocker=$true;Issues='test'}; $script:blockerRemains=$true
    . $w11Branch
    Assert-Test 'Windows 11 retains a real policy blocker after gpupdate timeout' ($policy.HasLegacyBlocker -and $policyGpUpdateTimedOut)
    Assert-Test 'Windows 11 does not replace upgrade result with gpupdate timeout' ($w11 -notmatch "status = 'GPUPDATE_TIMEOUT'" -and $w11 -match 'PolicyGpUpdateTimedOut')
} finally { Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue }
Write-Host "Summary: Passed=$passed; Failed=$failed"
if ($failed) { exit 1 }

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDh0Ryxe25THJw4
# mWJfW/aUziXiO6JTdvcvq5NE6r4HhqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBCtLTHIMG8H/u0Fk2NZcC1
# dcYZFooj1JHwqO03I4ki8zANBgkqhkiG9w0BAQEFAASCAYCoJLSg0PsBeKbN1IpF
# SnlBeTTVlro/web3TeR1HTyoBALRHFmf9YcBUgsFD3IPSD6/wKMLrdBHsImYLBXl
# uijcM6Yz1ZHYcnbNE3JztdrnX3CBG/pk6WwIGKXehwiV1D1LeseMPOVjZ3So5rPt
# qH65Rp4amSRdz7GWPBQfcH6FOBkoOQFoW10W0SjJXur7SIuPXyC1lfevzBSxdY4p
# K+hyNOSS/ztL6OL16YGXXAfO81J8xM0L7Z9Y31ySav51k/+483rH5WNuFe/QLpiY
# AUwhnYtRIjY1EsZQoAZ0Ow8LIL1bcY8nHSMnsbseYbxvzRguTuyQow1EXSjDiN4V
# sHjHEi77fUQA5dIx3GQj6PPasXfzPMdD0TQMgQsaHjkQC3NnhyuceApeKOs1tbNZ
# xNSlKpCwIN4KXk70Zt1TzTKsPv5xtwBYUAFIG4fQbFGW51ORe6aPWE4bbBB+Z7wL
# sOR3UEZi6eYEZL1LKXe6R+kuX824ZQhqIhHk4P2RUXeano0=
# SIG # End signature block
