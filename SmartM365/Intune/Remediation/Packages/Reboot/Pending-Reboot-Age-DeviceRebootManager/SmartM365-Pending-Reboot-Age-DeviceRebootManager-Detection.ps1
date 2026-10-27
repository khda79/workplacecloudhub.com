#Requires -Version 5.1
# Description: Tracks aged pending reboots and requests the installed Smart Device Reboot Manager interactive task.
<#
.SYNOPSIS
    Tracks aged pending reboots and requests the installed Smart Device Reboot Manager interactive task.
.DESCRIPTION
    Standalone Intune detection, SYSTEM, 64-bit Windows PowerShell.
    Detection: 0 healthy/not applicable/deferred, 1 actionable finding, 2 technical error.
    Remediation: 0 completed/requested/deferred, 1 unresolved or technical failure.
    Requires the separately deployed signed app and standard limited task. Installs nothing and preserves app policy. Task start does not prove UI display or reboot.
.VERSION
    1.1.0
#>
[CmdletBinding()]
param([ValidateRange(1,720)][int]$PendingAgeHours = 72, [ValidateRange(1,168)][int]$ReminderCooldownHours = 24)

$Scenario = 'Pending-Reboot-Age-DeviceRebootManager'
$Remediate = $false

$ErrorActionPreference = 'Stop'
function Write-Result {
    param([string]$Status, [string]$Reason, [string]$Detail = '')
    $line = ('{0} Status={1} Reason={2} {3}' -f $Scenario, $Status, $Reason, $Detail).Trim()
    if ($line.Length -gt 2000) { $line = $line.Substring(0, 2000) }
    if ($Remediate) {
        $logRoot = Join-Path $env:ProgramData ('SmartM365\IntuneRemediation\Logs\' + $Scenario)
        [IO.Directory]::CreateDirectory($logRoot) | Out-Null
        Add-Content -LiteralPath (Join-Path $logRoot ($Scenario + '.log')) -Encoding UTF8 -Value ((Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + ' ' + $line)
    }
    Write-Output $line
}


function ConvertTo-ObservationDate {
    param($Value)
    if ($Value -is [datetime] -or $Value -is [datetimeoffset]) { return [datetimeoffset]$Value }
    if ([string]$Value -notmatch '^\d{4}-\d{2}-\d{2}T.+(Z|[+-]\d{2}:\d{2})$') { throw 'Invalid observation date.' }
    return [datetimeoffset]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture)
}
function Save-RebootObservation {
    param($Observation)
    [IO.Directory]::CreateDirectory((Split-Path $StatePath -Parent)) | Out-Null
    # Atomic replacement; a partial state must never be interpreted as an old observation.
    $temp = $StatePath + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    try {
        [IO.File]::WriteAllText($temp, (ConvertTo-Json -InputObject $Observation -Compress), [Text.UTF8Encoding]::new($false))
        if ([IO.File]::Exists($StatePath)) { [IO.File]::Replace($temp, $StatePath, [System.Management.Automation.Language.NullString]::Value) }
        else { [IO.File]::Move($temp, $StatePath) }
    }
    finally { if ([IO.File]::Exists($temp)) { [IO.File]::Delete($temp) } }
}
function Get-RebootObservation {
    $now = [datetimeoffset]::UtcNow
    $boot = [datetimeoffset](Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime
    if ($boot -gt $now.AddMinutes(5)) { throw 'Boot time is in the future.' }
    $reasons = @(foreach ($entry in @(
        @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending','CBS'),
        @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired','WindowsUpdate')
    )) { if (Test-Path -LiteralPath $entry[0] -ErrorAction Stop) { $entry[1] } })
    $state = $null
    if (Test-Path -LiteralPath $StatePath -ErrorAction Stop) {
        $state = Get-Content -LiteralPath $StatePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($state.SchemaVersion -ne 1) { throw 'Unsupported reboot observation schema.' }
    }
    if ($reasons.Count -eq 0) {
        $state = [pscustomobject]@{SchemaVersion=1;BootUtc=$boot.ToUniversalTime().ToString('o');FirstObservedUtc=$null;Reasons=@();LastReminderUtc=$null}
        Save-RebootObservation $state
        return [pscustomobject]@{State=$state;Pending=$false;Due=$false;AgeHours=0}
    }
    $reset = -not $state -or -not $state.FirstObservedUtc
    if (-not $reset) {
        $recordedBoot = (ConvertTo-ObservationDate $state.BootUtc)
        $first = (ConvertTo-ObservationDate $state.FirstObservedUtc)
        if ($first -gt $now.AddMinutes(5) -or $first -lt $recordedBoot) { throw 'Invalid reboot observation timestamp.' }
        $reset = $recordedBoot -ne $boot -or (@($state.Reasons | Where-Object { $_ -in $reasons }).Count -eq 0)
    }
    if ($reset) {
        $state = [pscustomobject]@{SchemaVersion=1;BootUtc=$boot.ToUniversalTime().ToString('o');FirstObservedUtc=$now.ToString('o');Reasons=$reasons;LastReminderUtc=$null}
    }
    else { $state.Reasons = $reasons }
    Save-RebootObservation $state
    $age = ($now - (ConvertTo-ObservationDate $state.FirstObservedUtc)).TotalHours
    return [pscustomobject]@{State=$state;Pending=$true;Due=($age -ge $PendingAgeHours);AgeHours=[math]::Floor($age)}
}

function Get-ValidatedRebootManagerTask {
    $task = Get-ScheduledTask -TaskPath '\SmartM365\' -TaskName 'Device Reboot Manager' -ErrorAction Stop
    if ($task.State -eq 'Disabled') { return $null }
    if ([string]$task.Principal.LogonType -ne 'Group' -or [string]$task.Principal.RunLevel -ne 'Limited' -or [string]$task.Principal.GroupId -notin @('S-1-5-32-545','BUILTIN\Users')) {
        throw 'Expected a limited interactive Users-group task.'
    }
    if (@($task.Actions).Count -ne 1) { throw 'Unexpected task action count.' }
    # Program Files since app 1.1.0; ProgramW6432 keeps the 64-bit path from a 32-bit host.
    $appRoot = Join-Path $(if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }) 'WCH\SmartDeviceRebootManager'
    $gui = Join-Path $appRoot 'SmartM365-DeviceRebootManager-GUI.ps1'
    $hostPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $action = $task.Actions[0]
    # App 1.1.0 runs PowerShell through a headless console host so no console window flashes.
    $headlessPrefix = ''
    if ([string]$action.Execute -ieq (Join-Path $env:SystemRoot 'System32\conhost.exe')) { $headlessPrefix = '--headless\s+"' + [regex]::Escape($hostPath) + '"\s+' }
    elseif ([string]$action.Execute -ine $hostPath) { throw 'Unexpected task executable.' }
    $pattern = '^\s*' + $headlessPrefix + '-STA\s+-NoProfile\s+-WindowStyle\s+Hidden\s+-ExecutionPolicy\s+Bypass\s+-File\s+"' + [regex]::Escape($gui) + '"(?:\s+-ConfigPath\s+"' + [regex]::Escape((Join-Path $appRoot 'SmartM365-DeviceRebootManager-GUI.config.json')) + '")?\s*$'
    if ([string]$action.Arguments -notmatch $pattern) { throw 'Unexpected task arguments.' }
    if ((Get-AuthenticodeSignature -LiteralPath $gui -ErrorAction Stop).Status -ne 'Valid') { throw 'Installed GUI signature is not valid.' }
    return $task
}

try {

    $StatePath = Join-Path $env:ProgramData ('SmartM365\IntuneRemediation\Output\' + $Scenario + '\observation.json')
    $observation = Get-RebootObservation
    if (-not $observation.Pending) { Write-Result -Status Healthy -Reason NoPendingReboot; exit 0 }
    $detail = 'ObservedAgeHours={0} ThresholdHours={1} Sources={2}' -f $observation.AgeHours,$PendingAgeHours,($observation.State.Reasons -join ',')
    if (-not $observation.Due) { Write-Result -Status Observing -Reason PendingAgeBelowThreshold -Detail $detail; exit 0 }
    if (-not $Remediate) { Write-Result -Status RemediationRequired -Reason PendingRebootAged -Detail $detail; exit 1 }
    if ($observation.State.LastReminderUtc) {
        $last = (ConvertTo-ObservationDate $observation.State.LastReminderUtc)
        if ($last -gt [datetimeoffset]::UtcNow.AddMinutes(5)) { throw 'Invalid reminder timestamp.' }
        if (([datetimeoffset]::UtcNow - $last).TotalHours -lt $ReminderCooldownHours) {
            Write-Result -Status Deferred -Reason ReminderCooldown -Detail 'RebootVerified=False'; exit 0
        }
    }

    $task = Get-ValidatedRebootManagerTask
    if (-not $task) { Write-Result -Status Deferred -Reason RebootManagerTaskDisabled; exit 0 }
    if ($task.State -eq 'Running') { Write-Result -Status Deferred -Reason RebootManagerAlreadyRunning; exit 0 }
    # The installed app owns session selection, uptime thresholds, postpone choices and restart policy.
    Start-ScheduledTask -InputObject $task -ErrorAction Stop

    $observation.State.LastReminderUtc = [datetimeoffset]::UtcNow.ToString('o')
    Save-RebootObservation $observation.State
    Write-Result -Status Requested -Reason RebootManagerTaskStarted -Detail 'RebootVerified=False'; exit 0

}
catch {
    # Do not expose user identities, certificate subjects, job names or native output.
    Write-Result -Status Error -Reason OperationFailed -Detail ('ErrorType=' + $_.Exception.GetType().Name)
    if ($Remediate) { exit 1 } else { exit 2 }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCCwiPoR2ZE10pv
# wv7JEFDB4bCZS9a9/bWfQ1+mrheNcqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDVW5ve6JYkHlsjyQR4kGwHb1bFCxCnrq1iKGfN5NE8RMA0GCSqG
# SIb3DQEBAQUABIIBgEYYtUlBzrwRHHBWWiNOUy5seFUfr9PZYEYNz5H88o71ysb0
# hJCIbWs3BlrSKuqXAtMPSo+UONu4o9q2qq1fi+CLL0u+6nH5A08/SqFBJq6RbrJ7
# 0YAooA6Z1Opq43xOi5EXb7SAZlpOmn+9zPrVwa5dEkV1al5IIy9LgfDRoQiSmbCM
# GSmvs+ysPHH+24jSy+QDusSniAKn51p4YwQ93XhZ8OeY+odczj4aOXzz+mD2r6nv
# Rjk83i1WTHIYPN8uJMDAPcl18ABZqDa1ixsngyG1FI35C/kaQODqr8uTPoqmEMGX
# 0qp5su52lFPF+QKOUiTvH4Y7vNELyeGpLKmoqUi6KwHRTFqC/MF6CZm4hjN2tHzO
# Oq3K4b+eab3nvHHeSErjVpwIWHvUmxfEScQP42TzB3dvi9+34C6g9JQKYIUoVx/h
# 1tsJ3B5MnNNRa3335AcdhYFOKClXkux5+cTjwVi+NDe2bEK5WW/oHDozUQNGDIfa
# zrroQZDl7zZp0dEulKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxMjQy
# MjRaMC8GCSqGSIb3DQEJBDEiBCD5dAK3ecMwlklOZXUUAJRw4LIEZRyY7MvfXUEe
# 0f1n/TANBgkqhkiG9w0BAQEFAASCAgBglovlk54rlebxMqKtKuS1aEV/ykM4Xy9c
# AsCIODPslh4tF+AiL8wxs1qThXGKZCkiuCFxZ9fEL2Zo3VK0wFaeFcTavRrTZPUs
# bfLx8WTM+mx0LkVRwvGjer1UjqeYrgh1KqHBhSrvPuqbidXtO0r4KMPPwiJ1yItI
# yGezjeVtXfuNUhl+J/B7zicK0zAQp4t2rMGqNdJf/e/hYCIIBH/PS4Sh14aUoZVD
# xerpXeWuXGrOmWMi8xdrKI9dpum+YnGOd1HJQTqBUWuGPgWULuqWiGjvzbfLgdR3
# T/hIlcvX/4DfySmgPiZeIPvOey8lKkrVY5/eiaIl5fYsbRP1LvIoyIPGkNjYR33u
# wHfIWiu1lqTwvuYuHLuISb2mcdfy6DDGRpVyzbgKKw1Gv+aiGSFMm1Hm6EWKfEn/
# pLk6A5JNaPygO5tegB6XfvVtysIUXkHtBYr5gnqjYcznX/oimQAIiW9y/kBLgS03
# ic+8FZkecxdk9rSkY8txGDGzRs+4eAD83yFkA9/+ZK7irW/e0mDupOzypXF3396M
# M6wZattZNGQ3PmDXhjH2Z/bBZh5RJewrDvzPfyXkLraM4QDxypSxm4vOVkSQdMSh
# lwxTzUi9Ep153GK2kv+H+atrv/HJVH0+rwnqoks6Y6QhZgryiJczw18WC1tBUIEy
# xD63iTHjtA==
# SIG # End signature block
