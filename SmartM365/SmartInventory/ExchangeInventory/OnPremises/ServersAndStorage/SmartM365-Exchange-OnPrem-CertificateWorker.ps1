#Requires -Version 5.1
# Certificate-only client process and transport helpers, version 1.0.0.
# Dot-sourced library / machine-readable JSON-lines worker: no human-facing banners.
[CmdletBinding()]
param(
    [switch]$RunWorker,
    [ValidateSet('Native','Remote')][string]$WorkerMode = 'Native',
    [string]$WorkerRoot
)

function Start-SmartM365CertificateWorker {
    [CmdletBinding()]
    param(
        [ValidateSet('Native','Remote')][string]$Mode,
        [Parameter(Mandatory)][string]$Root,
        [string]$WorkerPath = (Join-Path $PSScriptRoot 'SmartM365-Exchange-OnPrem-CertificateWorker.ps1'),
        [string]$PowerShellPath = (Join-Path $PSHOME 'powershell.exe')
    )
    foreach ($path in @($WorkerPath,$PowerShellPath)) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Certificate worker dependency missing: $path" }
    }
    $launch = "& '{0}' -RunWorker -WorkerMode '{1}' -WorkerRoot '{2}'" -f $WorkerPath.Replace("'","''"), $Mode, $Root.Replace("'","''")
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $PowerShellPath
    $info.Arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand ' + [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($launch))
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardInput = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $info.StandardOutputEncoding = New-Object Text.UTF8Encoding($false)
    $info.StandardErrorEncoding = New-Object Text.UTF8Encoding($false)
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $info
    $started = $false
    try {
        [void]$process.Start()
        $started = $true
        return [pscustomobject]@{ Process=$process; ErrorTask=$process.StandardError.ReadToEndAsync(); Mode=$Mode; Closed=$false; ReadyLogged=$false }
    }
    catch {
        try { if ($started -and -not $process.HasExited) { $process.Kill(); [void]$process.WaitForExit(1000) } }
        finally { $process.Dispose() }
        throw
    }
}

function Stop-SmartM365CertificateWorker {
    [CmdletBinding()]
    param([AllowNull()][object]$State, [switch]$Force)
    if ($null -eq $State -or $State.Closed) { return }
    try {
        if (-not $State.Process.HasExited) {
            if (-not $Force) {
                try { $State.Process.StandardInput.WriteLine('{"Operation":"Close"}'); $State.Process.StandardInput.Flush() }
                catch { $Force = [switch]$true }
            }
            # A blocked RPC may ignore PowerShell cancellation. Never wait indefinitely for cleanup.
            if ($Force -or -not $State.Process.WaitForExit(1000)) {
                $State.Process.Kill()
                if (-not $State.Process.WaitForExit(1000)) { throw 'Owned certificate worker did not exit after termination.' }
            }
        }
    }
    finally { $State.Closed=$true; $State.Process.Dispose() }
}

function Invoke-SmartM365CertificateWorkerRead {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$State,
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][double]$TimeoutSeconds
    )
    $watch = [Diagnostics.Stopwatch]::StartNew()
    try {
        if ($TimeoutSeconds -le 0) { throw [TimeoutException]::new('Certificate server budget exhausted before request.') }
        if ($State.Closed -or $State.Process.HasExited) { throw 'Certificate worker is unavailable.' }
        $requestId = [guid]::NewGuid().ToString('N')
        $request = @{Operation='Read'; RequestId=$requestId; Server=$Server} | ConvertTo-Json -Compress
        $State.Process.StandardInput.WriteLine($request)
        $State.Process.StandardInput.Flush()
        $readTask = $State.Process.StandardOutput.ReadLineAsync()
        while (-not $readTask.IsCompleted) {
            if ($watch.Elapsed.TotalSeconds -ge $TimeoutSeconds) { throw [TimeoutException]::new(('Certificate {0} worker timeout after {1:F3}s.' -f $State.Mode,$TimeoutSeconds)) }
            Start-Sleep -Milliseconds 25
        }
        if ($watch.Elapsed.TotalSeconds -gt $TimeoutSeconds) { throw [TimeoutException]::new('Certificate response arrived after its deadline.') }
        $line = $readTask.GetAwaiter().GetResult()
        if ([string]::IsNullOrWhiteSpace($line)) { throw 'Certificate worker ended without a complete response.' }
        $reply = $line | ConvertFrom-Json -ErrorAction Stop
        foreach ($name in @('RequestId','Status','Items','Stage','Message','FatalSession','SessionReady')) {
            if ($null -eq $reply.PSObject.Properties[$name]) { throw "Invalid certificate worker response: missing $name." }
        }
        if ($reply.RequestId -ne $requestId -or $reply.Status -notin @('Success','Failed') -or $reply.Items -isnot [array]) { throw 'Invalid certificate worker response identity, status or items.' }
        if ($reply.FatalSession -isnot [bool] -or $reply.SessionReady -isnot [bool]) { throw 'Invalid certificate worker session flags.' }
        if ($reply.Status -eq 'Success') {
            foreach ($item in $reply.Items) {
                if ($null -eq $item -or $null -eq $item.PSObject.Properties['Thumbprint'] -or [string]::IsNullOrWhiteSpace([string]$item.Thumbprint)) { throw 'Invalid certificate worker snapshot.' }
            }
        }
        if ($watch.Elapsed.TotalSeconds -gt $TimeoutSeconds) { throw [TimeoutException]::new('Certificate response validation exceeded its deadline.') }
        return $reply
    }
    catch {
        if ($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]) { throw }
        return [pscustomobject]@{ Status=$(if ($_.Exception -is [TimeoutException]) {'Timeout'} else {'Failed'}); Items=@(); Stage='Transport'; Message=$_.Exception.Message; FatalSession=$true; SessionReady=$false }
    }
    finally { $watch.Stop() }
}

function ConvertTo-SmartM365CertificateSnapshot {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Certificate, [ValidateSet('Native','Remote')][string]$Mode)
    $snapshot = [ordered]@{Thumbprint=[string]$Certificate.Thumbprint}
    foreach ($name in @('Subject','CertificateDomains','Services','NotBefore','NotAfter','IsSelfSigned','Status','ServicesStringForm','DnsNameList')) {
        $property = $Certificate.PSObject.Properties[$name]
        if ($null -eq $property) { continue }
        $value = $property.Value
        if ($null -eq $value) { $snapshot[$name]=$null; continue }
        # Native values are rendered before IPC, while Exchange types are still alive.
        if ($Mode -eq 'Native') {
            $snapshot[$name] = if ($value -is [array]) { (@($value | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_ }) -join '; ') } else { [string]$value }
        }
        elseif ($value -is [Collections.IEnumerable] -and $value -isnot [string]) {
            $snapshot[$name] = @(foreach ($entry in $value) { if ($null -ne $entry -and $null -ne $entry.PSObject.Properties['Unicode']) { [string]$entry.Unicode } else { [string]$entry } })
        }
        else { $snapshot[$name] = if ($value -is [datetime] -or $value -is [enum]) { [string]$value } elseif ($value -is [string] -or $value -is [bool] -or $value -is [ValueType]) { $value } else { [string]$value } }
    }
    return [pscustomobject]$snapshot
}

function Invoke-SmartM365CertificateChildLoop {
    [CmdletBinding()]
    param([ValidateSet('Native','Remote')][string]$Mode, [Parameter(Mandatory)][string]$Root)
    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'
    [Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
    [Console]::InputEncoding = New-Object Text.UTF8Encoding($false)
    $ownedSession = $null
    $nativeCommand = $null
    $remoteDisabled = $false
    try {
        while ($null -ne ($line = [Console]::In.ReadLine())) {
            $request = $line | ConvertFrom-Json -ErrorAction Stop
            if ($request.Operation -eq 'Close') { break }
            if ($request.Operation -ne 'Read' -or [string]::IsNullOrWhiteSpace([string]$request.Server) -or [string]::IsNullOrWhiteSpace([string]$request.RequestId)) { throw 'Invalid certificate worker request.' }
            $stage = 'Initialize'
            $fatal = $false
            try {
                if ($Mode -eq 'Native') {
                    if ($null -eq $nativeCommand) {
                        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop -WarningAction SilentlyContinue
                        $nativeCommand = Get-Command -Name 'Microsoft.Exchange.Management.PowerShell.SnapIn\Get-ExchangeCertificate' -CommandType Cmdlet -ErrorAction Stop
                    }
                    $stage = 'Read'
                    $certificates = @(& $nativeCommand -Server $request.Server -ErrorAction Stop)
                }
                else {
                    if ($remoteDisabled) { throw 'Certificate remote session unavailable; no reconnect.' }
                    if ($null -eq $ownedSession) {
                        . (Join-Path $Root 'Modules/SmartM365.Core/SmartM365-ExchangeOnPremSession.ps1')
                        $ownedSession = Start-SmartM365ExchangeOnPremSession -RequiredCommands @('Get-ExchangeCertificate') 3>$null 4>$null 5>$null 6>$null
                    }
                    $command = Get-Command Get-ExchangeCertificate -ErrorAction SilentlyContinue
                    if ($null -eq $command -or $ownedSession.Session.State -ne 'Opened' -or $command.CommandType -ne 'Function' -or $command.ModuleName -ne $ownedSession.Module.Name) { $remoteDisabled=$true; throw 'Certificate remote session/proxy ownership lost; no reconnect.' }
                    $stage = 'Read'
                    $certificates = @(& $command -Server $request.Server -ErrorAction Stop)
                }
                $stage = 'Project'
                $items = @(foreach ($certificate in @($certificates | Sort-Object NotAfter)) { ConvertTo-SmartM365CertificateSnapshot -Certificate $certificate -Mode $Mode })
                $reply = @{RequestId=$request.RequestId; Status='Success'; Items=$items; Stage=$stage; Message=''; FatalSession=$false; SessionReady=($null -ne $ownedSession)}
            }
            catch {
                if ($Mode -eq 'Remote') {
                    $fatal = $remoteDisabled -or $null -eq $ownedSession -or $ownedSession.Session.State -ne 'Opened'
                    if ($fatal) { $remoteDisabled=$true }
                }
                $reply = @{RequestId=$request.RequestId; Status='Failed'; Items=@(); Stage=$stage; Message=$_.Exception.Message; FatalSession=$fatal; SessionReady=($null -ne $ownedSession)}
            }
            [Console]::Out.WriteLine(($reply | ConvertTo-Json -Depth 12 -Compress))
            [Console]::Out.Flush()
        }
    }
    finally {
        if ($null -ne $ownedSession) { Stop-SmartM365ExchangeOnPremSession -State $ownedSession }
    }
}

if ($RunWorker) { Invoke-SmartM365CertificateChildLoop -Mode $WorkerMode -Root $WorkerRoot }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCbio0aHZe48ySM
# yBFcHobMRAj1OsosXh8r2NS+IFREsKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMxxHbPnATPlBUeLXRjfu4E+xOlgsgWakdDhW2Rz5vnHMA0GCSqG
# SIb3DQEBAQUABIIBgAkbu1+2JCUqFL6mtV4/zQ34E1IyfoNmZXHsxh4y5Hr6daID
# +QaUo8D8ZFn6kt4kUxV6DE/Dn3iuawFDBGjI8khUJmHEJPUEOwkcog3CrYwGenSf
# ruITSgg29DE5u50/QfoCusfbb2lG73AR9+AMSfU91UhOjWpf8orF18Zy4n+m9oOA
# /uNgkeBM6ZyrvJv1di8rIZqCSq2zPxxHCnMug4jmoFsfaXEkVLQETz7OcSD4WLks
# sGy9KvMdoj289lmhzWgoHY+WrQvBWhZpaNGzNxap28/aMnw9Q7acgqxrwvwQXXw+
# 9TgHqztqNPar1vTE80mjJ4c69lbhFY1a1Jv6YriZXXYjU62Ir+vIoMqya/QY/RlZ
# W8j6aYiPJKKrJbMMDVfq3qyKkbqhbWD7sDo01mLhc3q6RJRusfzwaKqoQiaGyp8f
# bNmMqEdXhlXWsyO7bCbw2rH/3CGFD6rbRb7t+9Dg2NzxJz8B4cNh09GtfCFKcO6Y
# PUQaU9oqV9fOrJR5UKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMyMzE2
# MDhaMC8GCSqGSIb3DQEJBDEiBCBFltRr3yKx9iJCeVdU0pyzhrO7tVMnIexrWrhb
# XMnSUDANBgkqhkiG9w0BAQEFAASCAgAyjEZi+9jere6/LRt/GYTz95Wok6XmcLa3
# dBTpw38UifpgbHkwZngo5RXmkL9Vi13L9uikK1LUhPqrg73LmAciMbZhAga7ZgEe
# go62NF5AfAdbSxdY+0pZlheDz0zwSGBp0RFhMf3nwXBoH4T8ZFBNilDhZGCrfQH7
# rESY9zs0tt7YrK76pKRgim0VsLT8bmjkpsJC9dZx7ougiD03jQlaQORbfIE96bmH
# livNglaG6qBLYIDquQdhJgfO54D+EV/oH7mNUz1NcPwuTb4Rcx43xita67nk8xir
# sCzHikGSqK5uANZnVvcqAIRvhm9CwW/oFP6ZyRPc3pyh6yC0Z2MWf9drClG62qTZ
# 3nk5l5o6p/f7wJqlVPCf4d0TDvCZWrtG/baTaV7MRI+/oZAIoj/m9nUjT+sPk9JF
# RyevtVGEus45CYp333SJGW+TVkgSsyXQTR4yR90o/a1TdIH4nvUlH0IGUUyG9z23
# szHy9UU1fjtvnsmQPiA7u8+r9D2NOcVfTYYqoV9yUI0emIciPL9VBU1yOljO1ukw
# KWV8Mh7jNTRNZ1yT08nyRCH2GwMI18bSikHHjyj4ze4xcIO9u2N95UEv+XI0SMNF
# hyTMXGjNYscqnZIJhA/gG7/ZbJZWpsMcB3rh68FQA8Q3hX+v4wSm4Xk139vM7soU
# evQ0DIM92g==
# SIG # End signature block
