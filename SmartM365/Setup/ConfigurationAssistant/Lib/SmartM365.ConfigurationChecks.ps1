# SmartInventory local prerequisite check, version 0.11.0.
function Invoke-SiPrerequisiteAudit {
    param([Parameter(Mandatory)][string]$SmartM365Root,[scriptblock]$Runner)
    $scriptPath = Join-Path $SmartM365Root 'Setup/Install-SmartM365-SmartInventoryPrerequisites.ps1'
    if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) { throw 'Prerequisite audit script is missing.' }
    if ($null -eq $Runner) {
        $Runner = {
            param($path)
            $exe = (Get-Command pwsh -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
            $output = @(& $exe -NoProfile -NonInteractive -File $path -AuditOnly 2>&1 | ForEach-Object { [string]$_ })
            [pscustomobject]@{ ExitCode=$LASTEXITCODE; Lines=$output }
        }
    }
    $result = & $Runner $scriptPath
    if ($null -eq $result -or $null -eq $result.ExitCode -or $result.ExitCode -notin @(0,1)) { throw 'Audit process did not return an interpretable result.' }
    [pscustomobject]@{
        State = if ($result.ExitCode -eq 0) { 'Modules found; imports not verified' } else { 'Modules missing or versions need review' }
        Detail = (@($result.Lines) -join [Environment]::NewLine)
    }
}

function Test-SiSharedPathRead {
    param([Parameter(Mandatory)][string]$Path,[int]$TimeoutSeconds=10,[string]$ExpectedSid='')
    if (-not [IO.Path]::IsPathFullyQualified($Path)) { throw 'Enter an absolute local or UNC path.' }
    if ($TimeoutSeconds -lt 1 -or $TimeoutSeconds -gt 60) { throw 'Invalid timeout.' }
    if ($ExpectedSid -and $ExpectedSid -notmatch '^S-1-(?:\d+-)+\d+$') { throw 'Invalid expected SID.' }
    $exe = Join-Path $PSHOME 'pwsh.exe'
    $probe = Join-Path $PSScriptRoot 'SmartM365-CheckSharedPath.ps1'
    if (-not (Test-Path -LiteralPath $exe -PathType Leaf) -or -not (Test-Path -LiteralPath $probe -PathType Leaf)) { throw 'Shared path probe or PowerShell is missing.' }
    $probeSources=@($probe)
    $probeLocks=[Collections.Generic.List[object]]::new()
    try {
        foreach ($source in $probeSources) {
            $hash=(Get-FileHash -LiteralPath $source -Algorithm SHA256 -ErrorAction Stop).Hash
            $probeLocks.Add([IO.File]::Open($source,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read))
            if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -cne $hash) { throw 'Shared path probe changed before launch.' }
        }
    } catch { foreach ($lock in $probeLocks) { $lock.Dispose() }; throw }
    $info=[Diagnostics.ProcessStartInfo]::new()
    $info.FileName=$exe
    $info.UseShellExecute=$false
    $info.CreateNoWindow=$true
    $info.RedirectStandardOutput=$true
    $info.RedirectStandardError=$true
    foreach ($arg in @('-NoProfile','-NonInteractive','-File',$probe,'-Path',$Path)) { $info.ArgumentList.Add($arg) }
    if ($ExpectedSid) { $info.ArgumentList.Add('-ExpectedSid'); $info.ArgumentList.Add($ExpectedSid) }
    $proc=[Diagnostics.Process]::new(); $proc.StartInfo=$info
    try {
        if (-not $proc.Start()) { throw 'Unable to start shared path probe.' }
        $stdout=$proc.StandardOutput.ReadToEndAsync()
        $stderr=$proc.StandardError.ReadToEndAsync()
        if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
            try {$proc.Kill($true);$stopped=$proc.WaitForExit(3000)} catch {$stopped=$false}
            return [pscustomobject]@{State='Timeout'; Detail=$(if($stopped){'Timed out. Access cannot be determined.'}else{'Timed out. Process termination and access are uncertain.'})}
        }
        $streams=Read-SiBoundedProcessStreams -Stdout $stdout -Stderr $stderr
        $text=$streams.Stdout
        try { $result=$text | ConvertFrom-Json -ErrorAction Stop } catch { throw 'The shared path probe returned invalid output.' }
        if ($proc.ExitCode -eq 0 -and $result.State -eq 'Readable') {
            if ($ExpectedSid -and $result.ExpectedIdentity -ne $true) { throw 'The probe did not confirm the expected identity.' }
            if ($ExpectedSid) { return [pscustomobject]@{State='ReadableAsExpected'; Detail='Folder listing succeeded under a process token with the selected SID. Write access, task logon and differences in group membership or logon type remain unverified.'} }
            return [pscustomobject]@{State='Readable'; Detail='Folder listing is accessible to the current account. Write access and service-account access are not verified.'}
        }
        if ($proc.ExitCode -eq 2 -and $result.State -eq 'IdentityMismatch' -and $ExpectedSid) { return [pscustomobject]@{State='IdentityMismatch';Detail='The probe process does not run as the selected service identity. The folder was not contacted.'} }
        if ($proc.ExitCode -eq 1 -and $result.State -eq 'Unavailable') {
            return [pscustomobject]@{State='Unavailable'; Detail='Read failed or path unavailable to the current account. The cause and service-account rights need investigation.'}
        }
        throw 'The shared path probe returned an unexpected result.'
    } finally { $proc.Dispose(); foreach ($lock in $probeLocks) { $lock.Dispose() } }
}

function Test-SiSharedPathWrite {
    param([Parameter(Mandatory)][string]$Path,[int]$TimeoutSeconds=10,[string]$MarkerName=('.smartm365-setup-probe-'+[guid]::NewGuid().ToString('N')+'.tmp'),[string]$ExpectedSid='')
    if (-not [IO.Path]::IsPathFullyQualified($Path)) { throw 'Enter an absolute local or UNC path.' }
    if ($TimeoutSeconds -lt 1 -or $TimeoutSeconds -gt 60 -or $MarkerName -cnotmatch '^\.smartm365-setup-probe-[0-9a-f]{32}\.tmp$') { throw 'Invalid probe parameters.' }
    if ($ExpectedSid -and $ExpectedSid -notmatch '^S-1-(?:\d+-)+\d+$') { throw 'Invalid expected identity.' }
    $exe=Join-Path $PSHOME 'pwsh.exe'; $probe=Join-Path $PSScriptRoot 'SmartM365-CheckSharedWrite.ps1'
    if (-not (Test-Path -LiteralPath $exe -PathType Leaf) -or -not (Test-Path -LiteralPath $probe -PathType Leaf)) { throw 'Shared path probe or PowerShell is missing.' }
    $hash=(Get-FileHash -LiteralPath $probe -Algorithm SHA256).Hash
    $probeLock=[IO.File]::Open($probe,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
    try { if ((Get-FileHash -LiteralPath $probe -Algorithm SHA256).Hash -cne $hash) {throw 'Write probe changed before launch.'} } catch {$probeLock.Dispose();throw}
    $info=[Diagnostics.ProcessStartInfo]::new(); $info.FileName=$exe; $info.UseShellExecute=$false; $info.CreateNoWindow=$true
    $info.RedirectStandardOutput=$true; $info.RedirectStandardError=$true
    foreach ($arg in @('-NoProfile','-NonInteractive','-File',$probe,'-Path',$Path,'-MarkerName',$MarkerName)) { $info.ArgumentList.Add($arg) }
    if ($ExpectedSid) { $info.ArgumentList.Add('-ExpectedSid');$info.ArgumentList.Add($ExpectedSid) }
    $proc=[Diagnostics.Process]::new(); $proc.StartInfo=$info
    $marker=Join-Path $Path $MarkerName
    try {
        if (-not $proc.Start()) { throw 'Unable to start shared path probe.' }
        $stdout=$proc.StandardOutput.ReadToEndAsync(); $stderr=$proc.StandardError.ReadToEndAsync()
        if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
            try {$proc.Kill($true);$stopped=$proc.WaitForExit(3000)} catch {$stopped=$false}
            return [pscustomobject]@{State='Timeout';MarkerPath=$marker;Detail=$(if($stopped){'Timed out; result and cleanup are uncertain. Inspect the indicated marker before retrying.'}else{'Timed out; process termination, result and cleanup are uncertain. Inspect the indicated marker before retrying.'})}
        }
        $streams=Read-SiBoundedProcessStreams -Stdout $stdout -Stderr $stderr; $text=$streams.Stdout
        try { $result=$text | ConvertFrom-Json -ErrorAction Stop } catch { throw 'The write probe returned invalid output; inspect the marker before retrying.' }
        if ($proc.ExitCode -eq 0 -and $result.State -eq 'WriteVerified') {
            if ($ExpectedSid -and $result.ExpectedIdentity -ne $true) {throw 'The write probe did not confirm the selected identity.'}
            if ($ExpectedSid) { return [pscustomobject]@{State='WriteVerifiedAsExpected';MarkerPath=$marker;Detail='Write, readback and marker removal observed with the selected SID. The token is not proven to come from the orchestrator task; task group membership and effective access remain unverified.'} }
            return [pscustomobject]@{State='WriteVerified';MarkerPath=$marker;Detail='Write, readback and marker removal observed as the current account. Service-account access not verified.'}
        }
        if ($proc.ExitCode -eq 3 -and $result.State -eq 'IdentityMismatch' -and $ExpectedSid) {return [pscustomobject]@{State='IdentityMismatch';MarkerPath=$marker;Detail='The process SID differs from the selected identity. The folder was not contacted and no marker was created.'}}
        if ($result.State -in @('NotCreated','CleanupUncertain')) {
            return [pscustomobject]@{State='Uncertain';MarkerPath=$marker;Detail='Write or cleanup unconfirmed. Inspect the indicated marker; service-account rights remain unknown.'}
        }
        throw 'The write probe returned an unexpected result; inspect the marker before retrying.'
    } finally { $proc.Dispose();$probeLock.Dispose() }
}

function Get-SiOrchestratorTaskObservation {
    param([Parameter(Mandatory)][string]$Tenant,[string]$SmartM365Root='',[string[]]$AllowedExecutables=@(),[scriptblock]$ListTasks)
    if ($Tenant -cnotmatch '^[A-Za-z0-9][A-Za-z0-9_-]{0,63}\z') { throw 'Invalid profile key.' }
    if ($null -eq $ListTasks) { $ListTasks={ Get-ScheduledTask -ErrorAction Stop } }
    $name='SmartM365 Inventory Orchestrator - ' + $Tenant
    $matches=@(& $ListTasks | Where-Object { $_.TaskName -ieq $name -and $_.TaskPath -in @('\WCH\','\') })
    if ($matches.Count -eq 0) { return [pscustomobject]@{State='NotObserved'; Detail='Standard task not found; a custom task may exist.'} }
    if ($matches.Count -gt 1) { return [pscustomobject]@{State='Conflict'; Detail='Multiple standard tasks found. Review locations before making changes.'} }
    $task=$matches[0]
    $account=if ($null -ne $task.Principal) { [string]$task.Principal.UserId } else { '' }
    $detail=('Standard task found in {0}, configured account: {1}. ' -f $task.TaskPath,$account)
    if (-not $SmartM365Root) { return [pscustomobject]@{State='Observed'; Detail=($detail + 'Action, triggers, rights and execution require verification.')} }
    if (-not [IO.Path]::IsPathFullyQualified($SmartM365Root)) { throw 'An absolute SmartM365 root is required.' }
    if ($AllowedExecutables.Count -eq 0) {
        foreach ($base in @($env:ProgramFiles,[Environment]::GetEnvironmentVariable('ProgramFiles(x86)'))) {
            if ($base) { $AllowedExecutables += Join-Path $base 'PowerShell/7/pwsh.exe' }
        }
    }
    $actions=@($task.Actions)
    $scriptPath=Join-Path $SmartM365Root 'SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.ps1'
    $expectedArgs='-NoProfile -ExecutionPolicy Bypass -File "{0}" -Tenant {1} -Connect' -f $scriptPath,$Tenant
    $matched=($actions.Count -eq 1 -and [string]$actions[0].Execute -iin $AllowedExecutables -and ([string]$actions[0].Arguments).Trim() -ieq $expectedArgs)
    if (-not $matched) { return [pscustomobject]@{State='ActionMismatch'; Detail=($detail + 'The action differs from the standard format for this copy; no changes made. Triggers, rights and execution require verification.')} }
    [pscustomobject]@{State='Observed'; Detail=($detail + 'The action matches the standard format for this copy. Triggers, rights and execution require verification.')}
}

function Get-SiCertificateObservation {
    param([Parameter(Mandatory)][string]$Root,[Parameter(Mandatory)][string]$Tenant,[scriptblock]$ListCertificates)
    if ($Tenant -notmatch '^[a-zA-Z0-9][a-zA-Z0-9_-]{0,63}$') { throw 'Invalid profile key.' }
    $profile=Read-SiPrivateJson (Join-Path $Root ("Config/Tenants/$Tenant.local"))
    if ($profile.State -ne 'Present') { throw 'Tenant configuration is missing, invalid or in conflict.' }
    $first=([string]$profile.Value.Thumb).Replace(' ','').ToUpperInvariant()
    $second=([string]$profile.Value.Thumbprint).Replace(' ','').ToUpperInvariant()
    if ($first -and $second -and $first -ne $second) { throw 'Configured certificate thumbprints differ; review the profile.' }
    $thumb=if ($first) { $first } else { $second }
    if ($thumb -notmatch '^[0-9A-F]{40}$' -or $thumb -match '^0+$') { throw 'No usable certificate thumbprint is configured.' }
    if ($null -eq $ListCertificates) { $ListCertificates={ Get-ChildItem Cert:\CurrentUser\My,Cert:\LocalMachine\My -ErrorAction Stop } }
    $matches=@(& $ListCertificates | Where-Object { $_.Thumbprint -ieq $thumb })
    if ($matches.Count -eq 0) { return [pscustomobject]@{State='NotObserved'; Detail='Certificate not found in the personal stores checked for this account.'} }
    if ($matches.Count -gt 1) { return [pscustomobject]@{State='Multiple'; Detail='Multiple certificates match; review stores and access before use.'} }
    $cert=$matches[0]
    $expired=([datetime]$cert.NotAfter -le (Get-Date))
    $notYetValid=([datetime]$cert.NotBefore -gt (Get-Date))
    $detail='Local presence and dates observed; effective private-key access, service identity and application association not verified.'
    if ($expired -or $notYetValid) { return [pscustomobject]@{State='OutsideValidity'; Detail=$detail} }
    if (-not $cert.HasPrivateKey) { return [pscustomobject]@{State='NoPrivateKey'; Detail='No private key associated with this certificate in the inspected store.'} }
    return [pscustomobject]@{State='Observed'; Detail=$detail}
}
. (Join-Path $PSScriptRoot 'SmartM365.BoundedStreams.ps1')

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC+dWBVdOvyfHY5
# afRKADqypkVSxr/xUvH2xdKSOAjNo6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMKdY0mfcPRT/efuVbzBr8P1lCA1VN4n1iDfC8smGBF+MA0GCSqG
# SIb3DQEBAQUABIIBgJOcTy40dVMsMMv2GHoZYWOAslprD64EUAzYWv38pzLZ6suO
# 7I3JS+zntghsKctezJ0uDYfPq0oRWG8vosnAYOeT2r6HZtvhpreMfj5Ob4obZxeM
# 8+/jyHP0Hw6Lro5rvzBrteFTkaQStia5kw0Sw40Phcs91EnXY8HTj1cknEJEsUOT
# w06oF/7URMFThKEvWc0bYmbobbRo9bamIRf9RBxicWvyJg2YSxG6fAxJMDH6ZgoI
# biR9AMdAE5F8M7uJRj41dxUdXwe5JryONm1f8/fBZvKjcmvHE1JrOxPd4qxlXYNf
# RtBtWIK7+LV6zIOUGHPpAkM6Tpj3wVkNCYBy1cS6H/LDJXYpkWYTbkdMWUFovlE5
# vSqXzTd/5FeOF8cld4GcKc7BzsbnDutuOHrjYlwcflFvKkBOetjLIXGaeFvZ6oX4
# EL+/4S+uYKwa3g8NovGgw8WnOpDCnXRq1q9bpj6y6odZoB3mK/3g9tFakhzKCm3B
# QlrEUC5uAIzXq3WZpqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MjZaMC8GCSqGSIb3DQEJBDEiBCDjtgmitJbZibZV5uFsHxARFkNH7NlWoCf5zIPX
# crQKoTANBgkqhkiG9w0BAQEFAASCAgBDlazdScWi/aZfwrcOvCAE8P2++qYQbgSa
# czOxWeoYQk22iit1r9E/YMc6S8igS7wBkErS7k5oVK4pFQryNqsjKOduQfIhAmZZ
# wQyW+8l36Gfit/gdurXhMbFaX08AWu6l5OdicI3UJP7BR0MowxcbhBq/blnE3VV2
# mf5lYv6Y2p0wZ7dAC2Sv5WSgfgXxsJMrleh0ILsLn8H2g0zXQf3IrVH3bMQZLd3a
# CxjbW2YLIDCCiVOxstY7okqThWesnYvY8kpi+LvgImoIuWKuPYiYvPqHRcTrVz5G
# hbAazRPvZbf7H2iFSshTGTTXN1Kbfkt+WQLahWHuhwUXwlVedCHey831ITPc0Dnb
# RARGTBhow0KPANiV9cLzmRAqeDGYW1bqUbfGJJUcIDqINYre7S0XlJf0dJ1GuTRa
# SJjibKPoPX2BDkUf8vfeeiiUe+XbjgDaRxAuueQXRatj+Xy5erP10utBLd/28RP/
# eii4zZ2a37BXWwEIVNYQKUor2G8fEqmxrjfQtO79dWPrfP13v3CryxwHOVSHA9g6
# bl/XP9zV29uF10RcO5yzlPdGOj8c6+Gu7aIcfw8hGN46lOhnAZMJrMX6sJgGRdwG
# sXpByrFjZY/ySvDxtTyRKRnexS4dHWGKpmjxCwrFj1IQtG74XNUpda/nKjo/0P9C
# VJwozwDvrA==
# SIG # End signature block
