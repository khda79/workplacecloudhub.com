<#
.SYNOPSIS
    DNS, administrative access and PsExec helpers.

.DESCRIPTION
    DNS preflight on a random sample, connection target resolution with an optional domain
    suffix, administrative-share preflight, PsExec binary security evidence and per-argument
    quoting of the remote PowerShell command line.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

function Get-SemtDnsCandidates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [AllowNull()][string]$DomainSuffix
    )

    $name = ConvertTo-SemtComputerName -Line $ComputerName
    $candidates = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($name)) { $candidates.Add($name) }
    $suffix = ([string]$DomainSuffix).Trim().Trim('.')
    if ($name -notmatch '\.' -and -not [string]::IsNullOrWhiteSpace($suffix)) { $candidates.Add(('{0}.{1}' -f $name,$suffix)) }
    return @($candidates | Select-Object -Unique)
}

function Resolve-SemtConnectionTarget {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [AllowNull()][string]$DomainSuffix
    )

    $lastError = ''
    $candidates = @(Get-SemtDnsCandidates -ComputerName $ComputerName -DomainSuffix $DomainSuffix)
    foreach ($candidate in $candidates) {
        try {
            $entry = [System.Net.Dns]::GetHostEntry($candidate)
            return [pscustomobject]@{ Computer = $ComputerName; ConnectionTarget = $candidate; Resolved = $true; AddressList = (($entry.AddressList | ForEach-Object { $_.IPAddressToString }) -join ';'); Error = '' }
        }
        catch { $lastError = $_.Exception.Message }
    }
    $fallback = if ($candidates.Count -gt 0) { $candidates[-1] } else { $ComputerName }
    return [pscustomobject]@{ Computer = $ComputerName; ConnectionTarget = $fallback; Resolved = $false; AddressList = ''; Error = $lastError }
}

function Test-SemtDnsSample {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$ComputerNames,
        [AllowNull()][string]$DomainSuffix,
        [ValidateRange(1, 50)][int]$SampleSize = 5
    )

    $names = @($ComputerNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $sample = if ($names.Count -le $SampleSize) { @($names) } else { @($names | Get-Random -Count $SampleSize) }
    $results = @($sample | ForEach-Object { Resolve-SemtConnectionTarget -ComputerName $_ -DomainSuffix $DomainSuffix })
    $failed = @($results | Where-Object { -not $_.Resolved }).Count
    return [pscustomobject]@{
        Tested = $results.Count
        Resolved = $results.Count - $failed
        Failed = $failed
        AllFailed = ($results.Count -gt 0 -and $failed -eq $results.Count)
        Samples = $results
    }
}

function Resolve-SemtPsExecPath {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Path,
        [Parameter(Mandatory = $true)][string]$PackageRoot
    )

    $candidate = ([string]$Path).Trim()
    if (-not [string]::IsNullOrWhiteSpace($candidate)) {
        $embedded = @([regex]::Matches($candidate, '(?i)(?:[A-Z]:\\|\\\\)[^"<>|?*]+?PsExec(?:64)?\.exe') | ForEach-Object { $_.Value.Trim() })
        if ($embedded.Count -gt 1) { throw ("SEMT_PSEXEC_PATH_INVALID: several PsExec paths were provided: {0}" -f ($embedded -join ' | ')) }
        if ($embedded.Count -eq 1) { $candidate = $embedded[0] }
        $candidate = $candidate.Trim('"')
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return (Get-Item -LiteralPath $candidate).FullName }
        $command = Get-Command -Name $candidate -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($command) { return [string]$command.Source }
        throw ("SEMT_PSEXEC_PATH_INVALID: '{0}' was not found. Provide exactly one executable path." -f $Path)
    }
    foreach ($local in @((Join-Path -Path $PackageRoot -ChildPath 'Tools\PsExec.exe'),(Join-Path -Path $PackageRoot -ChildPath 'Scripts\PsExec.exe'),(Join-Path -Path $env:SystemRoot -ChildPath 'System32\PsExec.exe'))) {
        if (Test-Path -LiteralPath $local -PathType Leaf) { return (Get-Item -LiteralPath $local).FullName }
    }
    foreach ($name in @('PsExec.exe','PsExec64.exe')) {
        $command = Get-Command -Name $name -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($command) { return [string]$command.Source }
    }
    throw ("SEMT_PSEXEC_NOT_FOUND: place PsExec.exe in {0}\Tools, in %WINDIR%\System32 or in PATH, or pass -PsExecPath." -f $PackageRoot)
}

function Get-SemtPsExecSecurityEvidence {
    [CmdletBinding()]
    param([AllowNull()][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return [pscustomobject]@{ CheckStatus = 'SkippedDryRun'; Path = ''; FileName = ''; LengthBytes = ''; SHA256 = ''; SignatureStatus = ''; SignerSubject = ''; SignerThumbprint = ''; FileVersion = ''; IsMicrosoftSigned = $true; IsExpectedFileName = $true; IsCompliant = $true; Detail = 'Dry run: PsExec is not required.' }
    }
    $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    $signature = Get-AuthenticodeSignature -LiteralPath $item.FullName -ErrorAction Stop
    $subject = if ($signature.SignerCertificate) { [string]$signature.SignerCertificate.Subject } else { '' }
    $thumbprint = if ($signature.SignerCertificate) { [string]$signature.SignerCertificate.Thumbprint } else { '' }
    $microsoft = ([string]$signature.Status -eq 'Valid' -and $subject -match '(^|,\s*)(CN|O)=Microsoft Corporation(,|$)')
    $expectedName = ($item.Name -match '^(?i:PsExec|PsExec64)\.exe$')
    $compliant = ($microsoft -and $expectedName)
    return [pscustomobject]@{
        CheckStatus = if ($compliant) { 'Compliant' } else { 'Blocked' }
        Path = [string]$item.FullName
        FileName = [string]$item.Name
        LengthBytes = [string]$item.Length
        SHA256 = Get-SemtFileSha256 -Path $item.FullName
        SignatureStatus = [string]$signature.Status
        SignerSubject = $subject
        SignerThumbprint = $thumbprint
        FileVersion = [string]$item.VersionInfo.FileVersion
        IsMicrosoftSigned = [bool]$microsoft
        IsExpectedFileName = [bool]$expectedName
        IsCompliant = [bool]$compliant
        Detail = if ($compliant) { 'Microsoft-signed PsExec binary accepted.' } else { 'Blocked: PsExec must be named PsExec.exe or PsExec64.exe and carry a valid Microsoft Corporation Authenticode signature.' }
    }
}

function Assert-SemtPsExecSecurityEvidence {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Evidence)

    if (-not [bool]$Evidence.IsCompliant) {
        throw ("SEMT_PSEXEC_BLOCKED: Status={0}; Path={1}; Signature={2}; Signer={3}; SHA256={4}; {5}" -f $Evidence.CheckStatus,$Evidence.Path,$Evidence.SignatureStatus,$Evidence.SignerSubject,$Evidence.SHA256,$Evidence.Detail)
    }
}

function ConvertTo-SemtAdminSharePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][string]$LocalPath
    )

    $full = [System.IO.Path]::GetFullPath($LocalPath)
    if ($full -notmatch '^[A-Za-z]:\\') { throw "Only local drive paths can be converted to administrative share paths: $LocalPath" }
    return ('\\{0}\{1}${2}' -f $ComputerName,$full.Substring(0, 1),$full.Substring(2))
}

function Test-SemtTcpPort {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][int]$Port,
        [int]$TimeoutMilliseconds = 3000
    )

    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $async = $client.BeginConnect($ComputerName, $Port, $null, $null)
        if (-not $async.AsyncWaitHandle.WaitOne($TimeoutMilliseconds, $false)) { return $false }
        $client.EndConnect($async)
        return $true
    }
    catch { return $false }
    finally { try { $client.Close() } catch { } }
}

function Test-SemtRemoteAdminAccess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ConnectionTarget,
        [string]$LogPath
    )

    $state = [ordered]@{ DnsResolved = $false; DnsAddressList = ''; PingReachable = $false; SmbPort445Reachable = $false; AdminShareReachable = $false; FailureType = ''; Detail = '' }
    try {
        $entry = [System.Net.Dns]::GetHostEntry($ConnectionTarget)
        $state.DnsResolved = $true
        $state.DnsAddressList = (($entry.AddressList | ForEach-Object { $_.IPAddressToString }) -join ';')
    }
    catch { if ($LogPath) { Write-SemtLogLine -Path $LogPath -Level WARN -Message ("DNS resolution failed: {0}" -f $_.Exception.Message) } }
    $state.PingReachable = [bool](Test-Connection -ComputerName $ConnectionTarget -Count 1 -Quiet -ErrorAction SilentlyContinue)
    $state.SmbPort445Reachable = Test-SemtTcpPort -ComputerName $ConnectionTarget -Port 445
    $adminShare = Test-Path -LiteralPath ('\\{0}\ADMIN$' -f $ConnectionTarget) -ErrorAction SilentlyContinue
    $rootShare = Test-Path -LiteralPath ('\\{0}\C$' -f $ConnectionTarget) -ErrorAction SilentlyContinue
    $state.AdminShareReachable = [bool]($adminShare -and $rootShare)
    if (-not $state.AdminShareReachable) {
        $state.FailureType = if (-not $state.DnsResolved) { 'DNS_FAILED' } elseif (-not $state.SmbPort445Reachable) { 'SMB_PORT_445_UNREACHABLE' } elseif (-not $state.PingReachable) { 'PING_FAILED_ADMIN_SHARE_FAILED' } else { 'PING_OK_ADMIN_SHARE_FAILED' }
    }
    $state.Detail = ('DNS={0}; Addresses={1}; Ping={2}; Tcp445={3}; ADMIN$={4}; C$={5}; FailureType={6}' -f $state.DnsResolved,$state.DnsAddressList,$state.PingReachable,$state.SmbPort445Reachable,[bool]$adminShare,[bool]$rootShare,$state.FailureType)
    if ($LogPath) { Write-SemtLogLine -Path $LogPath -Message ("Administrative access preflight: {0}" -f $state.Detail) }
    return [pscustomobject]$state
}

function ConvertTo-SemtPsExecArgument {
    [CmdletBinding()]
    param([AllowNull()][string]$Value)

    if ($null -eq $Value -or $Value.Length -eq 0) { return '""' }
    if ($Value -match '^[A-Za-z0-9_:\\./@=-]+$') { return $Value }
    return ('"{0}"' -f ($Value -replace '"', '\"'))
}

function Get-SemtPsExecCommunicationFailureDetail {
    [CmdletBinding()]
    param([AllowEmptyCollection()][string[]]$Lines)

    $evidence = @($Lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' | '
    if ([string]::IsNullOrWhiteSpace($evidence)) { return '' }
    if ($evidence -match 'Error communicating with PsExec service on (?<Computer>[^:]+):\s*(?<Error>[^|\r\n]+)') { return ('PsExec communication lost after launch. Computer={0}; Error={1}' -f $Matches.Computer.Trim(),$Matches.Error.Trim()) }
    if ($evidence -match 'Descripteur non valide|The handle is invalid') { return 'PsExec communication lost after launch. Error=Invalid handle from the PsExec service.' }
    if ($evidence -match 'Erreur r.seau inattendue|unexpected network error') { return 'PsExec communication lost after launch. Error=Unexpected network error.' }
    if ($evidence -match 'Le chemin r.seau n.a pas .t. trouv.|network path was not found') { return 'PsExec communication lost after launch. Error=Network path was not found.' }
    return ''
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD41Zal9QW5YFCH
# BYy7tZxY7/tN820KwBun8xampBrUPKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIF1/T+NSfl4FAd2QzkkxlH+iBCCbyA4GGBtE+Ox12Yz+MA0GCSqG
# SIb3DQEBAQUABIIBgIJYl7iAT1f1gFsp4FISOdxZw2FdGjHx2/gDMig/T1uFijPS
# Me8zpobMW4OMbGd74vkuM82Cp6QfY7eyE+l45gCIlgxvtQClNmuVzRWeCaUtl7pM
# nTL29wGLdGIvtmWKgVSRbSIYBEWmXecre5Ry5QA44Wi9D5MJ8fiUIng9wNCnDDmf
# UWmpS4V0EwmkLJci1CUvoHB+41ZTOZbPFMVGjwmwSQzIVUiBj4XoRMx5sdL5qCXr
# zNba/LcwuGEG1AJLn/eyv8pcaWtRof5i+3i1CCenPv4MQQadPRo8n02isjxUocox
# S5yAM3wT9EUFXgibedpNjWQuTUpiSdgJB4Keh49qE7RqJxFdkk2NJSaDErz4rLO1
# 0h+U9pIogGDNBOJLDaDl/vEqPIWUOypOBNczEBIjiOIewitr96XNkg2wIaQCstbc
# jVISXWum9lUdExS0yW56quL7HV5gx+CGcPagz/j9th8a7mrYhCaFhZpavjXDZzNY
# eZcNaMljgfIsHd6x9qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMzE2
# NTJaMC8GCSqGSIb3DQEJBDEiBCCD0thdBRK7m/0zjnzDSpASHgkS1mlZA/WhoigO
# wSoPnjANBgkqhkiG9w0BAQEFAASCAgCvgr/H6AgGp9AZ42yck8v5MssheWOvh0kD
# 6LKbhNIbHxrl2VVLG13C1Sy9VGXXP51Xi9iZA27Y+UG91+IAm4KVqVHEb8htjT1Y
# PhcV+eB95OogIp75MiXA5ia72wyCvmWM1hpqzkIBvMj5r0DSHrs41AP44T0a4uLT
# NQJhEHwxU9czAjpXcYYYN7UHb8x3hSmSWTAU9okJYIiFr+6MEc2pRJFQT9/CoWIH
# +1MfKJ+6qQlEWtpTZphsgQSOAbvmZKFHpBhJu2LM9ewCv31pivC1fcv+Squ+4kwj
# YzfIZcWHZkYflSJ4ctDZapt125dq//rKSEyyqUsdvdRfc4FJJ/MObd81dR3KLi/I
# 96+Lydyku4TzZq/sAeywzh22LohmW5Jy4fFgQxfl2mQ0xgbOti1Idl9dXfRlb0p0
# jIk+PjB72AvKQkpjeHn3E1+chuATa/EDaq+S6HUkYQ4uXURvoZeEPzBmyGJ9igp2
# ugBffybh/ScvmJHute23U431MgkFwgivjYV8OVD+Iurc9BGokKXwhEUjPXAYM4kN
# RyvRC1V43s0SmyH/hg6C83Jiq7y3Ak0Rj5+AwFt3ppny0bu/vKDv4Qa91o8tDuWR
# CoeKrb+5i/CpW31tosnWV9edoDhPNzhDbekj8+Q8WFmYAH1cQeO+TM9b4nwpLSfP
# voufDPbGZA==
# SIG # End signature block
