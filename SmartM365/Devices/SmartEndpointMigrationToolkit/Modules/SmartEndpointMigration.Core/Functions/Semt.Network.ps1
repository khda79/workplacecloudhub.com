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

function ConvertTo-SemtWindowsCommandLineArgument {
    <# Quote one argument according to the Windows command-line parsing rules. #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Value)

    if ($null -eq $Value) { $Value = '' }
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('"')
    $slashes = 0
    foreach ($character in $Value.ToCharArray()) {
        if ($character -eq [char]92) { $slashes++; continue }
        if ($character -eq [char]34) {
            [void]$builder.Append(([string][char]92 * (2 * $slashes + 1)))
        } else {
            [void]$builder.Append(([string][char]92 * $slashes))
        }
        [void]$builder.Append($character)
        $slashes = 0
    }
    [void]$builder.Append(([string][char]92 * (2 * $slashes)))
    [void]$builder.Append('"')
    return $builder.ToString()
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB9yD0K1NURw8Qw
# K87Vk6JSnP3KpZWiLfyrsml6rfrktKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDtQ18VvQIuN5lDhHVr6ure
# tn8kAjGbVatZsIm+1K81hDANBgkqhkiG9w0BAQEFAASCAYAyun6Uhzvntj/14SM0
# 1rC7f9i7qx+a6gQJ3MqWqLR4ndXf8nPbsLdLAOQPez1H53NWCPPkFR+fqn/aY1DC
# JWNzm0zou3mqxoVh2GwT+1r/vV5a3JhSeN41PaL8mui14xjB2lx0Slg5iMi/h//e
# 2/ylOMtWRG3lgZ7+w9hPGNyYAaBdR5dgsm32ENufM3SHkiMWZDTOgJFIb5uTl9Ai
# OJTa/IdVckjHVwi3dJI8QPyVyfB3DlekZehylySFT1hmZ7DsaEMXIp0feF1FsKG8
# hFcvZiICOZ93EQEP6kNLqRe7Ldv2MM8HDFHpi/QOy29TSwkjYDDgroANXmgEJEV9
# MbEGs6dXslrNwxat7PjWuaCt0Ag7Z2Z5f+Qq6Bu+k3z15kL9AOR9ZTMrC3bytR7I
# AVSOXWSyit5eowhLZLfPRSjEYd2Q5ojtjfp4ZSQRROjK7coVSVfs1Qtk6esgU3nt
# 0BUYTTLb9l7Ucsix9vpkTyVA1D/kOSbbqBplyjx03uhpKco=
# SIG # End signature block
