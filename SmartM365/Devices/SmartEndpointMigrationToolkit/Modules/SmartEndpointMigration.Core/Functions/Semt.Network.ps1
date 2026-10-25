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
