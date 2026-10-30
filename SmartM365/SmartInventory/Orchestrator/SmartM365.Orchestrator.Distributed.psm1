Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot '../../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -MinimumVersion '1.0.2' -Global -ErrorAction Stop

$script:DistributedModuleVersion = '1.1.24'

function Resolve-DistributedJsonPath {
    param([Parameter(Mandatory)][string]$Path)
    $names = Get-SmartM365JsonNames $Path
    $leaf = [IO.Path]::GetFileName($names.Legacy)
    $parent = [IO.Path]::GetFileName([IO.Path]::GetDirectoryName($names.Legacy))
    $existing = Get-SmartM365JsonReadPath -Path $Path -Optional
    $isLease = $false
    if ($existing) { $isLease = $null -ne (Read-SmartM365JsonDocument $Path).Document.PSObject.Properties['LeaseId'] }
    $owner = if ($isLease) { 'Orchestrator concurrency lease' } else { 'Orchestrator distributed state' }
    $compatibleOwners = if ($isLease) { @('Orchestrator distributed state') } else { @() }
    Resolve-SmartM365OwnedJsonPath -Path $Path -Owner $owner -CompatibleJournalOwners $compatibleOwners -Validate {
        param($document)
        if (-not $document.PSObject.Properties['JobName'] -or -not $document.PSObject.Properties['OwnerServer'] -or
            (-not $document.PSObject.Properties['ClaimId'] -and -not $document.PSObject.Properties['LeaseId'])) { throw 'Distributed state owner/schema mismatch.' }
        if ($document.PSObject.Properties['ClaimId']) {
            $expectedLeaf = ([datetime]$document.OccurrenceUtc).ToUniversalTime().ToString('yyyyMMddTHHmmssfffZ') + '.json'
            if ($parent -ne (ConvertTo-SafeFileName $document.JobName) -or ($leaf -ne $expectedLeaf -and $leaf -ne ($expectedLeaf + '.stale.' + $document.ClaimId + '.json'))) { throw 'Distributed claim filename/identity mismatch.' }
        } else {
            $expectedLeaf = (ConvertTo-SafeFileName $document.ConcurrencyKey) + '.json'
            if ($leaf -ne $expectedLeaf -and $leaf -ne ($expectedLeaf + '.stale.' + $document.LeaseId + '.json')) { throw 'Distributed lease filename/identity mismatch.' }
        }
    }
}

function Convert-SmartM365OrchestratorDistributedHistory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ClaimsRootPath,[Parameter(Mandatory)][string]$LeasesRootPath,[scriptblock]$OnProgress)
    if ((Get-SmartM365JsonTransportPolicy).Mode -ne 'JsonText') { return }
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $lastReport = 0.0
    $checked = 0; $skipped = 0; $resolved = 0; $vanished = 0
    if ($OnProgress) { & $OnProgress 'Distributed JSON history scan started.' | Out-Null }
    $entries = @{}
    $claimPaths = [Collections.Generic.List[string]]::new()
    $paths = @()
    if (Test-Path -LiteralPath $ClaimsRootPath) {
        foreach ($folder in @(Get-ChildItem -LiteralPath $ClaimsRootPath -Directory -ErrorAction Stop)) {
            if ($folder.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Distributed history folder is a reparse point.' }
            foreach ($file in @(Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction Stop)) {
                $entries[$file.FullName] = $file
                if ($file.Name -match '^\d{8}T\d{9}Z\.json(?:\.stale\.[0-9a-f]{32}\.json)?(?:\.txt)?$') { $claimPaths.Add($file.FullName) }
            }
            if ($OnProgress -and ($timer.Elapsed.TotalSeconds - $lastReport) -ge 5) {
                & $OnProgress ("Distributed JSON history listing: claims={0}; folder={1}." -f $claimPaths.Count,$folder.Name) | Out-Null
                $lastReport = $timer.Elapsed.TotalSeconds
            }
        }
    }
    foreach ($path in $claimPaths) {
        $checked++
        $completed = $false
        if ($path.EndsWith('.json.txt', [StringComparison]::OrdinalIgnoreCase)) {
            $legacy = $path.Substring(0, $path.Length - 4)
            $journal = $legacy + '.migration.log'
            # Only historical claims qualify. Live leases always use the locked resolver.
            # Readers still validate the current JSON whenever a claim is consumed.
            if (-not $entries.ContainsKey($legacy) -and -not ($entries[$path].Attributes -band [IO.FileAttributes]::ReparsePoint)) {
                if (-not $entries.ContainsKey($journal)) { $completed = $true }
                elseif (-not ($entries[$journal].Attributes -band [IO.FileAttributes]::ReparsePoint)) {
                    # One sequential receipt read replaces repeated JSON reads, hashes,
                    # policy loads and lock operations for an already completed migration.
                    $lastLine = $null
                    foreach ($line in [IO.File]::ReadLines($journal)) {
                        if (-not [string]::IsNullOrWhiteSpace($line)) { $lastLine = $line }
                    }
                    try {
                        $receipt = ConvertFrom-Json -InputObject $lastLine -ErrorAction Stop
                        $completed = $receipt.Phase -eq 'Completed' -and $receipt.Owner -eq 'Orchestrator distributed state'
                    } catch { $completed = $false }
                }
            }
        }
        if ($completed) { $skipped++ } else { $paths += $path }
        if ($OnProgress -and ($timer.Elapsed.TotalSeconds - $lastReport) -ge 5) {
            & $OnProgress ("Distributed JSON history scan: checked={0}/{1}; alreadyCurrent={2}; pending={3}." -f $checked,$claimPaths.Count,$skipped,$paths.Count) | Out-Null
            $lastReport = $timer.Elapsed.TotalSeconds
        }
    }
    if (Test-Path -LiteralPath $LeasesRootPath) {
        $paths += @(Get-ChildItem -LiteralPath $LeasesRootPath -File -ErrorAction Stop | Where-Object { $_.Name -match '\.json(?:\.txt)?$' } | ForEach-Object { $_.FullName })
    }
    foreach ($path in @($paths | ForEach-Object { (Get-SmartM365JsonNames $_).Legacy } | Sort-Object -Unique)) {
        try { $null = Resolve-DistributedJsonPath $path }
        catch {
            # A peer can release a live lease (or retention can purge a claim) between the listing
            # and this read. A file that no longer exists has nothing left to migrate; any other
            # failure, or a file still present, stays fatal for the owner.
            if (Get-SmartM365JsonReadPath -Path $path -Optional) { throw }
            $vanished++
            if ($OnProgress) { & $OnProgress ("Distributed JSON history recovery: {0} disappeared after the listing (released by a peer); skipped." -f [IO.Path]::GetFileName($path)) | Out-Null }
            continue
        }
        $resolved++
        if ($OnProgress -and ($timer.Elapsed.TotalSeconds - $lastReport) -ge 5) {
            & $OnProgress ("Distributed JSON history recovery: resolved={0}; alreadyCurrent={1}; current={2}." -f $resolved,$skipped,[IO.Path]::GetFileName($path)) | Out-Null
            $lastReport = $timer.Elapsed.TotalSeconds
        }
    }
    if ($OnProgress) { & $OnProgress ("Distributed JSON history scan complete: checked={0}; alreadyCurrent={1}; resolved={2}; vanished={3}; seconds={4:N1}." -f $checked,$skipped,$resolved,$vanished,$timer.Elapsed.TotalSeconds) | Out-Null }
}

function ConvertTo-SafeFileName {
    param([Parameter(Mandatory = $true)][string]$Value)

    $safe = $Value
    foreach ($character in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$character, '_')
    }
    return $safe
}

function Write-JsonAtomically {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Value,
        [int]$Depth = 12,
        [string]$ExpectedSHA256 = ''
    )

    $parent = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $parent -Force -ErrorAction Stop)
    }

    $Path = Resolve-DistributedJsonPath $Path
    $json = $Value | ConvertTo-Json -Depth $Depth
    $null = Write-SmartM365JsonBytesAtomically -Path $Path -Bytes ([Text.UTF8Encoding]::new($false).GetBytes($json)) -ExpectedSHA256 $ExpectedSHA256 -Validate {
        param($document) if ($document -isnot [pscustomobject]) { throw 'Distributed state must be an object.' }
    }
}

$script:CertificateCapabilities = @('Graph', 'EXO', 'TeamsPowerShell')

function Get-MaskedThumbprint {
    param([string]$Thumbprint)

    if ($Thumbprint.Length -le 8) { return $Thumbprint }
    return '{0}...{1}' -f $Thumbprint.Substring(0, 4), $Thumbprint.Substring($Thumbprint.Length - 4)
}

function Get-SmartM365OrchestratorCertificateCandidate {
    param(
        [Parameter(Mandatory = $true)][string]$Thumbprint,
        [string[]]$StorePaths = @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My')
    )

    $normalized = ($Thumbprint -replace '\s', '').ToUpperInvariant()
    foreach ($storePath in $StorePaths) {
        foreach ($certificate in @(Get-ChildItem -LiteralPath $storePath -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -eq $normalized })) {
            [pscustomobject]@{ StorePath = $storePath; Certificate = $certificate }
        }
    }
}

function Test-CertificatePrivateKeyAccess {
    param([Parameter(Mandatory = $true)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate)

    if (-not $Certificate.HasPrivateKey) {
        return [pscustomobject]@{ Ready = $false; Reason = 'NoPrivateKey'; Detail = 'the certificate has no associated private key' }
    }
    # A real signature proves that this account can open the key, not only that a key is linked.
    $data = [Text.Encoding]::UTF8.GetBytes('SmartM365 orchestrator launch prerequisite')
    try {
        $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
        if ($null -ne $rsa) {
            try { $null = $rsa.SignData($data, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1) }
            finally { $rsa.Dispose() }
            return [pscustomobject]@{ Ready = $true; Reason = 'Ready'; Detail = 'private key usable' }
        }
        $ecdsa = [System.Security.Cryptography.X509Certificates.ECDsaCertificateExtensions]::GetECDsaPrivateKey($Certificate)
        if ($null -ne $ecdsa) {
            try { $null = $ecdsa.SignData($data, [System.Security.Cryptography.HashAlgorithmName]::SHA256) }
            finally { $ecdsa.Dispose() }
            return [pscustomobject]@{ Ready = $true; Reason = 'Ready'; Detail = 'private key usable' }
        }
        return [pscustomobject]@{ Ready = $false; Reason = 'NoPrivateKey'; Detail = 'the private key algorithm is not supported' }
    }
    catch {
        $message = if ($null -ne $_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $_.Exception.Message }
        return [pscustomobject]@{ Ready = $false; Reason = 'PrivateKeyInaccessible'; Detail = ("the private key cannot be used by {0}: {1}" -f [Environment]::UserName, $message) }
    }
}

function Get-CertificateReadiness {
    param(
        [string]$Thumbprint,
        [int]$ExpiryWarningDays = 0,
        [datetime]$NowUtc = [datetime]::UtcNow,
        [AllowNull()][object[]]$Candidates = $null
    )

    $result = [ordered]@{
        Ready = $false
        Reason = 'NotConfigured'
        Detail = 'Certificate thumbprint is not configured.'
        StorePath = ''
        NotAfterUtc = ''
        DaysRemaining = $null
        ExpiresSoon = $false
    }
    if ([string]::IsNullOrWhiteSpace($Thumbprint)) { return [pscustomobject]$result }

    $normalized = ($Thumbprint -replace '\s', '').ToUpperInvariant()
    $masked = Get-MaskedThumbprint -Thumbprint $normalized
    if ($null -eq $Candidates) { $Candidates = @(Get-SmartM365OrchestratorCertificateCandidate -Thumbprint $normalized) }
    $Candidates = @($Candidates | Where-Object { $null -ne $_ -and $null -ne $_.Certificate })
    if ($Candidates.Count -eq 0) {
        $result.Reason = 'NotFound'
        $result.Detail = "Certificate $masked was not found in Cert:\CurrentUser\My or Cert:\LocalMachine\My for account $([Environment]::UserName)."
        return [pscustomobject]$result
    }

    $valid = @($Candidates | Where-Object {
            $_.Certificate.NotAfter.ToUniversalTime() -gt $NowUtc -and $_.Certificate.NotBefore.ToUniversalTime() -le $NowUtc
        } | Sort-Object { $_.Certificate.NotAfter } -Descending)
    if ($valid.Count -eq 0) {
        $latest = @($Candidates | Sort-Object { $_.Certificate.NotAfter } -Descending)[0]
        $notAfterUtc = $latest.Certificate.NotAfter.ToUniversalTime()
        $result.Reason = 'Expired'
        $result.StorePath = [string]$latest.StorePath
        $result.NotAfterUtc = $notAfterUtc.ToString('o')
        $result.DaysRemaining = [math]::Floor(($notAfterUtc - $NowUtc).TotalDays)
        $result.Detail = if ($notAfterUtc -le $NowUtc) {
            "Certificate $masked in $($latest.StorePath) expired on $($notAfterUtc.ToString('yyyy-MM-dd HH:mm')) UTC."
        } else {
            "Certificate $masked in $($latest.StorePath) is not valid yet."
        }
        return [pscustomobject]$result
    }

    $lastFailure = $null
    foreach ($candidate in $valid) {
        $keyAccess = Test-CertificatePrivateKeyAccess -Certificate $candidate.Certificate
        $notAfterUtc = $candidate.Certificate.NotAfter.ToUniversalTime()
        $result.StorePath = [string]$candidate.StorePath
        $result.NotAfterUtc = $notAfterUtc.ToString('o')
        $result.DaysRemaining = [math]::Floor(($notAfterUtc - $NowUtc).TotalDays)
        if ($keyAccess.Ready) {
            $result.Ready = $true
            $result.Reason = 'Ready'
            $result.ExpiresSoon = $ExpiryWarningDays -gt 0 -and $notAfterUtc -le $NowUtc.AddDays($ExpiryWarningDays)
            $result.Detail = "Usable certificate $masked found in $($candidate.StorePath); private key usable; expires $($notAfterUtc.ToString('yyyy-MM-dd HH:mm')) UTC ($($result.DaysRemaining) day(s))."
            return [pscustomobject]$result
        }
        $lastFailure = $keyAccess
    }
    $result.Reason = $lastFailure.Reason
    $result.Detail = "Certificate $masked in $($result.StorePath) is not usable: $($lastFailure.Detail)."
    return [pscustomobject]$result
}

$script:PathAccessProbeScript = @'
param($Checks, $ServerName, $Results)
for ($index = 0; $index -lt $Checks.Count; $index++) {
    $check = $Checks[$index]
    $path = [string]$check.Path
    try {
        if ([string]::IsNullOrWhiteSpace($path)) { throw 'The path is not configured.' }
        if ([string]$check.PathType -eq 'Leaf') {
            if (-not [IO.File]::Exists($path)) { throw "File not found: $path" }
            $stream = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
            try { $null = $stream.Read([byte[]]::new(1), 0, 1) } finally { $stream.Dispose() }
        }
        else {
            if (-not [IO.Directory]::Exists($path)) {
                if ([string]$check.Access -ne 'ReadWrite') { throw "Folder not found: $path" }
                $null = [IO.Directory]::CreateDirectory($path)
            }
            $enumerator = [IO.Directory]::EnumerateFileSystemEntries($path).GetEnumerator()
            try { $null = $enumerator.MoveNext() } finally { $enumerator.Dispose() }
            if ([string]$check.Access -eq 'ReadWrite') {
                $probe = [IO.Path]::Combine($path, ('.smartm365-prerequisite.{0}.{1}.tmp' -f $ServerName, [guid]::NewGuid().ToString('N')))
                [IO.File]::WriteAllText($probe, [datetime]::UtcNow.ToString('o'))
                [IO.File]::Delete($probe)
            }
        }
        $Results[$index] = [pscustomobject]@{ Ready = $true; Detail = ('{0} access verified.' -f $check.Access) }
    }
    catch {
        $exception = $_.Exception
        while ($null -ne $exception.InnerException) { $exception = $exception.InnerException }
        $Results[$index] = [pscustomobject]@{ Ready = $false; Detail = $exception.Message }
    }
}
'@

function Test-SmartM365OrchestratorPathAccess {
    <# Checks folders or files with a hard time limit: an unreachable share must not stall a scheduler tick. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Checks,
        [string]$ServerName = $env:COMPUTERNAME,
        [int]$TimeoutSeconds = 10
    )

    if ($Checks.Count -eq 0) { return @() }
    $results = [System.Collections.Concurrent.ConcurrentDictionary[int, object]]::new()
    $powerShell = [powershell]::Create()
    $completed = $false
    try {
        $null = $powerShell.AddScript($script:PathAccessProbeScript).AddArgument(@($Checks)).AddArgument((ConvertTo-SafeFileName $ServerName)).AddArgument($results)
        $async = $powerShell.BeginInvoke()
        $completed = $async.AsyncWaitHandle.WaitOne([math]::Max(1, $TimeoutSeconds) * 1000)
        if ($completed) { try { $null = $powerShell.EndInvoke($async) } catch { Write-Verbose ("Path access probe ended with an error: {0}" -f $_.Exception.Message) } }
    }
    finally {
        # A hung SMB call cannot be interrupted: the probe is abandoned and finishes on its own.
        if ($completed) { $powerShell.Dispose() } else { $null = $powerShell.BeginStop($null, $null) }
    }

    for ($index = 0; $index -lt $Checks.Count; $index++) {
        $check = $Checks[$index]
        $probe = $null
        if (-not $results.TryGetValue($index, [ref]$probe)) {
            $probe = [pscustomobject]@{ Ready = $false; Detail = ("No answer within {0} second(s); the share is unreachable or hung." -f $TimeoutSeconds) }
        }
        [pscustomobject]@{
            Name = [string]$check.Name
            Path = [string]$check.Path
            Access = [string]$check.Access
            PathType = [string]$check.PathType
            Ready = [bool]$probe.Ready
            Detail = [string]$probe.Detail
        }
    }
}

$script:VolumeFreeSpaceProbeScript = @'
param($Path, $Answers)
try {
    [uint64]$available = 0; [uint64]$total = 0; [uint64]$volumeFree = 0
    # The API expects a folder; a UNC share root needs its trailing backslash.
    $directory = ([string]$Path).TrimEnd('\') + '\'
    if (-not [SmartM365Orchestrator.DiskFreeSpaceNative]::GetDiskFreeSpaceEx($directory, [ref]$available, [ref]$total, [ref]$volumeFree)) {
        $code = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        throw ('{0} (Win32 error {1})' -f ([ComponentModel.Win32Exception]::new($code)).Message, $code)
    }
    $Answers['Result'] = [pscustomobject]@{ Available = $available; Total = $total; VolumeFree = $volumeFree; Error = '' }
}
catch {
    $exception = $_.Exception
    while ($null -ne $exception.InnerException) { $exception = $exception.InnerException }
    $Answers['Result'] = [pscustomobject]@{ Error = $exception.Message }
}
'@

function Initialize-DiskFreeSpaceNative {
    # Compiled once per process; probe runspaces of the same process see the type.
    if ('SmartM365Orchestrator.DiskFreeSpaceNative' -as [type]) { return }
    Add-Type -Namespace 'SmartM365Orchestrator' -Name 'DiskFreeSpaceNative' -MemberDefinition @'
[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
[return: MarshalAs(UnmanagedType.Bool)]
public static extern bool GetDiskFreeSpaceEx(string lpDirectoryName, out ulong lpFreeBytesAvailable, out ulong lpTotalNumberOfBytes, out ulong lpTotalNumberOfFreeBytes);
'@
}

function Get-SmartM365OrchestratorVolumeFreeSpace {
    <# Free space of the volume behind a local or UNC folder, with a hard time limit: a hung share must not stall a scheduler tick.
       FreePercent uses the bytes the current account can write, so a share quota is honored. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Path,
        [int]$TimeoutSeconds = 10,
        # Tests only: replaces the native probe (same parameters: Path, Answers).
        [string]$ProbeScript = ''
    )

    $result = [ordered]@{
        Path = $Path; CheckedAtUtc = [datetime]::UtcNow.ToString('o'); Ready = $false
        FreePercent = $null; FreeGB = $null; TotalGB = $null; VolumeFreeGB = $null; Error = ''
    }
    if ([string]::IsNullOrWhiteSpace($Path)) { $result.Error = 'The path is not configured.'; return [pscustomobject]$result }
    if (-not $ProbeScript) {
        try { Initialize-DiskFreeSpaceNative }
        catch { $result.Error = 'Disk space API unavailable: ' + $_.Exception.Message; return [pscustomobject]$result }
    }

    $answers = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
    $powerShell = [powershell]::Create()
    $completed = $false
    try {
        $probeText = if ($ProbeScript) { $ProbeScript } else { $script:VolumeFreeSpaceProbeScript }
        $null = $powerShell.AddScript($probeText).AddArgument($Path).AddArgument($answers)
        $async = $powerShell.BeginInvoke()
        $completed = $async.AsyncWaitHandle.WaitOne([math]::Max(1, $TimeoutSeconds) * 1000)
        if ($completed) { try { $null = $powerShell.EndInvoke($async) } catch { Write-Verbose ("Disk space probe ended with an error: {0}" -f $_.Exception.Message) } }
    }
    finally {
        # A hung SMB call cannot be interrupted: the probe is abandoned and finishes on its own.
        if ($completed) { $powerShell.Dispose() } else { $null = $powerShell.BeginStop($null, $null) }
    }

    $probe = $null
    if (-not $answers.TryGetValue('Result', [ref]$probe)) {
        $result.Error = "No answer within $([math]::Max(1, $TimeoutSeconds)) second(s); the share is unreachable or hung."
    }
    elseif ([string]$probe.Error) { $result.Error = [string]$probe.Error }
    elseif ([double]$probe.Total -le 0) { $result.Error = 'The volume reported a total size of zero.' }
    else {
        $result.Ready = $true
        $result.FreePercent = [math]::Round([double]$probe.Available * 100 / [double]$probe.Total, 1)
        $result.FreeGB = [math]::Round([double]$probe.Available / 1GB, 1)
        $result.TotalGB = [math]::Round([double]$probe.Total / 1GB, 1)
        $result.VolumeFreeGB = [math]::Round([double]$probe.VolumeFree / 1GB, 1)
    }
    return [pscustomobject]$result
}

function Get-SmartM365OrchestratorDiskSpaceDecision {
    <# Launch decision with hysteresis: launches pause below BlockFreePercent and resume only at ResumeFreePercent.
       A failed measurement never pauses launches by itself, but never lifts an existing pause either. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Measurement,
        [AllowNull()]$PreviousState,
        [double]$BlockFreePercent = 10,
        [double]$ResumeFreePercent = 12,
        [double]$WarningFreePercent = 15,
        [datetime]$NowUtc = [datetime]::UtcNow
    )

    $previous = if ($PreviousState -is [System.Collections.IDictionary]) { [pscustomobject]$PreviousState } elseif ($null -ne $PreviousState) { $PreviousState } else { [pscustomobject]@{} }
    foreach ($name in @('State', 'Blocked')) { if (-not $previous.PSObject.Properties[$name]) { $previous | Add-Member -NotePropertyName $name -NotePropertyValue $(if ($name -eq 'Blocked') { $false } else { '' }) } }
    $wasBlocked = [bool]$previous.Blocked
    $nowText = $NowUtc.ToUniversalTime().ToString('o')
    # Figures stay in invariant culture: they end up in English logs, mails and the shared heartbeat.
    $inv = [Globalization.CultureInfo]::InvariantCulture
    $thresholdText = [string]::Format($inv, 'launches pause below {0} % and resume at {1} %', $BlockFreePercent, $ResumeFreePercent)
    $decision = [ordered]@{
        State = ''; Blocked = $false; Path = [string]$Measurement.Path; CheckedAtUtc = $nowText
        FreePercent = $null; FreeGB = $null; TotalGB = $null; MeasuredAtUtc = ''
        Error = [string]$Measurement.Error; Detail = ''; SinceUtc = $nowText; UnknownSinceUtc = ''
        BlockFreePercent = $BlockFreePercent; ResumeFreePercent = $ResumeFreePercent; WarningFreePercent = $WarningFreePercent
    }

    if (-not [bool]$Measurement.Ready) {
        if ($wasBlocked) {
            # Keep the last valid figures: only a real measurement may lift the pause.
            $decision.State = 'Low'; $decision.Blocked = $true
            foreach ($name in @('FreePercent', 'FreeGB', 'TotalGB', 'MeasuredAtUtc')) { if ($previous.PSObject.Properties[$name]) { $decision[$name] = $previous.$name } }
            $decision.Detail = [string]::Format($inv, 'DATA free space could not be measured ({0}); launches stay paused until a measurement reaches {1} %.', $Measurement.Error, $ResumeFreePercent)
        }
        else {
            $decision.State = 'Unknown'
            $decision.UnknownSinceUtc = if ([string]$previous.State -eq 'Unknown' -and $previous.PSObject.Properties['UnknownSinceUtc'] -and [string]$previous.UnknownSinceUtc) { [string]$previous.UnknownSinceUtc } else { $nowText }
            $decision.Detail = [string]::Format($inv, 'DATA free space could not be measured: {0}. Launches continue.', $Measurement.Error)
        }
    }
    else {
        $percent = [double]$Measurement.FreePercent
        $decision.FreePercent = $Measurement.FreePercent; $decision.FreeGB = $Measurement.FreeGB; $decision.TotalGB = $Measurement.TotalGB
        $decision.MeasuredAtUtc = $nowText
        $decision.Blocked = if ($wasBlocked) { $percent -lt $ResumeFreePercent } else { $percent -lt $BlockFreePercent }
        $decision.State = if ($decision.Blocked) { 'Low' } elseif ($percent -lt $WarningFreePercent) { 'Warning' } else { 'OK' }
        $figures = [string]::Format($inv, 'DATA free {0} % ({1} GiB of {2} GiB)', $Measurement.FreePercent, $Measurement.FreeGB, $Measurement.TotalGB)
        $decision.Detail = switch ($decision.State) {
            'Low' { '{0}; new launches paused, running jobs continue; {1}.' -f $figures, $thresholdText }
            'Warning' { [string]::Format($inv, '{0} is below the {1} % warning level; {2}.', $figures, $WarningFreePercent, $thresholdText) }
            default { '{0}.' -f $figures }
        }
    }
    if ([string]$previous.State -eq $decision.State -and $previous.PSObject.Properties['SinceUtc'] -and [string]$previous.SinceUtc) { $decision.SinceUtc = [string]$previous.SinceUtc }
    return [pscustomobject]$decision
}

function Test-PrerequisiteCacheEntry {
    param([hashtable]$Cache, [string]$Key, [int]$CacheMinutes, [datetime]$NowUtc)

    if ($CacheMinutes -le 0 -or -not $Cache.ContainsKey($Key)) { return $false }
    $age = ($NowUtc - [datetime]$Cache[$Key].CheckedAtUtc).TotalMinutes
    return ($age -ge 0 -and $age -lt $CacheMinutes)
}

function Get-SmartM365OrchestratorCachedCertificateReadiness {
    [CmdletBinding()]
    param(
        [string]$Thumbprint,
        [Parameter(Mandatory = $true)][hashtable]$Cache,
        [int]$CacheMinutes = 5,
        [int]$ExpiryWarningDays = 30,
        [datetime]$NowUtc = [datetime]::UtcNow,
        [scriptblock]$CertificateProvider = $null
    )

    $normalized = ([string]$Thumbprint -replace '\s', '').ToUpperInvariant()
    $key = 'Certificate|' + $normalized
    if (Test-PrerequisiteCacheEntry -Cache $Cache -Key $key -CacheMinutes $CacheMinutes -NowUtc $NowUtc) { return $Cache[$key].Result }
    $candidates = $null
    if ($null -ne $CertificateProvider -and $normalized) { $candidates = @(& $CertificateProvider $normalized) }
    $readiness = Get-CertificateReadiness -Thumbprint $normalized -ExpiryWarningDays $ExpiryWarningDays -NowUtc $NowUtc -Candidates $candidates
    $Cache[$key] = @{ CheckedAtUtc = $NowUtc; Result = $readiness }
    return $readiness
}

function Test-SmartM365OrchestratorLaunchPrerequisite {
    <#
    Fast pre-launch gate. The certificate is checked for jobs that need app-only cloud access;
    each path check is { Name; Path; Access = Read|ReadWrite; PathType = Container|Leaf; Capability }.
    A failure with a Capability affects every job needing it; without one it is specific to the job.
    #>
    [CmdletBinding()]
    param(
        [string[]]$RequiredCapabilities = @(),
        [string]$CertificateThumbprint = '',
        [AllowEmptyCollection()][object[]]$PathChecks = @(),
        [Parameter(Mandatory = $true)][hashtable]$Cache,
        [int]$CacheMinutes = 5,
        [int]$PathTimeoutSeconds = 10,
        [int]$CertificateExpiryWarningDays = 30,
        [string]$ServerName = $env:COMPUTERNAME,
        [datetime]$NowUtc = [datetime]::UtcNow,
        [scriptblock]$CertificateProvider = $null
    )

    $failures = [System.Collections.Generic.List[object]]::new()
    $certificate = $null
    if (@($RequiredCapabilities | Where-Object { $_ -in $script:CertificateCapabilities }).Count -gt 0) {
        $certificate = Get-SmartM365OrchestratorCachedCertificateReadiness -Thumbprint $CertificateThumbprint -Cache $Cache -CacheMinutes $CacheMinutes -ExpiryWarningDays $CertificateExpiryWarningDays -NowUtc $NowUtc -CertificateProvider $CertificateProvider
        if (-not $certificate.Ready) {
            $failures.Add([pscustomobject]@{
                Key = 'Certificate'; Kind = 'Certificate'; Name = 'Application certificate'; Path = ''; Access = ''; PathType = ''
                Capabilities = @($script:CertificateCapabilities); JobSpecific = $false; Detail = [string]$certificate.Detail
            })
        }
    }

    $entries = @(foreach ($check in @($PathChecks | Where-Object { $null -ne $_ })) {
            $pathType = if ([string]$check.PathType -eq 'Leaf') { 'Leaf' } else { 'Container' }
            $access = if ([string]$check.Access -eq 'ReadWrite') { 'ReadWrite' } else { 'Read' }
            [pscustomobject]@{
                Key = 'Path|{0}|{1}|{2}' -f $access, $pathType, ([string]$check.Path).TrimEnd('\').ToUpperInvariant()
                Name = [string]$check.Name; Path = [string]$check.Path; Access = $access; PathType = $pathType
                Capability = if ($check.PSObject.Properties['Capability']) { [string]$check.Capability } else { '' }
            }
        })
    $toProbe = @($entries | Where-Object { -not (Test-PrerequisiteCacheEntry -Cache $Cache -Key $_.Key -CacheMinutes $CacheMinutes -NowUtc $NowUtc) } | Group-Object Key | ForEach-Object { $_.Group[0] })
    if ($toProbe.Count -gt 0) {
        $probeResults = @(Test-SmartM365OrchestratorPathAccess -Checks $toProbe -ServerName $ServerName -TimeoutSeconds $PathTimeoutSeconds)
        for ($index = 0; $index -lt $toProbe.Count; $index++) {
            $Cache[$toProbe[$index].Key] = @{ CheckedAtUtc = $NowUtc; Result = $probeResults[$index] }
        }
    }
    foreach ($entry in $entries) {
        $probe = $Cache[$entry.Key].Result
        if ($probe.Ready) { continue }
        $failures.Add([pscustomobject]@{
            Key = $entry.Key; Kind = 'Path'; Name = $entry.Name; Path = $entry.Path; Access = $entry.Access; PathType = $entry.PathType
            Capabilities = @($entry.Capability | Where-Object { $_ }); JobSpecific = [string]::IsNullOrWhiteSpace($entry.Capability)
            Detail = ("{0} ({1}, {2}) is not accessible: {3}" -f $entry.Name, $entry.Path, $entry.Access, $probe.Detail)
        })
    }

    return [pscustomobject]@{
        Ready = ($failures.Count -eq 0)
        Failures = $failures.ToArray()
        Certificate = $certificate
    }
}

function Invoke-IsolatedCapabilityProbe {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$EnginePath,
        [Parameter(Mandatory = $true)][string]$ScriptText,
        [hashtable]$Environment = @{},
        [int]$TimeoutSeconds = 90
    )

    if (-not (Test-Path -LiteralPath $EnginePath -PathType Leaf)) {
        return [pscustomobject]@{ Name = $Name; Ready = $false; Roles = @(); Detail = "PowerShell engine not found: $EnginePath" }
    }

    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($ScriptText))
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $EnginePath
    $startInfo.Arguments = "-NoLogo -NoProfile -NonInteractive -EncodedCommand $encoded"
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    foreach ($key in $Environment.Keys) {
        $startInfo.Environment[[string]$key] = [string]$Environment[$key]
    }

    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    try {
        if (-not $process.Start()) {
            throw "Could not start $EnginePath."
        }
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit([math]::Max(1, $TimeoutSeconds) * 1000)) {
            try { $process.Kill($true) } catch { Write-Verbose ("Probe process kill failed: {0}" -f $_.Exception.Message) }
            return [pscustomobject]@{ Name = $Name; Ready = $false; Roles = @(); Detail = "Read-only probe timed out after $TimeoutSeconds seconds." }
        }
        $stdout = $stdoutTask.GetAwaiter().GetResult()
        $stderr = $stderrTask.GetAwaiter().GetResult()
        $match = [regex]::Match($stdout, '(?m)^SMARTM365_CAPABILITY_JSON:(?<Json>.+)$')
        if (-not $match.Success) {
            $detail = (($stderr + "`n" + $stdout).Trim() -replace '[\r\n]+', ' ')
            if ($detail.Length -gt 600) { $detail = $detail.Substring(0, 600) }
            if ([string]::IsNullOrWhiteSpace($detail)) { $detail = "Probe exited with code $($process.ExitCode) without a result." }
            return [pscustomobject]@{ Name = $Name; Ready = $false; Roles = @(); Detail = $detail }
        }
        $result = $match.Groups['Json'].Value | ConvertFrom-Json -ErrorAction Stop
        return [pscustomobject]@{
            Name = $Name
            Ready = [bool]$result.Ready
            Roles = @($result.Roles | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
            Detail = [string]$result.Detail
        }
    }
    catch {
        return [pscustomobject]@{ Name = $Name; Ready = $false; Roles = @(); Detail = $_.Exception.Message }
    }
    finally {
        $process.Dispose()
    }
}

function Get-CapabilityProbeResult {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [bool]$Ready,
        [string]$Detail,
        [string[]]$Roles = @()
    )

    return [pscustomobject]@{
        Name = $Name
        Ready = $Ready
        Roles = @($Roles)
        Detail = $Detail
    }
}

function Test-SmartM365OrchestratorCapabilityComponentMissing {
    <# True when a capability is not ready only because its component is absent from this server
       (snap-in not registered or module not installed), which is normal on servers that do not host it. #>
    param([Parameter(Mandatory = $true)]$Result)
    return (-not [bool]$Result.Ready) -and ([string]$Result.Detail -match '(?i)\bis not (registered|installed) on this server\.$')
}

function Get-SmartM365OrchestratorPythonCommand {
    <# Python command of WorkplaceEvidence-Prepare: PreparedPythonCommand of its configuration, read (never migrated)
       where the job reads it; python when the file or the key is absent. #>
    $path = ConvertTo-SmartM365InventoryConfigLocation -Path (Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'PreparedEvidence\SmartM365-WorkplaceEvidence-Prepare.local.json')
    if (-not (Get-SmartM365JsonReadPath -Path $path -Optional)) { return 'python' }
    $config = (Read-SmartM365JsonDocument -Path $path).Document
    if ($null -ne $config -and $config.PSObject.Properties['PreparedPythonCommand'] -and -not [string]::IsNullOrWhiteSpace([string]$config.PreparedPythonCommand)) {
        return ([string]$config.PreparedPythonCommand).Trim()
    }
    return 'python'
}

function Get-PythonCapabilityResult {
    <# Python 3.10+ for WorkplaceEvidence-Prepare, from the version printed by the command itself (read-only, bounded). #>
    param(
        [Parameter(Mandatory = $true)][string]$Command,
        [int]$TimeoutSeconds = 30
    )

    $python = Get-Command -Name $Command -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $python) {
        return Get-CapabilityProbeResult -Name 'Python' -Ready $false -Detail ("Python command '{0}' (PreparedPythonCommand) is not installed on this server." -f $Command)
    }
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new($python.Source, '-c "import sys; print(''%d.%d.%d'' % sys.version_info[:3])"')
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    try {
        if (-not $process.Start()) { throw "Could not start $($python.Source)." }
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit([math]::Max(1, $TimeoutSeconds) * 1000)) {
            try { $process.Kill($true) } catch { Write-Verbose ("Python probe kill failed: {0}" -f $_.Exception.Message) }
            return Get-CapabilityProbeResult -Name 'Python' -Ready $false -Detail ("{0} did not answer within {1} seconds." -f $python.Source, $TimeoutSeconds)
        }
        $version = ([string]$stdoutTask.GetAwaiter().GetResult()).Trim()
        $stderr = ([string]$stderrTask.GetAwaiter().GetResult()).Trim() -replace '[\r\n]+', ' '
        $exitCode = $process.ExitCode
    }
    catch {
        return Get-CapabilityProbeResult -Name 'Python' -Ready $false -Detail $_.Exception.Message
    }
    finally {
        $process.Dispose()
    }
    if ($version -notmatch '^\d+\.\d+\.\d+$') {
        # The Microsoft Store alias (WindowsApps) answers without a version: no Python is installed.
        if ($python.Source -match '(?i)\\WindowsApps\\') {
            return Get-CapabilityProbeResult -Name 'Python' -Ready $false -Detail ("Python command '{0}' resolves to the Microsoft Store alias {1}: Python is not installed on this server." -f $Command, $python.Source)
        }
        if ($stderr.Length -gt 300) { $stderr = $stderr.Substring(0, 300) }
        return Get-CapabilityProbeResult -Name 'Python' -Ready $false -Detail ("{0} gave no Python version (exit code {1}): {2}" -f $python.Source, $exitCode, $stderr)
    }
    if ([version]$version -lt [version]'3.10') {
        return Get-CapabilityProbeResult -Name 'Python' -Ready $false -Detail ("Python {0} at {1} is below 3.10, required by WorkplaceEvidence-Prepare." -f $version, $python.Source)
    }
    return Get-CapabilityProbeResult -Name 'Python' -Ready $true -Detail ("Python {0} at {1}." -f $version, $python.Source)
}

function Get-SmartM365OrchestratorServerCapability {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ServerName,
        [Parameter(Mandatory = $true)][string]$SharedDataFolderPath,
        [string]$TenantId,
        [string]$AppId,
        [string]$CertificateThumbprint,
        [string]$Organization,
        [ValidateSet('Static', 'ReadOnly')][string]$ProbeMode = 'ReadOnly',
        [int]$ProbeTimeoutSeconds = 90,
        [AllowEmptyCollection()][object[]]$DataPathChecks = @(),
        [int]$PathTimeoutSeconds = 10,
        [int]$CertificateExpiryWarningDays = 0,
        # Blank: PreparedPythonCommand of the WorkplaceEvidence-Prepare configuration, else python.
        [string]$PythonCommand = ''
    )

    $generatedAtUtc = [datetime]::UtcNow
    $certificate = Get-CertificateReadiness -Thumbprint $CertificateThumbprint -ExpiryWarningDays $CertificateExpiryWarningDays -NowUtc $generatedAtUtc
    $results = [System.Collections.Generic.List[object]]::new()

    $probeFolder = Join-Path -Path $SharedDataFolderPath -ChildPath 'Election\Probe'
    try {
        [void](New-Item -ItemType Directory -Path $probeFolder -Force)
        $probePath = Join-Path -Path $probeFolder -ChildPath ('{0}.{1}.{2}.tmp' -f (ConvertTo-SafeFileName $ServerName), $PID, [guid]::NewGuid().ToString('N'))
        [System.IO.File]::WriteAllText($probePath, $generatedAtUtc.ToString('o'), [System.Text.UTF8Encoding]::new($false))
        Remove-Item -LiteralPath $probePath -Force -ErrorAction Stop
        # Tenant data roots (DATA-ALL, DATA-LAST, LOG-ALL) can live on other shares than the election folder.
        $dataFailures = @(Test-SmartM365OrchestratorPathAccess -Checks $DataPathChecks -ServerName $ServerName -TimeoutSeconds $PathTimeoutSeconds | Where-Object { -not $_.Ready })
        if ($dataFailures.Count -gt 0) {
            $detail = ($dataFailures | ForEach-Object { "{0} ({1}, {2}) is not accessible: {3}" -f $_.Name, $_.Path, $_.Access, $_.Detail }) -join ' '
            $results.Add((Get-CapabilityProbeResult -Name 'SharedRuntime' -Ready $false -Detail $detail))
        }
        else {
            $results.Add((Get-CapabilityProbeResult -Name 'SharedRuntime' -Ready $true -Detail ("Shared election storage is writable; {0} tenant data root(s) verified." -f @($DataPathChecks).Count)))
        }
    }
    catch {
        $results.Add((Get-CapabilityProbeResult -Name 'SharedRuntime' -Ready $false -Detail $_.Exception.Message))
    }

    $cloudConfigurationReady = $certificate.Ready -and
        -not [string]::IsNullOrWhiteSpace($TenantId) -and
        -not [string]::IsNullOrWhiteSpace($AppId)
    $pwshPath = (Get-Process -Id $PID).Path
    $windowsPowerShellPath = Join-Path -Path $env:SystemRoot -ChildPath 'System32\WindowsPowerShell\v1.0\powershell.exe'

    if ($ProbeMode -eq 'Static') {
        $graphModule = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Sort-Object Version -Descending | Select-Object -First 1
        $results.Add((Get-CapabilityProbeResult -Name 'Graph' -Ready ($cloudConfigurationReady -and $null -ne $graphModule) -Detail ("Static check: module={0}; {1}" -f $(if ($graphModule) { $graphModule.Version } else { 'missing' }), $certificate.Detail)))
        $exoModule = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
        $results.Add((Get-CapabilityProbeResult -Name 'EXO' -Ready ($cloudConfigurationReady -and $null -ne $exoModule -and -not [string]::IsNullOrWhiteSpace($Organization)) -Detail ("Static check: module={0}; organizationConfigured={1}; {2}" -f $(if ($exoModule) { $exoModule.Version } else { 'missing' }), (-not [string]::IsNullOrWhiteSpace($Organization)), $certificate.Detail)))
        $adModule = Get-Module -ListAvailable -Name ActiveDirectory | Select-Object -First 1
        $results.Add((Get-CapabilityProbeResult -Name 'AD' -Ready ($null -ne $adModule) -Detail ("Static check: ActiveDirectory module {0}." -f $(if ($adModule) { 'available' } else { 'missing' }))))
        $exchangeSnapIn = & $windowsPowerShellPath -NoLogo -NoProfile -NonInteractive -Command "if (Get-PSSnapin -Registered -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue) { 'yes' } else { 'no' }" 2>$null
        $exchangeRegistered = ($exchangeSnapIn | Select-Object -Last 1) -eq 'yes'
        $results.Add((Get-CapabilityProbeResult -Name 'ExchangeOnPrem' -Ready $exchangeRegistered -Detail $(if ($exchangeRegistered) { 'Static check of the Exchange Management snap-in registration.' } else { 'Exchange Management Shell snap-in Microsoft.Exchange.Management.PowerShell.SnapIn is not registered on this server.' })))
        $sharePointSnapIn = & $windowsPowerShellPath -NoLogo -NoProfile -NonInteractive -Command "if (Get-PSSnapin -Registered -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) { 'yes' } else { 'no' }" 2>$null
        $sharePointRegistered = ($sharePointSnapIn | Select-Object -Last 1) -eq 'yes'
        $results.Add((Get-CapabilityProbeResult -Name 'SharePointOnPrem' -Ready $sharePointRegistered -Detail $(if ($sharePointRegistered) { 'Static check of the SharePoint Server PowerShell snap-in registration.' } else { 'SharePoint Server PowerShell snap-in Microsoft.SharePoint.PowerShell is not registered on this server.' })))
        $teamsModule = Get-Module -ListAvailable -Name MicrosoftTeams | Sort-Object Version -Descending | Select-Object -First 1
        $results.Add((Get-CapabilityProbeResult -Name 'TeamsPowerShell' -Ready ($cloudConfigurationReady -and $null -ne $teamsModule) -Detail $(if ($null -eq $teamsModule) { 'MicrosoftTeams PowerShell module is not installed on this server.' } else { "Static check: module={0}; {1}" -f $teamsModule.Version, $certificate.Detail })))
        $adaxesModule = Get-Module -ListAvailable -Name Adaxes | Sort-Object Version -Descending | Select-Object -First 1
        $results.Add((Get-CapabilityProbeResult -Name 'Adaxes' -Ready ($null -ne $adaxesModule) -Detail $(if ($null -eq $adaxesModule) { 'Adaxes PowerShell module is not installed on this server.' } else { "Static check: Adaxes module {0}." -f $adaxesModule.Version })))
    }
    else {
        $environment = @{
            SMART_TENANT_ID = $TenantId
            SMART_APP_ID = $AppId
            SMART_CERT_THUMBPRINT = $CertificateThumbprint
            SMART_ORGANIZATION = $Organization
        }
        # Cloud probes retry transient service errors (5xx, 408, 429, timeouts, connection resets) twice
        # within the probe timeout; authorization errors (401, 403, AADSTS) fail at once.
        $cloudProbeRetry = @'
function Invoke-ProbeWithRetry([scriptblock]$Probe) {
    for ($attempt = 1; ; $attempt++) {
        try { return & $Probe }
        catch {
            $message = [string]$_.Exception.Message
            $transient = $message -notmatch '(?i)\b(?:401|403)\b|Unauthorized|Forbidden|AADSTS' -and $message -match '(?i)\b(?:408|429|500|502|503|504)\b|InternalServerError|ServiceUnavailable|BadGateway|GatewayTimeout|TooManyRequests|throttl|timeout|timed out|temporarily unavailable|An error occurred while sending the request|connection (?:was )?(?:reset|closed|aborted|forcibly closed)'
            if (-not $transient) { throw }
            if ($attempt -ge 3) { throw ("Persistent transient error after {0} attempt(s): {1}" -f $attempt, $message) }
            Start-Sleep -Seconds (5 * $attempt)
        }
    }
}

'@
        $graphScript = $cloudProbeRetry + @'
$ErrorActionPreference = 'Stop'
try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Invoke-ProbeWithRetry { Connect-MgGraph -TenantId $env:SMART_TENANT_ID -ClientId $env:SMART_APP_ID -CertificateThumbprint $env:SMART_CERT_THUMBPRINT -NoWelcome }
    $roles = @((Get-MgContext).Scopes | Sort-Object -Unique)
    $result = @{ Ready = $true; Roles = $roles; Detail = "Read-only Graph authentication succeeded with $($roles.Count) application roles." }
}
catch { $result = @{ Ready = $false; Roles = @(); Detail = $_.Exception.Message } }
finally { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
"SMARTM365_CAPABILITY_JSON:$($result | ConvertTo-Json -Compress -Depth 5)"
'@
        $results.Add((Invoke-IsolatedCapabilityProbe -Name 'Graph' -EnginePath $pwshPath -ScriptText $graphScript -Environment $environment -TimeoutSeconds $ProbeTimeoutSeconds))

        $exoScript = $cloudProbeRetry + @'
$ErrorActionPreference = 'Stop'
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Invoke-ProbeWithRetry {
        Connect-ExchangeOnline -AppId $env:SMART_APP_ID -CertificateThumbprint $env:SMART_CERT_THUMBPRINT -Organization $env:SMART_ORGANIZATION -ShowBanner:$false
        $null = Get-AcceptedDomain -ResultSize 1 -ErrorAction Stop
    }
    $result = @{ Ready = $true; Roles = @(); Detail = 'Read-only Exchange Online probe succeeded.' }
}
catch { $result = @{ Ready = $false; Roles = @(); Detail = $_.Exception.Message } }
finally { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
"SMARTM365_CAPABILITY_JSON:$($result | ConvertTo-Json -Compress -Depth 5)"
'@
        $results.Add((Invoke-IsolatedCapabilityProbe -Name 'EXO' -EnginePath $pwshPath -ScriptText $exoScript -Environment $environment -TimeoutSeconds $ProbeTimeoutSeconds))

        $adScript = @'
$ErrorActionPreference = 'Stop'
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $rootDse = Get-ADRootDSE -ErrorAction Stop
    $result = @{ Ready = $true; Roles = @(); Detail = "Read-only AD probe succeeded for $($rootDse.defaultNamingContext)." }
}
catch { $result = @{ Ready = $false; Roles = @(); Detail = $_.Exception.Message } }
"SMARTM365_CAPABILITY_JSON:$($result | ConvertTo-Json -Compress -Depth 5)"
'@
        $results.Add((Invoke-IsolatedCapabilityProbe -Name 'AD' -EnginePath $windowsPowerShellPath -ScriptText $adScript -TimeoutSeconds $ProbeTimeoutSeconds))

        $exchangeScript = @'
$ErrorActionPreference = 'Stop'
try {
    if (-not (Get-PSSnapin -Registered -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) { throw 'Exchange Management Shell snap-in Microsoft.Exchange.Management.PowerShell.SnapIn is not registered on this server.' }
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
    $server = Get-ExchangeServer -ErrorAction Stop | Select-Object -First 1
    $result = @{ Ready = $true; Roles = @(); Detail = "Read-only Exchange on-premises probe succeeded against $($server.Name)." }
}
catch { $result = @{ Ready = $false; Roles = @(); Detail = $_.Exception.Message } }
"SMARTM365_CAPABILITY_JSON:$($result | ConvertTo-Json -Compress -Depth 5)"
'@
        $results.Add((Invoke-IsolatedCapabilityProbe -Name 'ExchangeOnPrem' -EnginePath $windowsPowerShellPath -ScriptText $exchangeScript -TimeoutSeconds $ProbeTimeoutSeconds))

        # Same snap-in and read-only farm call as the SharePoint on-premises collectors.
        $sharePointScript = @'
$ErrorActionPreference = 'Stop'
try {
    if (-not (Get-PSSnapin -Registered -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) { throw 'SharePoint Server PowerShell snap-in Microsoft.SharePoint.PowerShell is not registered on this server.' }
    Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
    $farm = Get-SPFarm -ErrorAction Stop
    $result = @{ Ready = $true; Roles = @(); Detail = "Read-only SharePoint Server probe succeeded against farm $($farm.Name) (build $($farm.BuildVersion))." }
}
catch { $result = @{ Ready = $false; Roles = @(); Detail = $_.Exception.Message } }
"SMARTM365_CAPABILITY_JSON:$($result | ConvertTo-Json -Compress -Depth 5)"
'@
        $results.Add((Invoke-IsolatedCapabilityProbe -Name 'SharePointOnPrem' -EnginePath $windowsPowerShellPath -ScriptText $sharePointScript -TimeoutSeconds $ProbeTimeoutSeconds))

        $teamsScript = @'
$ErrorActionPreference = 'Stop'
try {
    if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) { throw 'MicrosoftTeams PowerShell module is not installed on this server.' }
    Import-Module MicrosoftTeams -MinimumVersion 4.7.1 -ErrorAction Stop
    # Same app-only call as M365-TeamsPhonePstnUsage-Inventory, so the probe checks what the job uses.
    # Connect at script level before looking for the cmdlet, whether the module exports it at import or on connection.
    Connect-MicrosoftTeams -TenantId $env:SMART_TENANT_ID -ApplicationId $env:SMART_APP_ID -CertificateThumbprint $env:SMART_CERT_THUMBPRINT -ErrorAction Stop | Out-Null
    if (-not (Get-Command -Name Get-CsPhoneNumberAssignment -ErrorAction SilentlyContinue)) { throw 'Get-CsPhoneNumberAssignment is not available in the installed MicrosoftTeams module after Connect-MicrosoftTeams.' }
    $null = @(Get-CsPhoneNumberAssignment -Top 1 -ErrorAction Stop)
    $result = @{ Ready = $true; Roles = @(); Detail = "Read-only Teams PowerShell probe succeeded (MicrosoftTeams $((Get-Module -Name MicrosoftTeams).Version), Get-CsPhoneNumberAssignment)." }
}
catch { $result = @{ Ready = $false; Roles = @(); Detail = $_.Exception.Message } }
finally {
    try {
        if ($null -ne (Get-Command -Name Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue)) {
            Disconnect-MicrosoftTeams -ErrorAction Stop | Out-Null
        }
    }
    catch {}
}
"SMARTM365_CAPABILITY_JSON:$($result | ConvertTo-Json -Compress -Depth 5)"
'@
        $results.Add((Invoke-IsolatedCapabilityProbe -Name 'TeamsPowerShell' -EnginePath $pwshPath -ScriptText $teamsScript -Environment $environment -TimeoutSeconds $ProbeTimeoutSeconds))

        # The probe has no Adaxes service name: it checks that the collector module loads in PowerShell 7.
        $adaxesScript = @'
$ErrorActionPreference = 'Stop'
try {
    if (-not (Get-Module -ListAvailable -Name Adaxes)) { throw 'Adaxes PowerShell module is not installed on this server.' }
    Import-Module Adaxes -ErrorAction Stop
    if (-not (Get-Command -Name Get-AdmUser -ErrorAction SilentlyContinue)) { throw 'Adaxes PowerShell module does not expose Get-AdmUser.' }
    $result = @{ Ready = $true; Roles = @(); Detail = "Adaxes PowerShell module $((Get-Module -Name Adaxes).Version) loaded." }
}
catch { $result = @{ Ready = $false; Roles = @(); Detail = $_.Exception.Message } }
"SMARTM365_CAPABILITY_JSON:$($result | ConvertTo-Json -Compress -Depth 5)"
'@
        $results.Add((Invoke-IsolatedCapabilityProbe -Name 'Adaxes' -EnginePath $pwshPath -ScriptText $adaxesScript -TimeoutSeconds $ProbeTimeoutSeconds))
    }

    # Python is checked the same way in both modes: the command only prints its version.
    try {
        $pythonCommandToProbe = if ([string]::IsNullOrWhiteSpace($PythonCommand)) { Get-SmartM365OrchestratorPythonCommand } else { $PythonCommand.Trim() }
        $results.Add((Get-PythonCapabilityResult -Command $pythonCommandToProbe -TimeoutSeconds ([math]::Min(30, [math]::Max(1, $ProbeTimeoutSeconds)))))
    }
    catch {
        $results.Add((Get-CapabilityProbeResult -Name 'Python' -Ready $false -Detail ("WorkplaceEvidence-Prepare configuration could not be read for PreparedPythonCommand: {0}" -f $_.Exception.Message)))
    }

    foreach ($result in $results) { $result | Add-Member -NotePropertyName ComponentMissing -NotePropertyValue (Test-SmartM365OrchestratorCapabilityComponentMissing -Result $result) -Force }
    $readyCapabilities = @($results | Where-Object Ready | ForEach-Object Name | Sort-Object -Unique)
    $graphResult = $results | Where-Object Name -eq 'Graph' | Select-Object -First 1
    return [pscustomobject]@{
        SchemaVersion = 1
        ModuleVersion = $script:DistributedModuleVersion
        ServerName = $ServerName.ToUpperInvariant()
        GeneratedAtUtc = $generatedAtUtc.ToString('o')
        ProbeMode = $ProbeMode
        ReadyCapabilities = $readyCapabilities
        GraphAppRoles = @($graphResult.Roles | Sort-Object -Unique)
        Results = @($results)
        CertificateStatus = [pscustomobject]@{
            Ready = [bool]$certificate.Ready
            Reason = [string]$certificate.Reason
            NotAfterUtc = [string]$certificate.NotAfterUtc
            DaysRemaining = $certificate.DaysRemaining
            ExpiresSoon = [bool]$certificate.ExpiresSoon
        }
        BlockedJobs = @()
        PrerequisiteBlocks = @()
    }
}

function Test-SmartM365OrchestratorCapabilityMatch {
    param(
        [Parameter(Mandatory = $true)]$ServerCapabilities,
        [string[]]$RequiredCapabilities = @(),
        [string[]]$RequiredGraphAppRoles = @()
    )

    $ready = @($ServerCapabilities.ReadyCapabilities | ForEach-Object { [string]$_ })
    $roles = @($ServerCapabilities.GraphAppRoles | ForEach-Object { [string]$_ })
    $missingCapabilities = @($RequiredCapabilities | Where-Object { $_ -notin $ready } | Sort-Object -Unique)
    $missingRoles = @($RequiredGraphAppRoles | Where-Object { $_ -notin $roles } | Sort-Object -Unique)
    return [pscustomobject]@{
        Eligible = ($missingCapabilities.Count -eq 0 -and $missingRoles.Count -eq 0)
        MissingCapabilities = $missingCapabilities
        MissingGraphAppRoles = $missingRoles
    }
}

function Get-ScheduleRunsPerDay {
    param($Schedule)

    if ($null -eq $Schedule) { return 0.0 }
    $times = @($Schedule.Times).Count
    if ([string]$Schedule.Type -eq 'Weekly') {
        return [double]($times * @($Schedule.DaysOfWeek).Count) / 7.0
    }
    return [double]$times
}

function Get-DependencyComponent {
    param([Parameter(Mandatory = $true)][object[]]$Jobs)

    $jobsByName = @{}
    foreach ($job in $Jobs) { $jobsByName[[string]$job.Name] = $job }
    $adjacency = @{}
    foreach ($job in $Jobs) { $adjacency[[string]$job.Name] = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }
    foreach ($job in $Jobs) {
        foreach ($dependencyName in @($job.DependsOn)) {
            if ($jobsByName.ContainsKey([string]$dependencyName)) {
                [void]$adjacency[[string]$job.Name].Add([string]$dependencyName)
                [void]$adjacency[[string]$dependencyName].Add([string]$job.Name)
            }
        }
    }

    $visited = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $components = @()
    foreach ($name in @($jobsByName.Keys | Sort-Object)) {
        if ($visited.Contains($name)) { continue }
        $queue = [System.Collections.Generic.Queue[string]]::new()
        $queue.Enqueue($name)
        [void]$visited.Add($name)
        $members = @()
        while ($queue.Count -gt 0) {
            $current = $queue.Dequeue()
            $members += $current
            foreach ($neighbor in $adjacency[$current]) {
                if ($visited.Add($neighbor)) { $queue.Enqueue($neighbor) }
            }
        }
        $components += ,@($members | Sort-Object)
    }
    return $components
}

function Test-ServerJobPolicy {
    param(
        [Parameter(Mandatory = $true)][string]$ServerName,
        [Parameter(Mandatory = $true)][object[]]$Jobs,
        [hashtable]$ServerJobPolicies = @{}
    )

    if (-not $ServerJobPolicies.ContainsKey($ServerName)) { return $true }
    $policy = $ServerJobPolicies[$ServerName]
    $onlyJobsRequiring = @()
    if ($policy -is [System.Collections.IDictionary]) {
        if ($policy.Contains('OnlyJobsRequiring')) { $onlyJobsRequiring = @($policy['OnlyJobsRequiring']) }
    }
    elseif ($null -ne $policy -and $policy.PSObject.Properties['OnlyJobsRequiring']) {
        $onlyJobsRequiring = @($policy.OnlyJobsRequiring)
    }
    $onlyJobsRequiring = @($onlyJobsRequiring | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    if ($onlyJobsRequiring.Count -eq 0) { return $true }

    foreach ($job in $Jobs) {
        $jobCapabilities = @($job.RequiredCapabilities | ForEach-Object { [string]$_ })
        if (@($onlyJobsRequiring | Where-Object { $_ -notin $jobCapabilities }).Count -gt 0) { return $false }
    }
    return $true
}

function Get-SmartM365OrchestratorElectionPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$Jobs,
        [Parameter(Mandatory = $true)][object[]]$ServerCapabilities,
        [hashtable]$ServerWeights = @{},
        [hashtable]$ServerJobPolicies = @{},
        [hashtable]$DurationMinutesByJob = @{},
        [AllowNull()]$PreviousPlan = $null,
        [switch]$PreservePreviousOwners,
        # Servers set Inactive in the cluster configuration: never candidates (orchestrator 1.5.76).
        [string[]]$InactiveServers = @(),
        [datetime]$NowUtc = [datetime]::UtcNow
    )

    $electedJobs = @($Jobs | Where-Object { $_.Enabled -and [string]$_.AssignmentMode -eq 'Elected' })
    $inactive = @{}
    foreach ($inactiveServer in @($InactiveServers)) { if (-not [string]::IsNullOrWhiteSpace([string]$inactiveServer)) { $inactive[([string]$inactiveServer).Trim().ToUpperInvariant()] = $true } }
    $capabilitiesByServer = @{}
    foreach ($serverCapability in $ServerCapabilities) {
        $capabilityServerName = [string]$serverCapability.ServerName
        # Either source is enough: the shared cluster state, or the state the server itself published.
        $publishedState = if ($serverCapability.PSObject.Properties['ServerState']) { [string]$serverCapability.ServerState } else { '' }
        if ($publishedState -ieq 'Inactive') { $inactive[$capabilityServerName.ToUpperInvariant()] = $true }
        if ($inactive.ContainsKey($capabilityServerName.ToUpperInvariant())) { continue }
        $capabilitiesByServer[$capabilityServerName] = $serverCapability
    }
    $loads = @{}
    foreach ($serverName in $capabilitiesByServer.Keys) { $loads[$serverName] = 0.0 }

    $previousOwnerByJob = @{}
    if ($PreservePreviousOwners -and $null -ne $PreviousPlan -and $PreviousPlan.PSObject.Properties['Assignments']) {
        foreach ($assignment in @($PreviousPlan.Assignments)) {
            if (-not $assignment.PSObject.Properties['JobName'] -or -not $assignment.PSObject.Properties['OwnerServer']) { continue }
            $jobName = [string]$assignment.JobName
            $ownerServer = ([string]$assignment.OwnerServer).ToUpperInvariant()
            if (-not [string]::IsNullOrWhiteSpace($jobName) -and -not [string]::IsNullOrWhiteSpace($ownerServer)) {
                $previousOwnerByJob[$jobName] = $ownerServer
            }
        }
    }

    $groups = @()
    # Dependencies constrain execution order through shared occurrence claims,
    # not placement: cloud and on-premises prerequisites can live on different hosts.
    foreach ($electedJob in $electedJobs) {
        $component = @([string]$electedJob.Name)
        $componentJobs = @($electedJobs | Where-Object { $_.Name -in $component })
        $requiredCapabilities = @($componentJobs | ForEach-Object { @($_.RequiredCapabilities) } | Sort-Object -Unique)
        $requiredRoles = @($componentJobs | ForEach-Object { @($_.RequiredGraphAppRoles) } | Sort-Object -Unique)
        $loadMinutes = 0.0
        foreach ($job in $componentJobs) {
            $duration = if ($DurationMinutesByJob.ContainsKey([string]$job.Name)) {
                [double]$DurationMinutesByJob[[string]$job.Name]
            }
            elseif ($null -ne $job.EstimatedDurationMinutes -and [double]$job.EstimatedDurationMinutes -gt 0) {
                [double]$job.EstimatedDurationMinutes
            }
            else { 5.0 }
            $loadMinutes += $duration * (Get-ScheduleRunsPerDay -Schedule $job.Schedule)
        }
        $candidates = @()
        foreach ($serverName in @($capabilitiesByServer.Keys | Sort-Object)) {
            $serverCapability = $capabilitiesByServer[$serverName]
            # A launch-time prerequisite failure specific to a job (missing script folder, RequiredPaths)
            # excludes that server for the job only.
            if ($serverCapability.PSObject.Properties['BlockedJobs'] -and
                @($component | Where-Object { $_ -in @($serverCapability.BlockedJobs | ForEach-Object { [string]$_ }) }).Count -gt 0) { continue }
            $match = Test-SmartM365OrchestratorCapabilityMatch -ServerCapabilities $serverCapability -RequiredCapabilities $requiredCapabilities -RequiredGraphAppRoles $requiredRoles
            if ($match.Eligible -and (Test-ServerJobPolicy -ServerName $serverName -Jobs $componentJobs -ServerJobPolicies $ServerJobPolicies)) { $candidates += $serverName }
        }
        $previousOwners = @(
            $component |
                ForEach-Object { if ($previousOwnerByJob.ContainsKey([string]$_)) { $previousOwnerByJob[[string]$_] } } |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique
        )
        $preservedOwner = ''
        if ($PreservePreviousOwners -and $previousOwners.Count -eq 1 -and $previousOwners[0] -in $candidates) {
            $preservedOwner = [string]$previousOwners[0]
        }
        $groups += [pscustomobject]@{
            GroupKey = (@($component) -join '+')
            Jobs = @($component)
            RequiredCapabilities = $requiredCapabilities
            RequiredGraphAppRoles = $requiredRoles
            LoadMinutesPerDay = [math]::Round($loadMinutes, 4)
            Candidates = $candidates
            PreservedOwner = $preservedOwner
        }
    }

    $assignments = @()
    $unassigned = @()
    $orderedGroups = @($groups | Sort-Object @{ Expression = 'LoadMinutesPerDay'; Descending = $true }, GroupKey)
    foreach ($group in $orderedGroups) {
        if (@($group.Candidates).Count -eq 0) {
            $unassigned += [pscustomobject]@{ GroupKey = $group.GroupKey; Jobs = $group.Jobs; Reason = 'No live active server satisfies every required capability, Graph application role and server job policy.' }
            continue
        }
        if ([string]::IsNullOrWhiteSpace([string]$group.PreservedOwner)) { continue }
        $owner = [string]$group.PreservedOwner
        $loads[$owner] = [double]$loads[$owner] + [double]$group.LoadMinutesPerDay
        foreach ($jobName in $group.Jobs) {
            $assignments += [pscustomobject]@{
                JobName = $jobName
                OwnerServer = $owner
                GroupKey = $group.GroupKey
                LoadMinutesPerDay = $group.LoadMinutesPerDay
            }
        }
    }

    foreach ($group in $orderedGroups) {
        if (@($group.Candidates).Count -eq 0 -or -not [string]::IsNullOrWhiteSpace([string]$group.PreservedOwner)) { continue }
        $owner = @($group.Candidates | Sort-Object @{
                Expression = {
                    $weight = if ($ServerWeights.ContainsKey([string]$_) -and [double]$ServerWeights[[string]$_] -gt 0) { [double]$ServerWeights[[string]$_] } else { 1.0 }
                    [double]$loads[[string]$_] / $weight
                }
            }, @{ Expression = { [string]$_ } })[0]
        $loads[$owner] = [double]$loads[$owner] + [double]$group.LoadMinutesPerDay
        foreach ($jobName in $group.Jobs) {
            $assignments += [pscustomobject]@{
                JobName = $jobName
                OwnerServer = $owner
                GroupKey = $group.GroupKey
                LoadMinutesPerDay = $group.LoadMinutesPerDay
            }
        }
    }

    return [pscustomobject]@{
        SchemaVersion = 2
        PlanId = [guid]::NewGuid().ToString('N')
        GeneratedAtUtc = $NowUtc.ToString('o')
        EligibleServers = @($capabilitiesByServer.Keys | Sort-Object)
        InactiveServers = @($inactive.Keys | Sort-Object)
        Assignments = @($assignments | Sort-Object JobName)
        UnassignedGroups = @($unassigned)
        ServerLoads = @($loads.Keys | Sort-Object | ForEach-Object {
                [pscustomobject]@{
                    ServerName = $_
                    Weight = if ($ServerWeights.ContainsKey($_)) { [double]$ServerWeights[$_] } else { 1.0 }
                    LoadMinutesPerDay = [math]::Round([double]$loads[$_], 4)
                }
            })
    }
}

function Test-SmartM365OrchestratorCanPreserveOwners {
    [CmdletBinding()]
    param(
        [AllowNull()]$PreviousPlan = $null,
        [Parameter(Mandatory = $true)][object[]]$ServerCapabilities,
        [string[]]$InactiveServers = @()
    )

    if ($null -eq $PreviousPlan -or
        -not $PreviousPlan.PSObject.Properties['SchemaVersion'] -or
        [int]$PreviousPlan.SchemaVersion -lt 2 -or
        -not $PreviousPlan.PSObject.Properties['EligibleServers']) {
        return $false
    }
    if ($PreviousPlan.PSObject.Properties['UnassignedGroups'] -and @($PreviousPlan.UnassignedGroups).Count -gt 0) {
        return $false
    }

    $previousServers = @(
        $PreviousPlan.EligibleServers |
            ForEach-Object { ([string]$_).Trim().ToUpperInvariant() } |
            Where-Object { $_ } |
            Sort-Object -Unique
    )
    # Inactive servers are never in EligibleServers: compare the active ones only.
    $inactive = @($InactiveServers | ForEach-Object { ([string]$_).Trim().ToUpperInvariant() })
    $currentServers = @(
        $ServerCapabilities |
            Where-Object { -not ($_.PSObject.Properties['ServerState'] -and [string]$_.ServerState -ieq 'Inactive') } |
            ForEach-Object { ([string]$_.ServerName).Trim().ToUpperInvariant() } |
            Where-Object { $_ -and $_ -notin $inactive } |
            Sort-Object -Unique
    )
    if ($previousServers.Count -ne $currentServers.Count) { return $false }
    return (($previousServers -join '|') -ceq ($currentServers -join '|'))
}

function Get-SmartM365OrchestratorOccurrenceClaim {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ClaimsRootPath,
        [Parameter(Mandatory = $true)][string]$JobName,
        [Parameter(Mandatory = $true)][datetime]$Occurrence
    )

    $jobFolder = Join-Path -Path $ClaimsRootPath -ChildPath (ConvertTo-SafeFileName $JobName)
    $occurrenceUtc = $Occurrence.ToUniversalTime()
    $claimPath = Join-Path -Path $jobFolder -ChildPath ($occurrenceUtc.ToString('yyyyMMddTHHmmssfffZ') + '.json'); $claimPath = Resolve-DistributedJsonPath $claimPath
    if (-not (Test-Path -LiteralPath $claimPath -PathType Leaf)) { return $null }
    $claim = (Read-SmartM365JsonDocument $claimPath).Document
    return [pscustomobject]@{
        ClaimPath = $claimPath
        Claim = $claim
    }
}

function Set-DistributedClaimOwnerAtomically {
    # Compare-and-swap on the claim file: the write succeeds only if nobody changed the claim
    # since it was read, so two servers can never both win an offered retry.
    param(
        [Parameter(Mandatory = $true)][string]$ClaimPath,
        [Parameter(Mandatory = $true)]$ClaimDocument,
        [Parameter(Mandatory = $true)][string]$OwnerServer,
        [Parameter(Mandatory = $true)][string]$PlanId,
        [int]$SafeMinutes = 60,
        [string]$HandoffFrom = ''
    )
    $claim = $ClaimDocument.Document
    $values = [ordered]@{
        OwnerServer = $OwnerServer.ToUpperInvariant(); Status = 'Claimed'; OrchestratorPid = $PID; PlanId = $PlanId
        UpdatedAtUtc = [datetime]::UtcNow.ToString('o'); SafeUntilUtc = [datetime]::UtcNow.AddMinutes([math]::Max(1, $SafeMinutes)).ToString('o')
    }
    if (-not [string]::IsNullOrWhiteSpace($HandoffFrom)) { $values['HandoffFrom'] = $HandoffFrom.ToUpperInvariant(); $values['HandoffAtUtc'] = [datetime]::UtcNow.ToString('o') }
    foreach ($key in $values.Keys) {
        if ($claim.PSObject.Properties.Name -contains $key) { $claim.$key = $values[$key] }
        else { $claim | Add-Member -NotePropertyName $key -NotePropertyValue $values[$key] }
    }
    Write-JsonAtomically -Path $ClaimPath -Value $claim -ExpectedSHA256 $ClaimDocument.SHA256
    return $claim
}

function Enter-SmartM365OrchestratorOccurrenceClaim {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ClaimsRootPath,
        [Parameter(Mandatory = $true)][string]$JobName,
        [Parameter(Mandatory = $true)][datetime]$Occurrence,
        [Parameter(Mandatory = $true)][string]$OwnerServer,
        [Parameter(Mandatory = $true)][string]$PlanId,
        [int]$SafeMinutes = 60,
        [string]$HeartbeatRootPath = '',
        [int]$HeartbeatStaleMinutes = 5,
        # Server whose RetryScheduled claim this server may take over (orchestrator retry handoff).
        [string]$RetryHandoffFrom = ''
    )

    $jobFolder = Join-Path -Path $ClaimsRootPath -ChildPath (ConvertTo-SafeFileName $JobName)
    [void](New-Item -ItemType Directory -Path $jobFolder -Force)
    $occurrenceUtc = $Occurrence.ToUniversalTime()
    $claimPath = Join-Path -Path $jobFolder -ChildPath ($occurrenceUtc.ToString('yyyyMMddTHHmmssfffZ') + '.json'); $claimPath = Resolve-DistributedJsonPath $claimPath
    $claim = [ordered]@{
        SchemaVersion = 1
        ClaimId = [guid]::NewGuid().ToString('N')
        JobName = $JobName
        OccurrenceUtc = $occurrenceUtc.ToString('o')
        OwnerServer = $OwnerServer.ToUpperInvariant()
        PlanId = $PlanId
        OrchestratorPid = $PID
        Status = 'Claimed'
        Attempt = 0
        CreatedAtUtc = [datetime]::UtcNow.ToString('o')
        UpdatedAtUtc = [datetime]::UtcNow.ToString('o')
        SafeUntilUtc = [datetime]::UtcNow.AddMinutes([math]::Max(1, $SafeMinutes)).ToString('o')
    }
    $json = $claim | ConvertTo-Json -Depth 6

    try {
        $null = Write-SmartM365JsonBytesAtomically -Path $claimPath -Bytes ([Text.UTF8Encoding]::new($false).GetBytes($json)) -ExpectedSHA256 'ABSENT' -Validate { param($document) if (-not $document.PSObject.Properties['ClaimId']) { throw 'ClaimId missing.' } }
        return [pscustomobject]@{ Acquired = $true; Reused = $false; ClaimPath = $claimPath; Claim = [pscustomobject]$claim; Reason = '' }
    }
    catch [System.IO.IOException] {
        try {
            $existingDocument = Read-SmartM365JsonDocument $claimPath
            $existing = $existingDocument.Document
            $sameOwner = [string]$existing.OwnerServer -eq $OwnerServer.ToUpperInvariant()
            $sameOrchestratorProcess = $false
            try { $sameOrchestratorProcess = [int]$existing.OrchestratorPid -eq $PID } catch { $sameOrchestratorProcess = $false }
            $reusable = [string]$existing.Status -eq 'RetryScheduled' -or
                ([string]$existing.Status -eq 'Claimed' -and $sameOrchestratorProcess)
            if ($sameOwner -and $reusable) {
                if ([string]$existing.Status -eq 'RetryScheduled') {
                    # The retry may be offered to peers: take it back only if no peer took it meanwhile.
                    try { $existing = Set-DistributedClaimOwnerAtomically -ClaimPath $claimPath -ClaimDocument $existingDocument -OwnerServer $OwnerServer -PlanId $PlanId -SafeMinutes $SafeMinutes }
                    catch { return [pscustomobject]@{ Acquired = $false; Reused = $false; ClaimPath = $claimPath; Claim = $existing; Reason = "The scheduled retry changed while it was claimed (another server took it): $($_.Exception.Message)" } }
                }
                return [pscustomobject]@{ Acquired = $true; Reused = $true; ClaimPath = $claimPath; Claim = $existing; Reason = 'Existing non-terminal claim belongs to this server.' }
            }
            if (-not [string]::IsNullOrWhiteSpace($RetryHandoffFrom) -and [string]$existing.Status -eq 'RetryScheduled' -and [string]$existing.OwnerServer -eq $RetryHandoffFrom.ToUpperInvariant()) {
                try { $existing = Set-DistributedClaimOwnerAtomically -ClaimPath $claimPath -ClaimDocument $existingDocument -OwnerServer $OwnerServer -PlanId $PlanId -SafeMinutes $SafeMinutes -HandoffFrom $RetryHandoffFrom }
                catch { return [pscustomobject]@{ Acquired = $false; Reused = $false; ClaimPath = $claimPath; Claim = $existing; Reason = "The offered retry was taken by another server: $($_.Exception.Message)" } }
                return [pscustomobject]@{ Acquired = $true; Reused = $true; ClaimPath = $claimPath; Claim = $existing; Reason = "Retry taken over from $($RetryHandoffFrom.ToUpperInvariant())." }
            }

            $terminal = [string]$existing.Status -in @('Success', 'CompletedWithWarnings', 'Failed', 'TimedOut', 'Interrupted', 'Cancelled')
            $safeUntilUtc = [datetime]::MaxValue
            try { $safeUntilUtc = ([datetime]$existing.SafeUntilUtc).ToUniversalTime() } catch { $safeUntilUtc = [datetime]::MaxValue }
            if (-not $terminal -and [datetime]::UtcNow -gt $safeUntilUtc -and -not [string]::IsNullOrWhiteSpace($HeartbeatRootPath)) {
                $heartbeatFresh = $false
                try {
                    $heartbeatPath = Join-Path -Path (Join-Path -Path $HeartbeatRootPath -ChildPath ([string]$existing.OwnerServer)) -ChildPath 'Orchestrator-Heartbeat.json'
                    $heartbeat = if(Get-SmartM365JsonReadPath $heartbeatPath -Optional){(Read-SmartM365JsonDocument $heartbeatPath).Document}else{$null}
                    $heartbeatAgeMinutes = if($heartbeat){([datetime]::UtcNow - ([datetime]$heartbeat.Timestamp).ToUniversalTime()).TotalMinutes}else{[double]::PositiveInfinity}
                    $heartbeatFresh = $heartbeatAgeMinutes -le [math]::Max(1, $HeartbeatStaleMinutes)
                }
                catch { throw ('Cannot determine peer liveness; existing coordination state preserved: ' + $_.Exception.Message) }
                # A restarted orchestrator on the same server may replace its own
                # expired claim even though its new heartbeat is healthy.
                if ($sameOwner) { $heartbeatFresh = $false }

                if (-not $heartbeatFresh) {
                    $takeoverLockPath = (Get-SmartM365JsonNames $claimPath).Legacy + '.takeover.lock'
                    $takeoverStream = $null
                    try {
                        $takeoverStream = [System.IO.File]::Open($takeoverLockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
                        $confirmed = (Read-SmartM365JsonDocument $claimPath).Document
                        $confirmedSafeUntilUtc = ([datetime]$confirmed.SafeUntilUtc).ToUniversalTime()
                        if ([string]$confirmed.ClaimId -eq [string]$existing.ClaimId -and
                            [string]$confirmed.Status -notin @('Success', 'CompletedWithWarnings', 'Failed', 'TimedOut', 'Interrupted', 'Cancelled') -and
                            [datetime]::UtcNow -gt $confirmedSafeUntilUtc) {
                            $archivePath = '{0}.stale.{1}.json' -f (Get-SmartM365JsonNames $claimPath).Legacy, [string]$confirmed.ClaimId
                            if ((Get-SmartM365JsonTransportPolicy).Mode -eq 'JsonText') { $archivePath += '.txt' }
                            Move-Item -LiteralPath $claimPath -Destination $archivePath -ErrorAction Stop
                        }
                    }
                    catch [System.IO.IOException] {
                        return [pscustomobject]@{ Acquired = $false; Reused = $false; ClaimPath = $claimPath; Claim = $existing; Reason = 'Another orchestrator is evaluating takeover of the expired claim.' }
                    }
                    finally {
                        if ($null -ne $takeoverStream) {
                            $takeoverStream.Dispose()
                            Remove-Item -LiteralPath $takeoverLockPath -Force -ErrorAction SilentlyContinue
                        }
                    }
                    if (-not (Test-Path -LiteralPath $claimPath)) {
                        return Enter-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $ClaimsRootPath -JobName $JobName -Occurrence $Occurrence -OwnerServer $OwnerServer -PlanId $PlanId -SafeMinutes $SafeMinutes -HeartbeatRootPath $HeartbeatRootPath -HeartbeatStaleMinutes $HeartbeatStaleMinutes
                    }
                }
            }
            return [pscustomobject]@{ Acquired = $false; Reused = $false; ClaimPath = $claimPath; Claim = $existing; Reason = "Occurrence already claimed by $($existing.OwnerServer) with status $($existing.Status)." }
        }
        catch {
            return [pscustomobject]@{ Acquired = $false; Reused = $false; ClaimPath = $claimPath; Claim = $null; Reason = "Claim exists but cannot be read safely: $($_.Exception.Message)" }
        }
    }
}

function Set-SmartM365OrchestratorOccurrenceClaim {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
    param(
        [Parameter(Mandatory = $true)][string]$ClaimPath,
        [Parameter(Mandatory = $true)][string]$OwnerServer,
        [Parameter(Mandatory = $true)][ValidateSet('Claimed', 'Running', 'RetryScheduled', 'Success', 'CompletedWithWarnings', 'Failed', 'TimedOut', 'Interrupted', 'Cancelled')][string]$Status,
        [int]$Attempt = 0,
        [int]$SafeMinutes = 60,
        [string]$Detail = ''
    )

    $ClaimPath = Resolve-DistributedJsonPath $ClaimPath
    $claimDocument = Read-SmartM365JsonDocument $ClaimPath
    $claim = $claimDocument.Document
    if ([string]$claim.OwnerServer -ne $OwnerServer) {
        throw "Claim owner mismatch for $ClaimPath. Expected $OwnerServer, found $($claim.OwnerServer)."
    }
    $claim.Status = $Status
    $claim.Attempt = $Attempt
    $claim.UpdatedAtUtc = [datetime]::UtcNow.ToString('o')
    $claim.SafeUntilUtc = [datetime]::UtcNow.AddMinutes([math]::Max(1, $SafeMinutes)).ToString('o')
    if (-not [string]::IsNullOrWhiteSpace($Detail)) {
        if ($claim.PSObject.Properties.Name -contains 'Detail') { $claim.Detail = $Detail }
        else { $claim | Add-Member -NotePropertyName Detail -NotePropertyValue $Detail }
    }
    if ($PSCmdlet.ShouldProcess($ClaimPath, "Set occurrence claim status to $Status")) {
        Write-JsonAtomically -Path $ClaimPath -Value $claim -ExpectedSHA256 $claimDocument.SHA256
    }
    return $claim
}

function Get-SmartM365OrchestratorConcurrencyLease {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$LeasesRootPath,
        [Parameter(Mandatory = $true)][string]$ConcurrencyKey
    )

    $leasePath = Join-Path -Path $LeasesRootPath -ChildPath ((ConvertTo-SafeFileName $ConcurrencyKey) + '.json'); $leasePath = Resolve-DistributedJsonPath $leasePath
    if (-not (Test-Path -LiteralPath $leasePath -PathType Leaf)) { return $null }
    $lease = (Read-SmartM365JsonDocument $leasePath).Document
    return [pscustomobject]@{
        LeasePath = $leasePath
        Lease = $lease
    }
}

function Enter-SmartM365OrchestratorConcurrencyLease {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$LeasesRootPath,
        [Parameter(Mandatory = $true)][string]$ConcurrencyKey,
        [Parameter(Mandatory = $true)][string]$JobName,
        [Parameter(Mandatory = $true)][datetime]$Occurrence,
        [Parameter(Mandatory = $true)][string]$OwnerServer,
        [int]$SafeMinutes = 60,
        [string]$HeartbeatRootPath = '',
        [int]$HeartbeatStaleMinutes = 5
    )

    [void](New-Item -ItemType Directory -Path $LeasesRootPath -Force)
    $leasePath = Join-Path -Path $LeasesRootPath -ChildPath ((ConvertTo-SafeFileName $ConcurrencyKey) + '.json'); $leasePath = Resolve-DistributedJsonPath $leasePath
    $owner = $OwnerServer.ToUpperInvariant()
    $occurrenceUtc = $Occurrence.ToUniversalTime()
    $lease = [ordered]@{
        SchemaVersion = 1
        LeaseId = [guid]::NewGuid().ToString('N')
        ConcurrencyKey = $ConcurrencyKey
        JobName = $JobName
        OccurrenceUtc = $occurrenceUtc.ToString('o')
        OwnerServer = $owner
        OrchestratorPid = $PID
        CreatedAtUtc = [datetime]::UtcNow.ToString('o')
        UpdatedAtUtc = [datetime]::UtcNow.ToString('o')
        SafeUntilUtc = [datetime]::UtcNow.AddMinutes([math]::Max(1, $SafeMinutes)).ToString('o')
    }
    $json = $lease | ConvertTo-Json -Depth 6

    try {
        $null = Write-SmartM365JsonBytesAtomically -Path $leasePath -Bytes ([Text.UTF8Encoding]::new($false).GetBytes($json)) -ExpectedSHA256 'ABSENT' -Validate { param($document) if (-not $document.PSObject.Properties['LeaseId']) { throw 'LeaseId missing.' } }
        return [pscustomobject]@{ Acquired = $true; Reused = $false; LeasePath = $leasePath; Lease = [pscustomobject]$lease; Reason = '' }
    }
    catch [System.IO.IOException] {
        try {
            $existing = (Read-SmartM365JsonDocument $leasePath).Document
            $sameOwner = [string]$existing.OwnerServer -eq $owner
            $sameProcess = $false
            try { $sameProcess = [int]$existing.OrchestratorPid -eq $PID } catch { $sameProcess = $false }
            $existingOccurrenceUtc = [datetime]::MinValue
            try { $existingOccurrenceUtc = ([datetime]$existing.OccurrenceUtc).ToUniversalTime() } catch { $existingOccurrenceUtc = [datetime]::MinValue }
            $sameOccurrence = [string]$existing.JobName -eq $JobName -and [math]::Abs(($existingOccurrenceUtc - $occurrenceUtc).TotalSeconds) -le 1
            if ($sameOwner -and $sameProcess -and $sameOccurrence) {
                return [pscustomobject]@{ Acquired = $true; Reused = $true; LeasePath = $leasePath; Lease = $existing; Reason = 'Existing concurrency lease belongs to this process and occurrence.' }
            }

            $safeUntilUtc = [datetime]::MaxValue
            try { $safeUntilUtc = ([datetime]$existing.SafeUntilUtc).ToUniversalTime() } catch { $safeUntilUtc = [datetime]::MaxValue }
            if ([datetime]::UtcNow -gt $safeUntilUtc) {
                $heartbeatFresh = $false
                if (-not [string]::IsNullOrWhiteSpace($HeartbeatRootPath)) {
                    try {
                        $heartbeatPath = Join-Path -Path (Join-Path -Path $HeartbeatRootPath -ChildPath ([string]$existing.OwnerServer)) -ChildPath 'Orchestrator-Heartbeat.json'
                        $heartbeat = if(Get-SmartM365JsonReadPath $heartbeatPath -Optional){(Read-SmartM365JsonDocument $heartbeatPath).Document}else{$null}
                        $heartbeatAgeMinutes = if($heartbeat){([datetime]::UtcNow - ([datetime]$heartbeat.Timestamp).ToUniversalTime()).TotalMinutes}else{[double]::PositiveInfinity}
                        $heartbeatFresh = $heartbeatAgeMinutes -le [math]::Max(1, $HeartbeatStaleMinutes)
                    }
                    catch { throw ('Cannot determine peer liveness; existing coordination state preserved: ' + $_.Exception.Message) }
                }
                if ($sameOwner) { $heartbeatFresh = $false }

                if (-not $heartbeatFresh) {
                    $takeoverLockPath = (Get-SmartM365JsonNames $leasePath).Legacy + '.takeover.lock'
                    $takeoverStream = $null
                    try {
                        if ((Test-Path -LiteralPath $takeoverLockPath -PathType Leaf) -and
                            ([datetime]::UtcNow - (Get-Item -LiteralPath $takeoverLockPath).LastWriteTimeUtc).TotalMinutes -gt 5) {
                            Remove-Item -LiteralPath $takeoverLockPath -Force -ErrorAction SilentlyContinue
                        }
                        $takeoverStream = [System.IO.File]::Open($takeoverLockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
                        $confirmed = (Read-SmartM365JsonDocument $leasePath).Document
                        $confirmedSafeUntilUtc = ([datetime]$confirmed.SafeUntilUtc).ToUniversalTime()
                        if ([string]$confirmed.LeaseId -eq [string]$existing.LeaseId -and [datetime]::UtcNow -gt $confirmedSafeUntilUtc) {
                            $archivePath = '{0}.stale.{1}.json' -f (Get-SmartM365JsonNames $leasePath).Legacy, [string]$confirmed.LeaseId
                            if ((Get-SmartM365JsonTransportPolicy).Mode -eq 'JsonText') { $archivePath += '.txt' }
                            Move-Item -LiteralPath $leasePath -Destination $archivePath -ErrorAction Stop
                        }
                    }
                    catch [System.IO.IOException] {
                        return [pscustomobject]@{ Acquired = $false; Reused = $false; LeasePath = $leasePath; Lease = $existing; Reason = 'Another orchestrator is evaluating takeover of the expired concurrency lease.' }
                    }
                    finally {
                        if ($null -ne $takeoverStream) {
                            $takeoverStream.Dispose()
                            Remove-Item -LiteralPath $takeoverLockPath -Force -ErrorAction SilentlyContinue
                        }
                    }
                    if (-not (Test-Path -LiteralPath $leasePath)) {
                        return Enter-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $LeasesRootPath -ConcurrencyKey $ConcurrencyKey -JobName $JobName -Occurrence $Occurrence -OwnerServer $OwnerServer -SafeMinutes $SafeMinutes -HeartbeatRootPath $HeartbeatRootPath -HeartbeatStaleMinutes $HeartbeatStaleMinutes
                    }
                }
            }
            $reason = "ConcurrencyKey '$ConcurrencyKey' is held by job $($existing.JobName) on $($existing.OwnerServer) until $($existing.SafeUntilUtc)."
            return [pscustomobject]@{ Acquired = $false; Reused = $false; LeasePath = $leasePath; Lease = $existing; Reason = $reason }
        }
        catch {
            return [pscustomobject]@{ Acquired = $false; Reused = $false; LeasePath = $leasePath; Lease = $null; Reason = "Concurrency lease exists but cannot be read safely: $($_.Exception.Message)" }
        }
    }
}

function Set-SmartM365OrchestratorConcurrencyLease {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$LeasePath,
        [Parameter(Mandatory = $true)][string]$LeaseId,
        [Parameter(Mandatory = $true)][string]$OwnerServer,
        [Parameter(Mandatory = $true)][datetime]$SafeUntilUtc
    )

    $LeasePath = Resolve-DistributedJsonPath $LeasePath
    if (-not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return $false }
    $takeoverLockPath = (Get-SmartM365JsonNames $LeasePath).Legacy + '.takeover.lock'
    $lockStream = $null
    try {
        for ($attempt = 1; $attempt -le 10 -and $null -eq $lockStream; $attempt++) {
            try {
                $lockStream = [System.IO.File]::Open($takeoverLockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
            }
            catch [System.IO.IOException] {
                if ($attempt -eq 10) { throw }
                Start-Sleep -Milliseconds 100
            }
        }

        if (-not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return $false }
        $current = (Read-SmartM365JsonDocument $LeasePath).Document
        if ([string]$current.LeaseId -ne $LeaseId -or [string]$current.OwnerServer -ne $OwnerServer.ToUpperInvariant()) {
            return $false
        }

        $targetSafeUntilUtc = $SafeUntilUtc.ToUniversalTime()
        $currentSafeUntilUtc = [datetime]::MinValue
        $hasCurrentSafeUntil = [datetime]::TryParse(
            [string]$current.SafeUntilUtc,
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::RoundtripKind,
            [ref]$currentSafeUntilUtc
        )
        $sameDeadline = $hasCurrentSafeUntil -and
            [math]::Abs(($currentSafeUntilUtc.ToUniversalTime() - $targetSafeUntilUtc).TotalSeconds) -le 1
        $sameProcess = $false
        try { $sameProcess = [int]$current.OrchestratorPid -eq $PID } catch { $sameProcess = $false }
        if ($sameDeadline -and $sameProcess) { return $true }

        $current.OrchestratorPid = $PID
        $current.UpdatedAtUtc = [datetime]::UtcNow.ToString('o')
        $current.SafeUntilUtc = $targetSafeUntilUtc.ToString('o')
        Write-JsonAtomically -Path $LeasePath -Value $current
        return $true
    }
    finally {
        if ($null -ne $lockStream) {
            $lockStream.Dispose()
            Remove-Item -LiteralPath $takeoverLockPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Exit-SmartM365OrchestratorConcurrencyLease {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$LeasePath,
        [Parameter(Mandatory = $true)][string]$LeaseId,
        [Parameter(Mandatory = $true)][string]$OwnerServer
    )

    $LeasePath = Resolve-DistributedJsonPath $LeasePath
    if (-not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return $true }
    $takeoverLockPath = (Get-SmartM365JsonNames $LeasePath).Legacy + '.takeover.lock'
    $lockStream = $null
    try {
        for ($attempt = 1; $attempt -le 10 -and $null -eq $lockStream; $attempt++) {
            try {
                $lockStream = [System.IO.File]::Open($takeoverLockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
            }
            catch [System.IO.IOException] {
                if ($attempt -eq 10) { throw }
                Start-Sleep -Milliseconds 100
            }
        }
        if (-not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return $true }
        $current = (Read-SmartM365JsonDocument $LeasePath).Document
        if ([string]$current.LeaseId -ne $LeaseId -or [string]$current.OwnerServer -ne $OwnerServer.ToUpperInvariant()) {
            return $false
        }
        $consumedLease = Read-SmartM365JsonDocument $LeasePath
        if ($consumedLease.Document.LeaseId -ne $LeaseId -or $consumedLease.Document.OwnerServer -ne $OwnerServer.ToUpperInvariant()) { return $false }
        Complete-SmartM365JsonConsumption -Path $LeasePath -Owner 'Orchestrator concurrency lease' -ExpectedSHA256 $consumedLease.SHA256
        return $true
    }
    finally {
        if ($null -ne $lockStream) {
            $lockStream.Dispose()
            Remove-Item -LiteralPath $takeoverLockPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-SmartM365OrchestratorFreshProducerReceipt {
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)]$Job,
        [Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)][datetime]$Now,
        [Parameter(Mandatory)][double]$MaxAgeHours
    )

    if ($MaxAgeHours -le 0 -or [string]::IsNullOrWhiteSpace($TenantKey)) { return $null }
    $scriptName = [IO.Path]::GetFileName([string]$Job.ScriptPath)
    if ($scriptName -notmatch '^SmartM365-[A-Za-z0-9-]+\.ps1$') { return $null }
    # One script can back several jobs with different arguments; its receipt cannot identify those variants.
    if ([IO.Path]::GetFileNameWithoutExtension($scriptName) -cne ('SmartM365-' + [string]$Job.Name)) { return $null }
    if ([IO.Path]::GetFileName($SharedDataFolderPath.TrimEnd([char[]]'\/')) -cne 'Orchestrator') { return $null }
    $dataAll = Split-Path -Path $SharedDataFolderPath -Parent
    if ([IO.Path]::GetFileName($dataAll) -cne 'DATA-ALL') { return $null }
    $latest = Join-Path (Split-Path -Path $dataAll -Parent) 'DATA-LAST'
    $receiptPath = Join-Path $latest ('SmartInventory_{0}.current.json' -f [IO.Path]::GetFileNameWithoutExtension($scriptName))
    try {
        $readPath = Get-SmartM365JsonReadPath -Path $receiptPath -Optional
        if (-not $readPath) { return $null }
        if ((Get-Item -LiteralPath $readPath).Attributes -band [IO.FileAttributes]::ReparsePoint) { return $null }
        $receiptHash = (Get-FileHash -LiteralPath $readPath -Algorithm SHA256).Hash
        $receipt = (Read-SmartM365JsonDocument -Path $readPath).Document
        foreach ($field in @('Owner','Status','TenantKey','Producer','RunId','StartedAtUtc','CompletedAtUtc','Errors','IsPartialInventory','Files')) {
            if (-not $receipt.PSObject.Properties[$field]) { return $null }
        }
        $version = if ($receipt.PSObject.Properties['ContractVersion']) { [string]$receipt.ContractVersion } else { '' }
        $supported = ([string]$receipt.Owner -ceq 'SmartInventory-CmdbSourceReceipt' -and $version -in @('', '1.1')) -or
            ([string]$receipt.Owner -ceq 'SmartInventory-SourceReceipt' -and $version -ceq '1.2')
        if (-not $supported -or
            [string]$receipt.Status -cne 'Completed' -or
            [string]$receipt.TenantKey -cne $TenantKey -or
            [string]$receipt.Producer -cne $scriptName -or
            [string]::IsNullOrWhiteSpace([string]$receipt.RunId) -or
            [int]$receipt.Errors -ne 0 -or
            $receipt.IsPartialInventory -isnot [bool] -or $receipt.IsPartialInventory) { return $null }

        # Completion proof is not the latest attempt. Never release on a canonical replacement.
        $runPath = Join-Path $latest ('SmartInventory_{0}.run.json' -f [IO.Path]::GetFileNameWithoutExtension($scriptName))
        $runReadPath = Get-SmartM365JsonReadPath -Path $runPath -Optional
        $protocol = $receipt.PSObject.Properties['PublicationProtocol']
        if ($protocol -and (($protocol.Value -isnot [int] -and $protocol.Value -isnot [long]) -or $protocol.Value -ne 1)) { return $null }
        if ($protocol -and -not $runReadPath) { return $null }
        $runHash = ''
        if ($runReadPath) {
            if ((Get-Item -LiteralPath $runReadPath).Attributes -band [IO.FileAttributes]::ReparsePoint) { return $null }
            $runHash = (Get-FileHash -LiteralPath $runReadPath -Algorithm SHA256).Hash
            $run = (Read-SmartM365JsonDocument -Path $runPath).Document
            if ($run.Owner -cne 'SmartInventory-SourceRun' -or $run.ContractVersion -cne '1.0' -or
                $run.Status -notin @('Collecting','Publishing','Completed','Failed') -or
                $run.PublicationStarted -isnot [bool] -or $run.UnqualifiedPublication -isnot [bool] -or
                -not $run.RunId -or $run.Status -eq 'Publishing' -or $run.UnqualifiedPublication) { return $null }
            foreach ($field in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId','Producer')) {
                if (-not $receipt.PSObject.Properties[$field] -or $run.$field -cne $receipt.$field) { return $null }
            }
        }

        $started = [datetimeoffset]::MinValue
        $completed = [datetimeoffset]::MinValue
        $styles = [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
        if (-not [datetimeoffset]::TryParse([string]$receipt.StartedAtUtc, [Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$started) -or
            -not [datetimeoffset]::TryParse([string]$receipt.CompletedAtUtc, [Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$completed)) { return $null }
        $nowUtc = $Now.ToUniversalTime()
        if ($started -gt $completed -or $completed.UtcDateTime -gt $nowUtc.AddMinutes(5) -or
            $completed.UtcDateTime -lt $nowUtc.AddHours(-1 * $MaxAgeHours)) { return $null }

        $files = @($receipt.Files)
        if ($files.Count -eq 0) { return $null }
        $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($file in $files) {
            foreach ($field in @('File','RunId','Producer','Status','Errors','IsPartialInventory','SHA256')) {
                if (-not $file.PSObject.Properties[$field]) { return $null }
            }
            $fileName = [string]$file.File
            if ([string]::IsNullOrWhiteSpace($fileName) -or $fileName -match '[/\\:]' -or $fileName -in @('.', '..') -or
                $fileName -cne [IO.Path]::GetFileName($fileName) -or -not $seen.Add($fileName) -or
                [string]$file.RunId -cne [string]$receipt.RunId -or
                [string]$file.Producer -cne $scriptName -or
                [string]$file.Status -cne 'Success' -or [int]$file.Errors -ne 0 -or
                $null -eq $file.IsPartialInventory -or [bool]$file.IsPartialInventory -or
                [string]$file.SHA256 -notmatch '^[A-Fa-f0-9]{64}$') { return $null }
            $csvPath = Join-Path $latest $fileName
            $csv = Get-Item -LiteralPath $csvPath -ErrorAction SilentlyContinue
            if ($null -eq $csv -or $csv.PSIsContainer -or $csv.LinkType -in @('SymbolicLink', 'Junction')) { return $null }
            if ([string]$receipt.Owner -ceq 'SmartInventory-SourceReceipt' -and
                (Get-FileHash -LiteralPath $csvPath -Algorithm SHA256).Hash -ine [string]$file.SHA256) { return $null }
        }
        if ((Get-SmartM365JsonReadPath -Path $receiptPath -Optional) -ne $readPath) { return $null }
        if ((Get-FileHash -LiteralPath $readPath -Algorithm SHA256).Hash -ne $receiptHash) { return $null }
        if ((Get-SmartM365JsonReadPath -Path $runPath -Optional) -ne $runReadPath) { return $null }
        if ($runReadPath -and (Get-FileHash -LiteralPath $runReadPath -Algorithm SHA256).Hash -ne $runHash) { return $null }
        return $completed.UtcDateTime
    }
    catch { return $null }
}

function ConvertTo-DistributedEvidenceUtc {
    param([Parameter(Mandatory)]$Value)
    if ($Value -is [datetimeoffset]) { return $Value.ToUniversalTime() }
    if ($Value -is [datetime] -and $Value.Kind -ne [DateTimeKind]::Unspecified) { return [datetimeoffset]$Value.ToUniversalTime() }
    if ([string]$Value -notmatch '(Z|[+-]\d{2}:\d{2})$') { throw 'Evidence timestamp has no explicit time zone.' }
    return [datetimeoffset]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture).ToUniversalTime()
}

function Get-SmartM365OrchestratorPinnedDependencyReadiness {
    <# Read-only, fail-closed FreshSuccess gate for a uniquely pinned producer, local or remote. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)]$Job,
        [Parameter(Mandatory)][datetime]$Now,
        [Parameter(Mandatory)][double]$MaxAgeHours,
        [string]$ProfileKey = '',
        [double]$HeartbeatStaleMinutes = 5
    )
    $result = [ordered]@{ State = 'Waiting'; Detail = 'Pinned producer state unavailable'; LastSuccessUtc = $null }
    try {
        $owners = @($Job.AllowedServers | Sort-Object -Unique)
        if ($owners.Count -ne 1 -or [string]$owners[0] -notmatch '^[A-Za-z0-9_-]+$' -or $MaxAgeHours -le 0 -or $HeartbeatStaleMinutes -le 0) {
            $result.Detail = 'Pinned dependency requires one explicit owner and positive freshness limits'
            return [pscustomobject]$result
        }
        $folder = Join-Path $SharedDataFolderPath ([string]$owners[0])
        $heartbeatPath = Get-SmartM365JsonReadPath (Join-Path $folder 'Orchestrator-Heartbeat.json')
        $statePath = Get-SmartM365JsonReadPath (Join-Path $folder 'Orchestrator-State.json')
        $before = @{}
        foreach ($path in @($heartbeatPath, $statePath)) {
            if ((Get-Item -LiteralPath $path).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Linked dependency state is not accepted.' }
            $before[$path] = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        }
        $heartbeat = (Read-SmartM365JsonDocument $heartbeatPath).Document
        $nowUtc = $Now.ToUniversalTime()
        $stamp = ConvertTo-DistributedEvidenceUtc $heartbeat.Timestamp
        if ($stamp.UtcDateTime -gt $nowUtc.AddMinutes(1) -or ($nowUtc - $stamp.UtcDateTime).TotalMinutes -gt $HeartbeatStaleMinutes -or
            $heartbeat.Lifecycle -ne 'Running' -or $heartbeat.StatePersistenceHealthy -isnot [bool] -or -not $heartbeat.StatePersistenceHealthy -or
            ($ProfileKey -and $heartbeat.Tenant -cne $ProfileKey)) {
            $result.Detail = 'Pinned owner heartbeat is stale, unhealthy or belongs to another profile'
            return [pscustomobject]$result
        }
        $state = (Read-SmartM365JsonDocument $statePath).Document
        if ($state.SchemaVersion -ne 1) { throw 'Unsupported dependency state schema.' }
        $entry = $state.Jobs.PSObject.Properties[[string]$Job.Name]
        if (-not $entry) { return [pscustomobject]$result }
        $record = $entry.Value
        if (@($heartbeat.RunningJobs | Where-Object { $_.Name -eq $Job.Name }).Count -gt 0 -or $null -ne $record.Running -or $null -ne $record.PendingRetry) {
            $result.Detail = 'Pinned dependency is running or awaiting retry'
            return [pscustomobject]$result
        }
        if ($record.LastStatus -notin @('Success', 'CompletedWithWarnings')) {
            $result.Detail = 'Pinned dependency has no successful latest execution: ' + [string]$record.LastStatus
            return [pscustomobject]$result
        }
        $end = ConvertTo-DistributedEvidenceUtc $record.LastRunEnd
        $start = ConvertTo-DistributedEvidenceUtc $record.LastRunStart
        if ($start -gt $end -or $end.UtcDateTime -gt $nowUtc -or ($nowUtc - $end.UtcDateTime).TotalHours -gt $MaxAgeHours) {
            $result.Detail = 'Pinned dependency success is stale or has invalid timestamps'
            return [pscustomobject]$result
        }
        # A lease closes the interval before a newly started run reaches state/heartbeat files.
        $key = if ($Job.PSObject.Properties['ConcurrencyKey'] -and $Job.ConcurrencyKey) { [string]$Job.ConcurrencyKey } else { [string]$Job.Name }
        $leasePath = Join-Path $SharedDataFolderPath ('Election/Concurrency/' + (ConvertTo-SafeFileName $key) + '.json')
        if (Get-SmartM365JsonReadPath $leasePath -Optional) {
            $result.Detail = 'Pinned dependency has a concurrency lease; waiting for owner reconciliation'
            return [pscustomobject]$result
        }
        foreach ($path in $before.Keys) {
            if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $before[$path]) { throw 'Dependency state changed during inspection.' }
        }
        $result.State = 'Ready'; $result.Detail = 'Fresh pinned producer success: ' + [string]$record.LastStatus
        $result.LastSuccessUtc = $end.UtcDateTime
    }
    catch { $result.Detail = 'Pinned producer state is missing or invalid' }
    return [pscustomobject]$result
}

function Get-SmartM365OrchestratorSourceRunState {
    <# Read-only view of a producer run state (SmartInventory_<script>.run.json) beside the shared DATA-LAST. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)]$Job
    )
    $result = [ordered]@{ Known = $false; StagedPublication = $false; Status = ''; StartedAtUtc = $null; Detail = 'Producer run state unavailable' }
    try {
        $scriptName = [IO.Path]::GetFileName([string]$Job.ScriptPath)
        if ($scriptName -notmatch '^SmartM365-[A-Za-z0-9-]+\.ps1$') { $result.Detail = 'Producer script has no run state'; return [pscustomobject]$result }
        if ([IO.Path]::GetFileName($SharedDataFolderPath.TrimEnd([char[]]'\/')) -cne 'Orchestrator') { return [pscustomobject]$result }
        $dataAll = Split-Path -Path $SharedDataFolderPath -Parent
        if ([IO.Path]::GetFileName($dataAll) -cne 'DATA-ALL') { return [pscustomobject]$result }
        $latest = Join-Path (Split-Path -Path $dataAll -Parent) 'DATA-LAST'
        $runPath = Join-Path $latest ('SmartInventory_{0}.run.json' -f [IO.Path]::GetFileNameWithoutExtension($scriptName))
        $readPath = Get-SmartM365JsonReadPath -Path $runPath -Optional
        if (-not $readPath) { $result.Detail = 'Producer run state is absent'; return [pscustomobject]$result }
        if ((Get-Item -LiteralPath $readPath).Attributes -band [IO.FileAttributes]::ReparsePoint) { return [pscustomobject]$result }
        $run = (Read-SmartM365JsonDocument -Path $runPath).Document
        if ([string]$run.Owner -cne 'SmartInventory-SourceRun' -or [string]$run.Producer -cne $scriptName -or
            [string]$run.Status -notin @('Collecting', 'Publishing', 'Completed', 'Failed')) { $result.Detail = 'Producer run state is invalid'; return [pscustomobject]$result }
        $staged = $run.PSObject.Properties['StagedPublication']
        $result.Known = $true
        $result.StagedPublication = $null -ne $staged -and $staged.Value -is [bool] -and [bool]$staged.Value
        $result.Status = [string]$run.Status
        $started = [datetimeoffset]::MinValue
        if ([datetimeoffset]::TryParse([string]$run.StartedAtUtc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$started)) { $result.StartedAtUtc = $started.UtcDateTime }
        $result.Detail = 'Producer run {0}{1}' -f $result.Status, $(if ($result.StagedPublication) { ', staged publication' } else { ', direct publication' })
    }
    catch { $result.Known = $false; $result.Detail = 'Producer run state could not be read: ' + $_.Exception.Message }
    return [pscustomobject]$result
}

function Get-SmartM365OrchestratorDependencyActivity {
    <#
    FreshSuccess companion gate: may a downstream occurrence start while this dependency is active?
    - An active run (claim Claimed, Running or RetryScheduled within its safe window) expected to end
      by the downstream occurrence (occurrence + estimated duration) blocks until it ends.
    - A longer active run blocks unless its producer publishes through staging and is not promoting:
      such a run leaves DATA-LAST untouched, so the last fresh success stays consistent.
    - A due occurrence without a claim blocks during the claim grace only, if it was expected to end
      by the downstream occurrence; afterwards it is queued, skipped or blocked elsewhere.
    Otherwise the dependency is Clear and the freshness rule applies. Read-only.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClaimsRootPath,
        [Parameter(Mandatory)]$DependencyJob,
        [Parameter(Mandatory)][datetime]$DownstreamOccurrence,
        [Parameter(Mandatory)][datetime]$Now,
        [AllowNull()]$LatestPastOccurrence = $null,
        [int]$ClaimGraceMinutes = 15,
        [AllowNull()]$RunState = $null
    )
    $nowUtc = $Now.ToUniversalTime()
    $downstreamUtc = $DownstreamOccurrence.ToUniversalTime()
    $estimate = 60.0
    if ($DependencyJob.PSObject.Properties['EstimatedDurationMinutes'] -and [double]$DependencyJob.EstimatedDurationMinutes -gt 0) { $estimate = [double]$DependencyJob.EstimatedDurationMinutes }
    $folder = Join-Path -Path $ClaimsRootPath -ChildPath (ConvertTo-SafeFileName ([string]$DependencyJob.Name))
    $claimed = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $active = $null
    if (Test-Path -LiteralPath $folder -PathType Container) {
        foreach ($file in @(Get-ChildItem -LiteralPath $folder -File -Force -ErrorAction Stop)) {
            # Only canonical claim files; stale archives and takeover locks are not executions.
            if ($file.Name -notmatch '^(\d{8}T\d{9}Z)\.json(?:\.txt)?$') { continue }
            $occurrenceUtc = [datetime]::ParseExact($Matches[1], 'yyyyMMddTHHmmssfffZ', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal)
            $stamp = $Matches[1]
            if (-not $claimed.Add($stamp)) { continue }
            if ($occurrenceUtc -gt $nowUtc -or $occurrenceUtc -lt $nowUtc.AddDays(-8)) { continue }
            $claim = (Read-SmartM365JsonDocument -Path (Join-Path -Path $folder -ChildPath ($stamp + '.json'))).Document
            if ([string]$claim.JobName -ne [string]$DependencyJob.Name) { throw "Dependency claim identity mismatch: $($file.Name)" }
            if ([string]$claim.Status -notin @('Claimed', 'Running', 'RetryScheduled')) { continue }
            # An expired non-terminal claim belongs to a lost owner; claim takeover handles it.
            $safeUntil = [datetimeoffset]::MinValue
            if ([datetimeoffset]::TryParse([string]$claim.SafeUntilUtc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$safeUntil) -and $safeUntil.UtcDateTime -lt $nowUtc) { continue }
            if ($null -eq $active -or $occurrenceUtc -gt $active.OccurrenceUtc) {
                # JSON readers may return the claim time as a DateTime: keep a culture-neutral text.
                $claimedAt = if ($claim.PSObject.Properties['CreatedAtUtc']) { $claim.CreatedAtUtc } else { '' }
                if ($claimedAt -is [datetime]) { $claimedAt = $claimedAt.ToUniversalTime().ToString('o') }
                $active = [pscustomobject]@{ OccurrenceUtc = $occurrenceUtc; Status = [string]$claim.Status; OwnerServer = [string]$claim.OwnerServer; ClaimedAtUtc = [string]$claimedAt }
            }
        }
    }
    if ($null -ne $active) {
        $label = '{0} for occurrence {1}' -f $active.Status, $active.OccurrenceUtc.ToLocalTime().ToString('yyyy-MM-dd HH:mm')
        if ($active.OccurrenceUtc.AddMinutes($estimate) -le $downstreamUtc) {
            return [pscustomobject]@{ State = 'Waiting'; Detail = "$label, expected to finish before this occurrence"; OwnerServer = $active.OwnerServer; ClaimedAtUtc = $active.ClaimedAtUtc; Status = $active.Status }
        }
        if ($null -ne $RunState -and $RunState.Known -and $RunState.StagedPublication -and $RunState.Status -ne 'Publishing') {
            return [pscustomobject]@{ State = 'Clear'; Detail = "$label, staged publication leaves DATA-LAST unchanged" }
        }
        $reason = if ($null -ne $RunState -and $RunState.Known -and $RunState.Status -eq 'Publishing') { 'promoting its CSVs' } else { 'may rewrite DATA-LAST during the run' }
        return [pscustomobject]@{ State = 'Waiting'; Detail = "$label, $reason"; OwnerServer = $active.OwnerServer; ClaimedAtUtc = $active.ClaimedAtUtc; Status = $active.Status }
    }
    if ($null -ne $LatestPastOccurrence) {
        $dueUtc = ([datetime]$LatestPastOccurrence).ToUniversalTime()
        if (-not $claimed.Contains($dueUtc.ToString('yyyyMMddTHHmmssfffZ', [Globalization.CultureInfo]::InvariantCulture)) -and
            $dueUtc -le $nowUtc -and ($nowUtc - $dueUtc).TotalMinutes -le [math]::Max(1, $ClaimGraceMinutes) -and
            $dueUtc.AddMinutes($estimate) -le $downstreamUtc) {
            return [pscustomobject]@{ State = 'Waiting'; Detail = 'occurrence {0} is due and not started yet' -f $dueUtc.ToLocalTime().ToString('yyyy-MM-dd HH:mm') }
        }
    }
    return [pscustomobject]@{ State = 'Clear'; Detail = 'No active execution' }
}

Export-ModuleMember -Function @(
    'Get-SmartM365OrchestratorServerCapability',
    'Test-SmartM365OrchestratorCapabilityComponentMissing',
    'Test-SmartM365OrchestratorCapabilityMatch',
    'Test-SmartM365OrchestratorCanPreserveOwners',
    'Get-SmartM365OrchestratorElectionPlan',
    'Get-SmartM365OrchestratorOccurrenceClaim',
    'Enter-SmartM365OrchestratorOccurrenceClaim',
    'Set-SmartM365OrchestratorOccurrenceClaim',
    'Get-SmartM365OrchestratorConcurrencyLease',
    'Enter-SmartM365OrchestratorConcurrencyLease',
    'Set-SmartM365OrchestratorConcurrencyLease',
    'Exit-SmartM365OrchestratorConcurrencyLease',
    'Convert-SmartM365OrchestratorDistributedHistory',
    'Get-SmartM365OrchestratorFreshProducerReceipt',
    'Get-SmartM365OrchestratorPinnedDependencyReadiness',
    'Get-SmartM365OrchestratorSourceRunState',
    'Get-SmartM365OrchestratorDependencyActivity',
    'Get-SmartM365OrchestratorCertificateCandidate',
    'Get-SmartM365OrchestratorCachedCertificateReadiness',
    'Test-SmartM365OrchestratorPathAccess',
    'Test-SmartM365OrchestratorLaunchPrerequisite',
    'Get-SmartM365OrchestratorVolumeFreeSpace',
    'Get-SmartM365OrchestratorDiskSpaceDecision'
)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDRE7MqZ3gwKNmV
# sIe1HqZaaXpoUI85DMqWK/kcOQe+VaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAhL8uHrAR8npHnzpMulyQDh23XA80LpkZQWTeVspBfnMA0GCSqG
# SIb3DQEBAQUABIIBgBgI4i1SoDGDOHqV4JSBeiYSLVE2gwz/yoNMe0rugsrARsYN
# +RCeFmEtKugHFmD80gN/1nRzFZpJHXIhL7GvRfvklQKJrW/AgOKQRTFHd9ES90m/
# TGTzZN8+6xE9X9JhYG7xzQuHWxtx/1tWXtfgnSZfaZVcqGBrlXq7Y/v7Ac1ZTNxd
# KW8AINbs7eCrceqRsuYp/su39c0W1O3FOlKMXxpTT5lK9ITCVB1jEa3bbBW+17iR
# EDGYKtFmVr7fxZXY/vBmSgBbhXhMlbuXq3a4MnTUdsm4Xr1l+t+ePBxTNljGkx2s
# yBNYKoSbUVX+4aXih1mMdHGL77eVwfKt5YZDwQuL90Y58DcRzi+SJHMv5tqTPLSQ
# //kO6SoEpw+OvT+M+HpYqS1EkXRNdjhFC+bWBHuq3XE4FaMHSNxx+CKFV4U81Vfw
# ox3f5ZijEtBTu0Sz95BJWNq1te7BzXQ2LUC6H4Ol7mr4E0LYkhujKmXKBIcyVfWE
# 4l1j+LqoVFxv4/f2CaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzAxMDIx
# NDFaMC8GCSqGSIb3DQEJBDEiBCAtSdBXH0H6LKngPDo+47nUjnpMTvPqCCZXMYEM
# OBG2PDANBgkqhkiG9w0BAQEFAASCAgAQjx5FC6UEOpagDFVF0lhC4UxTDeMEcNwm
# TsGtvhuXNMPTQZ2oRY52TFVi/UdCBb+kZ++OExl64b8vi8zId8nYx60kuErKvq+i
# i1aipEvGFO0nc0ZPkIGXq0BCAUJf0vYs/edT+mq9r3Q0wa92GZsPeNyaFXxIMwoy
# dwzGo7+AC4Tu7PGH/Li2ErEoqUEguSsbl/77ckT36UDEVIGsdSIpfs5bqdbx50PV
# TtXy4d5iI4YqNb8l81IBln4d7xgdKxlhRjybLqr7v9ciLyyNi67ZE95T1CAs8xOG
# HdyJ15xcgMZaeYqyHfNu4npk4U0XtzmQuwoPUJaUZtvBqt14I3liCqc71NlwVMgw
# //gjrd5qnKw8xvNHqUzLl9oOG5tv5wg5tmu0Jyuz0lbtKguxccP+XCrEQIao/n7a
# TcblTahHs+sR95T1DFyYNnrm2D8znIm0lCb34zm6H9QW5lFphnkedAZRENHxiA+U
# AIBL1Gj2nEMBbfSMyNwRzz8oVolLT826OZur7qEzChmBFuXEBx4EMnuKp7saeROs
# ++ajyECxTLeVkbqEk4dUnkN1c8YsEsLhuUYkHRuOyjvEzMUi1VGfTs5l6AFUcFpz
# mMyfdyxmpMakvPNRLRSCODRB5vhaQjpMUWAZ7p16R7QA4xAripj1MHK+2PRODGEd
# dEorYPy5Gg==
# SIG # End signature block
