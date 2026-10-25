<#
.SYNOPSIS
    Worker-side remote operations: payload copy, PsExec execution, result correlation, evidence.

.DESCRIPTION
    Each endpoint attempt receives a CorrelationId. The endpoint writes
    <DataRoot>\State\Results\<CorrelationId>.json when it starts (State=Running) and when it
    ends (State=Completed). The worker only trusts a result whose CorrelationId matches the
    attempt: older LastRun.json files, retry-after-reboot results of previous attempts and
    evidence of other Runs can never be attributed to the current attempt.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

function Get-SemtComputerPathKey {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$ComputerName)

    $short = Get-SemtSafeName -Value ($ComputerName.Split('.')[0]) -Fallback 'Computer' -MaxLength 40
    $hash = (Get-SemtStringSha256 -Text $ComputerName.Trim().ToLowerInvariant()).Substring(0, 8).ToLowerInvariant()
    return ('{0}-{1}' -f $short,$hash)
}

function Protect-SemtRemotePathAcl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [switch]$Directory
    )

    $icacls = Join-Path -Path $env:SystemRoot -ChildPath 'System32\icacls.exe'
    if (-not (Test-Path -LiteralPath $icacls -PathType Leaf)) { return 'IcaclsNotFound' }
    $grants = if ($Directory) { @('*S-1-5-18:(OI)(CI)F','*S-1-5-32-544:(OI)(CI)F','*S-1-5-32-545:(OI)(CI)RX') } else { @('*S-1-5-18:F','*S-1-5-32-544:F','*S-1-5-32-545:RX') }
    $output = & $icacls $Path '/inheritance:r' '/grant:r' $grants 2>&1
    if ($LASTEXITCODE -ne 0) { throw ("ACL hardening failed. Path={0}; ExitCode={1}; Output={2}" -f $Path,$LASTEXITCODE,(($output | Out-String).Trim())) }
    return 'Protected'
}

function Copy-SemtRemotePayload {
    <#
    Copies the endpoint script through a staging name, verifies size, SHA256 and version, then
    renames it atomically. The expected SHA256 is supplied by the worker after its own
    verification against the distribution manifest or the launcher contract.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ConnectionTarget,
        [Parameter(Mandatory = $true)][string]$LocalScriptPath,
        [Parameter(Mandatory = $true)][string]$ExpectedSha256,
        [Parameter(Mandatory = $true)][string]$RemoteDataRoot,
        [Parameter(Mandatory = $true)][string]$RemoteProductRoot,
        [Parameter(Mandatory = $true)][string]$RemoteProgramRoot,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [string]$RemoteProgramFilesRoot = (Get-SemtTargetProgramFilesRoot)
    )

    # Data (results, logs, state) stays under ProgramData. The script that PsExec starts as SYSTEM
    # is stored under Program Files: the folder chain is created with a protected ACL (owner
    # Administrators) or verified, and any untrusted component stops the copy.
    $remoteProductShare = ConvertTo-SemtAdminSharePath -ComputerName $ConnectionTarget -LocalPath $RemoteProductRoot
    $remoteDataShare = ConvertTo-SemtAdminSharePath -ComputerName $ConnectionTarget -LocalPath $RemoteDataRoot
    $remoteProgramShare = ConvertTo-SemtAdminSharePath -ComputerName $ConnectionTarget -LocalPath $RemoteProgramRoot
    $remoteProgramFilesShare = ConvertTo-SemtAdminSharePath -ComputerName $ConnectionTarget -LocalPath $RemoteProgramFilesRoot
    New-SemtDirectory -Path $remoteProductShare
    [void](Protect-SemtRemotePathAcl -Path $remoteProductShare -Directory)
    New-SemtDirectory -Path $remoteDataShare
    [void](Initialize-SemtTrustedProgramFolder -Path $remoteProgramShare -BasePath $remoteProgramFilesShare)
    $scriptName = Split-Path -Path $LocalScriptPath -Leaf
    $finalPath = Join-Path -Path $remoteProgramShare -ChildPath $scriptName
    $stagingPath = Join-Path -Path $remoteProgramShare -ChildPath ('.{0}.{1}.tmp' -f $scriptName,([guid]::NewGuid().ToString('N')))
    $localItem = Get-Item -LiteralPath $LocalScriptPath -ErrorAction Stop
    $localVersion = Get-SemtScriptVersion -Path $LocalScriptPath
    try {
        Copy-Item -LiteralPath $LocalScriptPath -Destination $stagingPath -Force -ErrorAction Stop
        $stagingSha = Get-SemtFileSha256 -Path $stagingPath
        if ((Get-Item -LiteralPath $stagingPath).Length -ne $localItem.Length) { throw 'Staged endpoint script size mismatch.' }
        if ($stagingSha -ne $ExpectedSha256) { throw ("Staged endpoint script hash mismatch. Expected={0}; Staged={1}" -f $ExpectedSha256,$stagingSha) }
        Move-Item -LiteralPath $stagingPath -Destination $finalPath -Force -ErrorAction Stop
        $programTrust = Test-SemtTrustedProgramPath -Path $remoteProgramShare -Recurse
        if (-not $programTrust.Trusted) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $remoteProgramShare,$programTrust.State,$programTrust.Detail) }
        $remoteSha = Get-SemtFileSha256 -Path $finalPath
        if ($remoteSha -ne $ExpectedSha256) { throw ("Final endpoint script hash mismatch. Expected={0}; Remote={1}" -f $ExpectedSha256,$remoteSha) }
        $remoteVersion = Get-SemtScriptVersion -Path $finalPath
        if ($remoteVersion -ne $localVersion) { throw ("Final endpoint script version mismatch. Local={0}; Remote={1}" -f $localVersion,$remoteVersion) }
        Write-SemtLogLine -Path $LogPath -Message ("Endpoint script copied and verified: Version={0}; SHA256={1}; Path={2}" -f $remoteVersion,$remoteSha,$finalPath)
        return [pscustomobject]@{ LocalVersion = $localVersion; RemoteSha256 = $remoteSha; RemoteSharePath = $finalPath; RemoteDataShare = $remoteDataShare; RemoteScriptPath = (Join-Path -Path $RemoteProgramRoot -ChildPath $scriptName) }
    }
    finally {
        if (Test-Path -LiteralPath $stagingPath -PathType Leaf) { Remove-Item -LiteralPath $stagingPath -Force -ErrorAction SilentlyContinue }
    }
}

function Invoke-SemtPsExecEndpoint {
    <#
    Starts PsExec, records the PsExec PID in the worker reservation, refreshes the reservation
    heartbeat every 60 seconds and enforces the timeout. A timeout stops only the local PsExec
    process; the remote endpoint state stays unconfirmed until its result is read.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$PsExecPath,
        [Parameter(Mandatory = $true)][string]$ConnectionTarget,
        [Parameter(Mandatory = $true)][string]$RemoteScriptPath,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Arguments,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [int]$TimeoutMinutes = 0,
        [bool]$UseElevatedToken = $false,
        [AllowNull()][string]$LeasePath
    )

    $stdoutPath = '{0}.stdout.txt' -f $LogPath
    $stderrPath = '{0}.stderr.txt' -f $LogPath
    $psexecArguments = @(('\\{0}' -f $ConnectionTarget),'-accepteula','-nobanner','-s')
    if ($UseElevatedToken) { $psexecArguments += '-h' }
    $psexecArguments += @('powershell.exe','-NoProfile','-ExecutionPolicy','Bypass','-File',(ConvertTo-SemtPsExecArgument -Value $RemoteScriptPath))
    $psexecArguments += @($Arguments | ForEach-Object { ConvertTo-SemtPsExecArgument -Value $_ })
    Write-SemtLogLine -Path $LogPath -Message ("PsExec command: {0} {1}" -f $PsExecPath,($psexecArguments -join ' '))
    $process = Start-Process -FilePath $PsExecPath -ArgumentList $psexecArguments -PassThru -NoNewWindow -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
    $started = Get-Date
    [void](Update-SemtWorkerLease -LeasePath $LeasePath -Properties @{ PsExecProcessId = [int]$process.Id; PsExecProcessStartUtc = (Get-SemtProcessStartUtc -ProcessId $process.Id) })
    $timedOut = $false
    $lastHeartbeat = $started
    while (-not $process.WaitForExit(1000)) {
        $now = Get-Date
        if ($TimeoutMinutes -gt 0 -and ($now - $started).TotalMinutes -ge $TimeoutMinutes) {
            $timedOut = $true
            Write-SemtLogLine -Path $LogPath -Level ERROR -Message ("PsExec timed out after {0} minute(s). Stopping the local PsExec process; the remote endpoint state is unconfirmed." -f $TimeoutMinutes)
            try { & taskkill.exe /PID $process.Id /T /F 2>&1 | Out-Null } catch { }
            try { [void]$process.WaitForExit(5000) } catch { }
            break
        }
        if (($now - $lastHeartbeat).TotalSeconds -ge 60) {
            [void](Update-SemtWorkerLease -LeasePath $LeasePath)
            Write-SemtLogLine -Path $LogPath -Message ("PsExec still running for {0:N1} minute(s)." -f ($now - $started).TotalMinutes)
            $lastHeartbeat = $now
        }
    }
    try { $process.Refresh() } catch { }
    $exitCode = ''
    if (-not $timedOut) { try { if ($process.HasExited -and $null -ne $process.ExitCode) { $exitCode = [string]$process.ExitCode } } catch { $exitCode = '' } }
    $stdout = @()
    $stderr = @()
    if (Test-Path -LiteralPath $stdoutPath -PathType Leaf) { $stdout = @(Get-Content -LiteralPath $stdoutPath -ErrorAction SilentlyContinue) }
    if (Test-Path -LiteralPath $stderrPath -PathType Leaf) { $stderr = @(Get-Content -LiteralPath $stderrPath -ErrorAction SilentlyContinue) }
    if ([string]::IsNullOrWhiteSpace($exitCode)) {
        $nativeLine = $stderr | Where-Object { $_ -match 'with error code\s+-?\d+' } | Select-Object -Last 1
        if ($nativeLine -and $nativeLine -match 'with error code\s+(?<Code>-?\d+)') {
            $exitCode = $Matches.Code
            Write-SemtLogLine -Path $LogPath -Message ("PsExec exit code recovered from native STDERR: {0}" -f $exitCode)
        }
    }
    Add-Content -LiteralPath $LogPath -Value '----- PsExec STDOUT -----' -Encoding UTF8
    if ($stdout.Count -gt 0) { Add-Content -LiteralPath $LogPath -Value $stdout -Encoding UTF8 }
    Add-Content -LiteralPath $LogPath -Value '----- PsExec STDERR -----' -Encoding UTF8
    if ($stderr.Count -gt 0) { Add-Content -LiteralPath $LogPath -Value $stderr -Encoding UTF8 }
    return [pscustomobject]@{
        ExitCode = $exitCode
        TimedOut = $timedOut
        Stdout = $stdout
        Stderr = $stderr
        StdoutPath = $stdoutPath
        StderrPath = $stderrPath
        FinalStatus = ConvertFrom-SemtFinalStatusLine -Lines $stdout
        CommunicationFailureDetail = Get-SemtPsExecCommunicationFailureDetail -Lines $stderr
        ElapsedMinutes = [math]::Round(((Get-Date) - $started).TotalMinutes, 1)
    }
}

function ConvertFrom-SemtFinalStatusLine {
    [CmdletBinding()]
    param([AllowEmptyCollection()][string[]]$Lines)

    $line = @($Lines) | Where-Object { $_ -match '^FINAL_STATUS=' } | Select-Object -Last 1
    if (-not $line) { return $null }
    if ($line -match '^FINAL_STATUS=(?<Status>[^;]*);\s*EXIT_CODE=(?<ExitCode>[^;]*);\s*NEXT_ACTION=(?<NextAction>[^;]*);\s*(?:CORRELATION_ID=(?<CorrelationId>[^;]*);\s*)?DETAIL=(?<Detail>.*)$') {
        # The CorrelationId group is absent from lines written by older endpoint versions.
        $correlation = if ($Matches.ContainsKey('CorrelationId')) { ([string]$Matches['CorrelationId']).Trim() } else { '' }
        return [pscustomobject]@{ Status = $Matches.Status.Trim(); ExitCode = $Matches.ExitCode.Trim(); NextAction = $Matches.NextAction.Trim(); CorrelationId = $correlation; Detail = $Matches.Detail.Trim() }
    }
    return $null
}

function Read-SemtRemoteAttemptResult {
    <#
    Reads <DataRoot>\State\Results\<CorrelationId>.json from the target. Only a document whose
    CorrelationId matches is returned; anything else is reported as foreign evidence.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RemoteDataShare,
        [Parameter(Mandatory = $true)][string]$CorrelationId
    )

    $path = Join-Path -Path $RemoteDataShare -ChildPath ("State\Results\{0}.json" -f $CorrelationId)
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        return [pscustomobject]@{ Exists = $false; State = ''; Completed = $false; Result = $null; Path = $path; Detail = 'No result for this attempt on the target.' }
    }
    try {
        $document = Read-SemtJsonFile -Path $path -Required
        $documentCorrelation = [string](Get-SemtPropertyValue -InputObject $document -Name 'CorrelationId' -Default '')
        if ($documentCorrelation -ne $CorrelationId) {
            return [pscustomobject]@{ Exists = $true; State = 'Foreign'; Completed = $false; Result = $null; Path = $path; Detail = ('Foreign result ignored: CorrelationId={0}' -f $documentCorrelation) }
        }
        $state = [string](Get-SemtPropertyValue -InputObject $document -Name 'State' -Default '')
        return [pscustomobject]@{ Exists = $true; State = $state; Completed = ($state -eq 'Completed'); Result = $document; Path = $path; Detail = ('Attempt result State={0}' -f $state) }
    }
    catch {
        return [pscustomobject]@{ Exists = $true; State = 'Unreadable'; Completed = $false; Result = $null; Path = $path; Detail = ('Attempt result unreadable: {0}' -f $_.Exception.Message) }
    }
}

function Wait-SemtRemoteAttemptCompletion {
    <#
    Polls the target for the completed result of this attempt after a PsExec communication
    loss or timeout, refreshing the worker reservation heartbeat while waiting.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RemoteDataShare,
        [Parameter(Mandatory = $true)][string]$CorrelationId,
        [Parameter(Mandatory = $true)][int]$WaitMinutes,
        [Parameter(Mandatory = $true)][int]$PollMinutes,
        [AllowNull()][string]$LeasePath,
        [string]$LogPath
    )

    $deadline = (Get-Date).AddMinutes($WaitMinutes)
    $last = $null
    while ($true) {
        $last = Read-SemtRemoteAttemptResult -RemoteDataShare $RemoteDataShare -CorrelationId $CorrelationId
        if ($last.Completed) { return $last }
        if ((Get-Date) -ge $deadline) { return $last }
        if ($LogPath) { Write-SemtLogLine -Path $LogPath -Message ("Completed attempt result not found yet ({0}). Next check in {1} minute(s)." -f $last.Detail,$PollMinutes) }
        $sleepUntil = (Get-Date).AddMinutes([math]::Max(1, $PollMinutes))
        if ($sleepUntil -gt $deadline) { $sleepUntil = $deadline }
        while ((Get-Date) -lt $sleepUntil) {
            Start-Sleep -Seconds 30
            [void](Update-SemtWorkerLease -LeasePath $LeasePath)
        }
    }
}

function Copy-SemtEvidenceFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$TargetRoot,
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [long]$MaxBytes = 5MB,
        [string]$Mode = 'Standard'
    )

    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) { return $false }
    try {
        $item = Get-Item -LiteralPath $SourcePath -ErrorAction Stop
        if ($Mode -ne 'Full' -and $item.Length -gt $MaxBytes) {
            Add-Content -LiteralPath (Join-Path -Path $TargetRoot -ChildPath 'CentralLogCollection.skipped.txt') -Value ('{0} Skipped large file: {1}; SizeMB={2:N2}; Source={3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$RelativePath,($item.Length / 1MB),$SourcePath) -Encoding UTF8
            return $false
        }
        $destination = Join-Path -Path $TargetRoot -ChildPath $RelativePath
        New-SemtDirectory -Path (Split-Path -Path $destination -Parent)
        Copy-Item -LiteralPath $SourcePath -Destination $destination -Force -ErrorAction Stop
        return $true
    }
    catch {
        Add-Content -LiteralPath (Join-Path -Path $TargetRoot -ChildPath 'CentralLogCollection.skipped.txt') -Value ('{0} Copy failed: {1}; Error={2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$SourcePath,$_.Exception.Message) -Encoding UTF8
        return $false
    }
}

function ConvertTo-SemtRemoteSharePathFromLocal {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RemoteDataShare,
        [Parameter(Mandatory = $true)][string]$RemoteDataRoot,
        [AllowNull()][string]$EndpointLocalPath
    )

    if ([string]::IsNullOrWhiteSpace($EndpointLocalPath)) { return '' }
    try {
        $base = [System.IO.Path]::GetFullPath($RemoteDataRoot).TrimEnd('\')
        $local = [System.IO.Path]::GetFullPath($EndpointLocalPath)
        if (-not $local.StartsWith($base + '\', [System.StringComparison]::OrdinalIgnoreCase)) { return '' }
        return (Join-Path -Path $RemoteDataShare -ChildPath $local.Substring($base.Length).TrimStart('\'))
    }
    catch { return '' }
}

function New-SemtCentralLogTarget {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$CentralLogRoot,
        [Parameter(Mandatory = $true)][string]$ComputerPathKey,
        [Parameter(Mandatory = $true)][string]$Bucket,
        [Parameter(Mandatory = $true)][string[]]$BucketList,
        [Parameter(Mandatory = $true)][int]$CycleNumber,
        [bool]$KeepHistory = $false
    )

    if ($BucketList -notcontains $Bucket) { $Bucket = 'Errors' }
    if ($KeepHistory) {
        return (Join-Path -Path (Join-Path -Path (Join-Path -Path $CentralLogRoot -ChildPath $Bucket) -ChildPath $ComputerPathKey) -ChildPath ('Cycle{0}_{1}' -f $CycleNumber,(Get-Date -Format 'yyyyMMdd-HHmmss')))
    }
    foreach ($other in $BucketList) {
        $latest = Join-Path -Path (Join-Path -Path (Join-Path -Path $CentralLogRoot -ChildPath $other) -ChildPath $ComputerPathKey) -ChildPath 'Latest'
        if (Test-Path -LiteralPath $latest) { Remove-Item -LiteralPath $latest -Recurse -Force -ErrorAction SilentlyContinue }
    }
    return (Join-Path -Path (Join-Path -Path (Join-Path -Path $CentralLogRoot -ChildPath $Bucket) -ChildPath $ComputerPathKey) -ChildPath 'Latest')
}

function Copy-SemtCentralEvidence {
    <#
    Standard mode copies only current-attempt evidence with short destination names: the
    attempt result, LastRun.json, the endpoint log and CSV named by the attempt result, and
    the campaign evidence rules (files newer than the attempt start). Full mode mirrors the
    endpoint Logs, Output, State and Transcripts folders.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RemoteDataShare,
        [Parameter(Mandatory = $true)][string]$RemoteDataRoot,
        [Parameter(Mandatory = $true)][string]$TargetPath,
        [Parameter(Mandatory = $true)][string]$CorrelationId,
        [AllowNull()]$AttemptResult,
        [Parameter(Mandatory = $true)][datetime]$AttemptStartedUtc,
        [ValidateSet('Standard','Full')][string]$Mode = 'Standard',
        [object[]]$EvidenceRules = @()
    )

    if (-not (Test-Path -LiteralPath $RemoteDataShare -PathType Container)) { return 0 }
    if (Test-Path -LiteralPath $TargetPath) { Remove-Item -LiteralPath $TargetPath -Recurse -Force -ErrorAction SilentlyContinue }
    New-SemtDirectory -Path $TargetPath
    $count = 0
    if ($Mode -eq 'Full') {
        foreach ($child in @('Logs','Output','State','Transcripts','LastRun.json')) {
            $source = Join-Path -Path $RemoteDataShare -ChildPath $child
            if (Test-Path -LiteralPath $source) { Copy-Item -LiteralPath $source -Destination $TargetPath -Recurse -Force -ErrorAction SilentlyContinue; $count++ }
        }
        return $count
    }
    Add-Content -LiteralPath (Join-Path -Path $TargetPath -ChildPath 'CentralLogCollection.txt') -Value ('{0} Standard collection for CorrelationId={1}; full logs remain on the target under {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$CorrelationId,$RemoteDataShare) -Encoding UTF8
    if (Copy-SemtEvidenceFile -SourcePath (Join-Path -Path $RemoteDataShare -ChildPath ("State\Results\{0}.json" -f $CorrelationId)) -TargetRoot $TargetPath -RelativePath 'AttemptResult.json' -MaxBytes 2MB) { $count++ }
    if (Copy-SemtEvidenceFile -SourcePath (Join-Path -Path $RemoteDataShare -ChildPath 'LastRun.json') -TargetRoot $TargetPath -RelativePath 'LastRun.json' -MaxBytes 1MB) { $count++ }
    foreach ($named in @(@{ Property = 'LogPath'; Name = 'Endpoint.log'; Max = 10MB },@{ Property = 'CsvPath'; Name = 'Result.csv'; Max = 2MB },@{ Property = 'TranscriptPath'; Name = 'Transcript.txt'; Max = 10MB })) {
        $localPath = [string](Get-SemtPropertyValue -InputObject $AttemptResult -Name $named.Property -Default '')
        $sharePath = ConvertTo-SemtRemoteSharePathFromLocal -RemoteDataShare $RemoteDataShare -RemoteDataRoot $RemoteDataRoot -EndpointLocalPath $localPath
        if (-not [string]::IsNullOrWhiteSpace($sharePath) -and (Copy-SemtEvidenceFile -SourcePath $sharePath -TargetRoot $TargetPath -RelativePath $named.Name -MaxBytes $named.Max)) { $count++ }
    }
    $minimumUtc = $AttemptStartedUtc.AddMinutes(-5)
    foreach ($rule in @($EvidenceRules)) {
        $folder = Join-Path -Path $RemoteDataShare -ChildPath ([string]$rule.Folder)
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) { continue }
        $files = @(Get-ChildItem -LiteralPath $folder -File -Recurse:([bool](Get-SemtPropertyValue -InputObject $rule -Name 'Recurse' -Default $false)) -Filter ([string](Get-SemtPropertyValue -InputObject $rule -Name 'Filter' -Default '*')) -ErrorAction SilentlyContinue)
        if ([bool](Get-SemtPropertyValue -InputObject $rule -Name 'CurrentAttemptOnly' -Default $true)) { $files = @($files | Where-Object { $_.LastWriteTimeUtc -ge $minimumUtc }) }
        $names = @(Get-SemtPropertyValue -InputObject $rule -Name 'Names' -Default @())
        if ($names.Count -gt 0) { $files = @($files | Where-Object { $names -contains $_.Name }) }
        $files = @($files | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First ([int](Get-SemtPropertyValue -InputObject $rule -Name 'MaxFiles' -Default 20)))
        foreach ($file in $files) {
            $relative = Join-Path -Path ([string]$rule.TargetFolder) -ChildPath $file.Name
            if (Copy-SemtEvidenceFile -SourcePath $file.FullName -TargetRoot $TargetPath -RelativePath $relative -MaxBytes ([long](Get-SemtPropertyValue -InputObject $rule -Name 'MaxBytes' -Default 5MB))) { $count++ }
        }
    }
    return $count
}

function Publish-SemtLauncherEvidence {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$TargetPath,
        [Parameter(Mandatory = $true)][string]$WorkerLogPath
    )

    New-SemtDirectory -Path $TargetPath
    foreach ($pair in @(@($WorkerLogPath,'Worker.log'),@(('{0}.stdout.txt' -f $WorkerLogPath),'PsExec.stdout.txt'),@(('{0}.stderr.txt' -f $WorkerLogPath),'PsExec.stderr.txt'))) {
        if (Test-Path -LiteralPath $pair[0] -PathType Leaf) { Copy-Item -LiteralPath $pair[0] -Destination (Join-Path -Path $TargetPath -ChildPath $pair[1]) -Force -ErrorAction SilentlyContinue }
    }
    return $TargetPath
}

#region SEMT trusted program folder v1
# Scripts that run as SYSTEM (scheduled tasks, Intune runner, retry after reboot) are installed
# under %ProgramFiles%, which only administrators can write. ProgramData keeps only data (logs,
# state, results, setup media). The same region is copied unchanged into both endpoint scripts,
# the Intune installer and the Intune runner (a test checks that all copies are identical).
#
# A folder is trusted when it and every item below it are owned by SYSTEM, Administrators or
# TrustedInstaller and no other principal can write, append, delete, change permissions or take
# ownership. Permissions that cannot be read are never trusted. An existing folder that is not
# trusted is never repaired automatically: the operation stops (SEMT_PROGRAM_FOLDER_UNTRUSTED).

$script:SemtProgramTrustedSids = @('S-1-5-18','S-1-5-32-544','S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464')
$script:SemtProgramFolderSddl = 'O:BAD:PAI(A;OICI;FA;;;SY)(A;OICI;FA;;;BA)(A;OICI;0x1200a9;;;BU)'

function Get-SemtProgramRoot {
    param([Parameter(Mandatory = $true)][string]$Campaign, [string]$ProgramFilesPath = $env:ProgramFiles)
    if ([string]::IsNullOrWhiteSpace($ProgramFilesPath)) { throw 'SEMT_PROGRAM_FOLDER_UNTRUSTED: the Program Files folder is unknown.' }
    return (Join-Path $ProgramFilesPath ('WCH\SmartM365\SmartEndpointMigrationToolkit\{0}' -f $Campaign))
}

function Test-SemtTrustedProgramPath {
    # Read-only. State: Trusted, Untrusted (proven) or Unreadable (not proven either way).
    param([Parameter(Mandatory = $true)][string]$Path, [switch]$Recurse)
    # Write data/append/extended attributes, delete child, delete, WRITE_DAC, WRITE_OWNER and the
    # generic write/all bits that inherit-only entries may carry.
    $writeMask = 0x00000002 -bor 0x00000004 -bor 0x00000010 -bor 0x00000040 -bor 0x00010000 -bor 0x00040000 -bor 0x00080000 -bor 0x40000000 -bor 0x10000000
    try {
        $items = @(Get-Item -LiteralPath $Path -Force -ErrorAction Stop)
        if ($Recurse -and $items[0].PSIsContainer) { $items += @(Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction Stop) }
    }
    catch { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Program folder could not be read: {0}' -f $_.Exception.Message) } }
    foreach ($item in $items) {
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('Reparse point in the program folder: {0}' -f $item.FullName) } }
        try {
            $acl = Get-Acl -LiteralPath $item.FullName -ErrorAction Stop
            $owner = $acl.GetOwner([System.Security.Principal.SecurityIdentifier])
            $rules = @($acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))
        }
        catch { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Permissions of {0} could not be read: {1}' -f $item.FullName,$_.Exception.Message) } }
        if ($null -eq $owner) { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Owner of {0} could not be read.' -f $item.FullName) } }
        if ($script:SemtProgramTrustedSids -notcontains $owner.Value) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('Owner {0} of {1} is not SYSTEM, Administrators or TrustedInstaller.' -f $owner.Value,$item.FullName) } }
        foreach ($rule in $rules) {
            if ($rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) { continue }
            if (([int64]$rule.FileSystemRights -band $writeMask) -eq 0) { continue }
            $sid = $rule.IdentityReference.Value
            if ($script:SemtProgramTrustedSids -contains $sid) { continue }
            # CREATOR OWNER only applies to items created later, by principals that can already write.
            if ($sid -eq 'S-1-3-0' -and ($rule.PropagationFlags -band [System.Security.AccessControl.PropagationFlags]::InheritOnly) -ne 0) { continue }
            return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('{0} can modify {1}.' -f $sid,$item.FullName) }
        }
    }
    return [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = '' }
}

function Set-SemtProgramFolderAcl {
    # Owner Administrators; protected DACL: SYSTEM and Administrators full control, Users read and
    # execute. Every other entry is removed. Requires an elevated or SYSTEM context.
    param([Parameter(Mandatory = $true)][string]$Path)
    $security = New-Object System.Security.AccessControl.DirectorySecurity
    $security.SetSecurityDescriptorSddlForm($script:SemtProgramFolderSddl)
    Set-Acl -LiteralPath $Path -AclObject $security -ErrorAction Stop
}

function Initialize-SemtTrustedProgramFolder {
    # Creates the folder chain below Program Files with a protected ACL, or verifies an existing
    # chain. Throws SEMT_PROGRAM_FOLDER_UNTRUSTED when any component is not trusted afterwards.
    param([Parameter(Mandatory = $true)][string]$Path, [string]$BasePath = $env:ProgramFiles)
    $base = [System.IO.Path]::GetFullPath($BasePath).TrimEnd('\')
    $target = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    if (-not $target.StartsWith($base + '\', [System.StringComparison]::OrdinalIgnoreCase)) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} is not below {1}.' -f $target,$base) }
    $baseTrust = Test-SemtTrustedProgramPath -Path $base
    if (-not $baseTrust.Trusted) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $base,$baseTrust.State,$baseTrust.Detail) }
    $current = $base
    foreach ($segment in @($target.Substring($base.Length + 1).Split('\'))) {
        $current = Join-Path $current $segment
        if (-not (Test-Path -LiteralPath $current)) {
            New-Item -ItemType Directory -Path $current -ErrorAction Stop | Out-Null
            try { Set-SemtProgramFolderAcl -Path $current }
            catch {
                # A folder created here that cannot be protected is removed again.
                Remove-Item -LiteralPath $current -Force -ErrorAction SilentlyContinue
                throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: permissions of {0} could not be set: {1}' -f $current,$_.Exception.Message)
            }
        }
        elseif (-not (Test-Path -LiteralPath $current -PathType Container)) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} exists and is not a folder.' -f $current) }
        $trust = Test-SemtTrustedProgramPath -Path $current -Recurse:($current -eq $target)
        if (-not $trust.Trusted) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $current,$trust.State,$trust.Detail) }
    }
    return $target
}

function Install-SemtTrustedProgramFile {
    # Copies through a staging name inside the trusted folder, verifies the SHA256 against the
    # source, renames, then verifies the hash and the permissions of the installed file.
    param([Parameter(Mandatory = $true)][string]$Source, [Parameter(Mandatory = $true)][string]$Destination)
    $expected = (Get-FileHash -LiteralPath $Source -Algorithm SHA256 -ErrorAction Stop).Hash
    $folder = Split-Path -Path $Destination -Parent
    $staging = Join-Path $folder ('.{0}.{1}.tmp' -f (Split-Path -Path $Destination -Leaf),([guid]::NewGuid().ToString('N')))
    try {
        Copy-Item -LiteralPath $Source -Destination $staging -ErrorAction Stop
        if ((Get-FileHash -LiteralPath $staging -Algorithm SHA256).Hash -ne $expected) { throw ('SEMT_PROGRAM_FILE_HASH_MISMATCH: staged copy of {0} differs from the source.' -f $Source) }
        Move-Item -LiteralPath $staging -Destination $Destination -Force -ErrorAction Stop
    }
    finally { if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Force -ErrorAction SilentlyContinue } }
    $actual = (Get-FileHash -LiteralPath $Destination -Algorithm SHA256).Hash
    if ($actual -ne $expected) { throw ('SEMT_PROGRAM_FILE_HASH_MISMATCH: {0} Expected={1}; Actual={2}' -f $Destination,$expected,$actual) }
    $trust = Test-SemtTrustedProgramPath -Path $Destination
    if (-not $trust.Trusted) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $Destination,$trust.State,$trust.Detail) }
    return [pscustomobject]@{ Path = $Destination; Sha256 = $actual }
}

# ProgramData data folders (logs, state, results, setup media whose setup.exe runs as SYSTEM).
# Scope: only the SEMT components (SmartEndpointMigrationToolkit and the folders below it) are
# ever changed. The shared parents above it (WCH, WCH\SmartM365), used by other products and by
# the legacy caches reused in place, are verified read-only and never changed; when they are not
# protected, the operation stops. A read-only pass checks the whole scope before the first
# change: shared parents, SEMT components, the provenance of existing content and, when a
# permission replacement is planned, every reparse point and hard link below it.
# A permission replacement never makes existing content trusted: an item that another principal
# could modify, or any content of a SEMT component that was not protected, has no verified
# provenance; it is refused and left unchanged (SEMT_DATA_FOLDER_UNTRUSTED), never adopted.
$script:SemtDataProductFolderName = 'SmartEndpointMigrationToolkit'

function Get-SemtHardLinkCount {
    # Number of names of a file; -1 when it cannot be read.
    param([Parameter(Mandatory = $true)][string]$Path)
    $fsutil = Join-Path $env:SystemRoot 'System32\fsutil.exe'
    $output = @(& $fsutil hardlink list $Path 2>$null)
    if ($LASTEXITCODE -ne 0) { return -1 }
    return @($output | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count
}

function Assert-SemtNoLinkedItem {
    # Read-only. Before a DACL replacement propagates inheritance below Path: no reparse point
    # (never followed) and only files with a single name.
    param([Parameter(Mandatory = $true)][string]$Path)
    $pending = New-Object System.Collections.Generic.Queue[string]
    $pending.Enqueue([System.IO.Path]::GetFullPath($Path))
    while ($pending.Count -gt 0) {
        $folder = $pending.Dequeue()
        try { $children = @(Get-ChildItem -LiteralPath $folder -Force -ErrorAction Stop) }
        catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be listed: {1}; nothing was changed.' -f $folder,$_.Exception.Message) }
        foreach ($child in $children) {
            if (($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: reparse point {0}; nothing was changed.' -f $child.FullName) }
            if ($child.PSIsContainer) { $pending.Enqueue($child.FullName); continue }
            $links = Get-SemtHardLinkCount -Path $child.FullName
            if ($links -ne 1) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} has {1} hard link name(s); nothing was changed.' -f $child.FullName,$links) }
        }
    }
}

function Test-SemtCanonicalFolderAcl {
    # True when the folder already has exactly the protected ACL that this region sets.
    param([Parameter(Mandatory = $true)][string]$Path)
    try {
        $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
        if ($acl.GetOwner([System.Security.Principal.SecurityIdentifier]).Value -ne 'S-1-5-32-544') { return $false }
        if (-not $acl.AreAccessRulesProtected) { return $false }
        $rules = @($acl.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier]) | ForEach-Object { '{0}:{1}:{2}:{3}' -f $_.IdentityReference.Value,[int64]$_.FileSystemRights,[int]$_.InheritanceFlags,$_.AccessControlType })
        $expected = @('S-1-5-18:2032127:3:Allow','S-1-5-32-544:2032127:3:Allow','S-1-5-32-545:1179817:3:Allow')
        return (@($rules | Sort-Object) -join '|') -eq (@($expected | Sort-Object) -join '|')
    }
    catch { return $false }
}

function Test-SemtTrustedSharedParent {
    # Read-only; a shared parent is never changed. Trusted when its owner is SYSTEM,
    # Administrators or TrustedInstaller and no other principal can delete it, delete its
    # children, change its permissions or take ownership. Creating new items (other products) is
    # allowed: it cannot alter the existing SEMT folder below.
    param([Parameter(Mandatory = $true)][string]$Path)
    # Delete child, delete, WRITE_DAC, WRITE_OWNER and generic all.
    $mask = 0x00000040 -bor 0x00010000 -bor 0x00040000 -bor 0x00080000 -bor 0x10000000
    try {
        $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
        $acl = Get-Acl -LiteralPath $item.FullName -ErrorAction Stop
        $owner = $acl.GetOwner([System.Security.Principal.SecurityIdentifier])
        $rules = @($acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))
    }
    catch { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Shared folder {0} could not be read: {1}' -f $Path,$_.Exception.Message) } }
    if (-not $item.PSIsContainer) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('{0} is not a folder.' -f $item.FullName) } }
    if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('{0} is a reparse point.' -f $item.FullName) } }
    if ($null -eq $owner) { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Owner of {0} could not be read.' -f $item.FullName) } }
    if ($script:SemtProgramTrustedSids -notcontains $owner.Value) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('Owner {0} of {1} is not SYSTEM, Administrators or TrustedInstaller.' -f $owner.Value,$item.FullName) } }
    foreach ($rule in $rules) {
        if ($rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) { continue }
        # Inherit-only entries apply to new children only, not to this folder.
        if (($rule.PropagationFlags -band [System.Security.AccessControl.PropagationFlags]::InheritOnly) -ne 0) { continue }
        if (([int64]$rule.FileSystemRights -band $mask) -eq 0) { continue }
        if ($script:SemtProgramTrustedSids -contains $rule.IdentityReference.Value) { continue }
        return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('{0} can delete, rename or change the permissions of {1}.' -f $rule.IdentityReference.Value,$item.FullName) }
    }
    return [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = '' }
}

function Assert-SemtDataFolderAnchor {
    # The data folder must be below the base inside a SmartEndpointMigrationToolkit folder, and
    # the base and its parents must not be reparse points.
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$BasePath)
    $base = [System.IO.Path]::GetFullPath($BasePath).TrimEnd('\')
    $target = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    if (-not $target.StartsWith($base + '\', [System.StringComparison]::OrdinalIgnoreCase)) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} is not below {1}.' -f $target,$base) }
    $segments = @($target.Substring($base.Length + 1).Split('\'))
    $productIndex = -1
    for ($i = 0; $i -lt $segments.Count; $i++) { if ($segments[$i] -ieq $script:SemtDataProductFolderName) { $productIndex = $i; break } }
    if ($productIndex -lt 0 -or $productIndex -ge ($segments.Count - 1)) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} is not a campaign folder below {1}.' -f $target,$script:SemtDataProductFolderName) }
    $current = $base
    while (-not [string]::IsNullOrWhiteSpace($current)) {
        try { $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop }
        catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be read: {1}' -f $current,$_.Exception.Message) }
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} is a reparse point.' -f $current) }
        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) { break }
        $current = $parent
    }
    return [pscustomobject]@{ Base = $base; Target = $target; Segments = $segments; ProductIndex = $productIndex }
}

function Get-SemtDataFolderPlan {
    # Read-only pass over the whole scope. Returns one step per component (Keep, Create or
    # Replace); throws SEMT_DATA_FOLDER_UNTRUSTED, with nothing changed, when the operation
    # cannot be done within the SEMT scope or when existing content has no verified provenance.
    param([Parameter(Mandatory = $true)]$Anchor)
    $plan = New-Object System.Collections.Generic.List[object]
    $current = $Anchor.Base
    $missing = $false
    for ($index = 0; $index -lt $Anchor.Segments.Count; $index++) {
        $current = Join-Path $current $Anchor.Segments[$index]
        $shared = $index -lt $Anchor.ProductIndex
        if ($missing -or -not (Test-Path -LiteralPath $current)) {
            $missing = $true
            $plan.Add([pscustomobject]@{ Path = $current; Shared = $shared; Action = 'Create'; WasUntrusted = $false })
            continue
        }
        try { $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop }
        catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be read: {1}; nothing was changed.' -f $current,$_.Exception.Message) }
        if (-not $item.PSIsContainer) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} exists and is not a folder; nothing was changed.' -f $current) }
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} is a reparse point; nothing was changed.' -f $current) }
        if ($shared) {
            $trust = Test-SemtTrustedSharedParent -Path $current
            if (-not $trust.Trusted) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: shared folder {0} ({1}) {2} It is used by other products and is never changed; nothing was changed.' -f $current,$trust.State,$trust.Detail) }
            $plan.Add([pscustomobject]@{ Path = $current; Shared = $true; Action = 'Keep'; WasUntrusted = $false })
            continue
        }
        $trust = Test-SemtTrustedProgramPath -Path $current
        if ($trust.State -eq 'Unreadable') { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} (Unreadable) {1}; nothing was changed.' -f $current,$trust.Detail) }
        if (-not $trust.Trusted) {
            # Content of a SEMT folder that was not protected has no verified provenance.
            try { $children = @(Get-ChildItem -LiteralPath $current -Force -ErrorAction Stop) }
            catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be listed: {1}; nothing was changed.' -f $current,$_.Exception.Message) }
            if ($children.Count -gt 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} was not protected ({1}) and is not empty; its content has no verified provenance and is not used; nothing was changed.' -f $current,$trust.Detail) }
            $plan.Add([pscustomobject]@{ Path = $current; Shared = $false; Action = 'Replace'; WasUntrusted = $true })
            continue
        }
        $action = if (Test-SemtCanonicalFolderAcl -Path $current) { 'Keep' } else { 'Replace' }
        $plan.Add([pscustomobject]@{ Path = $current; Shared = $false; Action = $action; WasUntrusted = $false })
    }
    # Existing content must already be trusted, item by item: below the data root and, when a
    # permission replacement is planned, everywhere it propagates to (other SEMT campaigns
    # included), so that the replacement never makes untrusted content look trusted.
    $firstReplace = @($plan | Where-Object { $_.Action -eq 'Replace' } | Select-Object -First 1)
    $existing = @($plan | Where-Object { $_.Action -ne 'Create' -and -not $_.Shared })
    $checkFrom = if ($firstReplace.Count -gt 0) { $firstReplace[0].Path } elseif (-not $missing) { $Anchor.Target } else { '' }
    if (-not [string]::IsNullOrWhiteSpace($checkFrom)) {
        $chain = @($existing | ForEach-Object { $_.Path })
        $start = [array]::IndexOf([string[]]$chain, [string]$checkFrom)
        for ($i = $start; $i -lt $chain.Count; $i++) {
            try { $children = @(Get-ChildItem -LiteralPath $chain[$i] -Force -ErrorAction Stop) }
            catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be listed: {1}; nothing was changed.' -f $chain[$i],$_.Exception.Message) }
            foreach ($child in $children) {
                if ($i -lt ($chain.Count - 1) -and $child.FullName -ieq $chain[$i + 1]) { continue }
                $trust = Test-SemtTrustedProgramPath -Path $child.FullName -Recurse
                if (-not $trust.Trusted) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: existing content has no verified provenance and is not used ({0}) {1}; nothing was changed.' -f $trust.State,$trust.Detail) }
            }
        }
    }
    # A DACL replacement propagates inheritance below the folder: the whole subtree of the first
    # replaced component is checked for reparse points and hard links before any change.
    if ($firstReplace.Count -gt 0) { Assert-SemtNoLinkedItem -Path $firstReplace[0].Path }
    return $plan
}

function Initialize-SemtTrustedDataFolder {
    # Protects the SEMT data folder chain below ProgramData and verifies it. Throws
    # SEMT_DATA_FOLDER_UNTRUSTED when it cannot be made trusted within the SEMT scope.
    param([Parameter(Mandatory = $true)][string]$Path, [string]$BasePath = $env:ProgramData)
    $anchor = Assert-SemtDataFolderAnchor -Path $Path -BasePath $BasePath
    $plan = Get-SemtDataFolderPlan -Anchor $anchor
    foreach ($step in $plan) {
        if ($step.Action -eq 'Create') {
            New-Item -ItemType Directory -Path $step.Path -ErrorAction Stop | Out-Null
            if ($step.Shared) {
                # Created with inherited permissions only: a shared folder is never changed.
                $trust = Test-SemtTrustedSharedParent -Path $step.Path
                if (-not $trust.Trusted) {
                    Remove-Item -LiteralPath $step.Path -Force -ErrorAction SilentlyContinue
                    throw ('SEMT_DATA_FOLDER_UNTRUSTED: shared folder {0} created here is not protected ({1}) {2}' -f $step.Path,$trust.State,$trust.Detail)
                }
                continue
            }
            try { Set-SemtProgramFolderAcl -Path $step.Path }
            catch {
                Remove-Item -LiteralPath $step.Path -Force -ErrorAction SilentlyContinue
                throw ('SEMT_DATA_FOLDER_UNTRUSTED: permissions of {0} could not be set: {1}' -f $step.Path,$_.Exception.Message)
            }
        }
        elseif ($step.Action -eq 'Replace') {
            try { Set-SemtProgramFolderAcl -Path $step.Path }
            catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: owner and permissions of {0} could not be replaced: {1}' -f $step.Path,$_.Exception.Message) }
        }
        else { continue }
        # A folder created here, or one that was not protected (empty in the read-only pass), must
        # still be empty once protected: nothing written meanwhile is adopted.
        if ($step.Action -eq 'Create' -or $step.WasUntrusted) {
            if (@(Get-ChildItem -LiteralPath $step.Path -Force -ErrorAction Stop).Count -gt 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: content appeared in {0} before it was protected; it is not used.' -f $step.Path) }
        }
        $trust = Test-SemtTrustedProgramPath -Path $step.Path
        if (-not $trust.Trusted) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $step.Path,$trust.State,$trust.Detail) }
    }
    $final = Test-SemtTrustedDataFolder -Path $anchor.Target -BasePath $anchor.Base
    if (-not $final.Trusted) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $anchor.Target,$final.State,$final.Detail) }
    return $anchor.Target
}

function Test-SemtTrustedDataFolder {
    # Read-only verification used before anything is run: shared parents protected, every SEMT
    # component trusted, and the data root trusted with everything below it.
    param([Parameter(Mandatory = $true)][string]$Path, [string]$BasePath = $env:ProgramData)
    try { $anchor = Assert-SemtDataFolderAnchor -Path $Path -BasePath $BasePath }
    catch { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = $_.Exception.Message } }
    $current = $anchor.Base
    for ($index = 0; $index -lt $anchor.Segments.Count; $index++) {
        $current = Join-Path $current $anchor.Segments[$index]
        if ($index -lt $anchor.ProductIndex) { $trust = Test-SemtTrustedSharedParent -Path $current }
        else { $trust = Test-SemtTrustedProgramPath -Path $current -Recurse:($current -eq $anchor.Target) }
        if (-not $trust.Trusted) { return $trust }
    }
    return [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = '' }
}
#endregion SEMT trusted program folder v1
