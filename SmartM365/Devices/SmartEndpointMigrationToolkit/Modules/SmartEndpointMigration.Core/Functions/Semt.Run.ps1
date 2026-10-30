<#
.SYNOPSIS
    Waves, Runs, worker contract, component integrity, run guard history and night pause.

.DESCRIPTION
    Layout:
      <Package>\Waves\<Campaign>\<Wave>\Computers.txt + Wave.local.json + wrappers
      <Package>\Runs\<Campaign>\<Wave>\<RunId>\{Logs,PsExecLogs,Reports,CentralLogs,State,Inventory}
    A Single PC Run uses an ephemeral Wave created under Runs\<Campaign>\_SinglePC\<RunId>\Wave
    and goes through the same engine.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#region Waves and Runs

function Test-SemtWaveName {
    [CmdletBinding()]
    param([AllowNull()][string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    if ($Name.Length -gt 80) { return $false }
    if ($Name -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') { return $false }
    if ($Name.StartsWith('_')) { return $false }
    if ($Name -match '(?i)^(CON|PRN|AUX|NUL|COM\d|LPT\d)$') { return $false }
    return $true
}

function Get-SemtWaveRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$PackageRoot,
        [Parameter(Mandatory = $true)][string]$Campaign
    )
    [void](Assert-SemtCampaign -Campaign $Campaign)
    return (Join-Path -Path (Join-Path -Path $PackageRoot -ChildPath 'Waves') -ChildPath $Campaign)
}

function Get-SemtWaves {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$PackageRoot,
        [Parameter(Mandatory = $true)][string]$Campaign
    )

    $root = Get-SemtWaveRoot -PackageRoot $PackageRoot -Campaign $Campaign
    if (-not (Test-Path -LiteralPath $root -PathType Container)) { return @() }
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($folder in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Sort-Object Name)) {
        $computersPath = Join-Path -Path $folder.FullName -ChildPath 'Computers.txt'
        $count = 0
        if (Test-Path -LiteralPath $computersPath -PathType Leaf) { try { $count = @(Get-SemtComputerList -Path $computersPath).Count } catch { $count = 0 } }
        $wrappers = @('Start-Wave-Loop.cmd','Start-Wave-Once.cmd','Start-Wave-Loop-IgnoreRunGuard.cmd','Start-Wave-Once-IgnoreRunGuard.cmd')
        $missing = @($wrappers | Where-Object { -not (Test-Path -LiteralPath (Join-Path -Path $folder.FullName -ChildPath $_) -PathType Leaf) })
        $rows.Add([pscustomobject]@{
            Campaign = $Campaign
            Name = $folder.Name
            Path = $folder.FullName
            ComputersPath = $computersPath
            ComputerCount = $count
            WrappersReady = ($missing.Count -eq 0)
            MissingWrappers = ($missing -join ', ')
            Launchable = ($count -gt 0 -and $missing.Count -eq 0 -and (Test-SemtWaveName -Name $folder.Name))
        })
    }
    return @($rows.ToArray())
}

function Get-SemtWaveWrapperDefinitions {
    [CmdletBinding()]
    param()

    return @(
        [pscustomobject]@{ FileName = 'Start-Wave-Loop.cmd'; RunMode = 'Loop'; IgnoreRunGuard = $false; Description = 'Repeats cycles until the Wave is complete or stopped.' },
        [pscustomobject]@{ FileName = 'Start-Wave-Once.cmd'; RunMode = 'Once'; IgnoreRunGuard = $false; Description = 'Runs a single cycle.' },
        [pscustomobject]@{ FileName = 'Start-Wave-Loop-IgnoreRunGuard.cmd'; RunMode = 'Loop'; IgnoreRunGuard = $true; Description = 'Loop, bypassing the endpoint and technician run guards on the first cycle.' },
        [pscustomobject]@{ FileName = 'Start-Wave-Once-IgnoreRunGuard.cmd'; RunMode = 'Once'; IgnoreRunGuard = $true; Description = 'Once, bypassing the endpoint and technician run guards.' }
    )
}

function Write-SemtWaveWrappers {
    <#
    Writes the four thin Wave launchers. They only call Scripts\Run-EndpointMigrationWave.cmd
    with the Wave folder, the run mode and the run guard choice: every option is resolved by
    the engine from the catalog, the configuration files and the environment.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$WavePath)

    $written = New-Object System.Collections.Generic.List[string]
    $expected = @{}
    foreach ($wrapper in @(Get-SemtWaveWrapperDefinitions)) {
        $expected[$wrapper.FileName] = $true
        $extra = if ($wrapper.IgnoreRunGuard) { 'IgnoreRunGuard' } else { 'Default' }
        $lines = @(
            '@echo off',
            ('rem Smart Endpoint Migration Toolkit - {0}' -f $wrapper.Description),
            'rem Generated launcher: options are resolved by the engine, never by this file.',
            'setlocal DisableDelayedExpansion',
            ('call "%~dp0..\..\..\Scripts\Run-EndpointMigrationWave.cmd" "%~dp0." {0} {1}' -f $wrapper.RunMode,$extra),
            'exit /b %ERRORLEVEL%'
        )
        $path = Join-Path -Path $WavePath -ChildPath $wrapper.FileName
        [System.IO.File]::WriteAllText($path, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.ASCIIEncoding))
        $written.Add($path)
    }
    foreach ($obsolete in @(Get-ChildItem -LiteralPath $WavePath -Filter 'Start-Wave-*.cmd' -File -ErrorAction SilentlyContinue | Where-Object { -not $expected.ContainsKey($_.Name) })) {
        Remove-Item -LiteralPath $obsolete.FullName -Force -ErrorAction SilentlyContinue
    }
    return @($written.ToArray())
}

function New-SemtWave {
    <#
    Creates Waves\<Campaign>\<Name> with Computers.txt, Wave.local.json (from the campaign
    template) and the Loop/Once wrappers. It never launches the Wave.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$PackageRoot,
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowEmptyCollection()][string[]]$Computers = @(),
        [ValidateSet('Manual','Automatic')][string]$Source = 'Manual',
        [string]$Description = '',
        [string]$EvidencePath = ''
    )

    if (-not (Test-SemtWaveName -Name $Name)) { throw ("SEMT_WAVE_NAME_INVALID: '{0}' must start with a letter or digit and use only letters, digits, dot, dash or underscore (80 characters maximum)." -f $Name) }
    $root = Get-SemtWaveRoot -PackageRoot $PackageRoot -Campaign $Campaign
    $path = Join-Path -Path $root -ChildPath $Name
    if (Test-Path -LiteralPath $path) { throw ("SEMT_WAVE_EXISTS: Wave '{0}' already exists for {1}: {2}" -f $Name,$Campaign,$path) }
    New-SemtDirectory -Path $path
    $names = New-Object System.Collections.Generic.List[string]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($computer in @($Computers)) {
        $clean = ConvertTo-SemtComputerName -Line $computer
        if (-not [string]::IsNullOrWhiteSpace($clean) -and $seen.Add($clean)) { $names.Add($clean) }
    }
    Write-SemtTextLinesAtomic -Path (Join-Path -Path $path -ChildPath 'Computers.txt') -Lines @($names.ToArray())
    $config = Read-SemtWaveConfig -PackageRoot $PackageRoot -Campaign $Campaign -WavePath $path -CreateWhenMissing
    $data = $config.Data
    $data['Campaign'] = $Campaign
    $data['Source'] = $Source
    $data['Description'] = $Description
    $data['CreatedUtc'] = Get-SemtUtcNowText
    if (-not [string]::IsNullOrWhiteSpace($EvidencePath)) { $data['EvidencePath'] = $EvidencePath }
    Write-SemtJsonFileAtomic -Path $config.Path -InputObject $data
    [void](Write-SemtWaveWrappers -WavePath $path)
    return (@(Get-SemtWaves -PackageRoot $PackageRoot -Campaign $Campaign | Where-Object { $_.Name -eq $Name }) | Select-Object -First 1)
}

function New-SemtRunContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$PackageRoot,
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][string]$WavePath,
        [string]$RunId,
        [switch]$SinglePc,
        [string]$RunRoot
    )

    [void](Assert-SemtCampaign -Campaign $Campaign)
    if ([string]::IsNullOrWhiteSpace($RunId)) { $RunId = New-SemtRunId }
    $wavePathFull = [System.IO.Path]::GetFullPath($WavePath).TrimEnd('\')
    $waveName = if ($SinglePc) { '_SinglePC' } else { Split-Path -Path $wavePathFull -Leaf }
    if ([string]::IsNullOrWhiteSpace($RunRoot)) {
        $RunRoot = Join-Path -Path (Join-Path -Path (Join-Path -Path (Join-Path -Path $PackageRoot -ChildPath 'Runs') -ChildPath $Campaign) -ChildPath $waveName) -ChildPath $RunId
    }
    $context = [pscustomobject]@{
        PackageRoot = $PackageRoot
        Campaign = $Campaign
        WavePath = $wavePathFull
        WaveName = $waveName
        ComputerListPath = Join-Path -Path $wavePathFull -ChildPath 'Computers.txt'
        RunId = $RunId
        SinglePc = [bool]$SinglePc
        RunRoot = $RunRoot
        LogRoot = Join-Path -Path $RunRoot -ChildPath 'Logs'
        PsExecLogRoot = Join-Path -Path $RunRoot -ChildPath 'PsExecLogs'
        ReportRoot = Join-Path -Path $RunRoot -ChildPath 'Reports'
        CentralLogRoot = Join-Path -Path $RunRoot -ChildPath 'CentralLogs'
        StateRoot = Join-Path -Path $RunRoot -ChildPath 'State'
        ContractRoot = Join-Path -Path (Join-Path -Path $RunRoot -ChildPath 'State') -ChildPath 'Contracts'
        InventoryRoot = Join-Path -Path $RunRoot -ChildPath 'Inventory'
        LauncherLogPath = Join-Path -Path (Join-Path -Path $RunRoot -ChildPath 'Logs') -ChildPath ("SEMT-{0}-Launcher_{1}.log" -f $Campaign,$RunId)
        EffectiveOptionsPath = Join-Path -Path $RunRoot -ChildPath 'EffectiveOptions.json'
        RunManifestPath = Join-Path -Path $RunRoot -ChildPath 'RunManifest.json'
    }
    foreach ($folder in @($context.LogRoot,$context.PsExecLogRoot,$context.ReportRoot,$context.CentralLogRoot,$context.StateRoot,$context.ContractRoot,$context.InventoryRoot)) { New-SemtDirectory -Path $folder }
    return $context
}

function New-SemtEphemeralWave {
    <#
    Creates the one-computer Wave used by Single PC. It lives inside its Run folder so it can
    never be mistaken for an operational Wave.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$PackageRoot,
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][string]$RunId
    )

    $name = ConvertTo-SemtComputerName -Line $ComputerName
    if ([string]::IsNullOrWhiteSpace($name) -or $name -match '[\s;,|<>"]') { throw "SEMT_SINGLE_PC_INVALID: '$ComputerName' is not a valid computer name." }
    $runRoot = Join-Path -Path (Join-Path -Path (Join-Path -Path (Join-Path -Path $PackageRoot -ChildPath 'Runs') -ChildPath $Campaign) -ChildPath '_SinglePC') -ChildPath $RunId
    $wavePath = Join-Path -Path $runRoot -ChildPath 'Wave'
    New-SemtDirectory -Path $wavePath
    Write-SemtTextLinesAtomic -Path (Join-Path -Path $wavePath -ChildPath 'Computers.txt') -Lines @($name)
    return [pscustomobject]@{ WavePath = $wavePath; RunRoot = $runRoot; ComputerName = $name }
}

function Update-SemtRunHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$RunContext,
        [Parameter(Mandatory = $true)][string]$Status,
        [string]$ReportPath = '',
        [string]$Detail = '',
        [int]$Cycles = 0
    )

    $historyPath = Join-Path -Path (Split-Path -Path $RunContext.RunRoot -Parent) -ChildPath 'RunHistory.json'
    $lockPath = '{0}.lock' -f $historyPath
    Invoke-SemtWithFileLock -LockPath $lockPath -TimeoutSeconds 60 -ArgumentList @($historyPath,$RunContext,$Status,$ReportPath,$Detail,$Cycles) -ScriptBlock {
        param($LockedPath, $LockedContext, $LockedStatus, $LockedReport, $LockedDetail, $LockedCycles)
        $history = Read-SemtJsonFile -Path $LockedPath -Default ([pscustomobject]@{ SchemaVersion = 1; Runs = @() })
        $runs = New-Object System.Collections.Generic.List[object]
        $existing = $null
        foreach ($entry in @($history.Runs)) { if ([string]$entry.RunId -eq $LockedContext.RunId) { $existing = $entry } else { $runs.Add($entry) } }
        $startedUtc = if ($existing) { [string]$existing.StartedUtc } else { Get-SemtUtcNowText }
        $runs.Add([pscustomobject]@{
            RunId = $LockedContext.RunId
            Campaign = $LockedContext.Campaign
            WaveName = $LockedContext.WaveName
            Status = $LockedStatus
            Detail = $LockedDetail
            Cycles = $LockedCycles
            StartedUtc = $startedUtc
            UpdatedUtc = Get-SemtUtcNowText
            RunRoot = $LockedContext.RunRoot
            ReportPath = $LockedReport
            Technician = [Environment]::UserName
        })
        Write-SemtJsonFileAtomic -Path $LockedPath -InputObject ([pscustomobject]@{ SchemaVersion = 1; Runs = @($runs.ToArray() | Sort-Object StartedUtc) })
    } | Out-Null
}

function Get-SemtRuns {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$PackageRoot,
        [Parameter(Mandatory = $true)][string]$Campaign,
        [ValidateRange(1, 1000)][int]$Last = 50
    )

    $root = Join-Path -Path (Join-Path -Path $PackageRoot -ChildPath 'Runs') -ChildPath $Campaign
    if (-not (Test-Path -LiteralPath $root -PathType Container)) { return @() }
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($historyFile in @(Get-ChildItem -LiteralPath $root -Filter 'RunHistory.json' -File -Recurse -Depth 1 -ErrorAction SilentlyContinue)) {
        $history = $null
        try { $history = Read-SemtJsonFile -Path $historyFile.FullName } catch { $history = $null }
        if (-not $history) { continue }
        foreach ($entry in @($history.Runs)) { $rows.Add($entry) }
    }
    return @($rows.ToArray() | Sort-Object { [string]$_.StartedUtc } -Descending | Select-Object -First $Last)
}

function Invoke-SemtRunArchive {
    <#
    Compresses Runs of the same Wave older than the threshold. A Run folder is removed only
    after the archive was re-opened and its entry count and total size match the source.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$RunContext,
        [Parameter(Mandatory = $true)][int]$OlderThanDays
    )

    if ($OlderThanDays -le 0 -or $RunContext.SinglePc) { return @() }
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $waveRunsRoot = Split-Path -Path $RunContext.RunRoot -Parent
    $archiveRoot = Join-Path -Path $waveRunsRoot -ChildPath 'Archives'
    $threshold = (Get-Date).AddDays(-1 * $OlderThanDays)
    $results = New-Object System.Collections.Generic.List[object]
    foreach ($folder in @(Get-ChildItem -LiteralPath $waveRunsRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne 'Archives' -and $_.FullName -ne $RunContext.RunRoot -and $_.LastWriteTime -lt $threshold })) {
        if (@(Get-ChildItem -LiteralPath (Join-Path -Path $folder.FullName -ChildPath 'State') -Filter 'ActiveRun_*.json' -File -ErrorAction SilentlyContinue).Count -gt 0) { continue }
        New-SemtDirectory -Path $archiveRoot
        $zipPath = Join-Path -Path $archiveRoot -ChildPath ('{0}.zip' -f $folder.Name)
        try {
            if (Test-Path -LiteralPath $zipPath) { throw "Archive already exists: $zipPath" }
            $files = @(Get-ChildItem -LiteralPath $folder.FullName -Recurse -File -ErrorAction Stop)
            $bytes = [int64]0
            foreach ($file in $files) { $bytes += $file.Length }
            [System.IO.Compression.ZipFile]::CreateFromDirectory($folder.FullName, $zipPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
            $zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
            try {
                $entries = @($zip.Entries | Where-Object { -not $_.FullName.EndsWith('/') })
                $entryBytes = [int64]0
                foreach ($entry in $entries) { $entryBytes += $entry.Length }
            }
            finally { $zip.Dispose() }
            if ($entries.Count -ne $files.Count -or $entryBytes -ne $bytes) { throw ("Archive verification failed: Files={0}/{1}; Bytes={2}/{3}" -f $entries.Count,$files.Count,$entryBytes,$bytes) }
            Remove-Item -LiteralPath $folder.FullName -Recurse -Force -ErrorAction Stop
            $results.Add([pscustomobject]@{ Run = $folder.Name; Archive = $zipPath; Status = 'Archived'; Detail = ('Files={0}; Bytes={1}' -f $files.Count,$bytes) })
        }
        catch {
            $results.Add([pscustomobject]@{ Run = $folder.Name; Archive = $zipPath; Status = 'Kept'; Detail = $_.Exception.Message })
        }
    }
    return @($results.ToArray())
}

#endregion

#region Night pause

function Get-SemtNightPauseResumeTime {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][datetime]$Now,
        [Parameter(Mandatory = $true)][int]$StartHour,
        [Parameter(Mandatory = $true)][int]$EndHour
    )

    if ($StartHour -eq $EndHour) { return $null }
    $start = New-TimeSpan -Hours $StartHour
    $end = New-TimeSpan -Hours $EndHour
    $current = $Now.TimeOfDay
    if ($StartHour -lt $EndHour) {
        if ($current -ge $start -and $current -lt $end) { return $Now.Date.Add($end) }
        return $null
    }
    if ($current -ge $start) { return $Now.Date.AddDays(1).Add($end) }
    if ($current -lt $end) { return $Now.Date.Add($end) }
    return $null
}

#endregion

#region Component integrity

function Get-SemtComponentManifestPath {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$PackageRoot)
    return (Join-Path -Path $PackageRoot -ChildPath 'SmartEndpointMigrationToolkit.components.json.txt')
}

function Get-SemtComponentManifest {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$PackageRoot)

    $path = Get-SemtComponentManifestPath -PackageRoot $PackageRoot
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
    $manifest = Read-SemtJsonFile -Path $path -Required
    $byPath = New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($file in @($manifest.Files)) { $byPath[([string]$file.Path).Replace('/','\')] = $file }
    return [pscustomobject]@{ Path = $path; Manifest = $manifest; ByPath = $byPath }
}

function Test-SemtComponentFile {
    <#
    Verifies one package file against the distribution manifest (when present) and, for
    PowerShell files, its Authenticode signature. The expected hash never comes from the caller.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$PackageRoot,
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [AllowNull()]$Manifest,
        [ValidateSet('Auto','Enforce','Warn')][string]$Mode = 'Auto'
    )

    $relative = $RelativePath.Replace('/','\').TrimStart('\')
    $full = Join-Path -Path $PackageRoot -ChildPath $relative
    if (-not (Test-Path -LiteralPath $full -PathType Leaf)) {
        return [pscustomobject]@{ Path = $relative; Status = 'Missing'; Blocking = $true; Sha256 = ''; ExpectedSha256 = ''; Signature = ''; Detail = 'Component file is missing.' }
    }
    $sha = Get-SemtFileSha256 -Path $full
    $signatureStatus = ''
    $signer = ''
    if ($full -match '\.(ps1|psm1|psd1)$') {
        try {
            $signature = Get-AuthenticodeSignature -LiteralPath $full -ErrorAction Stop
            $signatureStatus = [string]$signature.Status
            if ($signature.SignerCertificate) { $signer = [string]$signature.SignerCertificate.Thumbprint }
        }
        catch { $signatureStatus = 'Unreadable' }
    }
    $expected = ''
    $status = 'Verified'
    $problems = New-Object System.Collections.Generic.List[string]
    if ($null -ne $Manifest) {
        $entry = $Manifest.ByPath[$relative]
        if ($null -eq $entry) { $problems.Add('Not listed in the distribution manifest') }
        else {
            $expected = [string]$entry.Sha256
            if ($expected -ne $sha) { $problems.Add(('SHA256 mismatch: Expected={0}; Actual={1}' -f $expected,$sha)) }
        }
    }
    if ($signatureStatus -and $signatureStatus -notin @('Valid','NotSigned')) { $problems.Add(('Signature={0}' -f $signatureStatus)) }
    $requireSignature = $true
    if ($null -ne $Manifest) { $requireSignature = -not ((Get-SemtPropertyValue -InputObject $Manifest.Manifest -Name 'RequireSignature' -Default $true) -eq $false) }
    if ($null -ne $Manifest -and $requireSignature -and $signatureStatus -eq 'NotSigned') { $problems.Add('Package component is not signed') }
    if ($problems.Count -gt 0) { $status = 'Mismatch' }
    $enforce = ($Mode -eq 'Enforce' -or ($Mode -eq 'Auto' -and ($null -ne $Manifest -or ($signatureStatus -and $signatureStatus -notin @('Valid','NotSigned')))))
    return [pscustomobject]@{
        Path = $relative
        Status = $status
        Blocking = ($status -ne 'Verified' -and $enforce)
        Sha256 = $sha
        ExpectedSha256 = $expected
        Signature = $signatureStatus
        SignerThumbprint = $signer
        Detail = ($problems -join '; ')
    }
}

function Test-SemtComponentIntegrity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$PackageRoot,
        [Parameter(Mandatory = $true)][string[]]$RelativePaths,
        [ValidateSet('Auto','Enforce','Warn')][string]$Mode = 'Auto'
    )

    $manifest = Get-SemtComponentManifest -PackageRoot $PackageRoot
    $results = @($RelativePaths | ForEach-Object { Test-SemtComponentFile -PackageRoot $PackageRoot -RelativePath $_ -Manifest $manifest -Mode $Mode })
    $signers = @($results | Where-Object { $_.Signature -eq 'Valid' } | ForEach-Object { $_.SignerThumbprint } | Select-Object -Unique)
    $mixedSigners = ($signers.Count -gt 1)
    return [pscustomobject]@{
        ManifestPath = if ($manifest) { $manifest.Path } else { '' }
        ManifestPresent = ($null -ne $manifest)
        Mode = $Mode
        Files = $results
        MixedSigners = $mixedSigners
        Blocking = (@($results | Where-Object { $_.Blocking }).Count -gt 0 -or ($mixedSigners -and $Mode -ne 'Warn'))
        Summary = ('Files={0}; Verified={1}; Mismatch={2}; Manifest={3}; Signers={4}' -f $results.Count,@($results | Where-Object { $_.Status -eq 'Verified' }).Count,@($results | Where-Object { $_.Status -ne 'Verified' }).Count,($null -ne $manifest),($signers -join ','))
    }
}

#endregion

#region Worker contract

function New-SemtWorkerContract {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$RunContext,
        [Parameter(Mandatory = $true)][string]$Computer,
        [Parameter(Mandatory = $true)][string]$ConnectionTarget,
        [Parameter(Mandatory = $true)][int]$CycleNumber,
        [Parameter(Mandatory = $true)][hashtable]$Endpoint,
        [Parameter(Mandatory = $true)][hashtable]$PsExec,
        [Parameter(Mandatory = $true)][hashtable]$Evidence,
        [Parameter(Mandatory = $true)][hashtable]$Flags,
        [hashtable]$Gate = @{},
        [hashtable]$CampaignData = @{},
        [string]$CorrelationId
    )

    $correlationId = if ([string]::IsNullOrWhiteSpace($CorrelationId)) { [guid]::NewGuid().ToString('N') } else { $CorrelationId }
    $contract = [ordered]@{
        ContractVersion = $script:SemtWorkerContractVersion
        ContractId = [guid]::NewGuid().ToString('N')
        Product = $script:SemtProductName
        EngineVersion = $script:SemtModuleVersion
        CreatedUtc = Get-SemtUtcNowText
        Campaign = $RunContext.Campaign
        WaveName = $RunContext.WaveName
        RunId = $RunContext.RunId
        CorrelationId = $correlationId
        Computer = $Computer
        ConnectionTarget = $ConnectionTarget
        CycleNumber = $CycleNumber
        PackageRoot = $RunContext.PackageRoot
        Endpoint = $Endpoint
        PsExec = $PsExec
        Evidence = $Evidence
        Flags = $Flags
        Gate = $Gate
        CampaignData = $CampaignData
    }
    $safe = Get-SemtSafeName -Value ($Computer.Split('.')[0]) -MaxLength 40
    $path = Join-Path -Path $RunContext.ContractRoot -ChildPath ('C{0}_{1}_{2}.json' -f $CycleNumber,$safe,$contract.ContractId.Substring(0, 8))
    Write-SemtJsonFileAtomic -Path $path -InputObject $contract -Depth 8
    return [pscustomobject]@{ Path = $path; CorrelationId = $correlationId; ContractId = $contract.ContractId }
}

function Read-SemtWorkerContract {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    $contract = Read-SemtJsonFile -Path $Path -Required
    $version = [int](Get-SemtPropertyValue -InputObject $contract -Name 'ContractVersion' -Default 0)
    if ($version -ne $script:SemtWorkerContractVersion) {
        throw ("SEMT_WORKER_CONTRACT_VERSION_MISMATCH: worker supports {0}, contract declares {1}. Launcher and worker come from different packages." -f $script:SemtWorkerContractVersion,$version)
    }
    foreach ($required in @('ContractId','Campaign','RunId','CorrelationId','Computer','ConnectionTarget','CycleNumber','PackageRoot','Endpoint','PsExec','Evidence','Flags')) {
        $value = Get-SemtPropertyValue -InputObject $contract -Name $required -Default $null
        if ($null -eq $value -or ($value -is [string] -and [string]::IsNullOrWhiteSpace($value))) { throw "SEMT_WORKER_CONTRACT_INVALID: missing '$required'." }
    }
    [void](Assert-SemtCampaign -Campaign ([string]$contract.Campaign))
    $engine = [string](Get-SemtPropertyValue -InputObject $contract -Name 'EngineVersion' -Default '')
    if ($engine -ne $script:SemtModuleVersion) {
        throw ("SEMT_WORKER_ENGINE_MISMATCH: worker engine {0} differs from launcher engine {1}." -f $script:SemtModuleVersion,$engine)
    }
    return $contract
}

#endregion

#region Canonical result

$script:SemtCoreResultFields = @(
    'Timestamp','Campaign','WaveName','RunId','CycleNumber','ComputerName','ConnectionTarget','CorrelationId',
    'Status','NextAction','LauncherStatus','RemoteStatus','RemoteNextAction','ExitCode','Detail',
    'DnsResolved','DnsAddressList','PingReachable','SmbPort445Reachable','AdminShareReachable','AdminShareFailureType',
    'RemoteExecutionMode','RemotePayloadCopyAttempts','LocalScriptVersion','LocalScriptSha256','RemoteScriptSha256','PsExecExitCode',
    'TargetGuardStatus','LegacyActivityStatus','LegacyActivityDetail','RemoteLogsPath','PsExecLogPath','JobErrorMessage','WorkerContractVersion'
)

function Get-SemtCoreResultFields {
    [CmdletBinding()]
    param()
    return @($script:SemtCoreResultFields)
}

function New-SemtResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [int]$CycleNumber = 0,
        [string]$WaveName = '',
        [string]$RunId = '',
        [string]$LauncherStatus = 'UNKNOWN',
        [string]$NextAction = '',
        [string]$Detail = '',
        [string[]]$ExtensionFields = @()
    )

    $result = [ordered]@{}
    foreach ($field in $script:SemtCoreResultFields) { $result[$field] = '' }
    foreach ($field in @($ExtensionFields)) { if (-not $result.Contains($field)) { $result[$field] = '' } }
    $result.Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $result.Campaign = $Campaign
    $result.WaveName = $WaveName
    $result.RunId = $RunId
    $result.CycleNumber = $CycleNumber
    $result.ComputerName = $ComputerName
    $result.LauncherStatus = $LauncherStatus
    $result.Status = $LauncherStatus
    $result.NextAction = $NextAction
    $result.Detail = $Detail
    $result.WorkerContractVersion = $script:SemtWorkerContractVersion
    return $result
}

function Complete-SemtResultStatus {
    <#
    Status is the effective status: the remote endpoint status when one was proven for the
    current run, otherwise the launcher status.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Result)

    $remote = [string]$Result.RemoteStatus
    $Result.Status = if (-not [string]::IsNullOrWhiteSpace($remote)) { $remote } else { [string]$Result.LauncherStatus }
    if ([string]::IsNullOrWhiteSpace([string]$Result.NextAction) -and -not [string]::IsNullOrWhiteSpace([string]$Result.RemoteNextAction)) { $Result.NextAction = [string]$Result.RemoteNextAction }
    return $Result
}

#endregion

#region Technician run guard history (campaign policies live in campaign modules)

function Get-SemtRunGuardHistoryPath {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Campaign)

    [void](Assert-SemtCampaign -Campaign $Campaign)
    return (Join-Path -Path (Join-Path -Path (Get-SemtLauncherStateRoot) -ChildPath $Campaign) -ChildPath 'RunGuardHistory.json')
}

function Invoke-SemtRunGuardHistoryTransaction {
    <#
    Reads the campaign history, lets the caller update its entries, prunes expired entries and
    saves it, all under the history lock. The caller's scriptblock receives the entry list.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][int]$SchemaVersion,
        [Parameter(Mandatory = $true)][scriptblock]$Prune,
        [scriptblock]$Update,
        [object[]]$ArgumentList = @()
    )

    $path = Get-SemtRunGuardHistoryPath -Campaign $Campaign
    return (Invoke-SemtWithFileLock -LockPath ('{0}.lock' -f $path) -TimeoutSeconds 60 -ArgumentList @($path,$SchemaVersion,$Prune,$Update,$ArgumentList,$Campaign) -ScriptBlock {
        param($LockedPath, $LockedSchema, $LockedPrune, $LockedUpdate, $LockedArguments, $LockedCampaign)
        $history = $null
        try { $history = Read-SemtJsonFile -Path $LockedPath } catch { $history = $null }
        $entries = New-Object System.Collections.Generic.List[object]
        if ($history -and [int](Get-SemtPropertyValue -InputObject $history -Name 'SchemaVersion' -Default 0) -eq $LockedSchema) {
            foreach ($entry in @($history.Entries)) { if ($null -ne $entry -and (& $LockedPrune $entry)) { $entries.Add($entry) } }
        }
        $output = $null
        if ($LockedUpdate) { $output = & $LockedUpdate $entries @LockedArguments }
        Write-SemtJsonFileAtomic -Path $LockedPath -InputObject ([pscustomobject]@{ SchemaVersion = $LockedSchema; Product = $script:SemtProductName; Campaign = $LockedCampaign; UpdatedUtc = Get-SemtUtcNowText; Entries = @($entries.ToArray()) })
        return $output
    })
}

function Get-SemtRunGuardFqdn {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [AllowNull()][hashtable]$AdInventoryMap
    )

    $name = ConvertTo-SemtComputerName -Line $ComputerName
    if ([string]::IsNullOrWhiteSpace($name)) { return '' }
    if ($name.Contains('.')) { return $name.ToLowerInvariant() }
    $key = Get-SemtComputerKey -ComputerName $name
    if ($AdInventoryMap -and $AdInventoryMap.ContainsKey($key)) {
        $dns = [string](Get-SemtPropertyValue -InputObject $AdInventoryMap[$key] -Name 'DNSHostName' -Default '')
        if (-not [string]::IsNullOrWhiteSpace($dns)) { return $dns.Trim().TrimEnd('.').ToLowerInvariant() }
    }
    return $name.ToLowerInvariant()
}

#endregion

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDfQS/Z6P1vUYxd
# cYRkYtkLNJFMQm2DyswCgMJsikkcxaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBq/d1w3OCIityaoH4l83ZY
# /YjjFhfOQSEgGdXbM1XZSDANBgkqhkiG9w0BAQEFAASCAYB7r5ei9hTGRmz3Gzt5
# Qa9bGryUvHz3zjEP+jYCviVK2tpxKEnOQcgy+Eyc2PTk7bB3XUqy5tT6H++iEKFs
# Ktt2nxopsvfn6K2Yr4RNyiUAle+IAyH9N4OyOpl6Zx2qUSNQ/iBn/NdJRkf502y2
# 7tU/17hv5fnNBYYRLzCbTXgYBFLlbqGEYmQUfrXNWuEpJJdYiVaWWj5CFXmgjpdm
# RJr8gXCQqrKGUj3QzqZ2ywTsdgU6dyQi/WW5OqWt0RbqF6YAKmtp5H8mOEV3deeG
# oVBT4VGOuYQDOCqJO5QbwHlPa50O/qpuhFuEQSmRciAW8p3/tHvFCxz1z3vQ34RS
# O5srWg0kDDkMViEl1vdCK2TnUj2rqTkigNZRlRKdZWZngpXgL+bLbb/19DqD+xRU
# imnc4b3dZtq84KFaRyXFxIwIM334wyF7Qvh1qFo9wFk+RcJVZ2WZJu0hjISRiy/f
# g2bNgf7lHktKjMfBDmxImMFiKhHZcowJc+2Ubk4d91TJ5wM=
# SIG # End signature block
