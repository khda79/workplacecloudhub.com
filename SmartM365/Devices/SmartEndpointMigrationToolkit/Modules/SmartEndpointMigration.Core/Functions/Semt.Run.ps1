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
    'RemotePayloadCopyAttempts','LocalScriptVersion','LocalScriptSha256','RemoteScriptSha256','PsExecExitCode',
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAumtjWN6H1uqKc
# +duk+2dDkMhmVDXR07jwHhUJ6ZTHkKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHLHGGxRhgLdwa5uzhLQYlcEw71lqSNklGMzTy7U8nfpMA0GCSqG
# SIb3DQEBAQUABIIBgFq7sS8Hz9uDjAUT6LrJAMjQh0UnVatZY1lrtUms3wf4etuB
# fmUqPvCI/pfEVTfa4fYmPDqE5+USQUiOBAzYS07tYJvE/XHVS+SDd4NTmbXzLneY
# nfewwafBzWkgt/e0kGzD6GCJHSJTtCUDS5yUyPu/s41JSRYLV/GPFsKWOE6DOpV1
# lGUjwdp0ndzGQFfP93LGMCXtcI2dse7wu3bkc3DgcLDPlqr9Tn6Z0CDnd5h/Or9m
# rLN5JKd/CzNQDMFrpU6Kcrc+AT/iC5/BSYa4gaFhSbc9Hi+m0DfX6oOEpphCVd2K
# Om7f8G3mLzEiq1WrCUW2IdgMXk1qVaZ301PGVDh0vVQ2djK6dnpSx86tF10dfOiY
# Z8lAaSqWDrj1vvg9bio8EZ3VSEFSXSsm4KsD7eFiHozNmo5KRRYTNwwSdtfSLfuq
# mDFxWVKcehJ7LTAsIItshWMwZt0qvjsX97zU5DJ40nqeKDa2M2JhKFdH82HHnfKv
# ZRA2Se3p6FM646XXE6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMzE2
# NTNaMC8GCSqGSIb3DQEJBDEiBCD27cXo+P+kisfZFxIaNXSCz7qxd84wwDO31u35
# 1HeeGDANBgkqhkiG9w0BAQEFAASCAgA/PVyc8vwiE8aKwZ/eh312WbBugYYW9JX6
# ZJ+KhTj6+Fdfyxb7OIpmdiljK1YlLLHheYRpeGECYpL4J/TfaOm9BkxqCHUPhvu6
# 77Cjg/BsSki5uJk5AFfP3s5BV2es+74NBMA4ea8Dz5ADYZ0KCM1LVPHr9F1BXuQ6
# Q49cbffJqQblj6ZWJ+5pdoMUqy8WQ5wqtiTU0S/z1H7d9APiFCkgL5stG87kzvme
# NMC1mpYBI8I3ooBexMPJuJZf8Umxvq6ug1kV6FmZEL1C9amz6gKXonOLOTHrF1i6
# emmz4jGjC0uOHIlF+FyDHIn42z0vj0/CZQaBWteKGOKiANyLeBO0rl5xMNl2PxZj
# dl5J73Yr3mxLjTbg4jFIVEFJxHhQKzsPC5oao0SkcXsPQK6l+UEG9nuq4xyWyx4/
# ofb1UP3H6Hx+G19MruTEABpZvNnnb0i8Vnl77jiQ3xyCDedv29k/7LzLYg+2eo0J
# PijleZNxGoXZ1iJ1sCIUrG2HAorBKVcMgylrijmIYsL0Mk85G5jelvpqvGXKsziD
# LGvltwjTkVWtV3axTwQpfkADMxan66+6cdgf0Z0fXlhYMtgwVODhRXKlZCvOM0oo
# sKxN+ksDKs2NEmzCYWGcN9vTBN9lVEScyY2moafl/RJf3Q22UBwBZDd0klbfx3p9
# WRLzkDYvrg==
# SIG # End signature block
