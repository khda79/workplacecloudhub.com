# SmartWorkplaceCMDB runtime, preflight, logging, and run-state helpers.
# Version: 1.0.2

function Get-SmartWorkplaceCMDBRuntimeSetting {
    [CmdletBinding()]
    param(
        [AllowNull()]$Configuration,
        [Parameter(Mandatory)][string]$Name,
        $DefaultValue
    )

    if ($null -eq $Configuration) { return $DefaultValue }
    if ($Configuration -is [Collections.IDictionary] -and $Configuration.Contains($Name)) {
        return $Configuration[$Name]
    }
    $property = $Configuration.PSObject.Properties[$Name]
    if ($null -ne $property) { return $property.Value }
    return $DefaultValue
}

function ConvertTo-SmartWorkplaceCMDBSafeFileName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Value)

    $name = [regex]::Replace($Value, '[^A-Za-z0-9._-]', '-')
    return $name.Trim('-')
}

function Write-SmartWorkplaceCMDBRuntimeLog {
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$Path,
        [AllowEmptyString()][string]$Message,
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')][string]$Level = 'INFO',
        [switch]$Console
    )

    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
    foreach ($line in @([string]$Message -split "`r?`n")) {
        $entry = '[{0}] [{1}] {2}' -f $timestamp, $Level, $line
        if (-not [string]::IsNullOrWhiteSpace($Path)) {
            $folder = Split-Path -Parent $Path
            if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
                New-Item -ItemType Directory -Path $folder -Force | Out-Null
            }
            [IO.File]::AppendAllText(
                $Path,
                ($entry + [Environment]::NewLine),
                (New-Object Text.UTF8Encoding($false))
            )
        }
        if ($Console) {
            $color = switch ($Level) {
                'ERROR' { 'Red' }
                'WARN' { 'Yellow' }
                'DEBUG' { 'DarkGray' }
                default { 'Gray' }
            }
            Microsoft.PowerShell.Utility\Write-Host $entry -ForegroundColor $color
        }
    }
}

function Remove-SmartWorkplaceCMDBRuntimeLogOverflow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$Filter,
        [ValidateRange(0, 36500)][int]$RetentionDays = 30,
        [ValidateRange(0, 100000)][int]$MaxFiles = 30,
        [string[]]$ExcludePath = @()
    )

    if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) { return }
    $excluded = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($path in @($ExcludePath)) {
        if (-not [string]::IsNullOrWhiteSpace($path)) {
            [void]$excluded.Add([IO.Path]::GetFullPath($path))
        }
    }
    $files = @(Get-ChildItem -LiteralPath $FolderPath -Filter $Filter -File -ErrorAction SilentlyContinue)
    if ($RetentionDays -gt 0) {
        $cutoff = (Get-Date).AddDays(-1 * $RetentionDays)
        foreach ($file in @($files | Where-Object {
                    $_.LastWriteTime -lt $cutoff -and -not $excluded.Contains($_.FullName)
                })) {
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue
        }
    }
    if ($MaxFiles -gt 0) {
        foreach ($file in @(Get-ChildItem -LiteralPath $FolderPath -Filter $Filter -File -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTimeUtc, Name -Descending |
                Select-Object -Skip $MaxFiles)) {
            if (-not $excluded.Contains($file.FullName)) {
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

function Start-SmartWorkplaceCMDBExecutionContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Context,
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][string]$ScriptVersion,
        [AllowEmptyString()][string]$Mode = '',
        [switch]$NoWrite
    )

    $started = [datetimeoffset]::Now
    $scriptName = [IO.Path]::GetFileNameWithoutExtension($ScriptPath)
    $logging = Get-SmartWorkplaceCMDBRuntimeSetting $Context.Configuration 'Logging' $null
    $enabled = [bool](Get-SmartWorkplaceCMDBRuntimeSetting $logging 'Enabled' $true)
    $parentRunId = [string]$env:SMARTWORKPLACECMDB_PARENT_RUN_ID
    if (-not [string]::IsNullOrWhiteSpace($parentRunId)) {
        return [pscustomobject]@{
            Status = 'Delegated'; StartedDateTime = $started; ScriptName = $scriptName
            ScriptVersion = $ScriptVersion; Mode = $Mode; LogPath = ''
            TranscriptPath = ''; TranscriptStarted = $false; ParentRunId = $parentRunId
        }
    }
    if (-not $enabled -or $NoWrite) {
        return [pscustomobject]@{
            Status = $(if ($NoWrite) { 'NoWrite' } else { 'Disabled' })
            StartedDateTime = $started; ScriptName = $scriptName
            ScriptVersion = $ScriptVersion; Mode = $Mode; LogPath = ''
            TranscriptPath = ''; TranscriptStarted = $false; ParentRunId = ''
        }
    }

    $retentionDays = [math]::Max(0, [int](Get-SmartWorkplaceCMDBRuntimeSetting $logging 'StepLogRetentionDays' 30))
    $maxFiles = [math]::Max(0, [int](Get-SmartWorkplaceCMDBRuntimeSetting $logging 'MaxStepLogsPerScript' 30))
    $folder = Join-Path (Join-Path $Context.Paths.LogRootPath 'Jobs') (ConvertTo-SmartWorkplaceCMDBSafeFileName $scriptName)
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    $computerName = if ([string]::IsNullOrWhiteSpace([string]$env:COMPUTERNAME)) { 'HOST' } else { [string]$env:COMPUTERNAME }
    $stamp = $started.ToString('yyyyMMdd-HHmmssfff')
    $baseName = '{0}_{1}_{2}_{3}' -f $scriptName, (ConvertTo-SmartWorkplaceCMDBSafeFileName $computerName), $stamp, $PID
    $logPath = Join-Path $folder ($baseName + '.log')
    $transcriptPath = Join-Path $folder ($baseName + '.transcript.txt')
    Remove-SmartWorkplaceCMDBRuntimeLogOverflow $folder '*.log' $retentionDays $maxFiles @($logPath)
    Remove-SmartWorkplaceCMDBRuntimeLogOverflow $folder '*.transcript.txt' $retentionDays $maxFiles @($transcriptPath)

    Start-Transcript -LiteralPath $transcriptPath -Force -ErrorAction Stop | Out-Null
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $lines = @(
        ('=' * 80),
        'SmartWorkplaceCMDB script started',
        ('Script={0}; Version={1}; Mode={2}' -f $scriptName, $ScriptVersion, $Mode),
        ('Tenant={0}; Organization={1}; Environment={2}' -f $Context.Paths.TenantKey, $Context.Paths.OrganizationKey, $Context.Paths.EnvironmentKey),
        ('Computer={0}; User={1}; PID={2}' -f $computerName, $identity, $PID),
        ('PowerShell={0}; Edition={1}; ScriptPath={2}' -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition, $ScriptPath),
        ('Log={0}; Transcript={1}' -f $logPath, $transcriptPath),
        ('=' * 80)
    )
    foreach ($line in $lines) {
        Write-SmartWorkplaceCMDBRuntimeLog -Path $logPath -Message $line -Console
    }
    return [pscustomobject]@{
        Status = 'Started'; StartedDateTime = $started; ScriptName = $scriptName
        ScriptVersion = $ScriptVersion; Mode = $Mode; LogPath = $logPath
        TranscriptPath = $transcriptPath; TranscriptStarted = $true; ParentRunId = ''
    }
}

function Complete-SmartWorkplaceCMDBExecutionContext {
    [CmdletBinding()]
    param(
        [AllowNull()]$RuntimeContext,
        [AllowNull()][System.Management.Automation.ErrorRecord]$ErrorRecord,
        [ValidateRange(0, 2147483647)][int]$WarningCount = 0
    )

    if ($null -eq $RuntimeContext -or $RuntimeContext.Status -ne 'Started') { return }
    $ended = [datetimeoffset]::Now
    $failed = $null -ne $ErrorRecord
    $status = if ($failed) { 'Failed' } elseif ($WarningCount -gt 0) { 'CompletedWithWarnings' } else { 'Completed' }
    if ($failed) {
        Write-SmartWorkplaceCMDBRuntimeLog -Path $ExecutionContext.LogPath `
            -Message $ErrorRecord.Exception.Message -Level ERROR -Console
    }
    foreach ($line in @(
            ('=' * 80),
            'SmartWorkplaceCMDB script completed',
            ('Status={0}; Duration={1}; Warnings={2}; Errors={3}' -f
                $status, ($ended - $RuntimeContext.StartedDateTime).ToString('hh\:mm\:ss'),
                $WarningCount, [int]$failed),
            ('Log={0}; Transcript={1}' -f $RuntimeContext.LogPath, $RuntimeContext.TranscriptPath),
            ('=' * 80)
        )) {
        Write-SmartWorkplaceCMDBRuntimeLog -Path $RuntimeContext.LogPath -Message $line -Console
    }
    if ($RuntimeContext.TranscriptStarted) {
        try { Stop-Transcript -ErrorAction Stop | Out-Null } catch {
            Write-SmartWorkplaceCMDBRuntimeLog -Path $RuntimeContext.LogPath `
                -Message ("Stop-Transcript failed: {0}" -f $_.Exception.Message) -Level WARN
        }
    }
}

function Test-SmartWorkplaceCMDBPreflight {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Context,
        [Parameter(Mandatory)][string]$ProjectRoot,
        [Parameter(Mandatory)][string]$Pipeline,
        [ValidateSet('Validate', 'Collect', 'Finalize', 'Fixture')][string]$Mode = 'Validate',
        [string[]]$ScriptPath = @(),
        [switch]$ThrowOnFailure
    )

    $checks = New-Object 'Collections.Generic.List[object]'
    $add = {
        param([string]$Name, [string]$Status, [string]$Details)
        $checks.Add([pscustomobject]@{Name=$Name;Status=$Status;Details=$Details})
    }
    & $add 'PowerShellVersion' $(if ($PSVersionTable.PSVersion.Major -ge 7) {'Passed'} else {'Failed'}) ([string]$PSVersionTable.PSVersion)
    & $add 'ProjectRoot' $(if (Test-Path -LiteralPath $ProjectRoot -PathType Container) {'Passed'} else {'Failed'}) $ProjectRoot

    # Fixture tests must remain deterministic and offline: external modules and
    # tenant credentials are validated only for real validate/collect runs.
    $requiresExternalDependency = $Mode -ne 'Fixture'
    $notifications = Get-SmartWorkplaceCMDBRuntimeSetting `
        $Context.Configuration 'Notifications' $null
    $sharePoint = Get-SmartWorkplaceCMDBRuntimeSetting `
        $Context.Configuration 'SharePoint' $null
    $mailMode = [string](Get-SmartWorkplaceCMDBRuntimeSetting `
        $notifications 'SendMailMode' 'Graph')
    $finalizeNeedsGraph = $Mode -eq 'Finalize' -and (
        [bool](Get-SmartWorkplaceCMDBRuntimeSetting `
            $sharePoint 'Enabled' $false) -or
        ([bool](Get-SmartWorkplaceCMDBRuntimeSetting `
                $notifications 'Enabled' $false) -and
            $mailMode.ToUpperInvariant() -in @('GRAPH', 'BOTH'))
    )
    $requiresGraph = $requiresExternalDependency -and
        ($finalizeNeedsGraph -or
            ($Mode -ne 'Finalize' -and
                $Pipeline -notin @('ActiveDirectory', 'CuratedOnly')))
    $requiresExchange = $requiresExternalDependency -and $Mode -ne 'Finalize' -and
        $Pipeline -in @('Full', 'ExchangeOnlineMailboxes')
    $requiresAd = $requiresExternalDependency -and $Mode -ne 'Finalize' -and
        $Pipeline -in @('Full', 'ActiveDirectory')
    if ($requiresGraph) {
        $graphModule = Get-Module -ListAvailable Microsoft.Graph.Authentication | Sort-Object Version -Descending | Select-Object -First 1
        & $add 'MicrosoftGraphModule' $(if ($graphModule) {'Passed'} else {'Failed'}) $(if ($graphModule) {[string]$graphModule.Version} else {'Not installed'})
        $graph = Get-SmartWorkplaceCMDBRuntimeSetting $Context.Configuration 'MicrosoftGraph' $null
        foreach ($name in @('TenantId','ClientId','CertificateThumbprint')) {
            $value = [string](Get-SmartWorkplaceCMDBRuntimeSetting $graph $name '')
            & $add ("MicrosoftGraph.{0}" -f $name) $(if ($value) {'Passed'} else {'Failed'}) $(if ($value) {'Configured'} else {'Missing'})
        }
    }
    if ($requiresExchange) {
        $exoModule = Get-Module -ListAvailable ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
        & $add 'ExchangeOnlineModule' $(if ($exoModule) {'Passed'} else {'Failed'}) $(if ($exoModule) {[string]$exoModule.Version} else {'Not installed'})
    }
    if ($requiresAd) {
        try {
            Add-Type -AssemblyName System.DirectoryServices.Protocols -ErrorAction Stop
            & $add 'ActiveDirectoryProtocols' 'Passed' 'System.DirectoryServices.Protocols available'
        }
        catch { & $add 'ActiveDirectoryProtocols' 'Failed' $_.Exception.Message }
    }
    if ($Mode -in @('Collect', 'Finalize')) {
        try {
            Initialize-SmartWorkplaceCMDBTenantFolder -Paths $Context.Paths | Out-Null
            $probe = Join-Path $Context.Paths.LogRootPath ('.preflight-{0}.tmp' -f [guid]::NewGuid().ToString('N'))
            [IO.File]::WriteAllText($probe, 'ok', (New-Object Text.UTF8Encoding($false)))
            Remove-Item -LiteralPath $probe -Force -ErrorAction Stop
            & $add 'OutputWriteAccess' 'Passed' $Context.Paths.LogRootPath
        }
        catch { & $add 'OutputWriteAccess' 'Failed' $_.Exception.Message }
    }

    $logging = Get-SmartWorkplaceCMDBRuntimeSetting $Context.Configuration 'Logging' $null
    $signaturePolicy = [string](Get-SmartWorkplaceCMDBRuntimeSetting $logging 'ScriptSignaturePolicy' 'Audit')
    if ($signaturePolicy -notin @('Disabled','Audit','Enforce')) {
        & $add 'ScriptSignaturePolicy' 'Failed' "Unsupported value '$signaturePolicy'."
    }
    elseif ($signaturePolicy -ne 'Disabled' -and $env:OS -eq 'Windows_NT') {
        foreach ($path in @($ScriptPath | Sort-Object -Unique)) {
            try {
                $signature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $path -ErrorAction Stop
                $valid = $signature.Status -eq 'Valid'
                $status = if ($valid) {'Passed'} elseif ($signaturePolicy -eq 'Enforce') {'Failed'} else {'Warning'}
                & $add ('Signature:' + [IO.Path]::GetFileName($path)) $status ([string]$signature.Status)
            }
            catch {
                # Some AllSigned hosts reject the inbox Security.types.ps1xml
                # while auto-loading Microsoft.PowerShell.Security. Audit must
                # report that condition without preventing collection; Enforce
                # still blocks because the signature could not be established.
                $status = if ($signaturePolicy -eq 'Enforce') {'Failed'} else {'Warning'}
                & $add ('SignatureEngine:' + [IO.Path]::GetFileName($path)) $status $_.Exception.Message
                break
            }
        }
    }

    $failed = @($checks | Where-Object Status -eq 'Failed')
    $warnings = @($checks | Where-Object Status -eq 'Warning')
    $result = [pscustomobject]@{
        Status = if ($failed.Count -gt 0) {'Failed'} elseif ($warnings.Count -gt 0) {'PassedWithWarnings'} else {'Passed'}
        Pipeline = $Pipeline; Mode = $Mode; FailedCount = $failed.Count
        WarningCount = $warnings.Count; Checks = $checks.ToArray()
    }
    if ($ThrowOnFailure -and $failed.Count -gt 0) {
        throw ('SmartWorkplaceCMDB preflight failed: ' + (@($failed | ForEach-Object {"$($_.Name)=$($_.Details)"}) -join '; '))
    }
    return $result
}

function Enter-SmartWorkplaceCMDBRunGuard {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Paths,
        [Parameter(Mandatory)][string]$Pipeline,
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)][datetimeoffset]$StartedDateTime,
        [switch]$NoWrite
    )

    if ($NoWrite) { return $null }
    $tenantName = ConvertTo-SmartWorkplaceCMDBSafeFileName ([string]$Paths.TenantKey)
    $pipelineName = ConvertTo-SmartWorkplaceCMDBSafeFileName $Pipeline
    $folder = Join-Path (Join-Path $Paths.LogRootPath 'Orchestration\State') $tenantName
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    $lockPath = Join-Path $folder ($pipelineName + '.lock')
    try {
        $handle = [IO.File]::Open($lockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    }
    catch {
        throw "Another SmartWorkplaceCMDB '$Pipeline' run already owns the tenant guard '$lockPath'."
    }
    $state = [ordered]@{
        RunId=$RunId; TenantKey=[string]$Paths.TenantKey; Pipeline=$Pipeline
        Status='Starting'; ComputerName=[string]$env:COMPUTERNAME; ProcessId=$PID
        StartedDateTime=$StartedDateTime.ToString('o'); LastHeartbeatDateTime=[datetimeoffset]::UtcNow.ToString('o')
        CurrentStep=''; CompletedStepCount=0; LogPath=''; TranscriptPath=''; Error=''
    }
    $ownerJson = $state | ConvertTo-Json -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($ownerJson)
    $handle.SetLength(0); $handle.Write($bytes, 0, $bytes.Length); $handle.Flush($true)
    $guard = [pscustomobject]@{
        Handle=$handle; LockPath=$lockPath; StatePath=(Join-Path $folder ($pipelineName + '.state.json.txt'))
        State=$state; Released=$false
    }
    Update-SmartWorkplaceCMDBRunState -RunGuard $guard -Status 'Running'
    return $guard
}

function Update-SmartWorkplaceCMDBRunState {
    [CmdletBinding()]
    param(
        [AllowNull()]$RunGuard,
        [AllowEmptyString()][string]$Status = '',
        [AllowEmptyString()][string]$CurrentStep = '',
        [int]$CompletedStepCount = -1,
        [AllowEmptyString()][string]$LogPath = '',
        [AllowEmptyString()][string]$TranscriptPath = '',
        [AllowEmptyString()][string]$Error = ''
    )

    if ($null -eq $RunGuard -or $RunGuard.Released) { return }
    if ($Status) { $RunGuard.State.Status = $Status }
    if ($PSBoundParameters.ContainsKey('CurrentStep')) { $RunGuard.State.CurrentStep = $CurrentStep }
    if ($CompletedStepCount -ge 0) { $RunGuard.State.CompletedStepCount = $CompletedStepCount }
    if ($PSBoundParameters.ContainsKey('LogPath')) { $RunGuard.State.LogPath = $LogPath }
    if ($PSBoundParameters.ContainsKey('TranscriptPath')) { $RunGuard.State.TranscriptPath = $TranscriptPath }
    if ($PSBoundParameters.ContainsKey('Error')) { $RunGuard.State.Error = $Error }
    $RunGuard.State.LastHeartbeatDateTime = [datetimeoffset]::UtcNow.ToString('o')
    Write-SmartWorkplaceCMDBJsonAtomically -InputObject $RunGuard.State -Path $RunGuard.StatePath
}

function Exit-SmartWorkplaceCMDBRunGuard {
    [CmdletBinding()]
    param(
        [AllowNull()]$RunGuard,
        [Parameter(Mandatory)][string]$Status,
        [AllowEmptyString()][string]$Error = ''
    )

    if ($null -eq $RunGuard -or $RunGuard.Released) { return }
    try {
        Update-SmartWorkplaceCMDBRunState -RunGuard $RunGuard -Status $Status -CurrentStep '' -Error $Error
    }
    finally {
        $RunGuard.Handle.Dispose()
        $RunGuard.Released = $true
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB7aUCqt3TypuLA
# B2JYBGNWEkJWcBjnn38PAJ5SI9/GSqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCwp+qC0TDxwjCwNjWrP0OA
# nXFJpWInZGfecMuZ8NMXDjANBgkqhkiG9w0BAQEFAASCAYAfsQlMoq1GqYdS+cFI
# nZXHhqvHDxSx8+7MDNvPObsZ1ZjbQBDAnQcBR96uKx2veUHYKHAHYgCDwFIIReIY
# vsEy9hMfb2Drx3Fvz18xs0Ir0yEl6XokugFyGnAKPSrI+nr963pXKK3p+f0edZAV
# TwItKq639osTeVRqx3Ixg8tOX/y9XOCPSP3UilBp59hCyu9wI6ZCogX9aS+l5iGT
# yex7OzUZkssn5M5c8PsJ70cnbGp9xw3uQHPCdKMBzXoFrDCOZfHCgoMUjwHcpDps
# 2tRJWXCSBtQDIeCM1Zxjb0mkM31xayJjUyphFKRBgKWvg8zo3k31QAyTBLi9qJWy
# 3P457LEb3QOtdASJhl4rn/SQMF7wJ21CUxMkM5LlP2xsayGA8T/J05ST3ee59vR3
# uncz0k1KyJbivp4SEZXA4po5QV0b1lVk8qM9hD3Zg7IWaf/OvkZXe8iA+bgMtXhm
# CzWYjSVPAVS1jX72+Uzh6d6sfSCQ+iGmGvRMRWQn5xQZqqI=
# SIG # End signature block
