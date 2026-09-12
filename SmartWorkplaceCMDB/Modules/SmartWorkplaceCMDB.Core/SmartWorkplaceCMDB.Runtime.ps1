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
        [ValidateSet('Validate', 'Collect', 'Fixture')][string]$Mode = 'Validate',
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
    $requiresGraph = $requiresExternalDependency -and $Pipeline -notin @('ActiveDirectory', 'CuratedOnly')
    $requiresExchange = $requiresExternalDependency -and $Pipeline -in @('Full', 'ExchangeOnlineMailboxes')
    $requiresAd = $requiresExternalDependency -and $Pipeline -in @('Full', 'ActiveDirectory')
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
    if ($Mode -eq 'Collect') {
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
        Handle=$handle; LockPath=$lockPath; StatePath=(Join-Path $folder ($pipelineName + '.state.json'))
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDv+H9igZWlXWEQ
# jlhSP4+dykY7BiQIq3nS7/ai91pMw6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJ+x9JY61onMk5aCHeaMnuobX25yqgGTCMwaTp/y+To6MA0GCSqG
# SIb3DQEBAQUABIIBgKPaXszqiod14+p1qnzrzc9Z/8PtH1/ooWsaiI2a/szlx42+
# zR1LH6/9WnG3sx87yMNU68sEA62CkKZNZqBzLS73MQaFkymC1xnShHMym/PN7/O0
# TjbfjEgJteUuld4jEjEwHWYeWqEUzX48sv76liCpXGDaIMg/Ui9pIrJIowcY8RZZ
# FLRauBcWJbuMwjw48DNOgubqswrbxNjecQvNP9OVhWqe34RnRlCv4huD1jrV0wCZ
# 7Bfl/II6ZPVdzieKPhpSerTwI5rlvOXn0bo8eet6FyRX5dMZJH91rry+OMhgSokO
# eIOu0124dOzNrnyqNgf7RJs0kV0SeApHRg+WQmsEvEdczw4EYh1SMe1ticCaMzVa
# YWcHK01hswL5aFac7nJz7+Cexn9xrOh430jHadaBNIeJsmCKQ/6Yxr/LVgk/Dbhd
# 4UEuptS8ATvXa5V7rvDcRABBONvAHw2YQkNzrju0+trrN1LjDP0PP+NrutFfDcqO
# 7HHz+3psNOAc6E9p9KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxNzE4
# MDhaMC8GCSqGSIb3DQEJBDEiBCAv4M0Z4U/08juKJhQwhAwgSYg4ojE6S3kqn9KK
# NT6LizANBgkqhkiG9w0BAQEFAASCAgAga95AUthnUlylLjHCOBKZljFxAbSukaO7
# V7NOQKJk9pBKysCEVPQ4Xh/JQ0Th6NJm1X1obNg/cSQZw54uh2pZ1XlWEJ6TifIN
# 3RTZ8t1B6I3Acb3w0dgWl4vXiIQBHR/Pf7whVfPt1Ac7ctOJgr2EIAT8o9kjeNna
# 2xhjrTWDCI2x/9Z1nXBL8EFqvVDhXB3vTEmlGMMh9Aut6d2L+n5i+qeP4Gxpkfqo
# JMxIl4XODQLXzbh5h4OGy8DtQN75H1yOGm7kRaiy2uvD5628ikLojVRYcxmAWd5s
# mayC3WtZGkS2b3xhVk/yNj0oJGkq9ojzQJMxngMH76yHRBbP34vdfXepHD8ZSTt1
# xX2eAQEEs8Yc74rciM/YMWpa23QF2PyygaO9Ehg1+bYp8fq7FMBtBdJUAT6OFc7w
# 2W8/YAtIY2NbMh1PMu1uV3UE5tmQ88EW+ygX5cL1Z4L9FQN5talPqQbG7Jk1Ap0t
# GBotuM/VjcyYyD8aYhhLo/tnnmQeonRGwixDxWDscQPKrEfIQXYUMRzVjxIuvbQI
# gJVcdMfyqcu8zJZND69yl0kGErDaUE/3A9OSmintvAAgrlPRQohjgdpVl0J6OP4b
# DPqUf5lIWJ/3qDCI2d8Jpa+jAohVlBNnWDm1cg4q0QlPei+upxaC3EHIl6jkcst1
# kQXEmZNxow==
# SIG # End signature block
