# Shared scheduling control. Importing this module never touches runtime data.
Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot '../../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -MinimumVersion '1.0.0' -Global -ErrorAction Stop

function Get-SmartM365OrchestratorMaintenancePaths {
    param([Parameter(Mandatory)][string]$SharedDataFolderPath)
    $config = Join-Path $SharedDataFolderPath 'Config'
    [pscustomobject]@{
        State = Join-Path $config 'Orchestrator-Maintenance.json.txt'
        Gate = Join-Path $config 'Orchestrator-Maintenance.guard'
        Audit = Join-Path $SharedDataFolderPath 'Audit/Orchestrator_Maintenance.csv'
        MailCheckpoint = Join-Path $config 'Orchestrator-Maintenance-Mail.json.txt'
        MailGate = Join-Path $config 'Orchestrator-Maintenance-Mail.guard'
    }
}

function New-MaintenanceDefaultState {
    [pscustomobject][ordered]@{
        SchemaVersion = 1; Revision = 0; Enabled = $false
        ChangedAtUtc = ''; ChangedBy = ''; ChangedFromServer = ''; Reason = ''
        ResumeAfterUtc = ''
    }
}

function Assert-MaintenanceState {
    param([Parameter(Mandatory)]$Document)
    foreach ($name in @('SchemaVersion','Revision','Enabled','ChangedAtUtc','ChangedBy','ChangedFromServer','Reason','ResumeAfterUtc')) {
        if (-not $Document.PSObject.Properties[$name]) { throw "Maintenance state is missing $name." }
    }
    if (($Document.SchemaVersion -isnot [int] -and $Document.SchemaVersion -isnot [long]) -or $Document.SchemaVersion -ne 1 -or $Document.Enabled -isnot [bool] -or
        $Document.Revision -isnot [long] -and $Document.Revision -isnot [int] -or $Document.Revision -lt 0) {
        throw 'Maintenance state schema, revision or enabled flag is invalid.'
    }
    foreach ($name in @('ChangedBy','ChangedFromServer','Reason')) {
        if ($Document.$name -isnot [string]) { throw "Maintenance state $name must be a string." }
    }
    foreach ($name in @('ChangedAtUtc','ResumeAfterUtc')) {
        # ConvertFrom-Json in recent PowerShell versions materializes ISO dates.
        if ($Document.$name -is [datetime]) { $Document.$name = $Document.$name.ToUniversalTime().ToString('o') }
        elseif ($Document.$name -is [datetimeoffset]) { $Document.$name = $Document.$name.UtcDateTime.ToString('o') }
        if ($Document.$name -isnot [string]) { throw "Maintenance state $name must be a timestamp string." }
        if ($Document.$name) {
            $parsed = [datetimeoffset]::MinValue
            if ($Document.$name -notmatch '(Z|\+00:00)$' -or -not [datetimeoffset]::TryParse($Document.$name, [ref]$parsed)) {
                throw "Maintenance state $name must be a valid UTC timestamp."
            }
        }
    }
    if ($Document.Revision -gt 0 -and (-not $Document.ChangedAtUtc -or -not $Document.ChangedBy -or -not $Document.Reason)) {
        throw 'Maintenance transition has no timestamp, actor or reason.'
    }
    if ($Document.Enabled -and $Document.Revision -eq 0) { throw 'Initial maintenance state cannot be enabled.' }
    if (-not $Document.Enabled -and $Document.Revision -gt 0 -and $Document.ResumeAfterUtc -ne $Document.ChangedAtUtc) {
        throw 'Maintenance resume cutoff does not match the disabled transition.'
    }
    if ($Document.ResumeAfterUtc -and $Document.ChangedAtUtc -and
        [datetimeoffset]::Parse($Document.ResumeAfterUtc) -gt [datetimeoffset]::Parse($Document.ChangedAtUtc)) {
        throw 'Maintenance resume cutoff is later than the transition.'
    }
}

function Get-SmartM365OrchestratorMaintenanceState {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath)
    $paths = Get-SmartM365OrchestratorMaintenancePaths $SharedDataFolderPath
    # Distinguish an unavailable share from a truly new deployment.
    $null = Get-Item -LiteralPath (Join-Path $SharedDataFolderPath 'Config') -ErrorAction Stop
    if (-not (Test-Path -LiteralPath $paths.State -PathType Leaf)) {
        if (Test-Path -LiteralPath $paths.Gate) { throw 'Previously initialized maintenance control is missing. Launches remain paused.' }
        return New-MaintenanceDefaultState
    }
    $state = (Read-SmartM365JsonDocument -Path $paths.State -Validate { param($d) Assert-MaintenanceState $d }).Document
    Assert-MaintenanceState $state
    return $state
}

function Enter-SmartM365OrchestratorMaintenanceGate {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [ValidateRange(0,60)][int]$TimeoutSeconds = 10)
    $paths = Get-SmartM365OrchestratorMaintenancePaths $SharedDataFolderPath
    $null = Get-Item -LiteralPath (Split-Path $paths.Gate -Parent) -ErrorAction Stop
    $deadline = [datetime]::UtcNow.AddSeconds($TimeoutSeconds)
    do {
        try { return [IO.File]::Open($paths.Gate, 'OpenOrCreate', 'ReadWrite', 'None') }
        catch [IO.IOException] { if ([datetime]::UtcNow -ge $deadline) { throw }; Start-Sleep -Milliseconds 100 }
    } while ([datetime]::UtcNow -lt $deadline)
    throw 'Maintenance launch gate is unavailable.'
}

function Initialize-SmartM365OrchestratorMaintenance {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath)
    $paths = Get-SmartM365OrchestratorMaintenancePaths $SharedDataFolderPath
    # Legacy, non-central deployments may have no Config folder yet.
    $null = Get-Item -LiteralPath $SharedDataFolderPath -ErrorAction Stop
    $null = New-Item -ItemType Directory -Path (Split-Path $paths.State -Parent) -Force -ErrorAction Stop
    $gate = Enter-SmartM365OrchestratorMaintenanceGate $SharedDataFolderPath
    try {
        if (-not (Test-Path -LiteralPath $paths.State)) {
            if ($gate.Length -ne 0) { throw 'Maintenance state disappeared after initialization; refusing automatic reset.' }
            # Persistent marker in the shared gate: deleting state must never disable maintenance.
            $marker = [Text.Encoding]::UTF8.GetBytes('MaintenanceProtocol=1')
            $gate.Write($marker,0,$marker.Length); $gate.Flush($true)
            $document = New-MaintenanceDefaultState
            $bytes = [Text.UTF8Encoding]::new($false).GetBytes(($document | ConvertTo-Json -Depth 10))
            $null = Write-SmartM365JsonBytesAtomically -Path $paths.State -Bytes $bytes -ExpectedSHA256 ABSENT -Validate { param($d) Assert-MaintenanceState $d }
        }
        $state = Get-SmartM365OrchestratorMaintenanceState $SharedDataFolderPath
        if ($gate.Length -eq 0) {
            $marker = [Text.Encoding]::UTF8.GetBytes('MaintenanceProtocol=1')
            $gate.Write($marker,0,$marker.Length); $gate.Flush($true)
        }
        return $state
    }
    finally { $gate.Dispose() }
}

function Get-SmartM365OrchestratorMaintenanceReadiness {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)]$ClusterDocument, [Parameter(Mandatory)]$State,
        [datetime]$Now = [datetime]::UtcNow)
    $stale = [double]$ClusterDocument.PeerHeartbeatStaleMinutes
    foreach ($server in @($ClusterDocument.ExpectedOrchestratorServers | Sort-Object -Unique)) {
        $status = 'Offline'; $version = ''; $running = 0
        try {
            $path = Join-Path (Join-Path $SharedDataFolderPath $server) 'Orchestrator-Heartbeat.json'
            $h = (Read-SmartM365JsonDocument -Path $path).Document
            $version = [string]$h.ScriptVersion; $running = @($h.RunningJobs).Count
            $timestamp = if ($h.Timestamp -is [datetime]) { $h.Timestamp.ToUniversalTime() }
                elseif ($h.Timestamp -is [datetimeoffset]) { $h.Timestamp.UtcDateTime }
                else { ([datetimeoffset]::Parse([string]$h.Timestamp, [Globalization.CultureInfo]::InvariantCulture)).UtcDateTime }
            $age = ($Now.ToUniversalTime() - $timestamp).TotalMinutes
            if ($age -ge -1 -and $age -le $stale) {
                if (-not $h.PSObject.Properties['MaintenanceProtocol'] -or
                    ($h.MaintenanceProtocol -isnot [int] -and $h.MaintenanceProtocol -isnot [long]) -or $h.MaintenanceProtocol -ne 1) { $status = 'Unsupported version' }
                elseif (-not $h.PSObject.Properties['MaintenanceHealthy'] -or $h.MaintenanceHealthy -isnot [bool] -or $h.MaintenanceHealthy -ne $true -or
                    -not $h.PSObject.Properties['MaintenanceEnabled'] -or $h.MaintenanceEnabled -isnot [bool] -or -not $h.PSObject.Properties['MaintenanceRevision'] -or
                    ($h.MaintenanceRevision -isnot [int] -and $h.MaintenanceRevision -isnot [long]) -or $h.MaintenanceRevision -lt 0) { $status = 'Control unavailable' }
                elseif ($h.Lifecycle -ne 'Running') { $status = 'Pending' }
                elseif ($h.MaintenanceRevision -eq $State.Revision -and $h.MaintenanceEnabled -eq $State.Enabled) { $status = 'Applied' }
                else { $status = 'Pending' }
            }
        }
        catch { $status = 'Offline' }
        [pscustomobject]@{ Server=[string]$server; Status=$status; Version=$version; Running=$running }
    }
}

function Set-SmartM365OrchestratorMaintenance {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)][bool]$Enabled, [Parameter(Mandatory)][long]$ExpectedRevision,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Reason)
    if ([string]::IsNullOrWhiteSpace($Reason) -or $Reason.Length -gt 1000) { throw 'Enter a maintenance reason (1 to 1000 characters).' }
    $paths = Get-SmartM365OrchestratorMaintenancePaths $SharedDataFolderPath
    $gate = Enter-SmartM365OrchestratorMaintenanceGate $SharedDataFolderPath
    try {
        # Initialization belongs to resident orchestrators, never to a GUI write.
        $current = Get-SmartM365OrchestratorMaintenanceState $SharedDataFolderPath
        if ($current.Revision -ne $ExpectedRevision) { throw 'Maintenance changed in another session. Refresh before changing it.' }
        if ($current.Enabled -eq $Enabled) { return $current }
        if ($Enabled) {
            $cluster = (Read-SmartM365JsonDocument -Path (Join-Path $SharedDataFolderPath 'Config/Orchestrator-Cluster.json')).Document
            $servers = @(Get-SmartM365OrchestratorMaintenanceReadiness $SharedDataFolderPath $cluster $current)
            if ($servers.Count -eq 0 -or @($servers | Where-Object Status -ne 'Applied').Count -ne 0) {
                throw 'Every expected server must be online, maintenance-compatible and acknowledge the current revision before activation.'
            }
        }
        $now = [datetime]::UtcNow.ToString('o')
        $actor = [Environment]::UserName
        try { $actor = [Security.Principal.WindowsIdentity]::GetCurrent().Name } catch { }
        $desired = [pscustomobject][ordered]@{
            SchemaVersion=1; Revision=([long]$current.Revision + 1); Enabled=$Enabled
            ChangedAtUtc=$now; ChangedBy=$actor; ChangedFromServer=[Environment]::MachineName; Reason=$Reason.Trim()
            ResumeAfterUtc=if ($Enabled) { $current.ResumeAfterUtc } else { $now }
        }
        Assert-MaintenanceState $desired
        $audit = [pscustomobject][ordered]@{
            ChangedAtUtc=$now; Revision=$desired.Revision; Enabled=$Enabled; ChangedBy=$actor
            ChangedFromServer=$desired.ChangedFromServer; Reason=$desired.Reason; Outcome='Requested'
        }
        $null = New-Item -ItemType Directory -Path (Split-Path $paths.Audit -Parent) -Force -ErrorAction Stop
        $audit | Export-Csv -LiteralPath $paths.Audit -Append -NoTypeInformation -Encoding utf8 -ErrorAction Stop
        $hash = (Get-FileHash -LiteralPath $paths.State -Algorithm SHA256 -ErrorAction Stop).Hash
        $bytes = [Text.UTF8Encoding]::new($false).GetBytes(($desired | ConvertTo-Json -Depth 10))
        $null = Write-SmartM365JsonBytesAtomically -Path $paths.State -Bytes $bytes -ExpectedSHA256 $hash -Validate { param($d) Assert-MaintenanceState $d }
        $audit.Outcome = 'Published'
        try { $audit | Export-Csv -LiteralPath $paths.Audit -Append -NoTypeInformation -Encoding utf8 -ErrorAction Stop }
        catch { throw "Maintenance revision $($desired.Revision) was published, but its final audit could not be written. Refresh to see the actual state. $($_.Exception.Message)" }
        return $desired
    }
    finally { $gate.Dispose() }
}

function Assert-MaintenanceMailCheckpoint {
    param($Document)
    if (-not $Document.PSObject.Properties['SchemaVersion'] -or $Document.SchemaVersion -ne 1 -or
        -not $Document.PSObject.Properties['LastSentRevision'] -or ($Document.LastSentRevision -isnot [int] -and $Document.LastSentRevision -isnot [long]) -or $Document.LastSentRevision -lt 0 -or
        -not $Document.PSObject.Properties['LastAttemptRevision'] -or ($Document.LastAttemptRevision -isnot [int] -and $Document.LastAttemptRevision -isnot [long]) -or $Document.LastAttemptRevision -lt 0 -or
        -not $Document.PSObject.Properties['LastAttemptUtc']) { throw 'Maintenance mail checkpoint is invalid.' }
    if ($Document.LastAttemptUtc -is [datetime]) { $Document.LastAttemptUtc=$Document.LastAttemptUtc.ToUniversalTime().ToString('o') }
    if ($Document.LastAttemptUtc -is [datetimeoffset]) { $Document.LastAttemptUtc=$Document.LastAttemptUtc.ToUniversalTime().ToString('o') }
    if ($Document.LastAttemptUtc -isnot [string]) { throw 'Maintenance mail attempt timestamp is invalid.' }
    if ($Document.LastAttemptUtc) { $null=[datetimeoffset]::Parse($Document.LastAttemptUtc,[Globalization.CultureInfo]::InvariantCulture) }
}

function Invoke-SmartM365OrchestratorMaintenanceNotification {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)][scriptblock]$SendAction, [switch]$Initialize,
        [ValidateRange(0,3600)][int]$RetrySeconds=60, [datetime]$Now=[datetime]::UtcNow)
    $paths=Get-SmartM365OrchestratorMaintenancePaths $SharedDataFolderPath
    # Dedicated lock: mail transport never holds the scheduling/GUI maintenance gate.
    try { $gate=[IO.File]::Open($paths.MailGate,'OpenOrCreate','ReadWrite','None') }
    catch [IO.IOException] { return [pscustomobject]@{Status='Busy';Sent=0} }
    try {
        $state=Get-SmartM365OrchestratorMaintenanceState $SharedDataFolderPath
        if (-not (Test-Path -LiteralPath $paths.MailCheckpoint)) {
            if (-not $Initialize) { throw 'Maintenance mail checkpoint is not initialized.' }
            # First deployment starts here, without replaying historical maintenance mail.
            $checkpoint=[pscustomobject]@{SchemaVersion=1;LastSentRevision=[long]$state.Revision;LastAttemptRevision=0;LastAttemptUtc=''}
            $null=Write-SmartM365JsonBytesAtomically -Path $paths.MailCheckpoint -Bytes ([Text.UTF8Encoding]::new($false).GetBytes(($checkpoint|ConvertTo-Json))) -ExpectedSHA256 ABSENT -Validate {param($d) Assert-MaintenanceMailCheckpoint $d}
            return [pscustomobject]@{Status='Initialized';Sent=0}
        }
        $checkpoint=(Read-SmartM365JsonDocument -Path $paths.MailCheckpoint -Validate {param($d) Assert-MaintenanceMailCheckpoint $d}).Document
        if ($checkpoint.LastSentRevision -gt $state.Revision) { throw 'Maintenance mail checkpoint is ahead of the control revision.' }
        $pending=@()
        if (Test-Path -LiteralPath $paths.Audit) {
            $pending=@(Import-Csv -LiteralPath $paths.Audit -ErrorAction Stop | Where-Object {
                $_.Outcome -eq 'Published' -and [long]$_.Revision -gt $checkpoint.LastSentRevision -and [long]$_.Revision -le $state.Revision
            } | Sort-Object { [long]$_.Revision })
        }
        # A published control remains authoritative even if its final audit append failed.
        if ($state.Revision -gt $checkpoint.LastSentRevision -and @($pending|Where-Object {[long]$_.Revision -eq $state.Revision}).Count -eq 0) { $pending+= $state }
        $sent=0
        foreach ($transition in $pending) {
            $revision=[long]$transition.Revision
            if ($revision -ne $checkpoint.LastSentRevision+1) { throw 'Maintenance transition evidence has a revision gap or duplicate; notification stopped.' }
            if ($checkpoint.LastAttemptRevision -eq $revision -and $checkpoint.LastAttemptUtc -and
                ($Now.ToUniversalTime()-([datetimeoffset]::Parse($checkpoint.LastAttemptUtc)).UtcDateTime).TotalSeconds -lt $RetrySeconds) {
                return [pscustomobject]@{Status='RetryPending';Sent=$sent}
            }
            $enabledText=[string]$transition.Enabled
            if ($enabledText -notin @('True','False')) { throw 'Maintenance transition enabled flag is invalid.' }
            $notificationEvent=[pscustomobject]@{Revision=$revision;Enabled=($enabledText -eq 'True');ChangedAtUtc=[string]$transition.ChangedAtUtc;ChangedBy=[string]$transition.ChangedBy;ChangedFromServer=[string]$transition.ChangedFromServer;Reason=[string]$transition.Reason}
            $checkpoint.LastAttemptRevision=$revision; $checkpoint.LastAttemptUtc=$Now.ToUniversalTime().ToString('o')
            $null=Write-SmartM365JsonBytesAtomically -Path $paths.MailCheckpoint -Bytes ([Text.UTF8Encoding]::new($false).GetBytes(($checkpoint|ConvertTo-Json))) -Validate {param($d) Assert-MaintenanceMailCheckpoint $d}
            $deliveryResult=@(& $SendAction $notificationEvent)
            if ($deliveryResult.Count -ne 1 -or $deliveryResult[0] -isnot [bool]) { throw 'Maintenance mail callback must return exactly one Boolean delivery result.' }
            $delivered=$deliveryResult[0]
            if (-not $delivered) { return [pscustomobject]@{Status='RetryPending';Sent=$sent} }
            $checkpoint.LastSentRevision=$revision
            $null=Write-SmartM365JsonBytesAtomically -Path $paths.MailCheckpoint -Bytes ([Text.UTF8Encoding]::new($false).GetBytes(($checkpoint|ConvertTo-Json))) -Validate {param($d) Assert-MaintenanceMailCheckpoint $d}
            $sent++
        }
        return [pscustomobject]@{Status='Completed';Sent=$sent}
    }
    finally { $gate.Dispose() }
}

function Test-SmartM365OrchestratorMaintenanceLaunch {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$State, [Parameter(Mandatory)][ValidateSet('due','retry','pipeline','forced','pipeline-retry')][string]$Origin,
        [Parameter(Mandatory)][datetime]$Occurrence)
    Assert-MaintenanceState $State
    if ($Origin -in @('pipeline','forced','pipeline-retry')) { return $true }
    if ($State.Enabled) { return $false }
    if ($State.ResumeAfterUtc -and $Occurrence.ToUniversalTime() -le ([datetimeoffset]::Parse($State.ResumeAfterUtc)).UtcDateTime) { return $false }
    return $true
}

Export-ModuleMember -Function Get-SmartM365OrchestratorMaintenancePaths, Get-SmartM365OrchestratorMaintenanceState,
    Enter-SmartM365OrchestratorMaintenanceGate, Initialize-SmartM365OrchestratorMaintenance,
    Get-SmartM365OrchestratorMaintenanceReadiness, Set-SmartM365OrchestratorMaintenance,
    Test-SmartM365OrchestratorMaintenanceLaunch, Invoke-SmartM365OrchestratorMaintenanceNotification

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCngBGON9Yiy6Vm
# icuSd9y1+PIJr2MJudmyf/1lmVchsqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINwePZn/46OCMXKNFuw2SZXZThuTsUmK8gQ+FB1e81hTMA0GCSqG
# SIb3DQEBAQUABIIBgFgSZ+NuDbKhAcqcEYgCiLK7lZ7irSleor+hsPgKhkh+8wmz
# pcGOnYNpRVtm1611hmiJKHqD/71ZdvlzvW+etGxFgPKaQDkLj2ZLi6/aa0llMTiv
# momXulvecCybcBeDoY2z0NcgRYQ9BzPxjVHGeCvT0YhkJAyJoNo5HDdOfSbi0g0u
# s70EV3a0S4Ggwbc9sQYVA3OBCf3F6DBB0E7mzM02UbI/fFZE6TaIEMEiYZKR4+qh
# 0publeISKBHvnTXPs7RpwQK0/rDDCWpqDgJjcgWMsm/E0l0eb/qCY1clTi0Wr+mg
# Q8F1yikF3kyU3eEKOyeoEDJhy4bI8hZjqvfPoPJR8js0q4Lk/9sZocKYoGL9lZVF
# tIL4Ue2rNRnUoR2oOG1QkWyJMG93BsqiR+zXN3eujQ074W0sfG4E8wpNzhKSqCa5
# +0u9byjofy4ZkiiGn7oGAelCpMkhdXB3dqe1WF/o9lxgFvhunzNijXQqnT0x3KX/
# uTcKFO2lPAKwkJiLoaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDQxODM3
# NDRaMC8GCSqGSIb3DQEJBDEiBCDM+VvS6RxRA8lh/6METBqFGeexJql2vwPoDPNk
# FFxCyDANBgkqhkiG9w0BAQEFAASCAgCDKMHKPvmmMsFAe5nua3uxnH0AekBQvLf7
# 40BleAYoAb/+ecYY9pXLA5s28UxAbdeCNkdISVUURAUhlJeFGhOcNGIapuXvhb97
# ig/LDma2T86OaIfVjuIwHr8mxo4mvzLenSbvlnQ8NZN3DzuBUYh5IDtURiIqOKlq
# QHFJmrpjh+7kB8ln23SiTvLFUc/H1IT2VJi89EooygCNRPCKnFlXURp6lH20CziE
# Bf5ep/O6j39w6xtY/ocZQlp+RqR/BOjkTdpOkPZmMExoy5rULnREGjv8ulUJVhmU
# Dmuh8ta0n9NsH46EM0YsSfL3AsQ10KA63iurAcE8a9d/UVDGKmbgm4FSwjmIPMnU
# ZWouP4MY5p3uelnF9qedWjPTDFvts6g+WLbaxWjvz2Xo71auXJLFG+heDWCrIE5C
# b4Ldvt6DvxTXZ1sJDGVxVHgm2wpz3gk1TWJ1oslTeVqEm3iGV7Cvxx+qB/C+pPM7
# PBLCdTCD0bk1JujbugLvrxIbgFzfUC5e2rD1Ns6d/Kx8y6iSLZ+sJgyLvtziN+Cx
# u3bgmYkIHQLYNCHPQ7cBiMb8PWO64iVcnE/EiyCICsP5sJSyXSYZz4q6wrXgpMCA
# UOtlovUIIEPrAJKfCwH2Pb64Ai6gy3fsV1EXH8jy9EEV/7kf2D0ugoWjBiCdOIG5
# z+A3wu6idw==
# SIG # End signature block
