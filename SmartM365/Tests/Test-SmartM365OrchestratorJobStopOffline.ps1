# Offline job-stop contract test. No inventory collector or production share is used.
$ErrorActionPreference = 'Stop'
$folder = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.Management.psm1') -Force
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.JobStop.psm1') -Force
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.Pipeline.psm1') -Force
$source = Join-Path $folder 'SmartM365-Inventory-Orchestrator.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors.Message -join '; ') }
$names = @('Update-OrchestratorJobStopRequests', 'Complete-JobRun')
$definitions = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -in $names }, $true))
if ($definitions.Count -ne $names.Count) { throw 'Required runtime functions were not found.' }
$testModule = New-Module -Name OrchestratorJobStopOffline -ScriptBlock ([scriptblock]::Create((@($definitions | ForEach-Object { $_.Extent.Text }) -join "`n") + @'
function Get-JobState { param($JobName) return $script:State.Jobs[$JobName] }
function ConvertTo-StateTime { param([datetime]$Value) return $Value.ToString('o') }
function Save-OrchestratorState { $script:Saved = $true }
function Test-OrchestratorStatePersistenceReady { return [bool]$script:Saved }
function Write-OrchestratorLog { param($Message, $Level) }
function Set-OrchestratorPipelineJobStatus { throw 'Unexpected pipeline status call.' }
function Add-JobRunCsvRow { param($JobName,$ScheduledTime,$StartTime,$EndTime,$DurationSec,$ExitCode,$Status,$RetryCount,$LogPath) $script:CsvStatus = $Status }
function Get-JobRunsCsvPath { return 'synthetic.csv' }
function Invoke-OrchestratorSharePointUpload { param($LocalFilePath,$Reason,[switch]$Force) return $false }
function Send-JobResultEmail { throw 'Unexpected mail send.' }
Export-ModuleMember -Function Update-OrchestratorJobStopRequests, Complete-JobRun
'@))
Import-Module $testModule -Force

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-jobstop-offline-' + [guid]::NewGuid().ToString('N'))
try {
    $server = 'TEST-SERVER'
    $jobName = 'Synthetic-Inventory'
    $pidValue = 8123
    $started = (Get-Date).AddMinutes(-10)
    $startedText = $started.ToString('o')
    $heartbeatPath = Join-Path (Join-Path $tempRoot $server) 'Orchestrator-Heartbeat.json'
    $heartbeat = [pscustomobject]@{
        Timestamp = (Get-Date).ToString('o'); Lifecycle = 'Running'; Tenant = 'test'; JobStopProtocol = 1
        RunningJobs = @([pscustomobject]@{ Name = $jobName; Pid = $pidValue; StartTime = $startedText })
    }
    Write-SmartM365OrchestratorJsonAtomically -Path $heartbeatPath -Document $heartbeat
    $request = Request-SmartM365OrchestratorJobStop -SharedDataFolderPath $tempRoot -Server $server -JobName $jobName -ProcessId $pidValue -StartTime $startedText -Reason 'Operator test stop' -Tenant test
    $savedRequest = Read-SmartM365OrchestratorJson -Path $request.Path
    if ($savedRequest.Status -ne 'Requested' -or $savedRequest.Reason -ne 'Operator test stop') { throw 'Stop request was not durably recorded.' }
    $rejected = $false
    try { Request-SmartM365OrchestratorJobStop -SharedDataFolderPath $tempRoot -Server $server -JobName $jobName -ProcessId 9999 -StartTime $startedText -Reason 'Wrong process' -Tenant test | Out-Null }
    catch { $rejected = $true }
    if (-not $rejected) { throw 'A changed PID was accepted.' }
    $rejected = $false
    try { Request-SmartM365OrchestratorJobStop -SharedDataFolderPath $tempRoot -Server $server -JobName $jobName -ProcessId $pidValue -StartTime $started.AddSeconds(-10).ToString('o') -Reason 'Wrong start' -Tenant test | Out-Null }
    catch { $rejected = $true }
    if (-not $rejected) { throw 'A changed process start time was accepted.' }
    $staleRoot = Join-Path $tempRoot 'stale'
    $staleHeartbeatPath = Join-Path (Join-Path $staleRoot $server) 'Orchestrator-Heartbeat.json'
    $staleHeartbeat = [pscustomobject]@{
        Timestamp = (Get-Date).AddMinutes(-10).ToString('o'); Lifecycle = 'Running'; Tenant = 'test'; JobStopProtocol = 1
        RunningJobs = @([pscustomobject]@{ Name = $jobName; Pid = $pidValue; StartTime = $startedText })
    }
    Write-SmartM365OrchestratorJsonAtomically -Path $staleHeartbeatPath -Document $staleHeartbeat
    $rejected = $false
    try { Request-SmartM365OrchestratorJobStop -SharedDataFolderPath $staleRoot -Server $server -JobName $jobName -ProcessId $pidValue -StartTime $startedText -Reason 'Stale owner' -Tenant test | Out-Null }
    catch { $rejected = $true }
    if (-not $rejected) { throw 'A stale owner heartbeat was accepted.' }

    & $testModule {
        param($root,$serverName,$name,$id,$start)
        $env:COMPUTERNAME = $serverName
        $script:Tenant = 'test'
        $script:Settings = @{ SharedDataFolderPath = $root; JobMailMode = 'Never' }
        $script:State = @{ Jobs = @{ $name = @{ Running = @{ Pid = $id; StartTime = $start }; LastStatus = 'Running'; PendingRetry = $null } } }
        $script:Manifest = @{ JobsByName = @{ $name = [pscustomobject]@{ MaxRetries = 3 } } }
        $script:RunningJobs = @{ $name = @{ Process = [pscustomobject]@{ Id = $id }; StartTime = [datetime]::Parse($start); Occurrence = Get-Date; LogPath = ''; Attempt = 0; ProcessName = 'pwsh' } }
        $script:Saved = $false
    } $tempRoot $server $jobName $pidValue $startedText
    Update-OrchestratorJobStopRequests
    $ack = Read-SmartM365OrchestratorJson -Path $request.Path
    if ($ack.Status -ne 'Stopping') { throw "Owner did not accept the stop request: $($ack.Status)" }
    & $testModule {
        param($name,$path)
        if (-not $script:Saved -or -not $script:RunningJobs[$name].OperatorStopRequested -or
            $script:State.Jobs[$name].Running.OperatorStopRequestPath -ne $path) { throw 'Stop intent was not persisted before acknowledgement.' }
    } $jobName $request.Path
    & $testModule {
        param($name)
        $run = $script:RunningJobs[$name]
        Complete-JobRun -JobName $name -RunInfo $run -StatusHint Cancelled -ExitCode $null -EndTime (Get-Date) -ErrorText 'Stopped by operator for offline test.'
        if ($script:State.Jobs[$name].LastStatus -ne 'Cancelled' -or $script:State.Jobs[$name].PendingRetry -or $script:CsvStatus -ne 'Cancelled') { throw 'Cancelled run was retried or recorded incorrectly.' }
    } $jobName
    $final = Read-SmartM365OrchestratorJson -Path $request.Path
    if ($final.Status -ne 'Stopped') { throw 'Stop result was not recorded.' }
    $selection = [pscustomobject]@{ SelectedJobs = @([pscustomobject]@{ Name = $jobName }) }
    $batch = New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $tempRoot -Tenant test -Pipeline Jobs -Selection $selection -ManifestHash synthetic
    $null = Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $tempRoot -BatchId $batch.BatchId -JobName $jobName -Status Running -OwnerServer $server
    $blocked = Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $tempRoot -BatchId $batch.BatchId -JobName $jobName -Status Cancelled -OwnerServer $server
    if ($blocked.Status -ne 'Running') { throw 'Ordinary batch cancellation changed a running job.' }
    $accepted = Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $tempRoot -BatchId $batch.BatchId -JobName $jobName -Status Cancelled -OwnerServer $server -AllowRunningCancellation
    $summary = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $tempRoot -BatchId $batch.BatchId
    if ($accepted.Status -ne 'Cancelled' -or $summary.OverallStatus -ne 'Failed') { throw 'Operator stop did not finalize the running pipeline job.' }
    'PASS: exact-run request, stale PID/start/owner rejection, owner acknowledgement, Cancelled result without retry, and running pipeline finalization.'
}
finally {
    $resolved = [IO.Path]::GetFullPath($tempRoot)
    $tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    if ($resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolved) -like 'SmartM365-jobstop-offline-*' -and
        (Test-Path -LiteralPath $resolved)) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDUbCTZYNU3oezb
# dalR353Bqzf0E1ZgNkY1n7+svRCtA6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDGipI9n7ocqVcrPe6skfJx
# ZWjXc8YDNLt96139j4GLUjANBgkqhkiG9w0BAQEFAASCAYBLDJqri5cQGAQmhmR/
# TVZLTOjP3yF/xTJsknTwbuUk5+qFGq8HjInKE95y4D7T5r/7No9DO3YZeYXjQaN4
# B4xeetJI3jULflqAonuQN9eBdFe47wfKpHkuISywH9Lx2SrragN9TVnwvFWp9WFD
# 7Hmq2R6rauuhUqDNvqbVcXH8L34K6d5TS5QGPXtY1/F2VqbxJLiCl8ZdQUqfIeZ1
# St7Ta5FAUFIZeszRAf3FNLeS2IUiv5lGm3zLO5IG21cqiEXuWazJ0R1LYgDBmRb5
# c8bMqoWXi4ZlxOfjroI9UrV58TVxaxW/gOhTy0RuqXVC4iVmvTRsexyG+8Z51Uwq
# Qr3UHT3am49u7PmS3NsTkXbUUCARXWzzxL5g2wtjzs/x9VgkjPeQ2Wc9dOgl7uMR
# NJ9WugIX93jM4VfOQFiB4xZ9OrTEf550KbEiL+fxZdKDoXXZSN6XvqXTg6VcmsJF
# R+PriMZyxp1/QUOwV4M5UBwfbUHBQnd0k6lUysKOQguaMLA=
# SIG # End signature block
