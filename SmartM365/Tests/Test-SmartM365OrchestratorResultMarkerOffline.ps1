#Requires -Version 7.0
<#
.SYNOPSIS
Runs a synthetic collector that calls exit and checks its durable completion result.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$source = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'Orchestrator parse failed.' }
$definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Start-InventoryJob' }, $true)
if (-not $definition) { throw 'Job launcher is missing.' }
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ResultMarker-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $root
$collectorPath = Join-Path $root 'SyntheticCollector.ps1'
[IO.File]::WriteAllText($collectorPath, 'param([string]$Tenant) exit 3', [Text.UTF8Encoding]::new($false))
$module = New-Module -ScriptBlock ([scriptblock]::Create($definition.Extent.Text))
try {
    & $module {
        param($folder)
        $script:Settings = @{ JobLogFolderPath=$folder; SharedDataFolderPath=$folder }
        $script:RunningJobs = @{}
        $script:State = @{ Jobs=@{ Synthetic=@{ Running=$null } } }
        $script:JobState = $script:State.Jobs.Synthetic
        $script:ScriptName = 'SyntheticOrchestrator'
        $script:Tenant = 'test'
        $script:Connect = $false
        function script:Get-JobState { param($JobName) $script:JobState }
        function script:Resolve-OrchestratorJobPath { param($Path) $Path }
        function script:Invoke-OrchestratorAuthenticodeValidation { param($Job,$ScriptFullPath) [pscustomobject]@{ Allowed=$true } }
        function script:Get-JobEngine { param($Job) [pscustomobject]@{Path=(Get-Command pwsh).Source;ProcessName='pwsh'} }
        function script:ConvertTo-StateTime { param($Value) ([datetime]$Value).ToString('o') }
        function script:Save-OrchestratorState {}
        function script:Write-OrchestratorLog { param($Message,$Level) }
        function script:Complete-JobRun { throw 'Synthetic launch failed before a child was started.' }
    } $root
    $job = [pscustomobject]@{ Name='Synthetic'; ScriptPath=$collectorPath; Arguments=''; PowerShellEdition='Core'; TimeoutMinutes=2 }
    & $module { param($job) Start-InventoryJob -Job $job -Occurrence (Get-Date) } $job
    $record = & $module { $script:State.Jobs.Synthetic.Running }
    if (-not $record.RunId -or -not $record.ResultPath) { throw 'The run identifier/result path was not persisted.' }
    $process = & $module { $script:RunningJobs.Synthetic.Process }
    if (-not $process.WaitForExit(15000)) { throw 'Synthetic collector did not exit.' }
    if ($process.ExitCode -ne 3) { throw "Supervisor returned exit $($process.ExitCode), expected 3." }
    if (-not (Test-Path -LiteralPath $record.ResultPath)) { throw 'Collector exit bypassed the result marker.' }
    $result = Get-Content -LiteralPath $record.ResultPath -Raw | ConvertFrom-Json
    if ($result.RunId -cne $record.RunId -or $result.JobName -ne 'Synthetic' -or $result.ProcessId -ne $record.Pid -or $result.ExitCode -ne 3) { throw 'Result marker does not identify the completed run.' }
    'PASS: a collector exit is captured by the supervisor and persisted for restart recovery.'
}
finally {
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCCj3Xzq4zSqBOp
# jYG32juWexEksxEjE9IDdi4lZepj7qCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAjImYIGgwK7HayHt03Dmg+
# axi7y2OnKrhfnUTEdGRvLzANBgkqhkiG9w0BAQEFAASCAYA3sPCx4CMRDb1F46Vz
# tRauO9bxM8vh2OIds4PbhLy2UDyj9iHRwWspvQLUovpaFnPIJqw5yTy+Qd/NoImK
# SYWyHQ1B+AgouDvULn2Inf+vyCFxS5e+IMJbvyIfw2rTj7LO4YmVonfGRpp836bU
# IpcNHiZ+BdbHUihVw5dwbbUaj4zMepUsiCEtoSOtpNvvJ+jvyqTFfaK9uN0zIksE
# rrPaFDjDWJAySee2oeJUDj+8vhiVkmnDuEN1kmhZehkUxbuwlhSDqM+J2VQPleW5
# ztDjAYw9KBy8/0OEwd2ar3h27la3niMa6derdfLdUMPDvjQ9xK91LS+whKyqYD5u
# SkW0P466kOWtREghS4pZDhVJrJA2rPwdsKO9jbXwSfJdtf31cJdD20ddWFfzV3cG
# a4jZpeSc817DDofLN9oQyXhxoIZxI+5QiFIw5mImFkTslYu/cz2KB9Z9G/eBoAH9
# Utt2qoo5Q5SDw5+/C18ZivJapsMWwILmV2lLb1FY0XkhErU=
# SIG # End signature block
