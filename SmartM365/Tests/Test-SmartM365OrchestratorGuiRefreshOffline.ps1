#Requires -Version 7.0
<#
.SYNOPSIS
Checks that GUI shared-data refresh runs in a worker and propagates read failures.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$folder = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.GuiWorker.psm1') -Force
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $folder 'SmartM365-Inventory-Orchestrator-GUI.ps1'), [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'GUI parse failed.' }
$definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-GuiRefreshWork' }, $true)
if (-not $definition) { throw 'Background refresh work is missing.' }
. ([scriptblock]::Create($definition.Extent.Text))
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-GuiRefresh-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $root
$managementPath = Join-Path $root 'Management.psm1'
$insightsPath = Join-Path $root 'Insights.psm1'
@'
function Get-SmartM365OrchestratorConfigurationSnapshot { param($SharedDataFolderPath) [pscustomobject]@{Jobs=[pscustomobject]@{Jobs=@()};Cluster=[pscustomobject]@{ExpectedOrchestratorServers=@('WorkerA')}} }
function Get-SmartM365OrchestratorServerStatus { param($SharedDataFolderPath,$ClusterDocument) [pscustomobject]@{Server='WorkerA';AssignedJobNames='JobA'} }
function Get-SmartM365OrchestratorHistory { param($SharedDataFolderPath,$From,$To,$Server,$JobName,$Status) [pscustomobject]@{JobName='JobA';Status='Skipped'} }
function Get-SmartM365OrchestratorConfigurationVersions { param($SharedDataFolderPath) @() }
function Read-SmartM365OrchestratorJson { param($Path) [pscustomobject]@{ExpectedOrchestratorServers=@('WorkerA')} }
Export-ModuleMember -Function *
'@ | Set-Content -LiteralPath $managementPath
@'
function Get-SmartM365OrchestratorRecentRuns { param($SharedDataFolderPath,$Days) @() }
function Get-SmartM365OrchestratorOperations {
    param($SharedDataFolderPath,$ClusterDocument,$MailFolderPath,$MailHours)
    Start-Sleep -Milliseconds 900
    if (Test-Path (Join-Path $SharedDataFolderPath 'fail.flag')) { throw 'Synthetic shared read failure' }
    [pscustomobject]@{Servers=@();Running=@();Pending=@();Incidents=@();Mails=@()}
}
function Get-SmartM365OrchestratorRecentPipelineRuns { param($SharedDataFolderPath,$Count) @() }
Export-ModuleMember -Function *
'@ | Set-Content -LiteralPath $insightsPath
function Receive-CompletedWorker {
    param($Context)
    $deadline = [datetime]::UtcNow.AddSeconds(10)
    while (-not $Context.Worker.Handle.IsCompleted -and [datetime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 30 }
    if (-not $Context.Worker.Handle.IsCompleted) { throw 'Background refresh timed out.' }
    Receive-SmartM365OrchestratorGuiWorker -Context $Context
}
try {
    foreach ($mode in @('Full', 'Live', 'History')) {
        $context = [pscustomobject]@{ Worker=$null; Progress=[hashtable]::Synchronized(@{Message=''}) }
        $inputData = [pscustomobject]@{Mode=$mode;Root=$root;MailFolderPath='';ManagementModulePath=$managementPath;InsightsModulePath=$insightsPath;HistoryFrom=(Get-Date).Date;HistoryTo=(Get-Date);HistoryServer='';HistoryJob='';HistoryStatus='';OnlyFailures=$false}
        $watch = [Diagnostics.Stopwatch]::StartNew()
        if (-not (Start-SmartM365OrchestratorGuiWorker -Context $context -Work (Get-GuiRefreshWork) -InputObject $inputData)) { throw 'Worker was not started.' }
        $watch.Stop()
        if ($watch.ElapsedMilliseconds -ge 700) { throw "$mode blocked the caller for $($watch.ElapsedMilliseconds) ms." }
        $received = Receive-CompletedWorker $context
        if ($received.Error -or @($received.Output).Count -ne 1 -or $received.Output[0].Mode -ne $mode) { throw "$mode result was lost: $($received.Error)" }
    }
    $null = New-Item -ItemType File -Path (Join-Path $root 'fail.flag')
    $context = [pscustomobject]@{ Worker=$null; Progress=[hashtable]::Synchronized(@{Message=''}) }
    $inputData.Mode = 'Live'
    $null = Start-SmartM365OrchestratorGuiWorker -Context $context -Work (Get-GuiRefreshWork) -InputObject $inputData
    $received = Receive-CompletedWorker $context
    if ($received.Error -notlike '*Synthetic shared read failure*') { throw 'A failed shared read looked successful.' }
    'PASS: full, live, and history reads run asynchronously; shared read failure reaches the caller.'
}
finally {
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDgXMqLlylD6zlM
# nEEnN1Fcncnah/ICTf5Y9XtOKWBxs6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBHwuv7kkZtmlXyMvkH0NWP
# 9FrV/qzgWgolGlw8e0Q3bzANBgkqhkiG9w0BAQEFAASCAYBoXyrpUbkkSUUQxeuA
# d4Z3DlUnd9Y0X9cIix4hlVdHsZmp80Vt8GAis2giiB+GRDo/szdOegHIbULZUIZa
# 2nWJkU2dhI5RI+hPzx6Zuw9lPgWGx4TPc5cAKBfL85NICOsorap63SV8hFKlaUmm
# nFM1XnZsQngJbyK25jyLtiRmdVMaMJwM9IDVTIQ3TZrgHMdeFW25yuAAqePAFVEH
# rj3Cxz9uvrkRAYrhPJfKY+RDkAo5cFcbSIV5nRqWYE7S+pfTmWea433IrHDjdHNG
# 4Ad5Ry4kI0Uk1gFZRi5qDEQlkZw+3yONuLtp6CNXqitITqVFj3EFory/ma4lDDLI
# pxV88pl6LhshjbBQlzSzC53bQFGTwA7HX7UgptJLI23/dh4Fc/0xThb4tUpW8wkr
# 9T5ebewJvnAJfPeUpMUwKEZ71XW3FFU23eqciUb6/0p/KKiuTYyBrGy/JRCK+kv/
# A8Ur19iH2eUccd/zZWMAwckxJeky8zqKLSZQ8mtSjaJ36oQ=
# SIG # End signature block
