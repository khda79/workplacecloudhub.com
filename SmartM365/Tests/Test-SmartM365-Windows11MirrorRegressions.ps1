#requires -Version 7.0
<#
.SYNOPSIS
Offline regression tests for Windows 11 readiness issues and the orchestrator SharePoint mirror.
.VERSION
1.0.1
#>
$ErrorActionPreference='Stop'
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-W11Mirror-'+[guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory $root
$script:passed=0
function Check($value,$message){if(-not $value){throw $message};$script:passed++}
function LoadFunctions($path,$names){
 $tokens=$null;$errors=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
 if($errors.Count){throw 'Parse failure'}
 foreach($name in $names){$node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true);$node.Extent.Text}
}
$collector=Join-Path $PSScriptRoot '../SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-Intune-Windows11-Readiness-Issues-Inventory.ps1'
. ([scriptblock]::Create((LoadFunctions $collector @('ExportIssues')) -join "`n"))
$OutputFolder=Join-Path $root 'output';$LatestFolder=Join-Path $root 'latest'
function Assert-SmartM365CsvDataCompleteness { param($Data,$TimestampedPath,$LatestPath) }
function Write-SmartM365CsvAtomically {param($Data,$Path,$Encoding) $Data | Export-Csv -LiteralPath $Path -NoTypeInformation }
function CopyCsv {param($s,$d) Copy-Item -LiteralPath $s -Destination $d }
$rows=[Collections.Generic.List[object]]::new()
foreach($id in @('a','b')){$rows.Add([pscustomobject]@{IssueCode='ISSUE-TEST';Area='OS';ObjectGUID_Norm=$id;Potential_Issue='First condition, second condition';IssueCategory='2.High';PriorityScore=2;IsBlocking=$false})}
$files=@(ExportIssues $rows.ToArray())
$summary=@(Import-Csv (Join-Path $OutputFolder 'Intune_Windows11_Readiness_Issues_Summary.csv'))
Check ($files.Count -eq 4) 'Missing current detail or summary publication'
Check (-not (Test-Path -LiteralPath (Join-Path $OutputFolder 'Archive'))) 'Daily archive was unexpectedly created'
Check ($summary.Count -eq 1 -and $summary[0].Count -eq 2) 'Grouping count lost'
Check ($summary[0].Potential_Issue -eq 'First condition, second condition') 'Comma label corrupted'
Check ($summary[0].PriorityScore -eq 2 -and $summary[0].IsBlocking -eq 'False') 'Typed summary fields corrupted'
Check (@(Import-Csv (Join-Path $LatestFolder 'Intune_Windows11_Readiness_Issues.csv')).Count -eq 2) 'Detail rows lost'
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Force
$orch=Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.ps1'
. ([scriptblock]::Create((LoadFunctions $orch @('Get-OrchestratorSharePointMirrorRelativePath','Test-OrchestratorSharePointMirrorFile','Get-OrchestratorSharePointMirrorSnapshot')) -join "`n"))
function Write-OrchestratorLog {param($Message,$Level)}
function Update-OrchestratorHeartbeatDuringLongOperation {}
$shared=Join-Path $root 'DATA-ALL/Orchestrator';$null=New-Item -ItemType Directory (Join-Path $shared 'Election/Concurrency'),(Join-Path $shared 'Config') -Force
$lease=Join-Path $shared 'Election/Concurrency/Test.json.txt';$config=Join-Path $shared 'Config/Settings.json.txt'
[IO.File]::WriteAllText($lease,'{}');[IO.File]::WriteAllText($config,'{}')
$script:vanish=$lease
function Read-SmartM365JsonDocument {param($Path)
 if($Path -eq $script:vanish){Remove-Item -LiteralPath $Path;throw "JSON file missing: $Path"}
 SmartM365.JsonTransport\Read-SmartM365JsonDocument $Path
}
$snapshot=Get-OrchestratorSharePointMirrorSnapshot $shared
Check ($snapshot.Files.Count -eq 1) 'Persistent file lost after lease disappeared'
Check ($snapshot.DeferredPaths.Count -eq 2) 'Both transport names must be deferred'
$next=Get-OrchestratorSharePointMirrorSnapshot $shared
Check ($next.DeferredPaths.Count -eq 0) 'Deferred lease never reconciled'
$script:vanish=$config;$rejected=$false
try{Get-OrchestratorSharePointMirrorSnapshot $shared | Out-Null}catch{$rejected=$true}
Check $rejected 'Missing persistent file silently ignored'
$script:vanish='';[IO.File]::WriteAllText($lease,'{');$rejected=$false
try{Get-OrchestratorSharePointMirrorSnapshot $shared | Out-Null}catch{$rejected=$true}
Check $rejected 'Invalid lease silently ignored'
[pscustomobject]@{Passed=$script:passed;FixtureRoot=$root;Evidence='Synthetic only; no collector startup or remote access'}
# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBpCgay5lhvk1cJ
# GFtPNqXF8AzRRAnvf36AhOuhdlwODKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCD/zyli4DXcxKKDlTQMzf7B
# x2qJLY4Cjc3Iw1RW588utzANBgkqhkiG9w0BAQEFAASCAYAWAA98DeH4FmSOFC3f
# /YOYxm4HGdKgFCoHp8fwKNAYgRSoW8vww0RQgO5X6R22Top8V02s9qVubKd8gn5z
# NXUQO6BNdKrMYyRNUw38rLcfigwVW6YKXgzX4OekxkBEl4LmJec+cRCtj2PJ2SJv
# 3bx7FhaBCBjt+TS7CjhVauabXvii4M8f7m0v5hbBYtxef6BNLKfcxX6ZiwaR5YeC
# sDYilHAO7XSiJEDflvmSlo3V3su4UqBCie9hw4D+w9mqLlNdz31Naa1Q9bVQP74C
# IfCcaNSj/lBrft9W3vlsxE2P2L1w79uSU01Seov0DV5ylUwX5MkS31kNieOunSO8
# geKirpfYrZjuQya+nWxb4ONyV5ga7L0tV6Un3a/LaWcLJ5LG2f2HUNu7QOHOVQpj
# SH2NrN1/lkjuGo1jrF8iQJr7evgCnY1H0Fb6XthNEHVdrc0EtejsEc8C+BEBRvvX
# af03kfrgj+hMPwhm/2yznFAItbd3H723til+YMrDE/OnVxk=
# SIG # End signature block
