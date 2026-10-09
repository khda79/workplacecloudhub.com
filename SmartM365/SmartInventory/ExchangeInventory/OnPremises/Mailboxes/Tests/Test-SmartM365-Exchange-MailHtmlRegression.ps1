[CmdletBinding()]
param([switch]$RecordBaseline, [string]$ExchangeSourcePath)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$folder = Split-Path $PSScriptRoot -Parent
$scriptPath = if ($ExchangeSourcePath) { $ExchangeSourcePath } else { Join-Path $folder 'SmartM365-Exchange-Local-Mailboxes-Inventory.ps1' }
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
if (@($errors).Count -gt 0) { throw 'Exchange collector does not parse.' }
$functionAst = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'New-SmartM365ExchangeLocalMailboxReportEmailBody' }, $true))
if ($functionAst.Count -ne 1) { throw 'Exchange mail renderer was not found.' }
. ([scriptblock]::Create($functionAst[0].Extent.Text))

function ConvertTo-SmartM365EmailHtmlText { param($Value) return [System.Net.WebUtility]::HtmlEncode([string]$Value) }
$corePath = Join-Path $folder '..\..\..\..\Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psm1'
$coreTokens = $null
$coreErrors = $null
$coreAst = [Management.Automation.Language.Parser]::ParseFile($corePath, [ref]$coreTokens, [ref]$coreErrors)
$gridAst = @($coreAst.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'New-SmartM365EmailKpiGridHtml' }, $true))
if ($gridAst.Count -ne 1) { throw 'Shared Exchange card grid helper was not found.' }
. ([scriptblock]::Create($gridAst[0].Extent.Text))
$bodyAst = @($coreAst.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'New-SmartM365EmailBody' }, $true))
if ($bodyAst.Count -ne 1) { throw 'Shared Exchange email frame was not found.' }
. ([scriptblock]::Create($bodyAst[0].Extent.Text.Replace('function New-SmartM365EmailBody {','function New-SmartM365EmailBodyFixture {')))
function Get-SmartM365MailboxReportCsvRowCount { param($Path) return 2 }
function New-SmartM365LocalMailboxIssueEmailSection { param($Issues, $IssuesCsvPath) return $null }
function New-SmartM365EmailBody { param($Title, $Category, $Severity, $Tenant, $Message, $Sections, $Footer) return New-SmartM365EmailBodyFixture -Title $Title -Category $Category -Severity $Severity -Tenant $Tenant -Message $Message -Sections $Sections -Footer $Footer -HostName 'fixture-host' -GeneratedAt '2026-10-09 00:00:00 +02:00' }

$Tenant = 'fixture'
$rows = @(
    [pscustomobject]@{ DomainName='a.example'; TotalMailboxCount=4; TotalLocalMailboxCount=1; TotalRemoteMailboxCount=3; EnabledAccounts=3; DisabledAccounts=1; TotalLocalMailboxSizeGB=12.5 },
    [pscustomobject]@{ DomainName='b.example'; TotalMailboxCount=6; TotalLocalMailboxCount=2; TotalRemoteMailboxCount=4; EnabledAccounts=5; DisabledAccounts=1; TotalLocalMailboxSizeGB=7.25 }
)
$rendered = New-SmartM365ExchangeLocalMailboxReportEmailBody -ReportRows $rows -DailyStatsCsv 'daily.csv' -SummaryCsv 'summary.csv' -Title 'Fixture' -ShowMailLinks:$false
$bytes = [Text.Encoding]::UTF8.GetBytes($rendered)
$hasher = [Security.Cryptography.SHA256]::Create()
try { $hash = ([BitConverter]::ToString($hasher.ComputeHash($bytes))).Replace('-', '') } finally { $hasher.Dispose() }
$baseline = '2A82EC5C7A72AABB3E7D345B98CDE3E4C04DA192E5984694114BA566C23E409B'
if ($RecordBaseline) { Write-Output "EXCHANGE_HTML_BASELINE=$hash"; return }
if ($hash -ne $baseline) { throw "Exchange HTML changed: expected $baseline, got $hash" }
Write-Output "PASS: Exchange HTML unchanged ($hash)."

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBzsx6T2ywD9trp
# AF52nYMa2+NQKzJb5HcSoVieLNLkiqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCvZq396tS9VUkRI5Uutbeq
# 9oMHjOy4l2R0O+UXUaOtUDANBgkqhkiG9w0BAQEFAASCAYATIWRBTL+TO4U18I4o
# hNWluPdHDxhWVkhKhHisgGsorxgh/wG5EcJqJnQtVrgSTq79UAU5wDfes8D+m3Xw
# UkQRxkiViK2ZOke4eEJnGAnhY6dodw64L2eTZFQh+wOxAu6NUF8grkxdh5RMd7cl
# WgYXLXFWe8ejmWEgLJIq+dA649WJt05D2RPeG1UmQLfjojaUQKc+oF9WuUOhVIK4
# xzKldkhWMvKKxDeCTCoY29n3J3Jr5hYEIIL/KG0/nW6Avp2h/Ehm7qViRLx1UqQM
# GruquHJG6Cu1TtbvdJ6xGsak4/Y2XW4fA3hrl1jspgZsl7btNOYYWHvjeu00caSl
# TFSMdWpZu1ux86NTX1FMAyowGSoizCx/NZpaB2c+7882mJhCB5WrxUx0ZxGSnoTD
# MagRFEdhxCsPePuVQ30vG9UqiwKs2d0K1oTrZKlDAWMutlvjWQmBJh9hntwWgpBe
# OD7BEJuVShZmztsfWICG5mcgVNe3d5byUXXnVtVdRG33hiI=
# SIG # End signature block
