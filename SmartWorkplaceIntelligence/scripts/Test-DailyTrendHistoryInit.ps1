[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force
$root = Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
$staging = Join-Path $root 'staging'
$output = Join-Path $root 'DATA-POWERBI'
$contract = Join-Path $root 'contract.json.txt'
New-Item -ItemType Directory -Path $staging -Force | Out-Null

function Write-ContractFixture([bool]$IncludeNew) {
    $tables = @(@{table='Existing Trend';file='ExistingTrend.csv';historyKey=@('Date Key');uniqueKey=@('Date Key');columns=@(@{name='Date Key';type='int64'})})
    if ($IncludeNew) {
        foreach ($table in @('Daily Directory Trends','Daily Mailbox Trends')) {
            $file = if ($table -eq 'Daily Directory Trends') { 'DailyDirectoryTrends.csv' } else { 'DailyMailboxTrends.csv' }
            $tables += @{table=$table;file=$file;historyKey=@('Date Key','Domain Key');historyKeyVersion=1;uniqueKey=@('Date Key','Domain Key');columns=@(@{name='Date Key';type='int64'},@{name='Domain Key';type='string'})}
        }
    }
    @{tables=$tables} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $contract
}
function Write-DailyFixture([bool]$IncludeOld) {
    [pscustomobject]@{'Date Key'=20260101} | Export-Csv -LiteralPath (Join-Path $staging 'ExistingTrend.csv') -NoTypeInformation
    foreach ($file in @('DailyDirectoryTrends.csv','DailyMailboxTrends.csv')) {
        $rows = @()
        if ($IncludeOld) { $rows += [pscustomobject]@{'Date Key'=20260101;'Domain Key'='contoso.example'} }
        $rows += [pscustomobject]@{'Date Key'=20260102;'Domain Key'='contoso.example'}
        $rows | Export-Csv -LiteralPath (Join-Path $staging $file) -NoTypeInformation
    }
}
$publish = @{StagingRoot=$staging;OutputRoot=$output;TenantKey='synthetic-test';Provenance=@{Mode='SyntheticTest'};ContractPath=$contract}
Write-ContractFixture $false
Write-DailyFixture $true
Publish-PreparedEvidenceBatch @publish | Out-Null
Write-ContractFixture $true
$second = Publish-PreparedEvidenceBatch @publish
$manifest = Get-Content -LiteralPath (Join-Path $second.BatchPath 'batch.json.txt') -Raw | ConvertFrom-Json
if (@($manifest.HistoryKeyInitializations).Count -ne 2 -or @($manifest.HistoryKeyResets).Count -ne 0) { throw 'First-appearance initialization was not recorded.' }
Write-DailyFixture $false
$rejected = $false
try { Publish-PreparedEvidenceBatch @publish | Out-Null } catch { $rejected = $_.Exception.Message -like '*Historical coverage would regress*' }
if (-not $rejected) { throw 'History regression after initialization was accepted.' }
$current = Get-Content -LiteralPath (Join-Path $output 'current.json.txt') -Raw | ConvertFrom-Json
if ($current.BatchId -ne $second.BatchId) { throw 'Last good batch pointer changed after rejection.' }
Write-Output 'PASS: both new daily trend tables initialize once; later coverage regression is rejected.'

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBqcp86n0khfXBt
# g67mAMq/bVgzWcxaXIfuhdSr1LPba6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDZGUIHib0xHRjXNXFBKHJP
# /Ln6W0Wi/uQ4Hg+MhJVePDANBgkqhkiG9w0BAQEFAASCAYBmlnn/YtvpAqFjKBAu
# f9eL4IRnSXCZ6s4Wn//BPzh/M5/CqwcWCHW34qKkCiQa3AMAvVLJEJBlShHCL3BY
# CNCSdWk4KCtiN77XxmXORarZZ/mQG88qi6i0UJl4mUWxaWnjKKKUaTITW4wKZUUg
# 4hLLQ+zhiqM/i0qOlLuQj8giM39LAdWUD0ZOkoMdk9vQ1CsK0mPAKIcU0My940Gk
# yjZnNrqDfvJu8VrLa6qvEcHnpt4R0XlhhBRi66ImfkAF0Yhe2If0wz762ExGBhm8
# 9S9m+GGpc4EJ25NY8j+KPrrlf/nIe4UsKRyOgQ29zQKQJP8FzIJAC6Em9klKznU+
# jo/dtnK7YNjPwdJfxeeJJAVW6GhlDDOnYyf+lWMjmohS6NGdKfZlRAYEcmYRI0d7
# p8n8/yRU/8SK//PsOQ6mB4RnJwrHwFN8yLQ0fuYyIwVshLXbphIWECPd+GoielDF
# 8q5AH2LxhcIyn4/KHxL+3v+aca0bS8+ab9WiMj/KaUnBSH0=
# SIG # End signature block
