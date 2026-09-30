<#
.SYNOPSIS
Offline tests for the seven-day SharePoint retention of timestamped CSV copies (SmartM365.Core and
the Windows PowerShell 5 compatibility module).
.VERSION
1.0.0
#>
[CmdletBinding()]
param([switch]$WindowsPowerShell5)

$ErrorActionPreference = 'Stop'
$smartM365Root = Split-Path -Path $PSScriptRoot -Parent
$modulePath = if ($WindowsPowerShell5) {
    Join-Path -Path $smartM365Root -ChildPath 'Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psd1'
}
else {
    Join-Path -Path $smartM365Root -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
}
$minimumVersion = if ($WindowsPowerShell5) { '1.0.45' } else { '1.0.61' }
$module = Import-Module -Name $modulePath -MinimumVersion $minimumVersion -Force -PassThru -ErrorAction Stop

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Assertion failed: $Message" }
}

$retention = & $module {
    param($Ps5)
    $script:Deletes = New-Object System.Collections.ArrayList
    $script:Lists = New-Object System.Collections.ArrayList
    $script:Logs = New-Object System.Collections.ArrayList
    function script:Connect-SmartM365GraphForSharePointUpload { param($AppId, $TenantId, $Thumbprint) return $true }
    function script:ConvertTo-SmartM365SharePointDataRootPath { param($TargetFolderPath) return 'SMART-M365/DATA' }
    function script:Get-SmartM365SharePointRelativeFilePath { param($LocalFilePath) return 'DATA-ALL/Intune/Alerts/Report_20260930_120000.csv' }
    function script:WriteLog { param($Message, $Level) [void]$script:Logs.Add(("{0}|{1}" -f $Level, $Message)) }
    $file = { param($Id, $Name) [pscustomobject]@{ id = $Id; name = $Name; file = [pscustomobject]@{ mimeType = 'text/csv' } } }
    $page1 = [pscustomobject]@{
        value = @(
            (& $file 'current' 'Report_20260930_120000.csv'),
            (& $file 'recent' 'Report_20260925_010000.csv'),
            (& $file 'expired1' 'Report_20260920_010000.csv'),
            (& $file 'latest' 'Report.csv'),
            (& $file 'otherprefix' 'Report_Detail_20260901_010000.csv'),
            (& $file 'otherext' 'Report_20260901_010000.xlsx'),
            [pscustomobject]@{ id = 'folder'; name = 'Report_20260901_010000.csv'; folder = [pscustomobject]@{ childCount = 1 } }
        )
        '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/next-page'
    }
    $page2 = [pscustomobject]@{ value = @((& $file 'expired2' 'Report_20260801-230000.csv'), (& $file 'failing' 'Report_20260810_010000.csv')) }
    $script:Pages = @{ first = $page1; second = $page2 }
    function script:Invoke-SmartM365GraphRestWithRetry {
        param($Method, $Uri, $Body, $ContentType, $Operation, $AdditionalHeaders)
        if ($Method -eq 'DELETE') {
            if ($Uri -like '*/items/failing') { throw 'Graph request failed. Status=403' }
            [void]$script:Deletes.Add(($Uri -replace '^.*/items/', ''))
            return $null
        }
        if ($Uri -like '*/sites/*/drives') { return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'drive1'; name = 'Documents' }) } }
        if ($Uri -like '*/sites/*') { return [pscustomobject]@{ id = 'site1' } }
        [void]$script:Lists.Add($Uri)
        if ($Uri -eq 'https://graph.microsoft.com/v1.0/next-page') { return $script:Pages.second }
        return $script:Pages.first
    }
    function script:Invoke-SmartM365GraphDeleteQuietly {
        param($Uri, $Operation)
        if ($Uri -like '*/items/failing') { return [pscustomobject]@{ Success = $false; NotFound = $false; Message = 'HTTP 403' } }
        [void]$script:Deletes.Add(($Uri -replace '^.*/items/', ''))
        return [pscustomobject]@{ Success = $true; NotFound = $false; Message = '' }
    }
    $result = Remove-SmartM365SharePointTimestampedCsvOlderThan -TimestampedPath 'C:\Data\DATA-ALL\Intune\Alerts\Report_20260930_120000.csv' -RetentionDays 7 -ReferenceTime ([datetime]'2026-09-30T13:00:00') -Enabled $true -SiteHostname 'tenant.sharepoint.test' -SitePath '/sites/S' -LibraryDisplayName 'Documents' -TargetFolderPath 'SMART-M365'
    $notTimestamped = Remove-SmartM365SharePointTimestampedCsvOlderThan -TimestampedPath 'C:\Data\DATA-LAST\Report.csv' -Enabled $true -SiteHostname 'h' -SitePath '/s' -LibraryDisplayName 'Documents' -TargetFolderPath 'T'
    New-Object psobject -Property @{
        Result = $result; NotTimestamped = $notTimestamped
        Deletes = @($script:Deletes | Sort-Object); Lists = @($script:Lists); Logs = @($script:Logs)
    }
} $WindowsPowerShell5.IsPresent

Assert-True (($retention.Deletes -join ',') -eq 'expired1,expired2') ("expired copies deleted: expected expired1,expired2, got {0}" -f ($retention.Deletes -join ','))
Assert-True ($retention.Result.Deleted -eq 2 -and $retention.Result.Failed -eq 1 -and $retention.Result.Kept -eq 1) ("counters Deleted/Failed/Kept = {0}/{1}/{2}" -f $retention.Result.Deleted, $retention.Result.Failed, $retention.Result.Kept)
Assert-True ($retention.Lists.Count -eq 2 -and $retention.Lists[0] -like '*/drives/drive1/root:/SMART-M365/DATA/DATA-ALL/Intune/Alerts:/children*') ("folder listing followed the next link from the CSV folder: {0}" -f ($retention.Lists -join ' | '))
Assert-True (@($retention.Logs | Where-Object { $_ -like 'WARNING|*Report_20260810_010000.csv*' }).Count -eq 1) 'a failed deletion is reported as a warning'
Assert-True ($retention.NotTimestamped.Deleted -eq 0 -and $retention.NotTimestamped.Failed -eq 0) 'a latest (non-timestamped) path never triggers retention'

$publish = & $module {
    $script:RetentionCalls = New-Object System.Collections.ArrayList
    $script:UploadTimestamped = $true
    function script:Invoke-SmartM365SharePointCsvUpload {
        param($LocalFilePath)
        if (-not $script:UploadTimestamped -and $LocalFilePath -match '_\d{8}_\d{6}\.csv$') { return $null }
        return [pscustomobject]@{ LocalFilePath = $LocalFilePath; SharePointPath = 'x'; WebUrl = '' }
    }
    function script:Remove-SmartM365SharePointTimestampedCsvOlderThan { param($TimestampedPath, $RetentionDays) [void]$script:RetentionCalls.Add(("{0}|{1}" -f [IO.Path]::GetFileName($TimestampedPath), $RetentionDays)) }
    function script:Invoke-SmartM365WeeklyInventoryHistoryForCsv { param($SourceFiles, $TimestampedPath) }
    function script:WriteLog { param($Message, $Level) }
    $root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-SpRetention-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path (Join-Path $root 'all'), (Join-Path $root 'last') -Force | Out-Null
    $global:RetentionMaxCSV = 0
    try {
        $data = @([pscustomobject]@{ Name = 'a'; Value = '1' })
        $run = {
            param($Stamp, [int]$Days = 7)
            # The PS5 module has no -NoWeeklyHistory; the history publication is stubbed in both modules.
            $extra = @{}
            if ((Get-Command Publish-SmartM365Csv).Parameters.ContainsKey('NoWeeklyHistory')) { $extra['NoWeeklyHistory'] = $true }
            Publish-SmartM365Csv -Data $data -TimestampedPath (Join-Path $root "all\Report_$Stamp.csv") -LatestPath (Join-Path $root 'last\Report.csv') -Columns @('Name', 'Value') -NoTenantKey -RetentionMaxCsv 0 -SharePointRetentionDays $Days @extra | Out-Null
        }
        & $run '20260930_120000'
        $script:UploadTimestamped = $false
        & $run '20260930_130000'
        $script:UploadTimestamped = $true
        & $run '20260930_140000' 0
        @($script:RetentionCalls)
    }
    finally { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }
}
Assert-True ((@($publish) -join ',') -eq 'Report_20260930_120000.csv|7') ("Publish-SmartM365Csv retention calls: expected only the uploaded run with 7 days, got '{0}'" -f (@($publish) -join ','))

$edition = if ($WindowsPowerShell5) { 'WindowsPowerShell5' } else { 'Core' }
Write-Host ("SHAREPOINT_TIMESTAMPED_RETENTION_TEST_OK Module={0}; Deleted={1}; Failed={2}; Kept={3}" -f $edition, $retention.Result.Deleted, $retention.Result.Failed, $retention.Result.Kept)
if (-not $WindowsPowerShell5 -and $PSVersionTable.PSEdition -eq 'Core') {
    $ps5 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path -LiteralPath $ps5) {
        & $ps5 -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath -WindowsPowerShell5
        if ($LASTEXITCODE -ne 0) { throw "Windows PowerShell 5 compatibility test failed (exit $LASTEXITCODE)." }
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA1pxQirgtZCAxm
# EI9QBLndmRrpZ0OE8oy+AC2A/pB6naCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEbGcfW084dbnYsxlQdCxlF0/1wYhlBt257iHgRCUuE0MA0GCSqG
# SIb3DQEBAQUABIIBgA8GysasE6uK6In9oX9g6VUv3cYdfK4hJfayQtKsKRfri8am
# 75gJmWiht8XzexuOYVaUfYnh7ZuRp/LJw0GKRPrAOvldu/FpLAWkmvxfIa9ADb/6
# sWlKw1wfhQq8skYbmmg68ol28uMaY4r3NGADPRgRQuDYSQLByEF+5rX8+5+9coIs
# TjV4VM+FZ/glGWtjIWnsB/Jnl9s6KvzuAPr8+R+NKsFrG7ca8OxJ7XKKpz1Oma2W
# zAbOowwA35hLCvlHE1OIf8TpFYnG6UkEY5gIPjSASfM852CODJtgvNtqWmnw9Yj7
# vgBzXrI9TI9aPk5v1js03mlOH9Lheq2MLhhSZ9hxKlJmCnJ+9MpLvTxiDExUwf3y
# IaOrsa9AA4Uc9S1tzAA7ooqYNpqgiE6VP+meAgPYGtlROIBJCNaEGnxH4gbP+6DD
# thPYcfZbZ0ae3M0SgeOfXGHk423gOUOpfcN9DgRTulclXLtWB1NwQFcTl5uVz5Ow
# a7x0aZ5U6bnTsf7koKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MzAyMDQ3
# MjhaMC8GCSqGSIb3DQEJBDEiBCD9pG4wH6TZ9gqUlUV0aEt1frUB5XxmRzXhbAb7
# 6RDpYzANBgkqhkiG9w0BAQEFAASCAgA/yex5IxJ5lLPHlg82+8zvhKSttrQLN/1O
# 7aoI7QRzsNac5bPNDiEd7f9S4D6FzQRfab0ObWZmOI1gLeCBEGGhp1Y5g5EiL99P
# sdybwen/vibWpVMDGkvWoCyryo5NAqNco4N3mv7OsXDDxcBz5gwDk6O0H1tX8N1m
# Ke+K5n96Zz7w/orN2sXwx4wIQfALYcrxBCogjJCqOOlvMxHnORpVX2X/ZmDc/2zS
# ebfWDdbtFHsmOHRBw7OmEfq8cawybXqk8OjdPyKdd2TWPGKqdC2VwqLTRDZDwUyb
# Owk3i2adl7rLzcjqjt48XBAqZhFiYSqTbWiewSHzMmFzn3HkMGo1Dan9NSdwqKZr
# ri6teD9b6GOMqXjIsIQF0mZCgwrzmLJTB7FP965sMmReGiXCLeaBjGp6wLgz9r2b
# UkK/k6R32drPR2slKeOeRG6h7iFVc11AiHiidCOhzzr7JnvhaOSIPadF1qqk8ww6
# Vv0eKRMSkNs6TxHs3icZ+rXG1GRgiktHcDihycUOF33jRdmj6KalDoLrEKuLV4Va
# vWXc6cFDdqd0YAVPPruTnZe/exLd9WgR2oEADJydmO7ywKUNdthQeQsk57eAx8qj
# 0UKO+3h3HBR0FLKCRJBncmTBzN701vkbio+Cnv222sq83nZT09A44CUsDqsZm0PM
# v5MhyLWQeA==
# SIG # End signature block
