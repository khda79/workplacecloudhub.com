<#
.SYNOPSIS
Offline tests of the weekly history JSON transition and changed-only SharePoint publication.

.VERSION
1.3
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Force
$transport = Get-Module SmartM365.JsonTransport
$policy = & $transport { (Get-Command Get-SmartM365JsonTransportPolicy).ScriptBlock }
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-JsonWeekly-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory $root
$script:passed = 0
function Check([bool]$Value, [string]$Message) { if (-not $Value) { throw $Message }; $script:passed++ }
function Reject([scriptblock]$Action, [string]$Message) { $failed=$false; try { & $Action | Out-Null } catch { $failed=$true }; Check $failed $Message }
function WriteLog { param($Message,$Level) }
function Get-SmartM365IsoWeekName { '2026-W39' }
function Get-SmartM365WeeklyHistoryMaxRetentionWeeks { 0 }
function Get-SmartM365WeeklyHistoryFileName { param($Path) [IO.Path]::GetFileName($Path) }
function Copy-SmartM365FileAtomically { param($SourcePath,$DestinationPath) Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath }
function Invoke-SmartM365SharePointCsvUpload { param($LocalFilePath) $script:uploads += $LocalFilePath; if ($script:failUploads) { return $null }; return 'receipt' }
try {
    foreach ($relative in @('../Modules/SmartM365.Core/SmartM365.Core.psm1','../Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1')) {
        $tokens=$null;$errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $relative),[ref]$tokens,[ref]$errors)
        if ($errors.Count) { throw ($errors | Out-String) }
        $definition=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Save-SmartM365WeeklyInventoryHistory'},$true)
        . ([scriptblock]::Create($definition.Extent.Text))
        $case = Join-Path $root ([guid]::NewGuid().ToString('N'))
        $history = Join-Path $case 'WeeklyHistory'
        $week = Join-Path $history '2026-W39'
        $oldWeek = Join-Path $history '2025-W01'
        $null = New-Item -ItemType Directory $week,$oldWeek -Force
        $source = Join-Path $case 'Inventory.csv'
        [IO.File]::WriteAllText($source,"Id`r`n42`r`n")
        Copy-Item $source $week
        Copy-Item $source $oldWeek
        foreach ($folder in @($week,$oldWeek)) {
            $doc=[ordered]@{Week=[IO.Path]::GetFileName($folder);HistoryLabel='Synthetic';HistoryRootPath=$history;Files=@('Inventory.csv');SnapshotCreatedAtUtc='2025-01-01T00:00:00Z';FileSnapshotCreatedAtUtc=@{};UpdatedAt='2025-01-01T00:00:00Z'}
            [IO.File]::WriteAllText((Join-Path $folder 'manifest.json'),($doc | ConvertTo-Json -Depth 7),[Text.UTF8Encoding]::new($true))
        }
        $legacy = Join-Path $week 'manifest.json'
        $oldLegacy = Join-Path $oldWeek 'manifest.json'
        $hash = (Get-FileHash $legacy).Hash
        $oldHash = (Get-FileHash $oldLegacy).Hash
        & $transport { function script:Get-SmartM365JsonTransportPolicy { @{Mode='Readers';QualifiedUncRoots=@();QualifiedLocalDataRoots=@()} } }
        $script:uploads=@()
        Save-SmartM365WeeklyInventoryHistory -SourceFiles $source -HistoryRootPath $history -HistoryLabel Synthetic -RetentionWeeks 0
        Check ((Get-FileHash $legacy).Hash -eq $hash) 'Reader rollout rewrote an unchanged historical manifest.'
        & $transport { function script:Get-SmartM365JsonTransportPolicy { @{Mode='JsonText';QualifiedUncRoots=@();QualifiedLocalDataRoots=@()} } }
        $script:uploads=@()
        Save-SmartM365WeeklyInventoryHistory -SourceFiles $source -HistoryRootPath $history -HistoryLabel Synthetic -RetentionWeeks 0 -UploadChangedFilesOnly
        Check ((Get-FileHash "$legacy.txt").Hash -eq $hash -and -not (Test-Path $legacy)) 'Current manifest bytes changed during migration.'
        Check ((Get-FileHash "$oldLegacy.txt").Hash -eq $oldHash -and -not (Test-Path $oldLegacy)) 'Older manifest was lost or recalculated.'
        Check ($script:uploads -contains "$oldLegacy.txt" -and $script:uploads -contains "$legacy.txt") 'Converted manifests absent from upload candidates.'
        Check (@($script:uploads | Where-Object { $_ -match '\.json$' }).Count -eq 0) 'Legacy manifest selected for upload.'
        $script:uploads=@()
        Save-SmartM365WeeklyInventoryHistory -SourceFiles $source -HistoryRootPath $history -HistoryLabel Synthetic -RetentionWeeks 0 -UploadChangedFilesOnly
        Check ($script:uploads -contains "$oldLegacy.txt") 'Restart after migration did not retry remote publication.'
        Check ((Get-FileHash "$legacy.txt").Hash -eq $hash) 'Idempotent execution rewrote historical metadata.'
        Copy-Item "$legacy.txt" $legacy
        [IO.File]::WriteAllText("$legacy.txt",'{')
        Reject { Save-SmartM365WeeklyInventoryHistory -SourceFiles $source -HistoryRootPath $history -HistoryLabel Synthetic -RetentionWeeks 1 } 'Invalid preferred manifest fell back or was rebuilt.'
        Check (Test-Path "$oldLegacy.txt") 'Retention ran after failed manifest validation.'
        Copy-Item $legacy "$legacy.txt" -Force
        $foreign=Get-Content "$legacy.txt" -Raw | ConvertFrom-Json
        $foreign.HistoryLabel='Foreign'
        [IO.File]::WriteAllText("$legacy.txt",($foreign | ConvertTo-Json -Depth 7))
        Reject { Resolve-SmartM365WeeklyManifestPaths -HistoryRootPath $history -HistoryLabel Synthetic } 'Divergent or foreign manifest accepted.'
        Check ((Get-FileHash $legacy).Hash -eq $hash) 'Failure destroyed the last valid legacy manifest.'
        # Changed-only publication recovers from a failed SharePoint upload through the pending marker.
        $pendingCase = Join-Path $root ([guid]::NewGuid().ToString('N'))
        $pendingHistory = Join-Path $pendingCase 'WeeklyHistory'
        $null = New-Item -ItemType Directory $pendingHistory -Force
        $pendingSource = Join-Path $pendingCase 'Inventory.csv'
        [IO.File]::WriteAllText($pendingSource, "Id`r`n7`r`n")
        $pendingMarker = Join-Path $pendingHistory '2026-W39\upload.pending'
        $pendingCsv = Join-Path $pendingHistory '2026-W39\Inventory.csv'
        Set-Variable -Name EnableSharePointUpload -Scope Global -Value $true
        try {
            $script:failUploads = $true; $script:uploads = @()
            Reject { Save-SmartM365WeeklyInventoryHistory -SourceFiles $pendingSource -HistoryRootPath $pendingHistory -HistoryLabel Synthetic -RetentionWeeks 0 -UploadChangedFilesOnly } 'Failed weekly publication was not reported.'
            Check ((Test-Path $pendingCsv) -and (Test-Path $pendingMarker)) 'Failed publication left no pending marker.'
            $script:failUploads = $false; $script:uploads = @()
            Save-SmartM365WeeklyInventoryHistory -SourceFiles $pendingSource -HistoryRootPath $pendingHistory -HistoryLabel Synthetic -RetentionWeeks 0 -UploadChangedFilesOnly
            Check ($script:uploads -contains $pendingCsv) 'Pending week was not republished.'
            Check (-not (Test-Path $pendingMarker)) 'Pending marker kept after a complete publication.'
            Check (@($script:uploads | Where-Object { $_ -like '*upload.pending' }).Count -eq 0) 'Pending marker selected for upload.'
            $script:uploads = @()
            Save-SmartM365WeeklyInventoryHistory -SourceFiles $pendingSource -HistoryRootPath $pendingHistory -HistoryLabel Synthetic -RetentionWeeks 0 -UploadChangedFilesOnly
            Check (-not ($script:uploads -contains $pendingCsv)) 'Unchanged week CSV uploaded again without a pending marker.'
        }
        finally { Set-Variable -Name EnableSharePointUpload -Scope Global -Value $false; $script:failUploads = $false }
        $disabledHistory = Join-Path (Join-Path $root ([guid]::NewGuid().ToString('N'))) 'WeeklyHistory'
        $null = New-Item -ItemType Directory $disabledHistory -Force
        Save-SmartM365WeeklyInventoryHistory -SourceFiles $pendingSource -HistoryRootPath $disabledHistory -HistoryLabel Synthetic -RetentionWeeks 0 -UploadChangedFilesOnly
        Check (-not (Test-Path (Join-Path $disabledHistory '2026-W39\upload.pending'))) 'Pending marker written while SharePoint publication is disabled.'
    }
    [pscustomobject]@{Passed=$script:passed;FixtureRoot=$root;Evidence='Synthetic; SharePoint upload mocked; no collector entry point'}
} finally { & $transport { param($original) Set-Item Function:script:Get-SmartM365JsonTransportPolicy -Value $original } $policy }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAKD+jZjKRGItaG
# 4/5WBeZ+LgC5WMwdh06fgMWxegw12KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEII0HOSBjEzTz+k4xE+PmbXtvuqUnwE+EzkTshddQiwpzMA0GCSqG
# SIb3DQEBAQUABIIBgD9HY+H2n5wImLyX4EMLXV8/Rf8h95qVB0Jj66nKsSWmF5Xc
# JRfaSBSMMeskgKC6Np/0lz/dz1y6EE2kKZdG02OVX3bC1TK3metxg/zZAENP+xXi
# AZ5YTtJ/B5fTcCERgz83Ou+OIDFjE/8z16rYHiFavj6+OAxCOuhC27i893IzxK16
# uDdX7txPjS5a7aObKQkQFJ6BZW/dufxpXOZTyZ/OG+DnXsFHaAw2P8t9QpmelLHs
# M2sNgCydtKhT6M/ZlFB/j2SqDkIGGWJ2hKK9r8vEjjoRavQzeCWuXK5PecIUQ6PZ
# /I+oHfVNJ9gnrnhg1u9MLVL++PARsM+vfvXocI0kHwLNWwktzg0MrFbKL1uuZCbS
# fPQMjcZxeENn6RGsi84Tchosh5VSyWSrpaa68GDIdTPta0sGi8jFQN6KJPzf94K1
# IOeDkp2Dr7L4hCW4BYCRysSStuoreua7KvGD0kcVOSk9ubdrL4bRfPWK7FO9NcTP
# EUDNjyx82oEWHUKc9qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkxOTM2
# MzhaMC8GCSqGSIb3DQEJBDEiBCDsfPN9rePKZQ3T5JtCf2zGYdIaoceZQFey9JGZ
# 9zDCajANBgkqhkiG9w0BAQEFAASCAgBkj0xkYHWkQuN1csWWh0ckd6HK0C8sXuC9
# XAiG6bn1gZCdZ1wP28QVN5XHFhjdC/HtyT6+kHhCNoW0Qjr/EJtrmCqDjx+9JMwG
# oBL3cHv9MtcB3kXV0exA4mP0HVam+KJBcoziWwg6tnl43RDkIBpA87EUAieDxHGx
# n3yritpzGoY6T60QuZ8OGZcd3XRP+84tSvv+4YIIUnLBtEUQmVphcZBkQLR3OpiV
# tZTiaOGK682AvKEAPSzTI2BWz5Cfj5s5Z+Dys8P3Z78KuaGP0pKUzomtDBH5UofM
# 4Zdjt49oOzCFqn2fSui3Oshy7mOB5ef2CmnJUOMnuXc6H4h0/SH+3KynyAWLSiJm
# g9eeodbb1UUzPU/PenM3PKM39X5W9IySj1ky5scHJcM/eMG1ApuKGW/rbycDrOso
# MvbRAB2eryrXiY0D1CI2H9vrFRdiWAFdqcIaGu8YuwJjI5gTFmygI3T2erBZqcUM
# vqFU3iRja0PRfFLCsQeE2nlZ5xPmhkyDLmue+B5+SgS0/kdCXxDS+NUZtNm3ND6n
# 6/oDceJSV0uuPjZYGfo1kVW2tdq2bz23uY5jyL2uZDQKPBqd9dywv98f/6iay9OA
# Cs2703cIqsrK46b9FIvmkB4Aw2ZQJZSB2qXB/8EYf0nQHez6FRXeW4FSXgawwdYr
# ChoZhPOGdw==
# SIG # End signature block
