<#
.SYNOPSIS
    Verify the shared upload gate without network access.
.VERSION
    1.0.1
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$manifest = if ($PSVersionTable.PSVersion.Major -lt 7) {
    Join-Path $root 'Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psd1'
} else { Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1' }
$savedVariable = Get-Variable SmartM365CollectorUploadPolicies -Scope Global -ErrorAction SilentlyContinue
$saved = if ($savedVariable) { @{Value=$savedVariable.Value} } else { $null }
$savedRootsVariable = Get-Variable SmartM365CollectorUploadRoots -Scope Global -ErrorAction SilentlyContinue
$savedRoots = if ($savedRootsVariable) { @{Value=$savedRootsVariable.Value} } else { $null }
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365UploadGate-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
try {
    $module = Import-Module $manifest -Force -PassThru
    $checks = & $module {
        param($Base)
        $script:Attempted = 0
        $script:Roots = [pscustomobject]@{DataAllRootPath=(Join-Path $Base 'DATA-ALL');LatestCsvFolderPath=(Join-Path $Base 'DATA-LAST');LogAllRootPath=(Join-Path $Base 'LOG-ALL')}
        function script:Get-SmartM365EffectiveModuleGlobalConfig { $script:Roots }
        function script:Test-SmartM365InterruptedRun { $false }
        function script:Connect-SmartM365GraphForSharePointUpload { param($AppId,$TenantId,$Thumbprint) $script:Attempted++; $false }
        function script:Set-SmartM365SharePointUploadOutcome { param($Path,[switch]$Failed) }
        function script:WriteLog { param($Message,$Level) }
        $paths = @{}
        foreach ($category in @('DataAll','DataLast','LogAll')) {
            $dir = switch ($category) { DataAll {$script:Roots.DataAllRootPath} DataLast {$script:Roots.LatestCsvFolderPath} LogAll {$script:Roots.LogAllRootPath} }
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            $paths[$category] = Join-Path $dir 'synthetic.csv'
            [IO.File]::WriteAllText($paths[$category], 'synthetic')
        }
        $receiptPath=Join-Path $script:Roots.LatestCsvFolderPath 'synthetic.current.json.txt'
        [IO.File]::WriteAllText($receiptPath, '{}')
        $count=0
        foreach ($mask in 0..7) {
            $policies=@{}; $index=0
            foreach ($category in @('DataAll','DataLast','LogAll')) {
                $policies[$category]=[pscustomobject]@{Enabled=[bool]($mask -band (1 -shl $index))}; $index++
            }
            $global:SmartM365CollectorUploadPolicies=$policies
            foreach ($category in @('DataAll','DataLast','LogAll')) {
                $script:Attempted=0
                Invoke-SmartM365SharePointCsvUpload -LocalFilePath $paths[$category] -Enabled $true -SiteHostname 'synthetic' -SitePath '/synthetic' -LibraryDisplayName 'synthetic' -TargetFolderPath 'synthetic' | Out-Null
                if (($script:Attempted -eq 1) -ne $policies[$category].Enabled) { throw "Incorrect gate: $mask/$category" }
                $count++
            }
            $script:Attempted=0
            Invoke-SmartM365SharePointCsvUpload -LocalFilePath $receiptPath -Enabled $true -SiteHostname 'synthetic' -SitePath '/synthetic' -LibraryDisplayName 'synthetic' -TargetFolderPath 'synthetic' | Out-Null
            if (($script:Attempted -eq 1) -ne $policies.DataLast.Enabled) { throw "Receipt did not follow DATA-LAST: $mask" }
            $count++
        }
        $custom = Join-Path $Base 'custom-output'
        New-Item -ItemType Directory -Path $custom -Force | Out-Null
        $customFile=Join-Path $custom 'synthetic.csv'
        [IO.File]::WriteAllText($customFile, 'synthetic')
        $global:SmartM365CollectorUploadRoots=[pscustomobject]@{DataAllRoot=$custom;DataLastRoot=$script:Roots.LatestCsvFolderPath;LogAllRoot=$script:Roots.LogAllRootPath}
        $global:SmartM365CollectorUploadPolicies=@{DataAll=[pscustomobject]@{Enabled=$false};DataLast=[pscustomobject]@{Enabled=$true};LogAll=[pscustomobject]@{Enabled=$true}}
        $script:Attempted=0
        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $customFile -Enabled $true -SiteHostname 'synthetic' -SitePath '/synthetic' -LibraryDisplayName 'synthetic' -TargetFolderPath 'synthetic' | Out-Null
        if ($script:Attempted -ne 0) { throw 'Custom DATA-ALL output bypassed the policy.' }
        $count++
        # A collector-specific DATA-ALL root may sit under the global DATA-LAST root.
        # Both module editions must prefer the collector's actual output classification.
        $global:SmartM365CollectorUploadRoots=[pscustomobject]@{DataAllRoot=$script:Roots.LatestCsvFolderPath;DataLastRoot=$script:Roots.DataAllRootPath;LogAllRoot=$script:Roots.LogAllRootPath}
        foreach ($enabled in @($false,$true)) {
            $global:SmartM365CollectorUploadPolicies.DataAll.Enabled=$enabled
            $global:SmartM365CollectorUploadPolicies.DataLast.Enabled=-not $enabled
            $script:Attempted=0
            Invoke-SmartM365SharePointCsvUpload -LocalFilePath $receiptPath -Enabled $true -SiteHostname 'synthetic' -SitePath '/synthetic' -LibraryDisplayName 'synthetic' -TargetFolderPath 'synthetic' | Out-Null
            if (($script:Attempted -eq 1) -ne $enabled) { throw "Collector-specific root did not override global root: DATA-ALL=$enabled" }
            $count++
        }
        # Registered run artifacts keep their roles outside the configured roots.
        $outsideLog=Join-Path $Base 'outside.log'
        $outsideReceipt=Join-Path $Base 'outside-receipt.json.txt'
        [IO.File]::WriteAllText($outsideLog, 'log')
        [IO.File]::WriteAllText($outsideReceipt, '{}')
        $global:LogTextFile=$outsideLog
        $script:SmartM365SourceReceiptUploadPaths=@($outsideReceipt)
        foreach ($enabled in @($false,$true)) {
            $global:SmartM365CollectorUploadPolicies.LogAll.Enabled=$enabled
            $global:SmartM365CollectorUploadPolicies.DataLast.Enabled=-not $enabled
            foreach ($case in @(@($outsideLog,$enabled), @($outsideReceipt,(-not $enabled)))) {
                $script:Attempted=0
                Invoke-SmartM365SharePointCsvUpload -LocalFilePath $case[0] -Enabled $true -SiteHostname 'synthetic' -SitePath '/synthetic' -LibraryDisplayName 'synthetic' -TargetFolderPath 'synthetic' | Out-Null
                if (($script:Attempted -eq 1) -ne $case[1]) { throw "Incorrect registered artifact gate: $($case[0])" }
                $count++
            }
        }
        Remove-Item -LiteralPath $outsideLog,$outsideReceipt -Force
        $count
    } $tempRoot
    Write-Host "SharePoint upload gate: $checks checks passed; no network calls."
} finally {
    if ($saved) { Set-Variable SmartM365CollectorUploadPolicies -Scope Global -Value $saved.Value }
    else { Remove-Variable SmartM365CollectorUploadPolicies -Scope Global -ErrorAction SilentlyContinue }
    if ($savedRoots) { Set-Variable SmartM365CollectorUploadRoots -Scope Global -Value $savedRoots.Value }
    else { Remove-Variable SmartM365CollectorUploadRoots -Scope Global -ErrorAction SilentlyContinue }
    if ($module) { Remove-Module $module -Force }
    foreach ($folder in @('DATA-ALL','DATA-LAST','LOG-ALL','custom-output')) {
        $path = Join-Path $tempRoot $folder
        $file = Join-Path $path 'synthetic.csv'
        if (Test-Path -LiteralPath $file) { Remove-Item -LiteralPath $file -Force }
        $receipt = Join-Path $path 'synthetic.current.json.txt'
        if (Test-Path -LiteralPath $receipt) { Remove-Item -LiteralPath $receipt -Force }
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    }
    if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Force }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCEW4sGOEMpdb4o
# 5vh1jAO1kddO1P/0vU+c5Mq50PsydaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDcIykNtVhobzLKyPIvi3MP5nSgSp3j3zePjiyTzop7gMA0GCSqG
# SIb3DQEBAQUABIIBgBj9gLk8jPPkIgqWatLAtaDNZX/W8Uctvvt/xADhOfirGqNj
# 8L48aeH3yUMIhKEyOK6sGkc14a1RaSzO4o8imHtKHDpXeg03v33Vvt7geNCp7Mol
# 2G6P5as7iyGdUTBei96sopTHHSjb2ewIKd4KTbARx0H6ZkJx8S9bggDK2ot+plcx
# mNEB78DvRDKR+Qq3toSgiI9TTFmH97jqxqLJlAH/6NOYO0w0ZhozXKh6pIEu04Qa
# vyq/rJor/N6gG64Zo1DwYrGszcdbWkXi3VZ8Q2PGBH2t6HaiSvR7U4Tl8sA4t30u
# JOj+/FeVt/3usNSXMAlNGRtF73SgPZazF5ON/KGmbAEmofU6xsKlvD1aSY3qmh41
# 73TcyzIzVwTVk2Ly73o75Ukfna38PW5MlKRtONgW+uCnWCsn+lcGJg+nndfKGB38
# IUltw1GOjFA7pedqMWEGHS1FoBAEf3fA3t8RMMe0aBa2hvi0DNZ5YzTcBG4hYzYS
# Om1rtVzoW8W8Uculb6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzExMzQz
# MjRaMC8GCSqGSIb3DQEJBDEiBCAgH//I1gzBgMiEBfFNsJy2k7RkVvRdLuGtleep
# YUrz6TANBgkqhkiG9w0BAQEFAASCAgAP4GEnjRbS2EzLC7RpLGhGww0ayZHpXGlw
# 7WFm2D5P1MCbsx0Vd4JW0DDUSE4U14rxaRA8BVaVtSYylJ/rS0w55WQxgtMPb7ip
# Op+NpjdMypC5kz/EuKo08bjWADQOnDPFa/k4KPXMCI14WfdVW/LEH8vqjW83Lmam
# vlmLXuX4NxkbCdsvW9nYKixjF6gTxAryUwx7f7lNbN1Tu50sR2XZQAXEYqe5kwRU
# DC/SLFrLVbt8MGBgdJo1VwCg4R6p76Y+cqvHxXShdCtGR+/cyfVvHcNNJa2MaVur
# 9ahRue91fZ3WjljtyEoB3t20gdIEuxVti1dguhlfbqVbmIrFFNAoTNElihZ1WLWV
# 95QYV4l8yKdXLI2hpYEfFXv5P3TB+TQfnW0cazcC+KQhf7iqFolID3/aN/m99coQ
# 4FszvSkuJhqSVossQdt2S3Aw4Sk84tLMarwlxt6ZfWwcdhLYO9E7BVeDVTz6UxWS
# bK+RNByyiUI1QBnF8P/kyompOLMlxE33Ujy5iGEXnTSIoXvJTbBB1+vHitPUgnq0
# Y8/Ns3L4+1JN/PYdnYhQe37J+qvuhergut8qqQjUeqE9fO2iRTHToCiQWqDYHbpf
# vbEIu4zQ1KA97t0AD6lsEVsoZC3s3W5TMMoK1k7bKjR6oO1OVwBiDe0a7elpBcge
# 6Inb4YhiHw==
# SIG # End signature block
