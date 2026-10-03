<#
.SYNOPSIS
Offline tests for fresh app/device acquisition and checkpoint preservation.
.DESCRIPTION
Executes only AST helpers and the resume-selection branch against temporary
synthetic files. No collector entry point, Graph, tenant config or transport.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$smartRoot = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Force
$sourcePath = Join-Path $smartRoot 'SmartInventory/M365Inventory/IntuneInventory/Applications/SmartM365-Intune-DiscoveredApps-Inventory.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($sourcePath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
foreach ($name in @('Assert-DiscoveredAppsFreshDeviceDetailsOptions','Test-DiscoveredAppsResumeStateCompatible',
    'Resolve-DiscoveredAppsResumePath','Get-DiscoveredAppsResumeState','Save-DiscoveredAppsPreviousCheckpoint',
    'Save-DiscoveredAppsResumeState','Write-DiscoveredAppsCsvRows','Repair-DiscoveredAppsResumePartialCsv')) {
    $node = $ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name}, $true)
    if ($null -eq $node) { throw "Missing helper: $name" }
    Invoke-Expression $node.Extent.Text
}
function WriteLog { param([string]$Message, [string]$Level) }
function Assert-True { param([bool]$Condition, [string]$Label) if (-not $Condition) { throw $Label } }
$testCount = 0
function Test-Case {
    param([string]$Name, [scriptblock]$Body)
    & $Body
    $script:testCount++
    Write-Output "PASS: $Name"
}
function Assert-Rejected {
    param([hashtable]$Options)
    $failed = $false
    try { Assert-DiscoveredAppsFreshDeviceDetailsOptions @Options } catch { $failed = $true }
    Assert-True $failed 'Unsafe fresh options were accepted.'
}
Test-Case 'Fresh full mode accepted; legacy options unchanged' {
    Assert-DiscoveredAppsFreshDeviceDetailsOptions -FreshDeviceDetails -Mode All -MaxApps 0 -MaxItems 0
    Assert-DiscoveredAppsFreshDeviceDetailsOptions -Mode Top -MaxApps 5 -MaxItems 5 -ResetResume
}
foreach ($mode in @('None','Top','NonZero')) {
    Test-Case "Fresh mode rejects $mode" { Assert-Rejected @{FreshDeviceDetails=$true; Mode=$mode; MaxApps=0; MaxItems=0} }
}
Test-Case 'Fresh mode rejects app limits' { Assert-Rejected @{FreshDeviceDetails=$true; Mode='All'; MaxApps=1; MaxItems=0} }
Test-Case 'Fresh mode rejects item limits' { Assert-Rejected @{FreshDeviceDetails=$true; Mode='All'; MaxApps=0; MaxItems=1} }
Test-Case 'Fresh mode rejects destructive reset' { Assert-Rejected @{FreshDeviceDetails=$true; Mode='All'; MaxApps=0; MaxItems=0; ResetResume=$true} }

$testRoot = Join-Path $env:TEMP ('SmartM365-FreshApps-' + [guid]::NewGuid().ToString('N'))
try {
    $null = New-Item -ItemType Directory -Path $testRoot
    $oldPartial = Join-Path $testRoot 'old.partial.csv'
    $oldCompleted = Join-Path $testRoot 'old.csv'
    Write-DiscoveredAppsCsvRows -Path $oldPartial -Rows @()
    [IO.File]::AppendAllText($oldPartial, '"tenant-test","app-old","device-old"' + [Environment]::NewLine)
    [IO.File]::WriteAllText($oldCompleted, 'preserved previous completed export')
    $TaskName = 'SmartM365-Intune-DiscoveredApps-Inventory v1.30'
    $script:DeviceDetailResumePath = Join-Path $testRoot 'relations.resume.json.txt'
    Save-DiscoveredAppsResumeState -Path $script:DeviceDetailResumePath -PartialPath $oldPartial -TimestampedPath $oldCompleted `
        -Mode All -TargetCount 2 -ResumeContractVersion 6 -TargetAppIdsHash 'same-set' `
        -ProcessedAppIds @('app-old') -ProcessedCount 1 -SkippedCount 0 -DetailRows 1 -ActualDeviceCounts @{'app-old'=1}
    $oldHashes = @{}
    foreach ($path in @($oldPartial,$oldCompleted,$script:DeviceDetailResumePath)) { $oldHashes[$path] = (Get-FileHash -LiteralPath $path).Hash }
    $resumeState = Get-DiscoveredAppsResumeState -Path $script:DeviceDetailResumePath
    Test-Case 'Compatible checkpoint still resumes without fresh switch' {
        Assert-True (Test-DiscoveredAppsResumeStateCompatible -State $resumeState -Mode All -TargetCount 2 -TargetAppIdsHash 'same-set' -ResumeContractVersion 6) 'Default resume regressed.'
    }
    Test-Case 'Fresh switch refuses even an identical compatible app set' {
        Assert-True (-not (Test-DiscoveredAppsResumeStateCompatible -State $resumeState -Mode All -TargetCount 2 -TargetAppIdsHash 'same-set' -ResumeContractVersion 6 -FreshDeviceDetails)) 'Old relations reused.'
    }
    $branch = $ast.Find({param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -eq '$resumeStateCompatible'}, $true)
    if ($null -eq $branch) { throw 'Resume selection branch not found.' }
    # Execute the actual collector branch: fresh mode must leave no processed IDs.
    $resumeStateCompatible = $false
    $OutputPath = $testRoot; $detailBaseFileName = 'relations'; $detailTimestamp = 'new-run'
    $streamingEnabled = $true; $DeviceDetailMode = 'All'; $script:UsePreviousDeviceDetailCache = $false
    $RefreshDeviceDetailCache = $false
    $processedAppIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $actualDeviceCountsByAppId = @{}
    Test-Case 'Fresh branch archives old checkpoint and leaves every app for Graph' {
        Invoke-Expression $branch.Extent.Text
        Assert-True ($processedAppIds.Count -eq 0 -and $actualDeviceCountsByAppId.Count -eq 0) 'Old identities/counts reused.'
        foreach ($path in $oldHashes.Keys) { Assert-True ((Get-FileHash -LiteralPath $path).Hash -eq $oldHashes[$path]) "Old bytes changed: $path" }
        Assert-True (@(Import-Csv -LiteralPath $script:DeviceDetailPartialPath).Count -eq 0) 'New partial contains old relations.'
        $archive = Join-Path (Join-Path $testRoot 'ResumeHistory') ($oldHashes[$script:DeviceDetailResumePath] + '.json.txt')
        Assert-True ((Get-FileHash -LiteralPath $archive).Hash -eq $oldHashes[$script:DeviceDetailResumePath]) 'Archived checkpoint bytes differ.'
    }
    Test-Case 'A new checkpoint does not remove old partial or saved checkpoint' {
        Save-DiscoveredAppsResumeState -Path $script:DeviceDetailResumePath -PartialPath $script:DeviceDetailPartialPath -TimestampedPath $script:DeviceDetailTimestampedPath `
            -Mode All -TargetCount 2 -ResumeContractVersion 6 -TargetAppIdsHash 'same-set' `
            -ProcessedAppIds @('app-new') -ProcessedCount 1 -SkippedCount 0 -DetailRows 0 -ActualDeviceCounts @{'app-new'=0}
        foreach ($path in @($oldPartial,$oldCompleted)) { Assert-True ((Get-FileHash -LiteralPath $path).Hash -eq $oldHashes[$path]) 'Old export changed.' }
        $archive = Join-Path (Join-Path $testRoot 'ResumeHistory') ($oldHashes[$script:DeviceDetailResumePath] + '.json.txt')
        Assert-True ((Get-FileHash -LiteralPath $archive).Hash -eq $oldHashes[$script:DeviceDetailResumePath]) 'Original checkpoint lost.'
    }
    Test-Case 'Runtime binds fresh option before acquisition and keeps strict receipt gate' {
        $text = $ast.Extent.Text
        Assert-True ($text -match '-FreshDeviceDetails:\$FreshDeviceDetails') 'Runtime fresh binding missing.'
        Assert-True ($text.IndexOf('Assert-DiscoveredAppsFreshDeviceDetailsOptions -FreshDeviceDetails') -lt $text.IndexOf('if ($streamingEnabled -and $ResetResume)')) 'Destructive path runs before option validation.'
        Assert-True ($text -match 'Stat_DetailAppsFromCache -eq 0 -and \$script:Stat_DetailAppsSkippedByResume -eq 0') 'Full fresh scope gate weakened.'
        Assert-True ($text -match 'UsePreviousDeviceDetailCache = -not \$RefreshDeviceDetailCache -and \$DeviceDetailMode -ne ''All''') 'All mode can reuse cache.'
    }
    Write-Output "Fresh application offline tests: $testCount passed. No external actions."
} finally {
    $resolvedTestRoot = [IO.Path]::GetFullPath($testRoot)
    $tempPrefix = [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\'
    if (-not $resolvedTestRoot.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolvedTestRoot) -notlike 'SmartM365-FreshApps-*') { throw 'Unsafe test cleanup target.' }
    if (Test-Path -LiteralPath $resolvedTestRoot) { Remove-Item -LiteralPath $resolvedTestRoot -Recurse -Force }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC+jTkInft8OnQO
# CrQ6+YmCB+5cCf1q/5mTfsX6gWHUHKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOvEoGEmm1HEDxCOg4BjMB2+Ts4BuMHIPyIoHM8XoDz7MA0GCSqG
# SIb3DQEBAQUABIIBgA9TYyROWz1a0xsVUgfImu3tC8arcuB9v95t1NmchR2dI1LO
# XbuFfnEj/HwdFCC6mFytDhipcTLgVRw72ro+Z6QM1sGlkjBhP3YlL5tdyzv0MUh1
# 0YClV2VFTPwXaXlBjIqBA43N8aiopsS1ZgHfAb/Y1AfaxBTb36LnImBaC9Dc8uA8
# KzYLVCmc1KfUMj7Kb+swk6ON9rtQPQ4BAca3hLC/adTuOJkr2TJEBNlIxtn19HnF
# pr5Nz+rYA4if74Jvxux0Dcj9HMbUX6k+3ElkK/iZCpPMdHXy9yVI+D4Md/JiWD57
# yx+9bjh+i6BWxDo+AepgLzFv+3r19dOooGEUqQ71w6sK9r8yVWDDkuBWvwvwPNKZ
# H07MEoJNZFQinMEMe+Uo4FwobZN4r3g4ZVUlDrk23yILc6e6Mag4ZZPMJKzFGLMi
# V3dFGHTrbgvA3yPa0AAnxaNpSJXSoQWhjFQrdXyNFFoPeCQWfT0qt/JhNeAebjQV
# r+ufaEkDmedlL4exKqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMwNjU3
# MzZaMC8GCSqGSIb3DQEJBDEiBCDEv0G2OT08oPb+qiR74RQ4KS1HlYqSBWGJ0VHW
# I0Kd1DANBgkqhkiG9w0BAQEFAASCAgB90Z5E2VIQ7oE0GiU+T1ELOK5i9DwEzIxF
# 0OO2B3/YEW5L7gS48FHk7FbFBmk7DiKSNr9phhFffBLIGRvlqIvqH/+FuGlzB9vQ
# UR78xg0ZmqExotrcn0Wi0TAf8jRc0VzK2h1KH/CllqRpb8H6fwHwHCNekL8p34Fh
# NOYhtqveYIpg6TcrSx9ywNG0D5WCyqoRwa6VmHTRyTSfzGrd91pU7hSzIgZg7RBk
# Uf26ddLuT0ey2w1pUq11rVCENrjj9JgvU0Gi16ts+rM0OCj5tZjlWVJAmbpDx74W
# 4nQZzZszDU/nqi1kNUIM64qLkjFR/gBaDKh8lJiIFM9dnRjCwvFBY6r94mRHARzj
# Y2tXizGAdIcRxiZZauQ2f0fvpLjK8skvy94Wk2wUgjROjfIenSUkvNj0DxAePreQ
# z7VgxKjGO8lEzUAf4LIqgeMNwuN2id0iB//WQ2R0EG1it2pVljylipedzI5snsFJ
# JWXV9mgO174aYNnMZxAnhx4JK1i5gNv5MOYyxns5o2dqDxzgqx+BhagTsj9DxXhq
# nnF0Dw6kal9AcHkrYwbcuMv8KdKB4Ij95+C6UB6WeGfNIf8aV8jWtByE7jjPDtrs
# K3btM9K7zjboxUg41lazjovRUsRcqCVDAIOnqfxHbq0yGCG9n7OHhS7TFmHJYbmr
# q1ZqppgsOw==
# SIG # End signature block
