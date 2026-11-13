<#
.SYNOPSIS
Checks the SmartInventory launch catalog and generates the one-line aliases of the legacy launcher names.

.DESCRIPTION
The .cmd launchers were replaced by the catalog (Orchestrator-Jobs.json.template) and the single
entry point SmartInventory.cmd. For one version, every legacy name stays as a one-line alias:
- Launchers\Cloud, Launchers\OnPremises and Launchers\Orchestrator: the LegacyLaunchers of the
  catalog entries, alias "SmartInventory.cmd <Entry>";
- Launchers\Jobs\Start-<Job>.cmd: "SmartInventory.cmd <Job>" (run now on this server);
- LaunchersByOrchestrator\Request-<Job>.cmd: "SmartInventory.cmd <Job> -Request", and
  Request-SmartM365-OrchestratorJob.cmd: "SmartInventory.cmd -Request <Job[,Job...]>".
Job aliases are generated for every non-manual job of the template, as the former launchers were.
The Windows 11 endpoint collector and interactive task uninstaller are standalone launchers,
not generated catalog aliases. They are checked but never rewritten or removed here.

With -Check the script writes nothing: it validates the catalog (scripts, parameters by AST,
execution, includes, tokens) and compares the aliases with the folders; it exits 1 on any
difference (used by the offline tests).

.PARAMETER Check
Validates and compares without writing.

.VERSION
1.2.1
#>
#requires -Version 7.2
[CmdletBinding()]
param([switch]$Check)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$tenantContextPath = Join-Path -Path $PSScriptRoot -ChildPath '..\..\Config\SmartM365-TenantContext.ps1'
if (-not (Test-Path -LiteralPath $tenantContextPath -PathType Leaf)) { throw "SmartM365 tenant context not found: $tenantContextPath" }
. $tenantContextPath
Write-SmartM365StartupBanner
Import-Module -Name (Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365.Inventory.Launch.psm1') -Force -ErrorAction Stop

$smartInventoryRoot = Split-Path -Path $PSScriptRoot -Parent
$catalog = Get-SmartM365InventoryCatalog -SmartInventoryRoot $smartInventoryRoot
$differences = [Collections.Generic.List[string]]::new()
foreach ($catalogError in @(Test-SmartM365InventoryCatalog -Catalog $catalog)) { $differences.Add("catalog: $catalogError") }

function Get-AliasContent {
    param([Parameter(Mandatory)][string]$RelativePath, [Parameter(Mandatory)][string]$Arguments)
    $depth = $RelativePath.Split('\').Count - 1
    return ('@call "%~dp0' + ('..\' * $depth) + 'SmartInventory.cmd" ' + $Arguments + ' %*' + "`r`n")
}

$expected = [ordered]@{}
foreach ($entry in @($catalog.Entries)) {
    $legacy = $entry.Definition.PSObject.Properties['LegacyLaunchers']
    if ($null -eq $legacy) { continue }
    foreach ($launcher in @($legacy.Value)) {
        if ($launcher -notmatch '^Launchers\\(Cloud|OnPremises|Orchestrator)\\[A-Za-z0-9._-]+\.cmd$') { $differences.Add("catalog: $($entry.Name): unexpected legacy launcher $launcher"); continue }
        $expected[[string]$launcher] = Get-AliasContent -RelativePath $launcher -Arguments $entry.Name
    }
}
foreach ($entry in @($catalog.Entries | Where-Object { $_.Kind -eq 'Job' } | Sort-Object Name)) {
    $mode = if ($entry.Definition.PSObject.Properties['AssignmentMode'] -and $entry.Definition.AssignmentMode) { [string]$entry.Definition.AssignmentMode } else { 'Legacy' }
    # Same job set as the former generated launchers; the legacy CMDB preparation never had one.
    if ($mode -eq 'Manual' -or $entry.Name -eq 'CmdbEvidence-Prepare') { continue }
    $expected["Launchers\Jobs\Start-$($entry.Name).cmd"] = Get-AliasContent -RelativePath "Launchers\Jobs\Start-$($entry.Name).cmd" -Arguments $entry.Name
    $expected["LaunchersByOrchestrator\Request-$($entry.Name).cmd"] = Get-AliasContent -RelativePath "LaunchersByOrchestrator\Request-$($entry.Name).cmd" -Arguments "$($entry.Name) -Request"
}
$expected['LaunchersByOrchestrator\Request-SmartM365-OrchestratorJob.cmd'] = Get-AliasContent -RelativePath 'LaunchersByOrchestrator\Request-SmartM365-OrchestratorJob.cmd' -Arguments '-Request'

# These entry points intentionally bypass the catalog: the first is an endpoint-side
# collector, the second invokes an interactive Windows PowerShell 5.1 uninstaller.
# Treat them as owned source files, not as obsolete generated aliases.
$standalone = [ordered]@{
    'Launchers\OnPremises\Start-SmartM365-Windows11-HardwareReadiness-Inventory.cmd' = @{
        Script = 'EndpointInventory\Windows11Readiness\SmartM365-Endpoint-Windows11Readiness-Inventory.ps1'
        Pattern = 'SmartM365-Endpoint-Windows11Readiness-Inventory\.ps1"\s+%\*'
    }
    'Launchers\Orchestrator\Start-SmartM365-Inventory-OrchestratorScheduledTask-Uninstaller.cmd' = @{
        Script = 'Orchestrator\Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1'
        Pattern = 'Install-SmartM365-Inventory-OrchestratorScheduledTask\.ps1"\s+-UninstallInteractive\s+%\*'
    }
}
foreach ($relative in $standalone.Keys) {
    $path = Join-Path $smartInventoryRoot $relative
    $target = Join-Path $smartInventoryRoot $standalone[$relative].Script
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { $differences.Add("standalone missing: $relative") }
    elseif ([IO.File]::ReadAllText($path) -notmatch $standalone[$relative].Pattern) { $differences.Add("standalone invalid: $relative") }
    if (-not (Test-Path -LiteralPath $target -PathType Leaf)) { $differences.Add("standalone target missing: $($standalone[$relative].Script)") }
}
# Refuse any generation when these source launchers or their targets are invalid.
if (-not $Check -and @($differences | Where-Object { $_ -like 'standalone *' }).Count -gt 0) {
    $differences | Where-Object { $_ -like 'standalone *' } | ForEach-Object { Write-Output $_ }
    exit 1
}

$existing = @()
foreach ($folder in @('Launchers', 'LaunchersByOrchestrator')) {
    $path = Join-Path $smartInventoryRoot $folder
    if (Test-Path -LiteralPath $path) { $existing += @(Get-ChildItem -LiteralPath $path -Filter '*.cmd' -File -Recurse | ForEach-Object { $_.FullName.Substring($smartInventoryRoot.Length + 1) }) }
}
foreach ($relative in $expected.Keys) {
    $path = Join-Path $smartInventoryRoot $relative
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { $differences.Add("missing: $relative") }
    elseif ([IO.File]::ReadAllText($path) -cne $expected[$relative]) { $differences.Add("outdated: $relative") }
    if (-not $Check) {
        New-Item -ItemType Directory -Path (Split-Path $path -Parent) -Force | Out-Null
        [IO.File]::WriteAllText($path, $expected[$relative], [Text.Encoding]::ASCII)
    }
}
foreach ($relative in @($existing | Where-Object { -not $expected.Contains($_) -and -not $standalone.Contains($_) })) {
    $differences.Add("obsolete: $relative")
    if (-not $Check) { Remove-Item -LiteralPath (Join-Path $smartInventoryRoot $relative) -Force }
}

if ($Check) {
    if ($differences.Count -gt 0) {
        $differences | ForEach-Object { Write-Output $_ }
        Write-Output ("SmartInventory catalog or legacy aliases are out of date ({0} difference(s)). Run New-SmartM365-OrchestratorJobLaunchers.ps1." -f $differences.Count)
        exit 1
    }
    Write-Output ("SmartInventory catalog valid; legacy aliases up to date: {0} alias(es)." -f $expected.Count)
    exit 0
}
$catalogErrors = @($differences | Where-Object { $_ -like 'catalog:*' })
if ($catalogErrors.Count -gt 0) { $catalogErrors | ForEach-Object { Write-Output $_ }; exit 1 }
Write-Output ("Legacy aliases generated: {0} alias(es); {1} change(s)." -f $expected.Count, $differences.Count)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCeiFkqs/Btl2EH
# CQl+m1DHNCUo//R0zj3NNDV8fFRBQqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIONr1IA+Zo0hAlp0KCXa801kmGGJH3stvhNl4HeBXMj+MA0GCSqG
# SIb3DQEBAQUABIIBgEGlPK4YPQ2QpS5nLv5muYnWDA66P6fyf/68bkApYMlIJ6jZ
# 6O1cfxCgCD7npZaQMTNnVzu1pBTTLCj+mMsauCNfZAGerV30bKY/cR9CcWxpGUDt
# OoD/D/izXQYQd1or79qIFe2AFhDpfGHhrWgOxWTqTTdOluiTZtdiFuysc7wWWriG
# 2lxOjjKflelgi+AyqzzQUAN1UoA/ITdQMfz8ePSG5sILHxDg9BlJjCOQqnY7uMoc
# 3jUbXG3SfGsZ/orjzLx9VVCG4xYacgy2Ddtl86kAIqiLgurK1LOc8f9dBQBQw8Mr
# KgNbU4RYGoVNzCIWjmpKbMMcXdLHLqGFP2k568rMH5Iop57xViIYJ3esv3w6XrP+
# IFRI4axYPgmyVQpj2twLenTslQHirV1gUVUZSsD2P0qPjVBuAvLRF6b9ce9O5FR4
# NZzhBeVswFOndBBtmK0x1aFrGpHDrI+KFHOeEWtoGlU8MCOyKGH3FS/kJNcADdi6
# rb2jbOlHK30iwOJ4uaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTMxMzE2
# MThaMC8GCSqGSIb3DQEJBDEiBCCeGqPkQqbi+rta17FjIfx0RhfnINkiCG1UgIJg
# UuoMXDANBgkqhkiG9w0BAQEFAASCAgCXJWK8kv144aWAt85J/anKocISD6c+7hyo
# XTmyj0EDErpU2f+ZfZy4LJQywyZ3kLV5/NhJDFZtQjaS3iea+myeuLiqhefI7iV7
# dA85PskHu56UEjue9d87kgcegrl6o1VZB+jz2QvVDFhS6jZM3vbeAzRISxg3qrH+
# 0sEeApc8jx0mES3kXuDdQH+1ht22T1h3QHFeNBi1hg8Xjiu4I/ghZiAd0rw5ZQ2L
# fJjlrQdQzz/IELEXbPHmyukEZW218BZrPApCEJOkNJ4DP4YKJhLFfF06pY7xaWuA
# a2Zu272HOLDL3MNArsHanLTuQQ3Y+KagIcCbOwrkscEGP6q3Bh76ZFQF1osECRmS
# UjIugRLarh5vntQ324DPyeS+ruf1EARUF+NXO44t8iMtAume5nqXeDgzESKPXlD9
# USZ2aM7jS6r7+4yCjTlNpk/3O/G+ZwxdgMlNJ9FkMk74zOR3loatAl6gjaWPXibh
# RBaWLnx8K8lyisnooOMvMBvwUSriSbQdYs9lPoVnvNBN0aJwCzhY9s7db6pRVKRq
# 196VqtZxV6lBhvc2i21XUF1x7fQ+0wfWhZTRrQE+g5FtJryvr+Elu6rvBcrG9zIg
# zTZ8nsAC5WAXXBoqa7rdpz34ULKOH3TqHKZeXcak6TiI9JqVdhTT5ctND80KMwbO
# +9QZoy5dhA==
# SIG # End signature block
