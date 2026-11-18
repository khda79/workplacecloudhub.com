<#
.SYNOPSIS
Offline checks of the orchestrator SharePoint token pause (orchestrator 1.5.95): when the Graph token
cannot be obtained, the next uploads are skipped for SharePointTokenPauseMinutes instead of waiting
for the same token retries again, then resume. The token check calls a SmartM365.Core function that the
module does not export, so it must run in the module scope; a failing check must never stop the
orchestrator.

.VERSION
1.1.0
#>
#requires -Version 7.2
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$checks = 0
function Check([bool]$Condition, [string]$Message) { if (-not $Condition) { throw "FAIL: $Message" }; $script:checks++ }

$orchestratorPath = Join-Path $PSScriptRoot '..\SmartInventory\Orchestrator\SmartM365-Inventory-Orchestrator.ps1'
$corePath = Join-Path $PSScriptRoot '..\Modules\SmartM365.Core\SmartM365.Core.psd1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Resolve-Path $orchestratorPath), [ref]$tokens, [ref]$errors)
Check (@($errors).Count -eq 0) 'Orchestrator does not parse.'
$text = Get-Content -LiteralPath $orchestratorPath -Raw
Check ($text -match '\.VERSION\r?\n1\.5\.96\r?\n') 'Orchestrator .VERSION is not 1.5.96.'
Check ($text -match '\$ScriptVersion = "1\.5\.96"') 'Orchestrator ScriptVersion is not 1.5.96.'
Check ($text -match '\$script:SharePointTokenPauseMinutes = 15') 'Token pause default is not 15 minutes.'

# The real Core keeps the token helper internal: a direct call from the orchestrator fails (1.5.94 bug).
$realCore = Import-Module $corePath -Force -PassThru -DisableNameChecking -WarningAction SilentlyContinue
try {
    Check ($null -eq (Get-Command Connect-SmartM365GraphForSharePointUpload -ErrorAction SilentlyContinue)) 'Connect-SmartM365GraphForSharePointUpload is exported; the module-scope call is no longer needed.'
    Check ($null -ne (& $realCore { Get-Command Connect-SmartM365GraphForSharePointUpload -ErrorAction SilentlyContinue })) 'Connect-SmartM365GraphForSharePointUpload is missing from the SmartM365.Core module scope.'
}
finally { Remove-Module $realCore -Force -ErrorAction SilentlyContinue }

$node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-OrchestratorSharePointUpload' }, $true)
Check ($null -ne $node) 'Invoke-OrchestratorSharePointUpload not found.'

# Fake SmartM365.Core: the token helper is internal, as in the real module.
$global:TpConnectCalls = 0
$global:TpTokenAvailable = $false
$global:TpThrow = $false
$fakeCore = New-Module -Name SmartM365.Core -ScriptBlock {
    function Connect-SmartM365GraphForSharePointUpload {
        param($AppId, $TenantId, $Thumbprint)
        $global:TpConnectCalls++
        if ($global:TpThrow) { throw 'synthetic pre-check failure' }
        return $global:TpTokenAvailable
    }
    function Get-TpCoreMarker { 'fake-core' }
    Export-ModuleMember -Function Get-TpCoreMarker
} | Import-Module -PassThru -Force

$stubs = @'
$script:Log = [Collections.Generic.List[string]]::new()
$script:UploadCalls = 0
$script:SharePointUploadedFileState = @{}
$script:SharePointTokenPausedUntilUtc = [datetime]::MinValue
$script:SharePointTokenPauseMinutes = 15
$script:SharePointTokenPausedSkips = 0
$script:SharePointTokenPrecheckWarned = $false
$script:Settings = [pscustomobject]@{
    SharedDataFolderPath = 'C:\none'; OrchestratorLogFolderPath = 'C:\none'
    SharePointSiteHostname = 'tenant.sharepoint.test'; SharePointSitePath = '/sites/x'
    SharePointLibraryDisplayName = 'Documents'; SharePointTargetFolderPath = 'X'
}
function Test-OrchestratorUploadCategoryEnabled { param($Category) $true }
function Test-OrchestratorSharePointUploadConfigured { $true }
function Get-SmartM365SharePointUploadCategory { param($Path, $DataAllRoot, $LogAllRoot) 'DataAll' }
function Write-OrchestratorLog { param($Message, $Level = 'INFO') $script:Log.Add("[$Level] $Message") }
function Invoke-SmartM365SharePointCsvUpload {
    [CmdletBinding()] param($LocalFilePath, $Enabled, $SiteHostname, $SitePath, $LibraryDisplayName, $TargetFolderPath, [switch]$EnsureParentFolders)
    $script:UploadCalls++
    [pscustomobject]@{ WebUrl = 'https://tenant.sharepoint.test/' + [IO.Path]::GetFileName($LocalFilePath); SharePointPath = '' }
}
'@
$module = New-Module -ScriptBlock ([scriptblock]::Create($stubs + [Environment]::NewLine + $node.Extent.Text))
function Get-TpState { & $module { [pscustomobject]@{ Connect = $global:TpConnectCalls; Upload = $script:UploadCalls; Skips = $script:SharePointTokenPausedSkips; Until = $script:SharePointTokenPausedUntilUtc; Log = @($script:Log) } } }
function Invoke-TpBatch([string[]]$Paths) { & $module { param($Paths) foreach ($p in $Paths) { Invoke-OrchestratorSharePointUpload -LocalFilePath $p -Reason 'periodic orchestrator upload' } } $Paths }

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-TokenPause-' + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $root -Force | Out-Null
    $files = foreach ($name in 'a.log', 'b.csv', 'c.json', 'd.json') {
        $path = Join-Path $root $name; Set-Content -LiteralPath $path -Value $name; $path
    }

    # 1. Token unavailable: one token attempt for the whole batch, no upload, then a pause.
    $results = Invoke-TpBatch $files
    $state = Get-TpState
    Check (@($results | Where-Object { $_ }).Count -eq 0) 'An upload reported success without a token.'
    Check ($state.Connect -eq 1) "Token requested $($state.Connect) times for one batch; expected 1 (through the module scope)."
    Check ($state.Upload -eq 0) 'Upload attempted without a token.'
    Check ($state.Skips -eq 3) "Expected 3 skipped uploads, got $($state.Skips)."
    Check ($state.Until -gt [datetime]::UtcNow.AddMinutes(14) -and $state.Until -le [datetime]::UtcNow.AddMinutes(15)) 'Pause is not about 15 minutes.'
    Check (@($state.Log | Where-Object { $_.StartsWith('[WARN] SharePoint uploads paused for 15 minutes') -and $_.Contains('a.log') }).Count -eq 1) 'Pause warning missing or not naming the first file.'

    # 2. Token back but pause not expired: still skipped, no token request.
    $global:TpTokenAvailable = $true
    $null = & $module { param($p) Invoke-OrchestratorSharePointUpload -LocalFilePath $p -Reason 'job-runs CSV' -Force } $files[0]
    $state = Get-TpState
    Check ($state.Connect -eq 1 -and $state.Upload -eq 0 -and $state.Skips -eq 4) 'An upload ran during the pause.'

    # 3. Pause expired and token available: uploads resume and the skips are reported once.
    & $module { $script:SharePointTokenPausedUntilUtc = [datetime]::UtcNow.AddSeconds(-1) }
    $results = Invoke-TpBatch $files
    $state = Get-TpState
    Check (@($results | Where-Object { $_ }).Count -eq 4) 'Uploads did not resume after the pause.'
    Check ($state.Upload -eq 4) "Expected 4 uploads after the pause, got $($state.Upload)."
    Check ($state.Skips -eq 0) 'Skip counter not reset after resume.'
    Check (@($state.Log | Where-Object { $_.Contains('SharePoint uploads resumed') -and $_.Contains('4 upload(s) were skipped') }).Count -eq 1) 'Resume message missing or wrong count.'

    # 4. Unchanged files are not re-uploaded and do not request a token.
    $before = $state.Connect
    $null = Invoke-TpBatch $files
    $state = Get-TpState
    Check ($state.Connect -eq $before -and $state.Upload -eq 4) 'Unchanged files triggered a token request or an upload.'

    # 5. The pre-check itself fails: no exception reaches the caller, uploads run, one warning only.
    $global:TpThrow = $true
    foreach ($f in $files) { Add-Content -LiteralPath $f -Value 'changed' }
    $results = Invoke-TpBatch $files
    $state = Get-TpState
    Check (@($results | Where-Object { $_ }).Count -eq 4) 'A failing token pre-check blocked the uploads.'
    Check ($state.Upload -eq 8) "Expected 8 uploads in total, got $($state.Upload)."
    Check (@($state.Log | Where-Object { $_.StartsWith('[WARN] SharePoint token pre-check unavailable') }).Count -eq 1) 'Pre-check failure not logged exactly once.'
    Check ($state.Until -lt [datetime]::UtcNow) 'A pre-check failure paused the uploads.'

    # 6. No SmartM365.Core loaded: uploads run as before.
    $global:TpThrow = $false
    Remove-Module $fakeCore -Force
    foreach ($f in $files) { Add-Content -LiteralPath $f -Value 'changed again' }
    $results = Invoke-TpBatch $files
    $state = Get-TpState
    Check (@($results | Where-Object { $_ }).Count -eq 4 -and $state.Upload -eq 12) 'Uploads did not run without SmartM365.Core loaded.'
}
finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Module -Name SmartM365.Core -Force -ErrorAction SilentlyContinue
    Remove-Variable -Name TpConnectCalls, TpTokenAvailable, TpThrow -Scope Global -ErrorAction SilentlyContinue
}

Write-Host "Orchestrator SharePoint token pause offline: $checks checks passed."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCQ4Ck3LNnjP86L
# ernnd6J1gN7kaXqT0Trwuqd41tHZnKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIA4rYKYnv8Oz2IHmIs01laIUaAshTpPnmDJIt0hpNEtlMA0GCSqG
# SIb3DQEBAQUABIIBgBODYkoAmy8dGg4Y1aadb1Y3frq1H2tlx61uCAZTpx5PZ4QD
# S4KpquNxBjqyuU20cowyZeob+QhkQJhkS2adGGLgsUbtZRnM5JG/2UCRSgo4jMa8
# Fhp05tOcWY5wVSf/gvAXAqppRTUL7zkPDpYnjICkRGBw6h13vbzEDgnYVDqPhSwl
# /86lmg5EiHsjGCJS5KUGR7abrAds8cGSWpl4G7zEvWASM8Cm3hlc/V2rWASYPc/T
# 3bB5BE8RIEqAjjOig6OlvYKlqRdBaVepCcqi0pl6Kr8epv47OaxBo+I1CcD2kmCd
# hf7uAvCFEk2iltttArRGHpP/AYFlDmvSu+GPoNVcgTsRR8I2aViFMwBIjJCgu2ea
# z/sWSxmj0R4pWK1Y3s5NdwFzAr3ibAEUc7ea5yy6aVamHf5W5lvKxRo0kxhPM/sm
# y6mN2dedFznfx49CmAxY7z87dx9y39b2Ylo80ihzCwAPeMxq8+Wz5XzLqgqJCXFg
# 0rdMIHtYbILvq2yDvaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTcxMjQz
# MzNaMC8GCSqGSIb3DQEJBDEiBCDzs3ad6f9uD+HFmwLs49viTB9psE69uqSgD1Y6
# OkPvkjANBgkqhkiG9w0BAQEFAASCAgCj1ywTxYxd/Z9LQgLsPcyjMV/ywu33RbqK
# K6oPNaSg4FZ4XRn+te8xYXrzlWvwxdP56obOuDQtFwCF98+lPsB3cBFYnnlEEJHA
# 02xhQ6j71I6unDMgw7wUqJExwsNdHXGVRkQpy0eVX8J9AbwhkzuLsXbs88qa9aXp
# E/jgQ67lbHGYb1BbK/PfnZBD2x3ZvMr7EtJywlQkbwauAok9zGlBOiZcS4LfUgJu
# n8PSvBstbTxjMGQaLCAz7B30X2ltgfX/xyvKTB8fGEjMExg6wd0B5gDJPvoRQjxm
# GrSU23g3zEzxdtm08efHqVThn45jy+cwB9e+Hi9Pk8dlOWufkjs6D7g6eUTGuDz2
# najQE1LAb1ysjb1fFSXWuaj7qDE2e2p9NUoCkT1EvRkTSGXJw2fyXS5hhNOm7Bdb
# FJiXgJFPjSi+Z/c4IqN41R187dpcJF4pzWtWtJBXLREIWC3+N81/3UNrJ1BvA4a7
# rtds2iBiwZ7p8OgR+EUEcr7/JZ5dniHZOPi5zru+SkjCuv2B5ZejOWAoAApHByOv
# D7N+xbMm4dQwhj9w8To7YrZcAVikf3uWE95vU0yLOW9yuzi8fAeQKW001dahnPYr
# AmQtb4DVPKFwEcIhZOivfij3x+4SJWdnie3fbchNxI4eLXtla8cys81j0MLzedTF
# qgy5wfa/lQ==
# SIG # End signature block
