#Requires -Version 7.0
<#
.SYNOPSIS
Synthetic checks of staged DATA-LAST publication for the Exchange hybrid identity issues: with the
previous run's different CSVs in DATA-LAST, a running or failed run neither changes nor uploads them,
and a complete run promotes this run's CSVs with the receipt, receipt last.
.VERSION
1.0.0
.NOTES
Synthetic TEMP data only. SharePoint upload and WeeklyHistory are replaced by recorders.
No directory, Exchange, tenant, API, mail or synchronized data access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic Core identity globals are saved and restored in finally.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','',Justification='Script variables are read by the extracted collector functions.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.95' -Force -PassThru
$producer = Get-ChildItem (Join-Path $root 'SmartInventory') -Recurse -File -Filter 'SmartM365-Exchange-HybridIdentity-Issues-Inventory.ps1' | Select-Object -First 1
$ast = [Management.Automation.Language.Parser]::ParseFile($producer.FullName, [ref]$null, [ref]$null)
foreach ($name in @('FlattenPath', 'CopyCsv', 'PublishLatest', 'Remove-DetailedArchiveFilesOlderThan', 'ExportIssues')) {
    $definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    if (-not $definition) { throw "$name is missing." }
    . ([scriptblock]::Create($definition.Extent.Text))
}
function Log([string]$m) { $null = $m }
function Warn([string]$m) { $null = $m }
$files = @('Exchange_HybridIdentity_Issues.csv', 'Exchange_HybridIdentity_Issues_Summary.csv')
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-HybridStaged-' + [guid]::NewGuid().ToString('N'))
$LatestFolder = Join-Path $temporary 'DATA-LAST'
$OutputFolder = Join-Path $temporary 'DATA-ALL/Exchange/Issues/HybridIdentity'
$DetailedArchiveRetentionDays = 7
New-Item -ItemType Directory -Path $LatestFolder, $OutputFolder | Out-Null
$saved = @{}
foreach ($name in @('SmartM365TenantKey', 'SmartM365OrganizationKey', 'SmartM365EnvironmentKey', 'SmartM365TenantId', 'csvGeneratedPaths')) {
    $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $saved[$name] = @{ Exists = ($null -ne $variable); Value = $(if ($variable) { $variable.Value } else { $null }) }
}
$global:SmartM365TenantKey = 'synthetic'; $global:SmartM365OrganizationKey = 'contoso'
$global:SmartM365EnvironmentKey = 'test'; $global:SmartM365TenantId = 'synthetic-tenant'
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
$receiptPath = Join-Path $LatestFolder 'SmartInventory_SmartM365-Exchange-HybridIdentity-Issues-Inventory.current.json.txt'
$runPath = Join-Path $LatestFolder 'SmartInventory_SmartM365-Exchange-HybridIdentity-Issues-Inventory.run.json.txt'
$stagingProducer = Join-Path $LatestFolder '.staging/SmartM365-Exchange-HybridIdentity-Issues-Inventory'
function Get-LatestHashTable { $result = @{}; foreach ($file in $files) { $path = Join-Path $LatestFolder $file; $result[$file] = if (Test-Path -LiteralPath $path) { (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash } else { '' } }; return $result }
function Read-Json([string]$Path) { Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json }
function Get-IssueRowSet([string]$Label, [int]$Count) {
    @(1..$Count | ForEach-Object { [pscustomobject]@{ TenantKey = 'synthetic'; IssueNumber = $_; ObjectGUID = [guid]::NewGuid().ToString(); Potential_Issue = "$Label issue $_"; IssueCategory = '3.Medium'; RecommendedAction = 'Review' } })
}
function Open-HybridRun([switch]$Staged) {
    $global:csvGeneratedPaths = $null
    $script:RunStamp = (Get-Date).ToString('yyyyMMdd-HHmmss') + [guid]::NewGuid().ToString('N').Substring(0, 4)
    Start-SmartM365SourceReceipt -ScriptPath $producer.FullName -SourceRootPath $LatestFolder
    if ($Staged) { Enable-SmartM365StagedSourcePublication }
}
try {
    Initialize-SmartM365DefaultCsvValidationRules
    # Static wiring: staging right after the receipt, inputs read from other producers only, one version.
    $text = [IO.File]::ReadAllText($producer.FullName)
    Check ($text -match "(?m)^\s*Start-SmartM365SourceReceipt -ScriptPath \`$PSCommandPath -SourceRootPath \`$LatestFolder\r?\n(\s*#[^\r\n]*\r?\n)*\s*Enable-SmartM365StagedSourcePublication") 'The collector does not enable staging right after its receipt.'
    $inputs = @([regex]::Matches($text, "Csv '([^']+)'") | ForEach-Object { $_.Groups[1].Value })
    Check ($inputs.Count -ge 10 -and @($inputs | Where-Object { $_ -in $files }).Count -eq 0) 'The collector reads its own DATA-LAST output during the run.'
    Check (@($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'CopyCsv' -and $n.Extent.Text -match '\$(latest|sl)\b' }, $true)).Count -eq 0) 'A DATA-LAST copy bypasses the staged helper.'
    $version = [regex]::Match($text, '\.VERSION\r?\n\s*([0-9.]+)').Groups[1].Value
    Check ($version -and $text -match ('\$ScriptVersion="' + [regex]::Escape($version) + '"')) 'Header and runtime versions differ.'
    Check ($text -match "-MinimumVersion '1\.0\.(?:9[5-9]|[1-9]\d\d)'") 'The collector does not require the staging Core.'

    & $core {
        $script:Recorded = New-Object 'Collections.Generic.List[string]'
        function script:Invoke-SmartM365SharePointCsvUpload { param([string]$LocalFilePath) $script:Recorded.Add($LocalFilePath) }
        function script:WriteLog { param($Message, $Level) $null = $Message, $Level }
    }
    function Get-DeferredUploadList { & $core { $script:Recorded.ToArray() } }

    # Previous run, published directly: DATA-LAST holds yesterday's issues and receipt.
    Open-HybridRun
    $null = ExportIssues -rows (Get-IssueRowSet 'yesterday' 3)
    Complete-SmartM365SourceReceipt -Status 'Completed' | Out-Null
    $yesterdayReceipt = Read-Json $receiptPath
    $yesterday = Get-LatestHashTable
    Check ($yesterdayReceipt.Status -eq 'Completed' -and @($yesterdayReceipt.Files).Count -eq 2) 'The previous hybrid identity run did not publish.'

    # Running run with different issues: DATA-LAST and the receipt keep yesterday's version, and the
    # upload list of the run never names a DATA-LAST file (it would upload yesterday's copy).
    Open-HybridRun -Staged
    $runFiles = @(ExportIssues -rows (Get-IssueRowSet 'today' 5))
    $during = Get-LatestHashTable
    Check (@($files | Where-Object { $during[$_] -ne $yesterday[$_] }).Count -eq 0) 'A running hybrid identity run changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $yesterdayReceipt.RunId) 'A running hybrid identity run changed the receipt.'
    Check (@($runFiles | Where-Object { $_ -like "$LatestFolder*" }).Count -eq 0) 'The run uploads DATA-LAST files before their promotion.'
    Check (@($runFiles | Where-Object { $_ -like "$OutputFolder*" }).Count -eq 3) 'The run lost its DATA-ALL outputs.'
    Check (@(Get-DeferredUploadList).Count -eq 0) 'A staged copy was uploaded before promotion.'

    # Failed run: nothing changes, staging is removed.
    Complete-SmartM365SourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null
    $after = Get-LatestHashTable
    Check (@($files | Where-Object { $after[$_] -ne $yesterday[$_] }).Count -eq 0) 'A failed hybrid identity run changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $yesterdayReceipt.RunId) 'A failed hybrid identity run replaced the receipt.'
    Check (@(Get-ChildItem -LiteralPath $stagingProducer -Force).Count -eq 0) 'A failed hybrid identity run left its staging folder.'
    Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof (Read-Json $receiptPath)
    $script:checks++

    # Complete run: this run's CSVs replace yesterday's together with the receipt; uploads follow.
    Open-HybridRun -Staged
    $null = ExportIssues -rows (Get-IssueRowSet 'today' 5)
    Complete-SmartM365SourceReceipt -Status 'Completed' | Out-Null
    $receipt = Read-Json $receiptPath
    $final = Get-LatestHashTable
    Check ($receipt.Status -eq 'Completed' -and $receipt.RunId -ne $yesterdayReceipt.RunId) 'A complete hybrid identity run did not publish its receipt.'
    foreach ($file in $files) {
        $entry = @($receipt.Files | Where-Object File -eq $file)
        Check ($final[$file] -ne $yesterday[$file] -and $entry.Count -eq 1 -and $entry[0].SHA256 -eq $final[$file]) "Promoted hybrid identity CSV and receipt disagree: $file"
    }
    Check (@(Import-Csv -LiteralPath (Join-Path $LatestFolder 'Exchange_HybridIdentity_Issues.csv') | Where-Object { $_.Potential_Issue -like 'today*' }).Count -eq 5) 'DATA-LAST does not hold this run issues.'
    $deferred = @(Get-DeferredUploadList)
    Check ($deferred.Count -eq 2 -and @($deferred | Where-Object { $_ -like '*.staging*' -or $_ -notlike "$LatestFolder*" }).Count -eq 0) 'Promoted copies were not uploaded by canonical DATA-LAST path.'
    $run = Read-Json $runPath
    Check ($run.Status -eq 'Completed' -and -not $run.UnqualifiedPublication -and $run.StagedPublication) 'Completed hybrid identity run state is not qualified.'
    Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof $receipt
    $script:checks++
}
finally {
    try { Complete-SmartM365SourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null } catch { Write-Verbose $_.Exception.Message }
    foreach ($name in $saved.Keys) {
        if ($saved[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $saved[$name].Value }
        else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
    }
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}
"PASS: $script:checks hybrid identity staged publication checks; PowerShell $($PSVersionTable.PSVersion)."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBkBqxVBcycTgKj
# 4fddKq0sS6fn+SAcPVo7KewOsClmhKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIADz/Ns8Gl+aQdl94L2PqrmzoIOu6o5p78m9l6XwnfkmMA0GCSqG
# SIb3DQEBAQUABIIBgGkMd8Xo6jFK15xura9b7Cj4tTLJ7re92DvODBSP8M4eQ9YS
# HMnb3aNOC9yRS9/zRV5VKatPYOHf1ClxPW4u7BPKGl3xAumwn5d3fy3v/MZsF1Ra
# NZPPynQR837FejiTuNVS9W82ozrS+GJp2qm1cNv/cpnxR3UQ+OTX/4vSvKJ9Nm95
# 0dVMUK4VSIGeI9mVij5pM6C8W9wWVmlhbCMlHjnD/1H6LeyjzjorgIC8mFCA+KYi
# LxOgQudcNpeVXjSc5kYPA90phnljiN7Mp642NBB1Z492O0ZYCN010QyB8IPLdyeN
# rdau1uowXtWEvrH5iHFeuG1rgdqkEhitlSBBke1aXdQnJdyffu7fcAz5g2b8xkWX
# C/0qBSkEaxaf57uxiXdwVZgst8kLgTTEmiAFNrUyF+6B03kxuwrWw+KnZnDvRAYK
# Hb6redc7xWKiIIEbWGxqWIlP6Eo4OgWSWDfsBHx8O2oapYGxoyYv24e9a/pgs6Ag
# UGXxx5tceaR0iiMK9aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjEwNjUx
# MDRaMC8GCSqGSIb3DQEJBDEiBCCK3gxyHS4ugIQmIJoCWxD3/0PJNQanRwHG2dde
# F/gsgTANBgkqhkiG9w0BAQEFAASCAgCtZHK0+p0ATsrOZZgmCC9K2ZYdMDTV/3OS
# nAaq49MliiYlu75lD+umz2rOgyQzMg4c3naGtQmpa/ZNHEnw75LyaJnvjW609fnw
# kkVms0iAY9sDw1t/G8HEHsH7dfMlWtKsFTtyFqC+qHN7t6AHDrpo+LmJ2Hr2y6FH
# iI1Gj4Ed65fnLnrmCd3FiusMJ+CyoOvBbIz4gEiCfr3RXLaYztJQlngOEK+UeaX8
# QllBuNnYus2B49vYJJ1J91iuC23G42FCldpscwI/UefXhj1wKavkz5UbZANFGMZw
# RTpwTKyyHMpV/GY2bh+7/aFioXCDrJGq7Py9TJq/lR+5wg540AjzrArGqu6mpgPD
# L66ox/ZNljskDOnsvv7UX7Rao6soBAo0cepbtSUaLG11qgLaJn/vTEb5X45RknZK
# K14tD+JQDBEyRV5ex3HLznUMx0rftH+mASMCEcFy4kVdHi1xowz6rcgOD6QTMKfZ
# /UqDIgZQ+dBZH4iz0mexEhWd/5YobZ0keppfAfuTV9mbR3uLe+ElAzD5o++T1J0k
# Mh/Ni4oV+P56MflbSFr7p5XMYyKMWBSs7ne5xLEghMjx51k2cj/puejnoOl/8zgI
# BXrHAnbEaLpTyv+fohwENtInqTqUig6hK/nYmM1Rx6Qb6NLSGWV5etHJcYNAe8DE
# PjJfuIQTAA==
# SIG # End signature block
