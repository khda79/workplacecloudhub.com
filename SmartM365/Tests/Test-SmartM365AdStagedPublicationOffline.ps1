#Requires -Version 7.0
<#
.SYNOPSIS
Synthetic checks of staged DATA-LAST publication for the Active Directory inventory: the combined
CSVs, duplicate-identity CSVs and daily copies reach DATA-LAST with the receipt, never before.
.VERSION
1.0.0
.NOTES
Synthetic TEMP data only. SharePoint upload and WeeklyHistory are replaced by recorders.
No directory, tenant, API, mail or synchronized data access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic Core identity globals are saved and restored in finally.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.93' -Force -PassThru
$producer = Get-ChildItem (Join-Path $root 'SmartInventory') -Recurse -File -Filter 'SmartM365-ActiveDirectory-Inventory.ps1' | Select-Object -First 1
$ast = [Management.Automation.Language.Parser]::ParseFile($producer.FullName, [ref]$null, [ref]$null)
$definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Publish-SmartM365AdLatestCopy' }, $true)
if (-not $definition) { throw 'Publish-SmartM365AdLatestCopy is missing.' }
. ([scriptblock]::Create($definition.Extent.Text))
$required = @('AD_Users_AllDomains.csv', 'AD_Computers_AllDomains.csv', 'AD_Domains_AllDomains.csv', 'AD_Groups_AllDomains.csv', 'AD_DirectoryObjects_AllDomains.csv', 'AD_GroupMemberships_AllDomains.csv')
$extra = @('AD_Users_DuplicateUPN.csv', 'AD_Users_DailyReport.csv')
$files = @($required + $extra)
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-AdStaged-' + [guid]::NewGuid().ToString('N'))
$latest = Join-Path $temporary 'DATA-LAST'
$output = Join-Path $temporary 'DATA-ALL/ActiveDirectory/Inventory'
New-Item -ItemType Directory -Path $latest, $output | Out-Null
$saved = @{}
foreach ($name in @('SmartM365TenantKey', 'SmartM365OrganizationKey', 'SmartM365EnvironmentKey', 'SmartM365TenantId', 'csvGeneratedPaths')) {
    $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $saved[$name] = @{ Exists = ($null -ne $variable); Value = $(if ($variable) { $variable.Value } else { $null }) }
}
$global:SmartM365TenantKey = 'synthetic'; $global:SmartM365OrganizationKey = 'contoso'
$global:SmartM365EnvironmentKey = 'test'; $global:SmartM365TenantId = 'synthetic-tenant'
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
$script:Uploads = New-Object 'Collections.Generic.List[string]'
function Invoke-SmartM365SharePointCsvUpload { param([string]$LocalFilePath) $script:Uploads.Add($LocalFilePath); return [pscustomobject]@{ LocalFilePath = $LocalFilePath } }
$receiptPath = Join-Path $latest 'SmartInventory_SmartM365-ActiveDirectory-Inventory.current.json.txt'
$runPath = Join-Path $latest 'SmartInventory_SmartM365-ActiveDirectory-Inventory.run.json.txt'
$stagingProducer = Join-Path $latest '.staging/SmartM365-ActiveDirectory-Inventory'
$scope = 'CMDB:ad_users,ad_computers,ad_domains,ad_groups,ad_objects,ad_members'
function Get-LatestHashTable { $result = @{}; foreach ($file in $files) { $path = Join-Path $latest $file; $result[$file] = if (Test-Path -LiteralPath $path) { (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash } else { '' } }; return $result }
function Read-Json([string]$Path) { Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json }
# Same publication as the collector: combined and report CSVs in DATA-ALL, then their DATA-LAST copies.
function Publish-AdSet([string]$Label) {
    foreach ($file in $files) {
        $source = Join-Path $output $file
        [IO.File]::WriteAllText($source, "TenantKey,DomainName,Name`r`nsynthetic,contoso.test,$Label-$file`r`n", [Text.UTF8Encoding]::new($false))
        Publish-SmartM365AdLatestCopy -SourcePath $source -DestinationPath (Join-Path $latest $file) | Out-Null
        [void]$global:csvGeneratedPaths.Add((Join-Path $latest $file))
    }
}
function Open-AdRun([switch]$Staged) {
    $global:csvGeneratedPaths = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    Start-SmartM365CmdbSourceReceipt -ScriptPath $producer.FullName -SourceRootPath $latest
    if ($Staged) { Enable-SmartM365StagedSourcePublication }
}
try {
    # Static wiring: staging follows the receipt, and every DATA-LAST copy uses the staged helper.
    $text = [IO.File]::ReadAllText($producer.FullName)
    Check ($text -match "(?m)^\s*Start-SmartM365CmdbSourceReceipt [^\r\n]*-ReadOnly:\(\`$DomainWorker -or \`$ReportOnly -or \`$DuplicateAnalysisOnly\)\r?\n(\s*#[^\r\n]*\r?\n)*\s*Enable-SmartM365StagedSourcePublication") 'AD does not enable staging right after its receipt.'
    $copies = @($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Copy-SmartM365FileAtomically' }, $true))
    # Atomic copies stay in the staged helper, the DATA-ALL combine retry and the DATA-ALL WeeklyHistory.
    $directLatest = @($copies | Where-Object {
        $parent = $_.Parent; while ($parent -and $parent -isnot [Management.Automation.Language.FunctionDefinitionAst]) { $parent = $parent.Parent }
        -not $parent -or $parent.Name -notin @('Publish-SmartM365AdLatestCopy', 'Copy-SmartM365AdFileWithRetry', 'Save-WeeklyInventoryHistory')
    })
    Check ($directLatest.Count -eq 0) 'AD copies a CSV to DATA-LAST outside the staged helper.'
    Check (@($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Publish-SmartM365AdLatestCopy' }, $true)).Count -ge 4) 'AD does not publish its DATA-LAST copies through the staged helper.'
    $version = [regex]::Match($text, '\.VERSION\r?\n\s*([0-9.]+)').Groups[1].Value
    Check ($version -and $text -match ('\$ScriptVersion = "' + [regex]::Escape($version) + '"')) 'Header and runtime versions differ.'
    Check ($text -match "-MinimumVersion '1\.0\.(?:9[3-9]|[1-9]\d\d)'") 'AD does not require the staging Core.'

    & $core {
        $script:Recorded = New-Object 'Collections.Generic.List[object]'
        function script:Invoke-SmartM365SharePointCsvUpload { param([string]$LocalFilePath) $script:Recorded.Add($LocalFilePath) }
        function script:WriteLog { param($Message, $Level) $null = $Message, $Level }
    }
    function Get-DeferredUploadList { & $core { $script:Recorded.ToArray() } }

    # Baseline: the previous run published directly.
    Open-AdRun
    Publish-AdSet 'baseline'
    Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope $scope
    Complete-SmartM365CmdbSourceReceipt -Status 'Success' | Out-Null
    $baselineReceipt = Read-Json $receiptPath
    $baselineHashes = Get-LatestHashTable
    Check ($baselineReceipt.Status -eq 'Completed' -and @($baselineReceipt.Files).Count -eq $files.Count) 'Baseline AD receipt not completed.'
    Check ($script:Uploads.Count -eq $files.Count) 'Direct AD copies were not uploaded immediately.'
    $script:Uploads.Clear()

    # Running run: DATA-LAST and the receipt are unchanged; no DATA-LAST upload yet.
    Open-AdRun -Staged
    Publish-AdSet 'running'
    $during = Get-LatestHashTable
    Check (@($files | Where-Object { $during[$_] -ne $baselineHashes[$_] }).Count -eq 0) 'A running AD run changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $baselineReceipt.RunId) 'A running AD run changed the receipt.'
    Check ($script:Uploads.Count -eq 0 -and @(Get-DeferredUploadList).Count -eq 0) 'A staged AD copy was uploaded before promotion.'
    Check ((Read-Json $runPath).Status -eq 'Collecting') 'A running AD run left the Collecting state.'

    # A failed run (for example a swallowed copy failure or a missing required file) changes nothing.
    Complete-SmartM365CmdbSourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null
    $after = Get-LatestHashTable
    Check (@($files | Where-Object { $after[$_] -ne $baselineHashes[$_] }).Count -eq 0) 'A failed AD run changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $baselineReceipt.RunId) 'A failed AD run replaced the receipt.'
    Check (@(Get-ChildItem -LiteralPath $stagingProducer -Force).Count -eq 0) 'A failed AD run left its staging folder.'
    Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof (Read-Json $receiptPath)
    $script:checks++

    # A run that lost one required copy is never promoted, even when declared complete.
    Open-AdRun -Staged
    Publish-AdSet 'missing'
    Remove-Item -LiteralPath (Join-Path (Get-ChildItem -LiteralPath $stagingProducer -Directory)[0].FullName 'AD_Groups_AllDomains.csv') -Force
    Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope $scope
    Complete-SmartM365CmdbSourceReceipt -Status 'Success' | Out-Null
    $after = Get-LatestHashTable
    Check (@($files | Where-Object { $after[$_] -ne $baselineHashes[$_] }).Count -eq 0) 'An AD run with a lost required copy changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $baselineReceipt.RunId) 'An AD run with a lost required copy published a receipt.'

    # Complete run: all copies and the receipt are promoted together, uploads follow the promotion.
    Open-AdRun -Staged
    Publish-AdSet 'complete'
    Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope $scope
    Complete-SmartM365CmdbSourceReceipt -Status 'Success' | Out-Null
    $receipt = Read-Json $receiptPath
    $final = Get-LatestHashTable
    Check ($receipt.Status -eq 'Completed' -and $receipt.RunId -ne $baselineReceipt.RunId) 'A complete AD run did not publish its receipt.'
    foreach ($file in $files) {
        $entry = @($receipt.Files | Where-Object File -eq $file)
        Check ($final[$file] -ne $baselineHashes[$file] -and $entry.Count -eq 1 -and $entry[0].SHA256 -eq $final[$file]) "Promoted AD CSV and receipt disagree: $file"
    }
    $deferred = @(Get-DeferredUploadList)
    Check ($deferred.Count -eq $files.Count -and @($deferred | Where-Object { $_ -like '*.staging*' }).Count -eq 0) 'Promoted AD copies were not uploaded by canonical path.'
    $run = Read-Json $runPath
    Check ($run.Status -eq 'Completed' -and -not $run.UnqualifiedPublication -and $run.StagedPublication) 'Completed AD run state is not qualified.'
    Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof $receipt
    $script:checks++
}
finally {
    try { Complete-SmartM365CmdbSourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null } catch { Write-Verbose $_.Exception.Message }
    foreach ($name in $saved.Keys) {
        if ($saved[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $saved[$name].Value }
        else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
    }
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}
"PASS: $script:checks AD staged publication checks; PowerShell $($PSVersionTable.PSVersion)."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCARN/YNtwwZaPcS
# ZRTctVfpufO0E1qco1UtvwM1Zv94GqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICajfMBgDCjH05Yq87fR+LANWc+orKnUcVKTitBlSz5mMA0GCSqG
# SIb3DQEBAQUABIIBgKrTQYoL7muQWOPaJ7GihAS4t3cNmfvGXgLFj7+OITn6v3+C
# aV0lV78uFfiLI00rtZ86I1AkRcj+FY8BEjCE23eUiJbzXv3y3HpMcwLSA5cXFNw0
# NDyPelAK2CwkV4DQlzo7lf8kUNhYRcsq4DFxtVnUMha9t08yTF36XogfqzccfdNs
# u+IH5HEs94V6T29eGbVDqxwL7FpmECT25HobZIBcMNNooLMyov/blcp24KrHZrxw
# HTWGLUQggo725lbgoiyrw/hWE/H7JcPjs98PkKg6k5yQijcRq9vpGkQz5wEWgEqi
# q0LkS8W+ZsnNTvOa0u5nH00RuvtSg4Jp3YkraqYP0DyPbFxfdZ6iHKxibZfSh+wZ
# sg4M/xu/AV3BgI/TYXOPB1leVa422eT9SBvTkBWraHW3zQmmmvlFfrH+6T7w0JdL
# i+IkiFo22vBLoTE8vBXaxltd18dZ2wc5PLyV08jIW+0QLP09vUIQ6BWKWAP00zjW
# 7X3h1IOPMYNwVGqYx6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjAxNDI2
# NDdaMC8GCSqGSIb3DQEJBDEiBCBsfd2btaPutLrr5xNxGhLpwQVXBP/F4IwOrvIh
# AKjxHzANBgkqhkiG9w0BAQEFAASCAgCf51wp3gV0q7AuSTx+PAmd32fi8T9DisFD
# t6DBK5XVldMcXid8zgFTm/3Xo/dxOhxnO8ropTy8692vN8jjJ/+6AgEXY1+MTy+8
# K5slKTN6+l4Ndn/0oNV6l4yjNCDkTSHnuHs2gKf9SJorYwxorJNebbktDPlxN1ja
# 68x3A4zsCxvL1chbaQ3g8delBUaroSoUnlhGqtIBtNfi8icM62wqk4xvueVrTK21
# yYPDu863Hg3RhONw88aBo8Q9DL7+7j7DL2V8et+8hj9MAU6oWYxlmhUkXn+Lbh/m
# U9We6lc+wI0+MxlRwW7uNdMer4Qx6qjNlgDpW+oUHDD/aL4XiBJvYfxnYSHBf7LM
# aCtXUM9SWVLffDdJS5+DQ+BXX+5+bFFjP26/xXUKUWj5i1aS/FbYMMbi4M5g7b7V
# 2n4Qt77lnOI8cPx4LXhGMNYqZFPVxXW6sFgi1GruvfvQbihMJzLw8kw7KjzPdOVq
# LPXCR7qRtgimBNXBLMpG7gUvNp7B/pW+/YXEnue6GwZNzWoxxzN553DY92b5QErN
# e2OEVdfkozgLtE8alX2JtiEyeFwtWyid44CmnGwAHA/i2xOVICxl8zCKYFKrcMU/
# efOpvyJ/6zt0v1IWOAyrPT38bpkSk4CceXGL5hf57j6QSqIc3hRDwJrXkJV1JRiE
# W+8WKk62Tg==
# SIG # End signature block
