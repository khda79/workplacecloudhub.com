<#
.SYNOPSIS
    Offline checks for independent SharePoint upload policy resolution.
.VERSION
    1.0.1
#>
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365-SharePointUploadPolicy.ps1')
$categories = @('DataAll','DataLast','DataPowerBI','LogAll')
$checks = 0
function Assert-Policy($Condition, $Message) {
    if (-not $Condition) { throw $Message }
    $script:checks++
}
$syntheticRoot = Join-Path ([IO.Path]::GetTempPath()) 'SmartM365-Policy-Root'
$allRoot = Join-Path $syntheticRoot 'DATA-ALL'
$lastRoot = Join-Path $syntheticRoot 'DATA-LAST'
$powerRoot = Join-Path $syntheticRoot 'DATA-POWERBI'
$logRoot = Join-Path $syntheticRoot 'LOG-ALL'
$roots = @{ DataAllRoot=$allRoot; DataLastRoot=$lastRoot; DataPowerBIRoot=$powerRoot; LogAllRoot=$logRoot }
foreach ($case in @(
    @((Join-Path $allRoot 'WeeklyHistory/2026-W44/manifest.json.txt'),'DataAll'),
    @((Join-Path $lastRoot 'AD.csv'),'DataLast'),
    @((Join-Path $powerRoot 'current.json.txt'),'DataPowerBI'),
    @((Join-Path $logRoot 'run.log'),'LogAll'),
    @((Join-Path ($allRoot + '-other') 'file.csv'),$null)
)) {
    $actual = Get-SmartM365SharePointUploadCategory -Path $case[0] @roots
    Assert-Policy ($actual -eq $case[1]) "Incorrect category for $($case[0])"
}
foreach ($mask in 0..15) {
    $config = @{}
    for ($i = 0; $i -lt $categories.Count; $i++) {
        $config['EnableSharePointUpload' + $categories[$i]] = [bool]($mask -band (1 -shl $i))
    }
    for ($i = 0; $i -lt $categories.Count; $i++) {
        $actual = Resolve-SmartM365SharePointUploadPolicy -Category $categories[$i] -GlobalConfig $config -LegacyEnabled $false
        Assert-Policy ($actual.Enabled -eq [bool]($mask -band (1 -shl $i)) -and $actual.Source -eq 'Global') "Incorrect matrix result: $mask/$($categories[$i])"
    }
}
foreach ($category in $categories) {
    $key = 'EnableSharePointUpload' + $category
    $global = @{}; $global[$key] = $true
    $local = @{}; $local[$key] = $false
    $result = Resolve-SmartM365SharePointUploadPolicy $category $local $global $true
    Assert-Policy (-not $result.Enabled -and $result.Source -eq 'Local') "Local override failed: $category"
    $local[$key] = '__USE_GLOBAL__'
    $result = Resolve-SmartM365SharePointUploadPolicy $category $local $global $false
    Assert-Policy ($result.Enabled -and $result.Source -eq 'Global') "Inheritance failed: $category"
    $result = Resolve-SmartM365SharePointUploadPolicy $category @{} @{} $true
    Assert-Policy ($result.Enabled -and $result.Source -eq 'Legacy') "Legacy fallback failed: $category"
    $local[$key] = 'false'
    $rejected = $false
    try { Resolve-SmartM365SharePointUploadPolicy $category $local $global $false | Out-Null } catch { $rejected = $true }
    Assert-Policy $rejected "Invalid Boolean accepted: $category"
}
foreach ($batch in @($false, $true)) {
    foreach ($logs in @($false, $true)) {
        $config = @{ EnableSharePointUploadDataPowerBI=$batch; EnableSharePointUploadLogAll=$logs }
        $result = Resolve-SmartM365PreparedUploadPolicy -LocalConfig $config -LegacyBatchEnabled $false
        Assert-Policy ($result.BatchEnabled -eq $batch -and $result.LogEnabled -eq $logs) 'Normal preparation policy mismatch.'
        $result = Resolve-SmartM365PreparedUploadPolicy -LocalConfig $config -LegacyBatchEnabled $false -Offline
        Assert-Policy (-not $result.BatchEnabled -and -not $result.LogEnabled) 'Offline mode allowed an upload.'
        foreach ($mode in @('ValidateOnly','WorkforceDiagnostic')) {
            $options = @{}; $options[$mode] = $true
            $result = Resolve-SmartM365PreparedUploadPolicy -LocalConfig $config -LegacyBatchEnabled $false @options
            Assert-Policy (-not $result.BatchEnabled -and $result.LogEnabled -eq $logs) "$mode policy mismatch."
        }
        $result = Resolve-SmartM365PreparedUploadPolicy -LocalConfig $config -LegacyBatchEnabled $false -TransferOnly
        Assert-Policy ($result.BatchEnabled -and $result.LogEnabled -eq $logs) 'TransferOnly policy mismatch.'
    }
}
$result = Resolve-SmartM365PreparedUploadPolicy -LegacyBatchEnabled $false
Assert-Policy (-not $result.BatchEnabled -and $result.LogEnabled) 'Historical preparation defaults changed.'
$rejected = $false
try { Resolve-SmartM365PreparedUploadPolicy -LegacyBatchEnabled $false -TransferOnly -Offline | Out-Null } catch { $rejected = $true }
Assert-Policy $rejected 'Conflicting prepared modes were accepted.'
foreach ($mask in 0..7) {
    $config = @{}
    $collectorCategories = @('DataAll','DataLast','LogAll')
    for ($i = 0; $i -lt $collectorCategories.Count; $i++) {
        $config['EnableSharePointUpload' + $collectorCategories[$i]] = [bool]($mask -band (1 -shl $i))
    }
    foreach ($allowed in @($false,$true)) {
        $result = Resolve-SmartM365CollectorUploadPolicies -GlobalConfig $config -LegacyEnabled $false -ExternalActionsAllowed $allowed
        for ($i = 0; $i -lt $collectorCategories.Count; $i++) {
            $expected = $allowed -and [bool]($mask -band (1 -shl $i))
            Assert-Policy ($result[$collectorCategories[$i]].Enabled -eq $expected) "Collector run guard failed: $mask/$allowed/$i"
            Assert-Policy ((Test-SmartM365CollectorUploadEnabled -Category $collectorCategories[$i] -Policies $result -RunEnabled $true) -eq $expected) "Collector effective policy failed: $mask/$allowed/$i"
            Assert-Policy (-not (Test-SmartM365CollectorUploadEnabled -Category $collectorCategories[$i] -Policies $result -RunEnabled $false)) "Runtime cancellation was bypassed: $mask/$allowed/$i"
        }
    }
}
Assert-Policy (Test-SmartM365CollectorUploadEnabled -Category DataAll -Policies $null -RunEnabled $true) 'Legacy upload behavior changed.'
$rejected = $false
try { Test-SmartM365CollectorUploadEnabled -Category DataAll -Policies @{} -RunEnabled $true | Out-Null } catch { $rejected = $true }
Assert-Policy $rejected 'Incomplete collector policy accepted.'
# Run the real entry-point policy assignment with synthetic settings and no preparation.
$preparePath = Join-Path $PSScriptRoot '../SmartInventory/PreparedEvidence/SmartM365-WorkplaceEvidence-Prepare.ps1'
$tokens = $null; $parseErrors = $null
$prepareAst = [Management.Automation.Language.Parser]::ParseFile($preparePath, [ref]$tokens, [ref]$parseErrors)
Assert-Policy ($parseErrors.Count -eq 0) 'Prepare entry point failed to parse.'
$assignments = @($prepareAst.FindAll({ param($node) $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$uploadPolicy' }, $true))
Assert-Policy ($assignments.Count -eq 1) 'Prepare entry point must have exactly one upload policy assignment.'
function ConfigValue([string]$Name) { if ($Name -eq 'EnableSharePointUpload') { return $false }; throw "Unexpected setting: $Name" }
$config = @{ EnableSharePointUploadDataPowerBI=$true; EnableSharePointUploadLogAll=$false }
$effective = @{}
$Offline=$false; $ValidateOnly=$false; $WorkforceDiagnostic=$false; $TransferOnly=$false
. ([scriptblock]::Create($assignments[0].Extent.Text))
Assert-Policy ($uploadPolicy.BatchEnabled -and -not $uploadPolicy.LogEnabled) 'Prepare policy assignment did not apply independent settings.'
$Offline=$true
. ([scriptblock]::Create($assignments[0].Extent.Text))
Assert-Policy (-not $uploadPolicy.BatchEnabled -and -not $uploadPolicy.LogEnabled) 'Prepare policy assignment bypassed Offline.'
# Exercise the Devices entry point's actual policy initialization without executing the collector.
$devicesPath = Join-Path $PSScriptRoot '../SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Inventory.ps1'
$deviceTokens=$null; $deviceErrors=$null
$deviceAst=[Management.Automation.Language.Parser]::ParseFile($devicesPath,[ref]$deviceTokens,[ref]$deviceErrors)
Assert-Policy ($deviceErrors.Count -eq 0) 'Devices entry point failed to parse.'
$policyAssignment=@($deviceAst.FindAll({param($node) $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$global:SmartM365CollectorUploadPolicies'},$true))
Assert-Policy ($policyAssignment.Count -eq 1) 'Devices policy assignment missing or ambiguous.'
$savedDevicePolicyVar=Get-Variable SmartM365CollectorUploadPolicies -Scope Global -ErrorAction SilentlyContinue
$savedDevicePolicy=if($savedDevicePolicyVar){@{Value=$savedDevicePolicyVar.Value}}else{$null}
$savedGlobalConfigVar=Get-Variable SmartM365GlobalConfig -Scope Global -ErrorAction SilentlyContinue
$savedGlobalConfig=if($savedGlobalConfigVar){@{Value=$savedGlobalConfigVar.Value}}else{$null}
try {
    $global:SmartM365GlobalConfig=@{}
    foreach($mask in 0..7) {
        $ScriptLocalConfig=@{}; $i=0
        foreach($category in @('DataAll','DataLast','LogAll')) { $ScriptLocalConfig['EnableSharePointUpload'+$category]=[bool]($mask -band (1 -shl $i)); $i++ }
        $legacySharePointUpload=$false
        . ([scriptblock]::Create($policyAssignment[0].Extent.Text))
        $i=0
        foreach($category in @('DataAll','DataLast','LogAll')) {
            Assert-Policy ($global:SmartM365CollectorUploadPolicies[$category].Enabled -eq [bool]($mask -band (1 -shl $i))) "Devices policy mismatch: $mask/$category"
            $i++
        }
    }
} finally {
    if($savedDevicePolicy){Set-Variable SmartM365CollectorUploadPolicies -Scope Global -Value $savedDevicePolicy.Value}else{Remove-Variable SmartM365CollectorUploadPolicies -Scope Global -ErrorAction SilentlyContinue}
    if($savedGlobalConfig){Set-Variable SmartM365GlobalConfig -Scope Global -Value $savedGlobalConfig.Value}else{Remove-Variable SmartM365GlobalConfig -Scope Global -ErrorAction SilentlyContinue}
}
# The two specialized initializers must reject ambiguous identical roots before publishing.
$oldRootsVariable=Get-Variable SmartM365CollectorUploadRoots -Scope Global -ErrorAction SilentlyContinue
$oldPoliciesVariable=Get-Variable SmartM365CollectorUploadPolicies -Scope Global -ErrorAction SilentlyContinue
$oldRoots=if($oldRootsVariable){@{Value=$oldRootsVariable.Value}}else{$null}
$oldPolicies=if($oldPoliciesVariable){@{Value=$oldPoliciesVariable.Value}}else{$null}
try {
    foreach($relative in @('M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Inventory.ps1','ActiveDirectoryInventory/SmartM365-ActiveDirectory-Inventory.ps1')) {
        $path=Join-Path (Join-Path $PSScriptRoot '../SmartInventory') $relative
        $tokens=$null;$errors=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
        Assert-Policy ($errors.Count -eq 0) "Specialized collector parser failed: $relative"
        $command=$ast.Find({param($node) $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Assert-SmartM365CollectorUploadRoots'},$true)
        Assert-Policy ($null -ne $command) "Root safety check missing: $relative"
        $global:SmartM365CollectorUploadRoots=[pscustomobject]@{DataAllRoot=$allRoot;DataLastRoot=$allRoot;LogAllRoot=$logRoot}
        $global:SmartM365CollectorUploadPolicies=@{DataAll=[pscustomobject]@{Enabled=$true};DataLast=[pscustomobject]@{Enabled=$false};LogAll=[pscustomobject]@{Enabled=$false}}
        $rejected=$false;try{& ([scriptblock]::Create($command.Extent.Text))}catch{$rejected=$true}
        Assert-Policy $rejected "Ambiguous category roots accepted: $relative"
        $global:SmartM365CollectorUploadPolicies.DataLast.Enabled=$true
        & ([scriptblock]::Create($command.Extent.Text))
        Assert-Policy $true "Equivalent category policies failed: $relative"
    }
} finally {
    if($oldRoots){Set-Variable SmartM365CollectorUploadRoots -Scope Global -Value $oldRoots.Value}else{Remove-Variable SmartM365CollectorUploadRoots -Scope Global -ErrorAction SilentlyContinue}
    if($oldPolicies){Set-Variable SmartM365CollectorUploadPolicies -Scope Global -Value $oldPolicies.Value}else{Remove-Variable SmartM365CollectorUploadPolicies -Scope Global -ErrorAction SilentlyContinue}
}
. (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1')
foreach ($mask in 0..7) {
    $local=@{};$i=0
    foreach($category in @('DataAll','DataLast','LogAll')) { $local['EnableSharePointUpload'+$category]=[bool]($mask -band (1 -shl $i));$i++ }
    $state=Start-SmartM365InventoryRun -ScriptPath 'synthetic.ps1'
    $global:EnableSharePointUpload=$false
    Initialize-SmartM365CollectorUploadContext -LocalConfig $local -GlobalConfig @{} -DataAllRoot $allRoot -DataLastRoot $lastRoot -LogAllRoot $logRoot -ExternalActionsAllowed $true
    Assert-Policy ($global:EnableSharePointUpload -eq ($mask -ne 0)) "Collector aggregate mismatch: $mask"
    $i=0;foreach($category in @('DataAll','DataLast','LogAll')) {
        Assert-Policy ($global:SmartM365CollectorUploadPolicies[$category].Enabled -eq [bool]($mask -band (1 -shl $i))) "Collector initialization mismatch: $mask/$category"; $i++
    }
    $state.ReachedTerminal=$true
    Complete-SmartM365InventoryRun -State $state
    Assert-Policy (-not $global:EnableSharePointUpload) 'Collector legacy flag was not restored.'
}
$state=Start-SmartM365InventoryRun -ScriptPath 'synthetic-guard.ps1'
$global:EnableSharePointUpload=$true
$allEnabled=@{EnableSharePointUploadDataAll=$true;EnableSharePointUploadDataLast=$true;EnableSharePointUploadLogAll=$true}
Initialize-SmartM365CollectorUploadContext -LocalConfig $allEnabled -GlobalConfig @{} -DataAllRoot $allRoot -DataLastRoot $lastRoot -LogAllRoot $logRoot -ExternalActionsAllowed $false
Assert-Policy (-not $global:EnableSharePointUpload -and @($global:SmartM365CollectorUploadPolicies.Values | Where-Object Enabled).Count -eq 0) 'External-action guard did not suppress collector uploads.'
$state.ReachedTerminal=$true;Complete-SmartM365InventoryRun -State $state
Assert-Policy $global:EnableSharePointUpload 'Legacy enabled flag was not restored after guarded run.'
$state=Start-SmartM365InventoryRun -ScriptPath 'synthetic-overlap.ps1';$global:EnableSharePointUpload=$false
$rejected=$false
try { Initialize-SmartM365CollectorUploadContext -LocalConfig @{EnableSharePointUploadDataAll=$true;EnableSharePointUploadDataLast=$false} -GlobalConfig @{} -DataAllRoot $allRoot -DataLastRoot $allRoot -LogAllRoot $logRoot -ExternalActionsAllowed $true } catch { $rejected=$true }
Assert-Policy $rejected 'Conflicting categories with identical roots were accepted.'
$state.ReachedTerminal=$true;Complete-SmartM365InventoryRun -State $state
Write-Host "SharePoint upload policy: $checks checks passed."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBiSNoHtAOQUCoG
# M4eYn9pB1sAtYI+cIF9ETQrfuSMV5aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGoF47JqUJFAAJDjKgymyPHFSGrD43LamEWM9BlpgCstMA0GCSqG
# SIb3DQEBAQUABIIBgCdh4mbSpo4ujjCWTjf4cKSkQtobGlPUY7eWcJ4yLqAq+0Nd
# F0eavW8yAylMlmPdp6BJOsTU5DzdIj3wTFXrYZ90D2Y8U1pF/mDerKlF8Nk4tre/
# KTwDruIEKW5QZOdp+tVpZVOc1eR0y/Uz1rxptPQrn7oawzkXrteawAgft0u2hStS
# AalfhqVkKjMZ9IfoY+Ag+/1Etyhi7XV44awpNtCvuAs9riP6PLahMVi7L46eRom/
# VRId0w3jyUDbSquB+3n6uvq1UbB+x2uORFSakIVb/5m2xAlvP/m3WTXMWuCN3AxC
# K7D/E0rzG6T0NH/9EpB8zkXou8QuzxBWes/ngE2BQz4RXLio2lfonXX2M6YI4whw
# kHIIEKSLwxNIB2v+ftjbUb5T0Xgzqtap2eG0055RB117fE6sLJ9TadlYuzP9S6Ip
# +Hb+TT3dX6WNewN0PyDFz8iL7eXO6eOaFecWb7Y/omlxPJZKicHfu/PQTXeF3emD
# kVVskgap1rmBi7p5uKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzExMzU5
# MjhaMC8GCSqGSIb3DQEJBDEiBCAkyHMhSLdsbRCY+d3w1CU+IiUIcaNMzWVb08cz
# YdCThzANBgkqhkiG9w0BAQEFAASCAgCMj+JpRKXTuWCkH/mZ6mPxWWOprRngtqoD
# 2jzNpzFhTkIqOBtfSbNOKK4Pv5tJKnWdNth0vndfEh8YJjC/8WSpAWMOnoXU5CL/
# lqC3r+aRJTueynrb+/WAfxeJhGaAwd2kNClo6tw9IVT/MSWY9wxk0tQvh/AETHmb
# UegYJUZ0R6KJA4WA6bEFz4twruchkW18j351f9k6uTWgcYIgFBdGA4rkbgptt19o
# iZT2Djts48wGoK66Bi4jrZF5++xDMbRTly0oWB0myPqnfg6gZPD3pjbF5utGmci1
# 1vT4ECjLwyk1ev1SauCuM2AJTY6IyMgBuhsUHtl5o4aHp4WgSiPYqJPD9WVp1lqk
# M76qVCnxdEFvh5QNZvig/L8p5ZLViwmzI2FCL1iDE/AMzzlSXrTM63p24IKu4qV2
# +Ro4Tg7h3NUXNmWOlq9ZNMgdkr9+eRTiJNIJm9D6mNmWGySJ0wLI/ZPdzAu7//dd
# x8SeuPlULm7ED78O+JUEUfPFytkTFZ4mhkv+kxNSrz8uooBG5ZGt+rLmyhpF9rT0
# Yr1siIslJ/qNSCCUe6S0udFiOX9YPTH7WFIqFWvWJ7M8l218MKT2Vnmj1lCcWyzd
# 18raXQAvigePaHoYRBYwLTzWVosqELxW9dRyanE4M4sva/zc2Go6+DM7IV5/mkjv
# yZEuk1BVxA==
# SIG # End signature block
