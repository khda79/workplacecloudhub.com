#Requires -Version 7.0
<#
Synthetic and local-file checks only. No authentication or Graph mutation.
.VERSION
1.0.2
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$smartM365Root = Join-Path $repoRoot 'SmartM365'
$deviceRoot = Join-Path $smartM365Root 'SmartInventory/M365Inventory/IntuneInventory/Devices'
$creatorPath = Join-Path $deviceRoot 'SmartM365-Create-DeviceSystemRemediation.ps1'
$detectionPath = Join-Path $deviceRoot 'SmartM365-Detect-DeviceSystemInfo.ps1'
$catalogPath = Join-Path $smartM365Root 'SmartInventory/Orchestrator/Orchestrator-Jobs.json.template'
$checks = 0
function Assert([bool]$Condition, [string]$Message) { $script:checks++; if (-not $Condition) { throw $Message } }

$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($creatorPath, [ref]$tokens, [ref]$errors)
Assert ($errors.Count -eq 0) 'The creator does not parse.'
foreach ($name in @('Get-GraphProperty','Get-GraphCollection','Get-DeviceSystemPackageBody','Invoke-CreateDeviceSystemPackage',
    'Get-DeviceSystemCollectorConfigPath','Get-ConfiguredDeviceHealthScriptId','Set-DeviceSystemCollectorPackageId','Confirm-ExistingDeviceSystemPackage')) {
    $definition = $ast.Find({ param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    }, $true)
    Assert ($null -ne $definition) "Missing function $name."
    . ([scriptblock]::Create($definition.Extent.Text))
}

$body = Get-DeviceSystemPackageBody -ScriptPath $detectionPath -Name 'Synthetic Device System' -CheckSignature $false
Assert ($body.displayName -eq 'Synthetic Device System' -and $body.runAsAccount -eq 'system' -and -not $body.runAs32Bit) 'The package execution context changed.'
Assert ($body.remediationScriptContent -ceq '' -and $body.deviceHealthScriptType -eq 'deviceHealthScript') 'The package is not detection only.'
Assert (-not $body.Contains('assignments') -and -not $body.Contains('runSchedule')) 'The creation body contains an assignment or schedule.'
$detectionVersion=[regex]::Match([IO.File]::ReadAllText($detectionPath),'(?m)^\s*Version:\s*([0-9]+(?:\.[0-9]+)+)\s*$').Groups[1].Value
Assert ($detectionVersion -and $body.version -eq $detectionVersion) 'The package version does not match the detection script.'
Assert ([Convert]::FromBase64String($body.detectionScriptContent).Length -eq [IO.File]::ReadAllBytes($detectionPath).Length) 'Detection script content was not encoded.'
Assert ([Linq.Enumerable]::SequenceEqual([byte[]][Convert]::FromBase64String($body.detectionScriptContent),[byte[]][IO.File]::ReadAllBytes($detectionPath))) 'Detection script bytes changed in the package.'
Assert ($body.roleScopeTagIds.Count -eq 1 -and $body.roleScopeTagIds[0] -eq '0') 'The default scope tag changed.'
$bodySigned = Get-DeviceSystemPackageBody -ScriptPath $detectionPath -Name 'Synthetic Device System' -CheckSignature $true
Assert ($bodySigned.enforceSignatureCheck -eq $true) 'Signature enforcement option was ignored.'

$script:graphPages = @(
    @{value=@(@{id='one'});'@odata.nextLink'='next'},
    @{value=@(@{id='two'})}
)
$script:graphCalls = 0
function Invoke-MgGraphRequest {
    param($Method,$Uri,$ErrorAction)
    $page = $script:graphPages[$script:graphCalls]
    $script:graphCalls++
    return $page
}
$items = @(Get-GraphCollection -Uri 'first')
Assert ($items.Count -eq 2 -and $script:graphCalls -eq 2) 'Paged Graph results were not collected.'
$script:graphPages = @(@{value=@()})
$script:graphCalls = 0
Assert (@(Get-GraphCollection -Uri 'assignments').Count -eq 0) 'A valid empty assignment list was rejected.'
$script:graphPages = @(@{value=@(@{id='one'});'@odata.nextLink'='first'})
$script:graphCalls = 0
$failed = $false
try { $null = Get-GraphCollection -Uri 'first' } catch { $failed = $true }
Assert $failed 'A repeated Graph pagination link was accepted.'
$script:graphPages = @(@{wrong=@()})
$script:graphCalls = 0
$failed = $false
try { $null = Get-GraphCollection -Uri 'first' } catch { $failed = $true }
Assert $failed 'A malformed Graph collection was accepted.'

function Write-Step { param($Message) }
$script:requests = [System.Collections.Generic.List[object]]::new()
$script:existingPackages = @()
$script:assignmentItems = @()
$script:createdId = '00000000-0000-0000-0000-000000000111'
function Invoke-MgGraphRequest {
    param($Method,$Uri,$Body,$ContentType,$ErrorAction)
    $script:requests.Add([pscustomobject]@{Method=$Method;Uri=$Uri;Body=$Body})
    if ($Method -eq 'POST') { return @{id=$script:createdId} }
    if ($Uri -like '*/assignments') { return @{value=$script:assignmentItems} }
    return @{value=$script:existingPackages}
}
$id = Invoke-CreateDeviceSystemPackage -Body $body -GraphRoot 'https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts'
Assert ($id -eq $script:createdId) 'The created package ID was not returned.'
Assert (@($script:requests | Where-Object Method -eq 'POST').Count -eq 1) 'Creation did not make exactly one POST.'
Assert (@($script:requests | Where-Object { $_.Method -eq 'POST' -and $_.Uri -like '*/assign*' }).Count -eq 0) 'The creator made an assignment POST.'
Assert (@($script:requests | Where-Object { $_.Method -eq 'GET' -and $_.Uri -like '*/assignments' }).Count -eq 1) 'The creator did not verify assignments.'
$posted = ($script:requests | Where-Object Method -eq 'POST' | Select-Object -First 1).Body | ConvertFrom-Json
Assert ($posted.remediationScriptContent -ceq '' -and $posted.runAsAccount -eq 'system') 'The posted package differs from the safe body.'

$script:requests.Clear()
$script:existingPackages = @(@{id='00000000-0000-0000-0000-000000000222'})
$failed = $false
try { $null = Invoke-CreateDeviceSystemPackage -Body $body -GraphRoot 'https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts' } catch { $failed = $true }
Assert $failed 'An existing package did not block creation.'
Assert (@($script:requests | Where-Object Method -eq 'POST').Count -eq 0) 'An existing package was overwritten or duplicated.'

$script:requests.Clear()
$script:existingPackages = @()
$script:assignmentItems = @(@{id='unexpected'})
$failed = $false
try { $null = Invoke-CreateDeviceSystemPackage -Body $body -GraphRoot 'https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts' } catch { $failed = $_.Exception.Message -match 'assignment' }
Assert $failed 'An unexpected assignment was not reported.'

$existingId = '00000000-0000-0000-0000-000000000333'
$script:requests.Clear()
$script:existingPackage = @{id=$existingId;displayName='Synthetic Device System';publisher='SmartM365 by WorkplaceCloudHub'}
function Invoke-MgGraphRequest {
    param($Method,$Uri,$Body,$ContentType,$ErrorAction)
    $script:requests.Add([pscustomobject]@{Method=$Method;Uri=$Uri;Body=$Body})
    return $script:existingPackage
}
Confirm-ExistingDeviceSystemPackage -GraphRoot 'https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts' -PackageId $existingId -Name 'Synthetic Device System'
Assert ($script:requests.Count -eq 1 -and $script:requests[0].Method -eq 'GET') 'Existing-package registration changed Intune.'
$script:existingPackage.displayName = 'Different package'
$failed = $false
try { Confirm-ExistingDeviceSystemPackage -GraphRoot 'https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts' -PackageId $existingId -Name 'Synthetic Device System' } catch { $failed = $true }
Assert $failed 'A mismatched existing package was accepted.'

function Resolve-SmartM365JsonConfigurationPath { param($Path) return "$Path.txt" }
function Read-SmartM365JsonDocument {
    param($Path,$Validate)
    $document = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    if ($Validate) { & $Validate $document }
    return [pscustomobject]@{Path=$Path;Document=$document;SHA256=(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash}
}
$script:configWrites = 0
function Write-SmartM365JsonBytesAtomically {
    param($Path,$Bytes,$ExpectedSHA256,$Validate)
    if ((Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash -ne $ExpectedSHA256) { throw 'Synthetic lost update.' }
    $document = [Text.Encoding]::UTF8.GetString($Bytes) | ConvertFrom-Json
    if ($Validate) { & $Validate $document }
    [IO.File]::WriteAllBytes($Path, $Bytes)
    $script:configWrites++
}
$tempDir = Join-Path $PSScriptRoot ('.device-system-config-test-' + [guid]::NewGuid().ToString('N'))
$configPath = Join-Path $tempDir 'SmartInventory\M365Inventory\IntuneInventory\Devices\SmartM365-Device-System-Inventory.local.json.txt'
try {
    [void][IO.Directory]::CreateDirectory((Split-Path $configPath -Parent))
    [IO.File]::WriteAllText($configPath, '{"Keep":"value"}', [Text.UTF8Encoding]::new($false))
    Assert ((Get-DeviceSystemCollectorConfigPath -Root $tempDir) -eq $configPath) 'The creator did not select the collector JSON text file under ConfigRoot.'
    Set-DeviceSystemCollectorPackageId -Path $configPath -PackageId $existingId
    $saved = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
    Assert ($saved.DeviceHealthScriptId -eq $existingId -and $saved.Keep -eq 'value') 'Package ID was not saved in the collector JSON text file.'
    Assert ((Get-ConfiguredDeviceHealthScriptId -Path $configPath) -eq $existingId) 'Collector package ID read-back failed.'
    $beforeHash = (Get-FileHash -LiteralPath $configPath -Algorithm SHA256).Hash
    Set-DeviceSystemCollectorPackageId -Path $configPath -PackageId $existingId
    Assert ((Get-FileHash -LiteralPath $configPath -Algorithm SHA256).Hash -eq $beforeHash -and $script:configWrites -eq 1) 'Registering the same package changed the configuration.'
    $failed = $false
    try { Set-DeviceSystemCollectorPackageId -Path $configPath -PackageId '00000000-0000-0000-0000-000000000444' } catch { $failed = $true }
    Assert ($failed -and (Get-FileHash -LiteralPath $configPath -Algorithm SHA256).Hash -eq $beforeHash) 'A different collector package ID was overwritten.'
}
finally {
    $resolvedTempDir = [IO.Path]::GetFullPath($tempDir)
    $resolvedTempRoot = [IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\') + '\'
    if ($resolvedTempDir.StartsWith($resolvedTempRoot, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolvedTempDir) -like '.device-system-config-test-*') {
        [IO.Directory]::Delete($resolvedTempDir, $true)
    }
}

$source = Get-Content -LiteralPath $creatorPath -Raw
Assert ($source -match 'Connect-MgGraph -TenantId \$tenantId -Scopes' -and $source -match 'DeviceManagementScripts.ReadWrite.All') 'Delegated interactive authentication was not requested.'
Assert ($source -match 'Invoke-MgGraphRequest -Method POST -Uri \$GraphRoot' -and $source -notmatch 'Invoke-MgGraphRequest -Method POST -Uri .*\/assign') 'The creator calls an assignment action.'
Assert ($source -match '\$GraphRoot/\$createdId/assignments' -and $source -match '\$assignments.Count -ne 0') 'Post-creation zero-assignment verification is missing.'
Assert ($source -match 'Set-DeviceSystemCollectorPackageId -Path \$collectorConfigPath -PackageId \$packageId') 'Package creation does not register the collector ID.'
$catalogEntry = @((Get-Content -LiteralPath $catalogPath -Raw | ConvertFrom-Json).Launches | Where-Object Name -eq 'Intune-DeviceSystemRemediation-Create')
Assert ($catalogEntry.Count -eq 1 -and $catalogEntry[0].Manual -and [string]$catalogEntry[0].ScriptPath -match 'SmartM365-Create-DeviceSystemRemediation\.ps1$') 'The catalog manual entry does not target the creator.'
Assert ([IO.File]::Exists([IO.Path]::GetFullPath((Join-Path (Join-Path $smartM365Root 'SmartInventory') $catalogEntry[0].ScriptPath)))) 'The catalog script path does not resolve.'

[pscustomobject]@{Passed=$true;Checks=$checks;GraphCalled=$false;GraphMutation=$false;InteractiveAuth=$false} | ConvertTo-Json

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAFS3QkP11tHO8P
# URIkDJrXgyd2kNsAcApc40PRySzHPKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIblXi2zd+cxV+B0DEwa2Xt78QFn5Kt0anYJ9J7Qctv9MA0GCSqG
# SIb3DQEBAQUABIIBgKu+sSnBCCXnNx0HqFqH8u1A3mgdTOzQCHBVUVDQQGmulvB5
# 1ekqVDEElTgJ02RIr+h3zYrfwZFl+OnxrLCfLRMINjk+G6z0mRt+y+Gy0kjsHML9
# yjzUalvtGAW55sOWL5plNAviFYNnCe3LfVkCaNQPI0kPjH9g6raSMJwqCoQrkz1O
# HmIh8yhxxvSBLr7uIQn8Fb/F0MGM3xcArDqy4yP23ODZ22x4PeKFX/pbGvO6BEhS
# nPOyv4Wjk1nE92LnOdCoo9DBiHr+RXxwm+d4ek27gqqHxy4PlfFE6OlojNyYD6+I
# DfLdOOllxYdEobyUdhwFDXCGeQdC8OKfhlojqIGt+nhd8q0RaDFs2BDQ1ZnBHkdO
# WvbOzW1XGomu8wKJVsLKJsv31m0OhFjuiIiEcoYvcB3ftXRxFQo0gStii3mhjMLK
# CMLCwnROPbFDZjbd5HH0KEpjsCpaJArTNNcQBLxBlwWNLju8mm+GOvcor1BCzwvY
# DrDWNWfEzoMZWWB3bqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTQxNTMy
# MjlaMC8GCSqGSIb3DQEJBDEiBCDmAIzXabRH6+UxQddSSK9m+GsyaOWKvdJ8/6Oi
# 0Kg6qDANBgkqhkiG9w0BAQEFAASCAgCR1c2YaKx0SL+UMg5QbBiWPVQuKYrajOJV
# lVhi+9RBvsByxPAxwdfNPj4lcP22Uw6qbNyRFTbZov3CKBJ8CATOi5s8njKwdjyu
# RI9t4tj6vUKCIU6K3/js53MgiSw1ypUFuuZjexqw3HTJNJAREt+WdaHc5qY42A7w
# tLjWQXvOKcNYMqeUaeN6+zOMQafaajhurSkQMXJuD+ZC4H0cfCOeuJxmFjVaHuSy
# p11mOyFhhg1C6X7ywIkib60HvNMVLLs/Y2sdrWBQNxqKwencTFkh58WyURk4L1eo
# f2gT4SkiwvLgwws62penPxdrrumoKlEV3E/I7+2CfhhtIO8XcDWTmJmm9fEU+mo+
# vV5Ys/VObZASnbV1PQ9LfYzaBZF5k25wkO24FQwDxjJ2ZjObcSbuCOxqb1SzpihN
# H7gztx+xGSITrg4WqBRS7jYotl5XujHmDlYG0pbjzYQ78AGtnBJhzcXLZsXzQLER
# tgtM7ngEA/b0TrRfgR46KHjolwMAeLjfJW9tsu24znNpwwEaGFIhr2enboHyb0sV
# 0jAJDWq3CmhO085VFno+Brrxj0Cv3ZjXcfqoWfqZPHn4tL5e2p5ejx8zQYFN+kL9
# tZ6vmyIW7p+PZZ9m+vY9pt8UjAlVgaNhv4Yu/Xagt+FNV2eiRAVyzcDBdc8CHYq+
# AybgeMFKkg==
# SIG # End signature block
