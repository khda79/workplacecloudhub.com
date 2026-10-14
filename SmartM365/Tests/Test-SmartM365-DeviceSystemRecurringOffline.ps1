#Requires -Version 7.0
<# Synthetic-only tests. Never runs a collector, a device probe, or Graph. #>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$deviceRoot = Join-Path $root 'SmartM365/SmartInventory/M365Inventory/IntuneInventory/Devices'
$probePath = Join-Path $deviceRoot 'SmartM365-Detect-DeviceSystemInfo.ps1'
$collectorPath = Join-Path $deviceRoot 'SmartM365-Device-System-Inventory.ps1'
$checks = 0
function Assert([bool]$Condition, [string]$Message) { $script:checks++; if (-not $Condition) { throw $Message } }
function Get-FunctionSource([string]$Path, [string]$Name) {
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    Assert ($errors.Count -eq 0) "$Path did not parse."
    $definition = $ast.Find({ param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name
    }, $true)
    Assert ($null -ne $definition) "$Name was not found."
    return $definition.Extent.Text
}
foreach ($name in @('Get-LastConsoleLogon')) { . ([scriptblock]::Create((Get-FunctionSource $probePath $name))) }
foreach ($name in @('ConvertTo-DeviceSystemGraphUtcTimestamp','Get-DeviceHealthScriptRunStates','Parse-PlatformScriptStdout','Test-DeviceSystemRecurringPayload','Get-DeviceSystemLocalFreshness',
    'Test-GraphResponseProperty','Get-GraphResponsePropertyValue')) { . ([scriptblock]::Create((Get-FunctionSource $collectorPath $name))) }

$script:eventMode = 'found'
$script:observedXPath = ''
function New-FakeEvent([string]$Sid, [string]$Name, [string]$Domain, [string]$Time) {
    $xml = "<Event><System><TimeCreated SystemTime='$Time'/></System><EventData><Data Name='TargetUserSid'>$Sid</Data><Data Name='TargetUserName'>$Name</Data><Data Name='TargetDomainName'>$Domain</Data></EventData></Event>"
    $securityEvent = [pscustomobject]@{ Xml = $xml }
    $securityEvent | Add-Member -MemberType ScriptMethod -Name ToXml -Value { $this.Xml }
    return $securityEvent
}
function Get-WinEvent {
    param($LogName, $FilterXPath, $MaxEvents, $ErrorAction)
    $script:observedXPath = $FilterXPath
    Assert ($LogName -eq 'Security' -and $MaxEvents -le 512) 'The local Security query is unbounded.'
    if ($script:eventMode -eq 'error') { throw 'Synthetic Security log read failure.' }
    if ($script:eventMode -eq 'systemOnly') { return (New-FakeEvent 'S-1-5-18' 'SYSTEM' 'NT AUTHORITY' '2026-10-13T11:00:00Z') }
    @(
        (New-FakeEvent 'S-1-5-18' 'SYSTEM' 'NT AUTHORITY' '2026-10-13T11:00:00Z'),
        (New-FakeEvent 'S-1-12-1-123' 'person' 'AzureAD' '2026-10-13T10:00:00Z')
    )
}
$last = Get-LastConsoleLogon
Assert ($last.Status -eq 'Observed' -and $last.User -eq 'AzureAD\person' -and $last.Utc -eq '2026-10-13T10:00:00.0000000+00:00') 'The last console logon was not selected.'
Assert ($last.Sid -eq 'S-1-12-1-123') 'The event SID was not preserved.'
Assert ($script:observedXPath -match "LogonType']='2'" -and $script:observedXPath -match "LogonType']='11'" -and $script:observedXPath -notmatch "LogonType']='10'") 'The query includes remote or network sessions.'
$script:eventMode = 'systemOnly'
Assert ((Get-LastConsoleLogon).Status -eq 'NotFoundInScan') 'A system-only Security log was presented as a user.'
$script:eventMode = 'error'
Assert ((Get-LastConsoleLogon).Status -eq 'Error') 'A Security log failure was presented as a known user.'

$parsed = Parse-PlatformScriptStdout -ResultMessage 'SecureBoot:Enabled|LastConsoleLogonStatus:Observed|LastConsoleLogonUser:AzureAD\person|LastConsoleLogonUtc:2026-10-13T10:00:00Z'
Assert ($parsed.LastConsoleLogonStatus -eq 'Observed' -and $parsed.LastConsoleLogonUser -eq 'AzureAD\person') 'Console logon fields were not parsed.'
Assert ((Parse-PlatformScriptStdout -ResultMessage '').LastConsoleLogonStatus -eq 'NotCollected') 'Missing console evidence was inferred.'
Assert ($null -eq (Parse-PlatformScriptStdout -ResultMessage '').LastConsoleLogonSid) 'A missing SID was invented.'
$sidParsed=Parse-PlatformScriptStdout -ResultMessage 'LastConsoleLogonStatus:Observed|LastConsoleLogonSid:S-1-5-21-1-2-3-1001'
Assert ($sidParsed.LastConsoleLogonSid -eq 'S-1-5-21-1-2-3-1001') 'The SID payload was not parsed.'
$completePayload = 'SecureBoot:Enabled|BIOSVersion:1|BIOSDate:2026-01-01|FirmwareType:Uefi|LastBootUpTime:2026-01-01 10:00:00|ThirdPartyAVStatus:Detected|ThirdPartyAVProducts:Example|WindowsUpdateRebootRequired:False|BitLockerOSProtectionStatus:On|BitLockerOSConversionStatus:FullyEncrypted|LastConsoleLogonStatus:Observed|LastConsoleLogonUser:AzureAD\person|LastConsoleLogonUtc:2026-10-13T10:00:00Z|DeviceSystemPayloadEnd:1.5'
Assert (Test-DeviceSystemRecurringPayload $completePayload) 'A complete recurring payload was rejected.'
$sidPayload=$completePayload.Replace('LastConsoleLogonUtc:','LastConsoleLogonSid:S-1-5-21-1-2-3-1001|LastConsoleLogonUtc:').Replace('DeviceSystemPayloadEnd:1.5','DeviceSystemPayloadEnd:1.6')
Assert (Test-DeviceSystemRecurringPayload $sidPayload) 'The SID-capable payload was rejected.'
Assert (-not (Test-DeviceSystemRecurringPayload ($completePayload.Replace('DeviceSystemPayloadEnd:1.5','DeviceSystemPayloadEnd:1.6')))) 'A 1.6 payload without its SID field was accepted.'
Assert (-not (Test-DeviceSystemRecurringPayload ($completePayload -replace '\|DeviceSystemPayloadEnd:1.5$',''))) 'A truncated recurring payload was accepted.'

$asOf = [datetimeoffset]'2026-10-13T12:00:00Z'
Assert ((Get-DeviceSystemLocalFreshness -ObservationUtc '2026-10-13T10:00:00Z' -IsRecurring $true -HasOutput $true -DetectionState success -AsOfUtc $asOf) -eq 'Fresh') 'Fresh recurring evidence was rejected.'
Assert ((Get-DeviceSystemLocalFreshness -ObservationUtc '2026-10-01T10:00:00Z' -IsRecurring $true -HasOutput $true -DetectionState success -AsOfUtc $asOf) -eq 'Stale') 'Stale recurring evidence was accepted.'
foreach ($case in @(@('pending','Pending'),@('notApplicable','NotApplicable'),@('unknown','Unknown'),
    @('unknownFutureValue','Unknown'),@('future-state','Unknown'),@('fail','Failed'),@('scriptError','ScriptError'))) {
    foreach ($hasOutput in @($false,$true)) {
        $state=Get-DeviceSystemLocalFreshness -ObservationUtc '2026-10-13T10:00:00Z' -IsRecurring $true -HasOutput $hasOutput -DetectionState $case[0] -AsOfUtc $asOf
        Assert ($state -eq $case[1]) "Detection state $($case[0]) was misclassified."
    }
}
Assert ((Get-DeviceSystemLocalFreshness -ObservationUtc '2026-10-13T10:00:00Z' -IsRecurring $true -HasOutput $true -DetectionState ' PENDING ' -AsOfUtc $asOf) -eq 'Pending') 'Case or whitespace altered pending classification.'
Assert ((Get-DeviceSystemLocalFreshness -ObservationUtc '' -IsRecurring $true -HasOutput $true -DetectionState '' -AsOfUtc $asOf) -eq 'Unknown') 'Output without a detection verdict was qualified.'
Assert ((Get-DeviceSystemLocalFreshness -ObservationUtc '' -IsRecurring $true -HasOutput $false -DetectionState success -AsOfUtc $asOf) -eq 'NotCollected') 'Success without output became usable evidence.'
Assert ((Get-DeviceSystemLocalFreshness -ObservationUtc '' -IsRecurring $true -HasOutput $true -DetectionState success -AsOfUtc $asOf) -eq 'InvalidTimestamp') 'Success without an observation time became fresh.'
Assert ((Get-DeviceSystemLocalFreshness -ObservationUtc '2026-10-13T13:00:00Z' -IsRecurring $true -HasOutput $true -DetectionState success -AsOfUtc $asOf) -eq 'FutureObservation') 'A future successful observation became fresh.'
$collectorSource=Get-Content -LiteralPath $collectorPath -Raw
$runtimeVersion=[regex]::Match($collectorSource,'\$ScriptVersion = "([0-9.]+)"').Groups[1].Value
Assert ($runtimeVersion -and $collectorSource -match ('(?m)^\.VERSION\s*\r?\n\s+'+[regex]::Escape($runtimeVersion)+'\b') -and $collectorSource -match ('Version\s*:\s*'+[regex]::Escape($runtimeVersion)+'\b')) 'Collector version metadata disagrees.'
Assert ((Get-DeviceSystemLocalFreshness -ObservationUtc '2026-10-13T10:00:00Z' -IsRecurring $true -HasOutput $true -DetectionState success -PayloadValid $false -AsOfUtc $asOf) -eq 'InvalidOutput') 'A truncated output was accepted.'
Assert ((Get-DeviceSystemLocalFreshness -ObservationUtc '' -IsRecurring $true -HasOutput $false -DetectionState '' -AsOfUtc $asOf) -eq 'NotCollected') 'Missing output was accepted.'
Assert ((Get-DeviceSystemLocalFreshness -ObservationUtc '2026-10-13T10:00:00Z' -IsRecurring $false -HasOutput $true -DetectionState '' -AsOfUtc $asOf) -eq 'LegacySnapshot') 'Legacy output was mislabelled recurring.'

$RunStatePageSize = 500; $RunStatePagePauseMilliseconds = 0; $MaxRetries = 0; $BaseDelaySeconds = 0
function WriteLog { param($Message, $Level) }
$script:pages = @()
function Invoke-GraphSafe { param($Uri, $MaxRetries, $BaseDelaySeconds)
    $script:pages[0]
}
$script:pages = @(@{ value = @(@{managedDevice=@{id='device-one'};preRemediationDetectionScriptOutput='SecureBoot:Enabled';detectionState='success';lastStateUpdateDateTime='2026-10-13T10:00:00Z'}) })
$map = Get-DeviceHealthScriptRunStates -ScriptId '00000000-0000-0000-0000-000000000001'
Assert ($map.Count -eq 1 -and $map['device-one'].ResultMessage -eq 'SecureBoot:Enabled') 'Recurring Graph output was not mapped by managed device ID.'
$script:pages = @(@{ value = @(
    @{managedDevice=@{id='device-one'};preRemediationDetectionScriptOutput='first';detectionState='success';lastStateUpdateDateTime='2026-10-13T10:00:00Z'},
    @{managedDevice=@{id='device-one'};preRemediationDetectionScriptOutput='latest';detectionState='success';lastStateUpdateDateTime='2026-10-13T11:00:00Z'},
    @{managedDevice=@{id='device-two'};preRemediationDetectionScriptOutput='same';detectionState='success';lastStateUpdateDateTime='2026-10-13T10:00:00Z'},
    @{managedDevice=@{id='device-two'};preRemediationDetectionScriptOutput='same';detectionState='success';lastStateUpdateDateTime='2026-10-13T10:00:00Z'},
    @{managedDevice=@{id='device-three'};preRemediationDetectionScriptOutput='old';detectionState='success';lastStateUpdateDateTime='2026-10-13T10:00:00Z'},
    @{managedDevice=@{id='device-three'};preRemediationDetectionScriptOutput='failed';detectionState='scriptError';lastStateUpdateDateTime='2026-10-13T11:00:00Z'}
) })
$map = Get-DeviceHealthScriptRunStates -ScriptId '00000000-0000-0000-0000-000000000001'
Assert ($map.Count -eq 3) 'Three duplicate device run states were not collapsed.'
Assert ($map['device-one'].ResultMessage -eq 'latest') 'The latest device run state was not selected.'
Assert ($map['device-two'].ResultMessage -eq 'same') 'An identical run state was rejected.'
Assert ($map['device-three'].DetectionState -eq 'scriptError') 'A newer detection failure was hidden by an older success.'
$script:pages = @(@{ value = @(
    @{managedDevice=@{id='device-one'};preRemediationDetectionScriptOutput='newest';detectionState='success';lastStateUpdateDateTime='2026-10-13T11:00:00Z'},
    @{managedDevice=@{id='device-one'};preRemediationDetectionScriptOutput='older';detectionState='success';lastStateUpdateDateTime='2026-10-13T10:00:00Z'}
) })
$map = Get-DeviceHealthScriptRunStates -ScriptId '00000000-0000-0000-0000-000000000001'
Assert ($map['device-one'].ResultMessage -eq 'newest') 'An older run state replaced a newer one.'
$script:pages = @(@{ value = @(
    @{managedDevice=@{id='device-one'};preRemediationDetectionScriptOutput='first';detectionState='success';lastStateUpdateDateTime='2026-10-13T10:00:00Z'},
    @{managedDevice=@{id='device-one'};preRemediationDetectionScriptOutput='conflict';detectionState='success';lastStateUpdateDateTime='2026-10-13T10:00:00Z'}
) })
$failed = $false
try { $null = Get-DeviceHealthScriptRunStates -ScriptId '00000000-0000-0000-0000-000000000001' } catch { $failed = $true }
Assert $failed 'Conflicting run states with the same timestamp were accepted.'
$script:pages = @(@{ value = @(
    @{managedDevice=@{id='device-one'};preRemediationDetectionScriptOutput='first';detectionState='success';lastStateUpdateDateTime='not-a-date'},
    @{managedDevice=@{id='device-one'};preRemediationDetectionScriptOutput='second';detectionState='success';lastStateUpdateDateTime='2026-10-13T10:00:00Z'}
) })
$failed = $false
try { $null = Get-DeviceHealthScriptRunStates -ScriptId '00000000-0000-0000-0000-000000000001' } catch { $failed = $true }
Assert $failed 'Duplicate run states with an invalid timestamp were accepted.'
$script:pages = @(@{ value = @(@{managedDevice=@{id=''};preRemediationDetectionScriptOutput='SecureBoot:Enabled'}) })
$failed = $false
try { $null = Get-DeviceHealthScriptRunStates -ScriptId '00000000-0000-0000-0000-000000000001' } catch { $failed = $true }
Assert $failed 'An unjoinable recurring run state was accepted.'
$script:pages = @(@{ value = @() })
$failed = $false
try { $null = Get-DeviceHealthScriptRunStates -ScriptId '00000000-0000-0000-0000-000000000001' } catch { $failed = $true }
Assert $failed 'An empty recurring collection could replace the existing CSV.'

[pscustomobject]@{Passed=$true;Checks=$checks;Sources='Synthetic only';CollectorExecuted=$false;DeviceProbeExecuted=$false;GraphCalled=$false} | ConvertTo-Json

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC05lOm+a1Br9e+
# IwXfZpRiNnlPG2e3BS3psdt2uvxH36CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICRm8ZT65mOF6Wxs3eZ4gfWOp4XyzsVpdf8VGBTBeXDTMA0GCSqG
# SIb3DQEBAQUABIIBgJ+qdl6zPhy9F0kNAYey5q3qJXxx/8+VbfcVu6//wCjP9tta
# pvTGqyo3XRe2v53mJwZhDvqxyFrs7REL5tlxc1v9bX47VEY3GT9OjF/pnLb00QNI
# RLFVyjUEgg9oKnBH5hzSCLrT/IeeCXvihSfb3uOgqUUU1Jq13ujw3LVL4nAuyMMA
# 3ZqcbPfW9yYlPf7xlA/j8Xnxg9nAguF22h0l91ZUm/yqTxwvFeTvoRtmzeiQvZRG
# fO0aS3+6Mt1h6TxuQUqlLc4k7YxLIoJaGWH1otOIK5nxUM9/kzU1QQmyklYOokxx
# F6dkQSAJOb5gJIzp2uU0bf1+tpRHKXPAtNhn9oXrNGlQvWKeNtfu2HpGuuxARotq
# 1siQDnQ01I70+PJd9bpwF1Y9pavbaOrpEK33DOBZvYx4QR81Cc844U69ofvp4z9F
# M1GKxNGNEwx+UrUH+JhQ5kKk7doAno1gIdyZK99Qwz8twjdwJW7n9vWjMYCoPXCN
# oxjFFozcPFDjotLXHaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTQxNTMw
# NDFaMC8GCSqGSIb3DQEJBDEiBCClM7EvTg+WtyJlbrHfL39JwmjDE5eYIGr3llyX
# dpG0WzANBgkqhkiG9w0BAQEFAASCAgB1KDapE2bWdt0maTZoC+SZFf1LO/OAdzqP
# GG0wXhvg2nM58dDQNMZYU06RN5Qi6TzygZLqqosljpTPvY9WKSHaYM86sespsw5a
# LCRRI7Pblhc86IjV/4aW90UHVAGeuxJ9fDxiS4nJSpZ/oSjqYRWM3tsA07ObsNij
# L3TV1KBE+oJjIYYP4TiC2K2AdRS99bJ1EiTCx1F9NuiXOopT/VMjUfRSETOqObwM
# qcnpC8V20ZufntHfL/zvvt6oz6C+I3g0dTMyb01JSPuDsTafLz8SF9Ab8svmRNde
# gd8ecmVHfe79zxGaaDhZh7F1JN+HP7favBbshdQ8bNcuKtierMFQukGVkZnmmCD4
# AQoEthuJfTQGr20/2wST6V3Jg/RBC5mJ8gaWZntWKKpVECn5mtH6Ndho+NGVzE/3
# VhdJPVwiwNLJ/lk7yUDsOQLJW3yxRCynHLQxsobumLs+ksYGkJRw5s9d6U2mM0Yb
# qRB7BRRpOT3eALQp27uD00xcI2hTVkvSCPKbAKw3N0aQfY3yWKMJ58n58SwAVDdQ
# hRhppaw/NOGuOwWC6dzokGZGiTjtK11z7Kb2yyKOXA3Y6mm93pd4cBNyHU/KuVcV
# XyPYTZ6PilwNN8bcn8Ui+fYPGU4a1S4QxCpsK7+HUuFur3KtwAbKsXgVYUGbGFKk
# 1vMoF5kzsA==
# SIG # End signature block
