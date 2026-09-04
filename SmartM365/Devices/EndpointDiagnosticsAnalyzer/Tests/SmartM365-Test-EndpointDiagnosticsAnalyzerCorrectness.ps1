#Requires -Version 5.1
# Correctness contracts for 0.3.0. Synthetic, local-only fixtures; no tenant or AI calls.
[CmdletBinding()]param([string]$ReportPath='')
$ErrorActionPreference='Stop'
$app=Split-Path -Parent $PSScriptRoot
$gui=Join-Path $app 'SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1'
$tokens=$null;$errors=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($gui,[ref]$tokens,[ref]$errors)
if($errors.Count){throw ($errors.Message -join '; ')}
foreach($node in $ast.EndBlock.Statements){
 if($node -is [Management.Automation.Language.FunctionDefinitionAst]){Invoke-Expression $node.Extent.Text}
 elseif($node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -in @('$script:AppName','$script:AppVersion','$script:ImeThemes','$script:ImeThemeLabels','$script:MdmErrorCodes')){Invoke-Expression $node.Extent.Text}
}
function Write-SedaLog {param($Level,$Message,$Exception)}
$script:AiFocusContext=''
$root=Join-Path ([IO.Path]::GetTempPath()) ('seda_correctness_'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$script:LocalDataRoot=$root
$checks=New-Object 'System.Collections.Generic.List[object]'
function check([string]$Name,[bool]$Pass){$checks.Add([pscustomobject]@{Name=$Name;Passed=$Pass});if(-not$Pass){Write-Warning "FAIL: $Name"}}
function fixture([string]$Name,[string]$Text,[Text.Encoding]$Encoding=[Text.Encoding]::UTF8){$path=Join-Path $root $Name;New-Item -ItemType Directory -Path (Split-Path -Parent $path) -Force|Out-Null;[IO.File]::WriteAllText($path,$Text,$Encoding);return $path}
$bad=fixture 'results.xml' '<DiagnosticsResults><broken>'
check 'Malformed results are reported as failures' (@((Get-SedaResultsXml -Paths @($bad)).Errors).Count -gt 0)
$good=fixture 'valid-results.xml' '<DiagnosticsResults><CollectedFile><Name>test</Name><HRESULT>0</HRESULT></CollectedFile></DiagnosticsResults>'
check 'Valid results retain success' (@((Get-SedaResultsXml -Paths @($good)).Errors).Count -eq 0)
$badMdm=fixture 'MDMDiagReport.xml' '<MDMEnterpriseDiagnosticsReport><broken>'
check 'Invalid MDM XML is not Parsed' (-not(Get-SedaMdmDiagReport -XmlFiles @($badMdm) -CabFiles @() -HtmlFiles @() -WorkingDirectory $root).Parsed)
$unknownMdm=fixture 'unknown\MDMDiagReport.xml' '<Unrelated />'
check 'Unrelated XML schema is not MDM evidence' (-not(Get-SedaMdmDiagReport -XmlFiles @($unknownMdm)).Parsed)
$mdm1=fixture 'capture1\MDMDiagReport.xml' '<MDMEnterpriseDiagnosticsReport><SystemInformation><OSVersion>capture-one</OSVersion></SystemInformation></MDMEnterpriseDiagnosticsReport>'
$mdm2=fixture 'capture2\MDMDiagReport.xml' '<MDMEnterpriseDiagnosticsReport><SystemInformation><OSVersion>capture-two</OSVersion></SystemInformation></MDMEnterpriseDiagnosticsReport>'
$multi=Get-SedaMdmDiagReport -XmlFiles @($mdm1,$mdm2,$badMdm)
check 'Every MDM report has retained provenance' ($multi.SourceReports.Count -eq 3 -and $multi.CoverageState -eq 'Partial')
check 'Conflicting MDM values are retained' (($multi.DeviceInfo.Values -contains 'capture-one') -and ($multi.DeviceInfo.Values -contains 'capture-two'))
$copyNode=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'copy-collect-files'},$true)
Invoke-Expression $copyNode.Extent.Text
[void](fixture 'copy-source\one\same.log' 'first source')
[void](fixture 'copy-source\two\same.log' 'second source')
$copyCount=copy-collect-files (Join-Path $root 'copy-source') (Join-Path $root 'copied')
check 'Collection preserves identical basenames in different folders' ($copyCount -eq 2 -and @(Get-ChildItem (Join-Path $root 'copied') -Recurse -File -Filter '*.log').Count -eq 2)
$mixed=fixture 'WindowsUpdate.log' '2026/09/04 10:00:00.000 1 2 Agent Initialization completed successfully; Update failed 0x80070005'
$wu=@(Get-SedaWindowsUpdateLogEvents $mixed)
check 'Mixed WU success/failure retained with nonzero code' ($wu.Count -eq 1 -and $wu[0].ErrorCode -eq '0X80070005')
$zero=fixture 'wu-zero.log' '2026/09/04 10:00:00.000 1 2 Agent completed successfully; error = 0x00000000'
check 'WU success control produces no error' (@(Get-SedaWindowsUpdateLogEvents $zero).Count -eq 0)
$wifi=fixture 'wifi.log' 'All User Profile : AuditWiFi'
check 'Standard netsh profile is parsed' (@(Get-SedaWifiProfiles $wifi).Count -eq 1)
$battery=fixture 'battery.html' '<html><body>Collection unavailable</body></html>'
check 'Unrelated HTML is not a battery report' (-not(Get-SedaBatteryReport $battery).Parsed)
$neutral=fixture 'readiness.reg' "Windows Registry Editor Version 5.00`r`n`r`n[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\TargetVersionUpgradeExperienceIndicators\GE24H2]`r`n`"UpEx`"=`"Green`"`r`n`"RedReason`"=`"None`"`r`n`"GatedBlockId`"=dword:00000000`r`n`"SysReqIssue`"=dword:00000000"
check 'Neutral readiness fields do not block' (@((Get-SedaWin11CompatibilityIndicators $neutral).BlockingIndicators).Count -eq 0)
$device=[pscustomobject]@{Sections=@{'Device State'=@{AzureAdJoined='YES'};'Device Details'=@{DeviceAuthStatus='SUCCESS'};'User State'=@{NgcSet='YES';WamDefaultSet='YES'};'SSO State'=@{AzureAdPrt='YES'}};RawText='';DeviceInfo=@{}}
$assessment=Get-SedaComplianceSummary -DsReg $device -Enrollments ([pscustomobject]@{Enrollments=@()}) -Firewall ([pscustomobject]@{Profiles=@{}}) -Results ([pscustomobject]@{Errors=@()})
check 'Entra authentication is not MDM enrollment compliance' (@($assessment.PolicyStatuses | Where-Object { $_.Area -eq 'MDM Enrollment' -and $_.Status -eq 'COMPLIANT' }).Count -eq 0)
check 'Hello is read from User State' (@($assessment.PolicyStatuses | Where-Object Area -eq 'Hello for Business').Count -eq 1)
$empty=[pscustomobject]@{CriticalIssues=@();ImeEvents=@();Compliance=[pscustomobject]@{PolicyStatuses=@();OverallStatus='INSUFFICIENT_DATA'};WindowsUpdate=[pscustomobject]@{Issues=@();Info=@{};ReportingEvents=@();EtlEvents=@();Policies=@()};Win11Compatibility=[pscustomobject]@{Status='NoIndicators';BlockingIndicators=@();HardwareReadiness=[pscustomobject]@{Checks=@()}};Health=[pscustomobject]@{Findings=@()};EventLogs=[pscustomobject]@{Events=@()};DeviceSummary=[pscustomobject]@{ComputerName='Unknown'};DsReg=$device;ZipInfo=[pscustomobject]@{AllFiles=@()};ResultsXml=[pscustomobject]@{Items=@();Errors=@();Unavailable=@()};ErrorSummary=[pscustomobject]@{ScannedFiles=0;SkippedFiles=@()};MdmDiagnostics=[pscustomobject]@{Parsed=$false;Issues=@()}}
$insights=Get-SedaInsights $empty
check 'No evidence cannot produce 100/Healthy' ($insights.Status -eq 'Not assessed' -and $null -eq $insights.Score)
$failure=fixture 'extended\ps_pending_reboot.error.txt' 'Access denied'
$healthInput=[pscustomobject]@{ZipInfo=[pscustomobject]@{AllFiles=@($failure)};DsReg=$device;DeviceSummary=[pscustomobject]@{};EventLogs=[pscustomobject]@{Events=@()};Applications=@();Drivers=@();Hardware=[pscustomobject]@{FirewallProfiles=@();Battery=[pscustomobject]@{Parsed=$false};Certificates=@()}}
check 'Failed reboot collection is never OK' (@((Get-SedaHealthReport $healthInput).Findings | Where-Object { $_.Area -eq 'Restart' -and $_.Severity -eq 'OK' }).Count -eq 0)
$date=(Get-Date).ToString('M-d-yyyy',[Globalization.CultureInfo]::InvariantCulture)
$id1='11111111-2222-3333-4444-555555555555';$id2='aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
$imeText='<![LOG[Failed installing application '+$id1+']LOG]!><time="10:00:00.000" date="'+$date+'" component="App" type="3">'+"`r`n"+'<![LOG[Failed installing application '+$id2+']LOG]!><time="10:01:00.000" date="'+$date+'" component="App" type="3">'
$imePath=fixture 'IntuneManagementExtension.log' $imeText
$ime=Get-SedaImeLogEvents -ImeThemes @{intunemanagementextension=@($imePath)}
check 'Separate app identities retained' (@($ime.Events | Where-Object Severity -eq 'ERROR').Count -eq 2)
$success='<![LOG[Successfully installed application '+$id1+']LOG]!><time="10:02:00.000" date="'+$date+'" component="App" type="1">'
$imePath=fixture 'IntuneManagementExtension.log' ($imeText+"`r`n"+$success)
$ime=Get-SedaImeLogEvents -ImeThemes @{intunemanagementextension=@($imePath)}
check 'Matching later success resolves only its app' (@($ime.Events | Where-Object { $_.Severity -eq 'ERROR' -and $_.ResolvedByLaterSuccess }).Count -eq 1)
check 'Information records participate in correlation' ($ime.Summary.InformationalRecordsRead -eq 1)
$empty.ErrorSummary=$ime.Summary
check 'Real parser summary is accepted by coverage' ($null -ne (Get-SedaInsights -Analysis $empty).Coverage)
$empty.ImeEvents=@([pscustomobject]@{Severity='WARNING';IsExpected=$true;IsActionable=$false;IsRecent=$false;ResolvedByLaterSuccess=$false;Message='EXCLUDED_EXPECTED_WARNING';Category='IME'})
check 'AI excludes expected historical warning from issue list' ((Build-SedaAIPrompt $empty) -notmatch 'EXCLUDED_EXPECTED_WARNING')
function Read-SedaEventXml {param($Path,$Maximum)
 $script:fixtureEvents|Select-Object -First $(if($Maximum -gt 0){$Maximum}else{$script:fixtureEvents.Count})
}
$script:fixtureEvents=@(('<Event><System><Provider Name="Audit"/><EventID Qualifiers="0">1002</EventID><Level>2</Level><TimeCreated SystemTime="2026-09-04T10:00:00Z"/><Channel>Application</Channel><EventRecordID>1</EventRecordID></System><EventData><Data>'+('e'*650)+'</Data></EventData></Event>'),'<Event><System><Provider Name="Audit"/><EventID>1003</EventID><Level>4</Level><TimeCreated SystemTime="2026-09-04T10:01:00Z"/><Channel>Application</Channel><EventRecordID>2</EventRecordID></System><EventData><Data>Information context</Data></EventData></Event>')
$evtx=fixture 'Application.evtx' 'native boundary fixture'
$scan=Get-SedaEventLogScan -Inventory @{event_logs=@($evtx)}
check 'Qualified EventID is numeric text' (@($scan.Events | Where-Object Id -eq '1002').Count -eq 1)
check 'Full event message retained' (@($scan.Events | Where-Object { $_.Message.Length -eq 650 }).Count -eq 1)
check 'Informational event retained' (@($scan.Events | Where-Object Level -eq 'Information').Count -eq 1)
$limited=Get-SedaEventLogScan -Inventory @{event_logs=@($evtx)} -MaxEventsPerLog 1
check 'Explicit event cap is marked partial' ($limited.Sampled -and $limited.TotalEvents -eq 1)
Add-Type -AssemblyName System.IO.Compression.FileSystem
$privacy=Join-Path $root 'privacy';New-Item -ItemType Directory -Path $privacy|Out-Null
[IO.File]::WriteAllText((Join-Path $privacy 'audit.person@example.test.reg'),('User=audit.person@example.test '+$id1),[Text.Encoding]::Unicode)
$zip=Join-Path $root 'source.zip';[IO.Compression.ZipFile]::CreateFromDirectory($privacy,$zip)
$hash=(Get-FileHash $zip).Hash
$out=Join-Path $root 'redacted.zip';[void](Export-SedaAnonymizedZip $zip $out)
$archive=[IO.Compression.ZipFile]::OpenRead($out)
try{$entry=$archive.Entries[0];$reader=[IO.StreamReader]::new($entry.Open());try{$text=$reader.ReadToEnd()}finally{$reader.Dispose()};check 'UTF16 identifiers and entry names redacted' ($text -notmatch 'audit.person|11111111' -and $entry.FullName -notmatch 'audit.person')}finally{$archive.Dispose()}
check 'Original archive unchanged' ((Get-FileHash $zip).Hash -eq $hash)
[IO.File]::WriteAllText((Join-Path $privacy 'binary.evtx'),'unsupported')
$binary=Join-Path $root 'binary.zip';[IO.Compression.ZipFile]::CreateFromDirectory($privacy,$binary)
$blocked=$false;$blockedPath=Join-Path $root 'must-not-exist.zip'
try{Export-SedaAnonymizedZip $binary $blockedPath|Out-Null}catch{$blocked=$true}
check 'Unverifiable binary export fails closed' ($blocked -and -not(Test-Path $blockedPath))
$result=[pscustomobject]@{Result=if(@($checks|Where-Object {-not$_.Passed}).Count){'FAIL'}else{'PASS'};Engine=[string]$PSVersionTable.PSVersion;Checks=$checks.ToArray();FixturePath=$root}
if($ReportPath){$result|ConvertTo-Json -Depth 6|Set-Content -LiteralPath $ReportPath -Encoding UTF8}
$result
if($result.Result -ne 'PASS'){throw 'Correctness regression tests failed.'}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAMeFczuz2jS7v1
# yfUvDVSiSZFVxOW7Yih62lnFMYSmDKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJbPEc3DZS20vkZ3z8pxGArHfTnVl78MEZ7Q8mgauGjBMA0GCSqG
# SIb3DQEBAQUABIIBgGapl6mvkR9CC8ahsyJAsaKKgPVzCC8iU0GLUEylI7tPtnZ1
# Ksccv/WAYuu1TkPpfXIfO8lwjPVq0Up8b0C1GO+KJrvhGvZCspRxH2f1ZlNs81Is
# xfzNxt8V6Vnd37A32BIbeOqw907P7gJtLTQqbz91bkd6I7yVoUSbrow6wn9x2t/e
# hCq+XmyX8wYNl6p9aC1cKhTwnr+cvJPAZym1Q5RhdhM1A7LX0k3ttQGxt9sOqs1j
# PmPOfG4mSYxfCriolMEFm6diwsUgw/s8HfHEESS/Ph3v/SG4nYTMnT9as6Grn/z1
# cU93G9H/WxXGLckKcpqPIlbC4GzUiXl1QZewuGArRaU2iQjl1nBvyu2abg0RKPMV
# Vj4lms7kncRJ8/LhjVPA5k1mH30+rtPu0EOe+evDhccEEZaufqIDmnM0Qu1UuzOl
# fIZLqblIHCMrSirAjsqwKzZDpdTMvtJbvs3/sqiTRqA7Vqq/d0xOOj1kbf4t+J9V
# 9dAwk0X0fbEqnnwiZqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MDQxNTQw
# MzhaMC8GCSqGSIb3DQEJBDEiBCDgAobkG0fySFHsT4D/jvrCMxcxXAhyGbm6WHYe
# Tz5zozANBgkqhkiG9w0BAQEFAASCAgATsWJL0SN3ttXdGEVSXv09ayfnglZC9qru
# KFxwcZg2ZY3kDLzdtvvejkj6+vobUd7w72ESUAAZ650G4BdGBx1wN8T3IJc+SfZ/
# FSrQwVoa7qWm+rjc7uOnl0QvIInrC/yd56BjjWCR+0bkbnHFbHP4CyWUCKhvI94h
# NdB2m95A9NL6fceLRzjd2+4eI53PJg0kb5x+GJezqzXVLrvNaPpqrDNAbtUKywxI
# i0BRFSEoMEMm3y6nuZoJT9zxP0AOtfWzWjE8VlxtU0UEhsLnqNkPL3/2jmcv1Pin
# bjRkS7eDF2oqQw0bia65T5VrSVbqK+43DpxNCvdy4PWufed5geIDDoXfihmP4bM5
# UTpJ71wHa2PqMhuJPx8AOlB/4Q1vPcrMGWwf647cChy7IYInYefsol6b3AXnd3Yk
# mbFQL/c4SC142gQbqDJrshsNWZEJcmmWT1zQm0QrisxQQj6EORsbqFyeDvg8Y/ng
# RX0g6D5t9M/JYOYxWVzsX06f/fY6GPA47J5ZDND+uLgeire+NMHf2dAv7iVZ8WfG
# gTb/8ZxuTHypvVKMzaKwJRz8XjmC92cTmQe45wmNLqe49bKsLQVMFGyiXj1fU2J9
# kGb8WnVD6C+JHXpOoRJix4zNrWgAPrOCd0U7woqWgfRN0iC0Xwgq/+H+8dGBbrcm
# oMd8+b2beA==
# SIG # End signature block
