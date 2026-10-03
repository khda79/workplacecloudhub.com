<#
.SYNOPSIS
    Offline contract test for the approval-gated five-item ShareGate pilot.
.VERSION
    1.0.3
#>
#Requires -Version 5.1
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$testsRoot = (Resolve-Path -LiteralPath $PSScriptRoot).ProviderPath
$scratch = Join-Path $testsRoot ('.pilot-test-' + [guid]::NewGuid().ToString('N'))
if (-not $scratch.StartsWith($testsRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe test path.' }
$oldModulePath = $env:PSModulePath
$oldCalls = $env:SMARTM365_PILOT_TEST_CALLS
$oldErrorId = $env:SMARTM365_PILOT_TEST_ERROR_ID

try {
    . (Join-Path $testsRoot '..\Scripts\Diagnostics\SmartM365-SharePointMigration-FarmMaintenance.ps1')
    $farmZone = [TimeZoneInfo]::FindSystemTimeZoneById('W. Europe Standard Time')
    foreach ($case in @(
        @{ Local='2026-10-03 23:44:59'; Blocked=$false },
        @{ Local='2026-10-03 23:45:00'; Blocked=$true },
        @{ Local='2026-10-04 00:00:00'; Blocked=$true },
        @{ Local='2026-10-04 00:14:59'; Blocked=$true },
        @{ Local='2026-10-04 00:15:00'; Blocked=$false }
    )) {
        $local = [datetime]::SpecifyKind([datetime]::ParseExact($case.Local, 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture), [DateTimeKind]::Unspecified)
        $utc = [TimeZoneInfo]::ConvertTimeToUtc($local, $farmZone)
        $state = Get-SmartM365FarmMaintenanceState -FarmTimeZone $farmZone -UtcNow $utc
        if ($state.IsBlocked -ne $case.Blocked -or $state.FarmTime -ne $local) { throw "Maintenance boundary failed at $($case.Local)." }
    }
    $moduleDir = Join-Path $scratch 'Modules\ShareGate'
    $analysis = Join-Path $scratch 'Project\ShareGate\Diagnostics\Analysis'
    $witness = Join-Path $scratch 'Project\ShareGate\Diagnostics\Witness-Positive'
    $reports = Join-Path $witness 'Reports'
    New-Item -ItemType Directory -Path $moduleDir,$analysis,$reports -Force | Out-Null
    @'
@{ RootModule='ShareGate.psm1'; ModuleVersion='99.0.0'; GUID='2b63af52-c537-4e71-875e-fdf16ec46924'; FunctionsToExport=@('Connect-Site','Get-List','Get-File','Get-ListItem','Copy-Content','New-CopySettings','Export-Report') }
'@ | Set-Content -LiteralPath (Join-Path $moduleDir 'ShareGate.psd1') -Encoding UTF8
    @'
function Connect-Site { [CmdletBinding()] param([string]$Url,[switch]$Browser) if ($Url -match 'target' -and -not $Browser) { throw 'Destination Browser missing' }; if ($Url -match 'source' -and $Browser) { throw 'Source must use current Windows user' }; [pscustomobject]@{ Url=$Url } }
function Get-List { [CmdletBinding()] param($Site,[string[]]$Name) [pscustomobject]@{ Title=$Name[0]; Site=$Site; RootFolder=('/sites/a/' + $Name[0]) } }
function Get-File { [CmdletBinding()] param($List,[string]$Path) [pscustomobject]@{ Address=($List.Site.Url.TrimEnd('/') + '/' + $List.Title + '/' + $Path) } }
function Get-ListItem { [CmdletBinding()] param($List,[int]$Id) [pscustomobject]@{ Address=($List.Site.Url.TrimEnd('/') + '/' + $List.Title + '/item-' + $Id) } }
function New-CopySettings { [CmdletBinding()] param([string]$OnContentItemExists) if ($OnContentItemExists -ne 'IncrementalUpdate') { throw 'IncrementalUpdate missing' }; [pscustomobject]@{ OnContentItemExists=$OnContentItemExists } }
function Copy-Content { [CmdletBinding()] param($SourceList,$DestinationList,[int[]]$SourceItemId,$CopySettings,[string]$TaskName) if ($SourceItemId.Count -ne 1 -or $CopySettings.OnContentItemExists -ne 'IncrementalUpdate' -or $TaskName -notmatch '^SmartM365 401 pilot 260930-6 ') { throw 'Pilot copy scope or settings were changed' }; Add-Content -LiteralPath $env:SMARTM365_PILOT_TEST_CALLS -Value ($SourceItemId[0].ToString() + '|' + $TaskName); [pscustomobject]@{ Id=('261003-'+$SourceItemId[0]); SourceId=$SourceItemId[0]; Marker='synthetic' } }
function Export-Report { [CmdletBinding()] param($CopyResult,[string]$Path) $id=$CopyResult.SourceId; $result=if ($env:SMARTM365_PILOT_TEST_ERROR_ID -eq [string]$id) { 'Error' } else { 'Success' }; $itemPath=if ($id -eq 1) { 'Home.aspx' } else { 'image-' + $id + '.jpg' }; [pscustomobject]@{ Result=$result; 'Source ID'=$id; 'Destination path'=$itemPath; 'Destination ID'=(1000 + $id); 'Destination site address'='https://target.example/sites/a'; Error=$(if ($result -eq 'Error') { 'Synthetic item failure' } else { '' }) } | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 }
Export-ModuleMember -Function *
'@ | Set-Content -LiteralPath (Join-Path $moduleDir 'ShareGate.psm1') -Encoding UTF8

    $source = 'https://source.example/sites/a'
    $target = 'https://target.example/sites/a'
    $analysisRows = @()
    foreach ($id in @(88,185,300,467,600,868)) {
        $analysisRows += [pscustomobject]@{ SessionId='260930-6'; RuleId='SG-ACCESS-SOURCE'; State='To fix'; AccessSide='Source'; ItemKey=('photo-' + $id); SourceUrl=$source; SourceList='Photos'; SourceItemId=$id; DestinationUrl=$target; DestinationList='Photos'; ObjectType='File'; 'Raw: Source path'=('Album/image-' + $id + '.jpg'); 'Raw: Destination path'=('Album/image-' + $id + '.jpg') }
    }
    $analysisRows += [pscustomobject]@{ SessionId='260930-6'; RuleId='SG-ACCESS-SOURCE'; State='To fix'; AccessSide='Source'; ItemKey='page-1'; SourceUrl=$source; SourceList='Pages'; SourceItemId=1; DestinationUrl=$target; DestinationList='Pages'; ObjectType='File'; 'Raw: Source path'='Home.aspx'; 'Raw: Destination path'='Home.aspx' }
    $analysisPath = Join-Path $analysis 'ClassifiedRows.csv'
    $analysisRows | Export-Csv -LiteralPath $analysisPath -NoTypeInformation -Encoding UTF8
    $hash = (Get-FileHash -LiteralPath $analysisPath -Algorithm SHA256).Hash
    ('Actor=TEST; Session=260930-6; AnalysisSHA256=' + $hash) | Set-Content -LiteralPath (Join-Path $witness 'Witness.log') -Encoding UTF8
    $witnessRows = @()
    foreach ($role in @('WitnessShortcut','WitnessModernLink','AccessDominantFirst','AccessDominantLast','AccessOtherList')) {
        $id = switch ($role) { 'WitnessShortcut' { 2 } 'WitnessModernLink' { 25 } 'AccessDominantFirst' { 88 } 'AccessDominantLast' { 868 } 'AccessOtherList' { 1 } }
        $positive = $role -like 'Witness*'
        $list = if ($role -eq 'AccessOtherList') { 'Pages' } else { 'Photos' }
        $path = Join-Path $reports ($role + '.csv')
        if ($positive) { [pscustomobject]@{ Result='Warning'; Error='known permanent issue' } | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8 }
        else { Set-Content -LiteralPath $path -Value 'Result,Error' -Encoding UTF8 }
        $witnessRows += [pscustomobject]@{ Role=$role; SessionId='260930-6'; ItemKey=$(if ($role -eq 'AccessOtherList') { 'page-1' } else { 'photo-' + $id }); SourceUrl=$source; SourceList=$list; SourceItemId=$id; DestinationUrl=$target; DestinationList=$list; Status=$(if ($positive) { 'Pre-check warning' } else { 'Undetermined - empty report' }); ReportRows=$(if ($positive) { 1 } else { 0 }); ReportPath=$path; Error='' }
    }
    $witnessRows | Export-Csv -LiteralPath (Join-Path $witness 'Witness-Results.csv') -NoTypeInformation -Encoding UTF8
    $witnessHash = (Get-FileHash -LiteralPath (Join-Path $witness 'Witness-Results.csv') -Algorithm SHA256).Hash
    $env:PSModulePath = (Join-Path $scratch 'Modules') + [IO.Path]::PathSeparator + $oldModulePath
    $env:SMARTM365_PILOT_TEST_CALLS = Join-Path $scratch 'calls.txt'
    $scriptPath = Join-Path $testsRoot '..\Scripts\Diagnostics\SmartM365-SharePointMigration-ShareGatePilot.ps1'
    $project = Join-Path $scratch 'Project'
    $testZone = @([TimeZoneInfo]::GetSystemTimeZones() | Where-Object {
        $hour = [TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow, $_).Hour
        $hour -ge 10 -and $hour -le 14
    } | Select-Object -First 1)[0]
    if (-not $testZone) { throw 'No safe daytime zone is available for the offline pilot test.' }

    foreach ($switches in @(@{ Run=$true }, @{ Run=$true; ConfirmPilot=$true })) {
        try { & $scriptPath -ProjectRoot $project -AnalysisDirectory $analysis -WitnessDirectory $witness -SessionId '260930-6' -FarmTimeZoneId $testZone.Id @switches | Out-Null; throw 'Expected real-copy block.' }
        catch { if ($_.Exception.Message -notmatch 'remains disabled') { throw } }
    }
    try { & $scriptPath -ProjectRoot $project -AnalysisDirectory $analysis -WitnessDirectory $witness -SessionId '260930-6' -FarmTimeZoneId $testZone.Id -ExpectedAnalysisHash ('0' * 64) -DryRun | Out-Null; throw 'Expected analysis hash guard.' }
    catch { if ($_.Exception.Message -notmatch 'hash differs') { throw } }
    $dry = @(& $scriptPath -ProjectRoot $project -AnalysisDirectory $analysis -WitnessDirectory $witness -SessionId '260930-6' -FarmTimeZoneId $testZone.Id -ExpectedAnalysisHash $hash -DryRun)
    if (@($dry | Where-Object { $_ -match ': ID=(88|185|467|868|1);' }).Count -ne 5) { throw 'DryRun did not select the five intended 401 items.' }
    if (@($dry | Where-Object { $_ -match 'destinationFolder=Album' }).Count -ne 4 -or
        @($dry | Where-Object { $_ -match 'destinationFolder=<library root>' }).Count -ne 1) {
        throw 'DryRun did not route the four photos to Album and Home.aspx to the library root.'
    }
    if (Test-Path -LiteralPath $env:SMARTM365_PILOT_TEST_CALLS) { throw 'DryRun invoked Copy-Content.' }
    if (@(Get-ChildItem -LiteralPath (Join-Path $project 'ShareGate\Diagnostics') -Directory -Filter 'Pilot-*').Count) { throw 'DryRun created pilot output.' }
    Write-Output 'ShareGate pilot offline path plan and real-copy block passed.'
}
finally {
    Remove-Item Function:\Read-Host -ErrorAction SilentlyContinue
    $env:PSModulePath = $oldModulePath
    $env:SMARTM365_PILOT_TEST_CALLS = $oldCalls
    $env:SMARTM365_PILOT_TEST_ERROR_ID = $oldErrorId
    if ((Test-Path -LiteralPath $scratch) -and $scratch.StartsWith($testsRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $scratch -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBQJ1bynMC/T0Pg
# KTuleYLHIzVFeyvIpg8owSSWWQGKt6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJGRb0zyz6B2EaGYfpxPRjqsAwXgCvk6qWIql3MOymOxMA0GCSqG
# SIb3DQEBAQUABIIBgISuChwLvAYrMAkQdaDuUHi1aiAtAZN90z/N2lhg51t2bB0N
# 8p9Y+E1uTJdgmHgl3ZcVUYRCyEac+xqTc0k8/OAEka1IBKYM4cAmwVjbwUNPJgLS
# GeX6+0G0TwjMZmUesfE7x6CAYFhWSzIpb4WxRX+NhvAiSWgyaWflD16FIoR/eAbT
# lo0nW+VkbhrIEkp1yuNyGhCdQUX1EvdBypZtSLitEGAmgy7MmpjL2k0A23AU4iEp
# kauAR8OmGIQHckdy3Zj1Zg60geyM1psMzbEa+LqJavYZcrm0/VzIu3KA8dBoXMHK
# xax0QwZMSMDYD8mhSj58E1+D0oV0OeAb9BxAu7d/FJuvuTTYt+rYAj7mcPSZVs6s
# FlsMg0M34xymcdQ8t+MM33tP08C9qZ98HZTcbvy87zqNTYg+XCWibrIqY7hfWHey
# uxwkjzjBb4k03crgXib96Pgdt36haTxssntfNIs0hDwcvxxyUxNo/KOKXUjRGJ0r
# EgWY9xl3ezVcb2eoJ6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxNTQ5
# MjFaMC8GCSqGSIb3DQEJBDEiBCAJqGAJUZ+ySE2ATUqQp2vfmD+JFI3YIbCupKTa
# G55gPTANBgkqhkiG9w0BAQEFAASCAgCoNYg+3sye1SasK8IPUZisS86PqQdIcmKe
# zJly8xRGTLbpqGbLfxRAprHHf1d/6/OYR2T7myen67Ruz/GaRZbfc/LWFRt5yHS3
# VzYs3OIOpDBEMctV/AWmdN+fi/wAZrSk6ZVxOJVBJ9h/afJmxcC4icB9i8iTdLy7
# 2JZ2xMsOx2CI2fKOIeS7S8Wed5W2UAMUbu7PwiNhn+CD58h87+2hgX9/UIfcNvdk
# SjX7dSVwgr8R8ynhK5jx+9e5EmqQisn0vcvJkAZcIkJYO1daMLlzO96bVJYS29Ys
# Xv6Vu+X5ESC4kM4+gSyDyPtqHASBpQtYrxgah5O/JNOEHfvDHbp4a+a1Tk/h1fG5
# 2QA1knF88Ak7zHm6nyh8jgfxSvvG7XSWDHQjvRjZFvJcHVilQchnembYTjagyAK3
# frXe0BxIuKN+KdSuFf46tmVRIWsP6dQx3qw0VfPZ0hFqrkOPxMpxkLETz8PkmI3x
# JzOLORgHc/4Oe9L3gRERXiu87cqFcphpZkAcJ5XalnKGEvKDIAPtlgD/nGcMvYA6
# cGkHzXDHB1xgiyrS3mckfvBJdqHgi35X5Uo7HnavhEqAFFps8/bblwbTiEZI3jd0
# HaPtSwlSvPAZyx/nOPucawpCulv2TtUYUWBiwgTXFqh1T0UbWUFAxJ95c0OMUDz2
# kE24ESA0tw==
# SIG # End signature block
