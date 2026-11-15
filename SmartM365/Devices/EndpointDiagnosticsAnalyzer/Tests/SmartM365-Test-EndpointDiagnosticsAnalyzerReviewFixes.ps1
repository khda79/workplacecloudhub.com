#Requires -Version 5.1
<#
.SYNOPSIS
Exercises archive safety and extraction ownership with synthetic offline fixtures.
.VERSION
1.0.1
#>
[CmdletBinding()]param()
$ErrorActionPreference = 'Stop'
$gui = Join-Path (Split-Path -Parent $PSScriptRoot) 'SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($gui,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw ($errors.Message -join '; ') }
foreach ($name in @('New-SedaObject','Get-SedaZipCategory','Assert-SedaSafeZip','Expand-SedaSelectedZipEntries','Expand-SedaDiagnosticZip','Remove-SedaAnalysisExtraction','Remove-SedaStaleExtractDirectories','Set-SedaCurrentAnalysis','Expand-SedaCabFile')) {
    $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name },$true)
    if (-not $node) { throw "Missing function: $name" }
    Invoke-Expression $node.Extent.Text
}
function Write-SedaLog { param($Level,$Message,$Exception) }
function Assert-Review([bool]$condition,[string]$message) { if (-not $condition) { throw "FAIL: $message" }; $script:count++ }
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression
$root = Join-Path ([IO.Path]::GetTempPath()) ('seda reviewfix ' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root | Out-Null
$script:count = 0
try {
    foreach ($names in @(@('a.txt','a.txt'),@('A.txt','a.txt'),@('dir/../a.txt','a.txt'),@('dir/name.','dir/name'))) {
        $path = Join-Path $root ([guid]::NewGuid().ToString('N') + '.zip')
        $archive = [IO.Compression.ZipFile]::Open($path,[IO.Compression.ZipArchiveMode]::Create)
        try { foreach ($name in $names) { $writer = New-Object IO.StreamWriter($archive.CreateEntry($name).Open()); try { $writer.Write($name) } finally { $writer.Dispose() } } } finally { $archive.Dispose() }
        $rejected = $false
        try { Assert-SedaSafeZip -Path $path -DestinationRoot (Join-Path $root 'out') } catch { $rejected = $true }
        Assert-Review $rejected "Conflicting ZIP names rejected: $($names -join ', ')"
    }
    $invalidCab = Join-Path $root 'invalid.cab'; [IO.File]::WriteAllText($invalidCab,'not a cabinet')
    $destination = Join-Path $root 'existing'; New-Item -ItemType Directory $destination | Out-Null
    [IO.File]::WriteAllText((Join-Path $destination 'preexisting.txt'),'preserve')
    $cabResult = Expand-SedaCabFile -CabPath $invalidCab -DestinationPath $destination
    Assert-Review (-not $cabResult.Success -and $cabResult.Status -match 'validation failed') 'Invalid CAB rejected'
    Assert-Review (@(Get-ChildItem $destination -File).Count -eq 1) 'Existing destination preserved'
    [IO.File]::WriteAllText($invalidCab,'MSCFbad')
    Assert-Review (-not (Expand-SedaCabFile -CabPath $invalidCab -DestinationPath $destination).Success) 'Truncated CAB with matching signature rejected'
    $fakeCab = New-Object byte[] 36
    [Array]::Copy([Text.Encoding]::ASCII.GetBytes('MSCF'),0,$fakeCab,0,4)
    [Array]::Copy([BitConverter]::GetBytes([uint32]36),0,$fakeCab,8,4)
    [Array]::Copy([BitConverter]::GetBytes([uint16]1),0,$fakeCab,26,2)
    [Array]::Copy([BitConverter]::GetBytes([uint16]1),0,$fakeCab,28,2)
    [IO.File]::WriteAllBytes($invalidCab,$fakeCab)
    Assert-Review (-not (Expand-SedaCabFile -CabPath $invalidCab -DestinationPath $destination).Success) 'Copied corrupt CAB is not reported as extracted'
    $report = Join-Path $root 'MDMDiagReport.xml'; [IO.File]::WriteAllText($report,'<MDMEnterpriseDiagnosticsReport/>')
    $cab = Join-Path $root 'valid.cab'
    Push-Location $root
    try { & "$env:WINDIR\System32\makecab.exe" $report $cab | Out-Null } finally { Pop-Location }
    if ($LASTEXITCODE -ne 0) { throw 'Could not create CAB fixture.' }
    $cabResult = Expand-SedaCabFile -CabPath $cab -DestinationPath $destination
    if (-not $cabResult.Success) { throw "Valid CAB extraction failed: $($cabResult.Status)" }
    Assert-Review ($cabResult.Files.Count -eq 1 -and ([IO.File]::ReadAllText($cabResult.Files[0]) -eq '<MDMEnterpriseDiagnosticsReport/>')) 'Valid CAB extracts into an isolated child'
    Assert-Review ((Test-Path (Join-Path $destination 'preexisting.txt')) -and $cabResult.ExtractPath -ne $destination) 'Preexisting files excluded from result'
    $old = Join-Path ([IO.Path]::GetTempPath()) ('seda_diag_' + [guid]::NewGuid().ToString('N'))
    $new = Join-Path ([IO.Path]::GetTempPath()) ('seda_diag_' + [guid]::NewGuid().ToString('N'))
    try {
        New-Item -ItemType Directory $old,$new | Out-Null
        $script:CurrentAnalysis = [pscustomobject]@{ExtractDir=$old}
        Set-SedaCurrentAnalysis ([pscustomobject]@{ExtractDir=$new})
        Assert-Review (-not(Test-Path $old) -and (Test-Path $new)) 'Replacing analysis releases only the old extraction'
        Set-SedaCurrentAnalysis $null
        Assert-Review (-not(Test-Path $new)) 'Reset releases current extraction'
    } finally { Remove-SedaAnalysisExtraction $old; Remove-SedaAnalysisExtraction $new }
    $source = [IO.File]::ReadAllText($gui)
    Assert-Review ($source.Contains('Identifiers may remain. Manually review all content and filenames before sharing; anonymity is not guaranteed.')) 'Export message accurately describes limitations'
    Assert-Review ($source -notmatch 'Anonymized ZIP saved:') 'Export message does not promise anonymity'
    $zip = Join-Path $root 'cli.zip'
    $archive = [IO.Compression.ZipFile]::Open($zip,[IO.Compression.ZipArchiveMode]::Create)
    try { $writer = New-Object IO.StreamWriter($archive.CreateEntry('Application.evtx').Open()); try { $writer.Write('synthetic') } finally { $writer.Dispose() } } finally { $archive.Dispose() }
    $savedTemp = $env:TEMP; $savedTmp = $env:TMP; $savedLocal = $env:LOCALAPPDATA
    $savedGui = $env:SEDA_FIX_GUI; $savedZip = $env:SEDA_FIX_ZIP; $savedExtract = $env:SEDA_FIX_EXTRACT; $savedOutput = $env:SEDA_FIX_OUTPUT; $savedMode = $env:SEDA_FIX_MODE
    try {
        $env:TEMP=$root; $env:TMP=$root; $env:LOCALAPPDATA=$root; $env:SEDA_FIX_GUI=$gui; $env:SEDA_FIX_ZIP=$zip
        $exe = (Get-Process -Id $PID).Path
        $command = '$ErrorActionPreference="Stop"; $body=[IO.File]::ReadAllText($env:SEDA_FIX_GUI).Replace(''$MyInvocation.MyCommand.Path'',''$env:SEDA_FIX_GUI''); $p=@{Cli=$true;ZipPath=$env:SEDA_FIX_ZIP;AnalysisExtractionPath=$env:SEDA_FIX_EXTRACT;KeepExtractedFiles=$true;SkipEventLogScan=$true;SkipCabExtraction=$true}; if($env:SEDA_FIX_MODE -eq "fail"){$p.ExportAnonymizedZipPath=$env:SEDA_FIX_OUTPUT}; & ([scriptblock]::Create($body)) @p'
        $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
        foreach ($mode in @('fail','keep')) {
            $env:SEDA_FIX_MODE=$mode; $env:SEDA_FIX_EXTRACT=Join-Path $root ('seda_diag_'+[guid]::NewGuid().ToString('N')); $env:SEDA_FIX_OUTPUT=Join-Path $root 'redacted.zip'
            $info=New-Object Diagnostics.ProcessStartInfo; $info.FileName=$exe; $info.Arguments="-NoProfile -EncodedCommand $encoded"; $info.UseShellExecute=$false; $info.CreateNoWindow=$true; $info.RedirectStandardOutput=$true; $info.RedirectStandardError=$true
            $child=[Diagnostics.Process]::Start($info)
            try { $stdout=$child.StandardOutput.ReadToEndAsync(); $stderr=$child.StandardError.ReadToEndAsync(); if(-not $child.WaitForExit(120000)){ $child.Kill(); throw 'CLI fixture timed out.' }; $null=$stdout.Result; $errorText=$stderr.Result; $code=$child.ExitCode } finally { $child.Dispose() }
            if ($mode -eq 'fail') { Assert-Review ($code -ne 0 -and $errorText -match 'binary or unsupported' -and -not(Test-Path $env:SEDA_FIX_EXTRACT) -and -not(Test-Path $env:SEDA_FIX_OUTPUT)) 'Failed export cleans extracted evidence even with KeepExtractedFiles' }
            else { if ($code -ne 0) { throw "CLI success fixture failed: $errorText" }; Assert-Review (Test-Path $env:SEDA_FIX_EXTRACT -PathType Container) 'Successful KeepExtractedFiles retains extraction' }
        }
    } finally {
        $env:TEMP=$savedTemp; $env:TMP=$savedTmp; $env:LOCALAPPDATA=$savedLocal; $env:SEDA_FIX_GUI=$savedGui; $env:SEDA_FIX_ZIP=$savedZip; $env:SEDA_FIX_EXTRACT=$savedExtract; $env:SEDA_FIX_OUTPUT=$savedOutput; $env:SEDA_FIX_MODE=$savedMode
    }
    $app=Split-Path -Parent $PSScriptRoot
    foreach($pair in @(@((Join-Path $app 'SmartM365.GuiDemo.ps1'),'Get-SmartM365DemoAst'),@((Join-Path $app 'SmartM365-EndpointDiagnosticsAnalyzer-Demo.ps1'),'script:Export-SmartM365DemoProfileOutputs'))){
        $parsedTokens=$null;$parsedErrors=$null
        $parsed=[Management.Automation.Language.Parser]::ParseFile($pair[0],[ref]$parsedTokens,[ref]$parsedErrors)
        if($parsedErrors.Count){throw 'Demo source parse failed.'}
        $functionName=$pair[1]
        $functionNode=$parsed.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $functionName},$true)
        if(-not $functionNode){throw "Demo function missing: $functionName"}
        Invoke-Expression $functionNode.Extent.Text
    }
    $demoReport=Join-Path $root 'demo-report.html'
    function Export-SmartM365DemoNativeReport {param($Kind,$SourcePath,$Name,$Data);[IO.File]::WriteAllText($demoReport,'<title>Analyzer DEMO v - sample</title><h1>Analyzer DEMO v</h1>');return $demoReport}
    $script:SmartM365Demo=[pscustomobject]@{GuiPath=$gui;Data=[pscustomobject]@{Analysis=$null}}
    $generated=@(Export-SmartM365DemoProfileOutputs)
    $content=[IO.File]::ReadAllText($generated[0])
    $expectedVersion=(Get-Content -LiteralPath (Join-Path $app 'SmartM365-EndpointDiagnosticsAnalyzer.version.json.txt') -Raw|ConvertFrom-Json).PackageVersion
    $escaped=[regex]::Escape([string]$expectedVersion)
    Assert-Review ($content -match ('<title>[^<]* v'+$escaped+' - sample</title>') -and $content -match ('<h1>[^<]* v'+$escaped+'</h1>')) 'Demo report title and heading contain the current version'
    [pscustomobject]@{Result='PASS';Checks=$script:count;Engine=[string]$PSVersionTable.PSVersion}
} finally { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAMifsVIR792gHF
# bFHkEH/I+cLM/XuKmnoV9h8cUEUuCKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOLyCM07MtmgZgcHtLtrWriNewifr/owkCXC0UNmKddhMA0GCSqG
# SIb3DQEBAQUABIIBgAcwZLoGH4KG00m8hPAjDfzC/0lBYN9dlLuy4QmAeAz+247C
# 0cv6cDq7DLxf+QLd38vJBklze9yxT4m3oALgWxS7JCvdL176YzoVdb5vG88vH8Vy
# MJkuJZRdvHDHKK2WeuUiRKu1Ez0xqScvY6ER8/qnpnn3R7k1UMpRnxPjVekQFszU
# m3RZ/Ng6iyFa0S24R2CCCh6Edfebb4Z6nZccAHjMPR5FUNlc1BDeiSl72oCuuQNv
# 9FFLRoiOMEkPk3cCgQInwnHfiFsWRT0vHcGld6j+d485rGtz/4MSTzLTR6w96plK
# 0rgYzRK3boRZcKyosU3QkXw9eWa1lTScbDM9t6/ssZECXIiXrAVTjTrowv8cpOrd
# eeO5KyhmD5bvsN5Oz8MPHfx1RO1yVM8qVN1ssqOTfZbU3bEK8C0Viqykm7uwVh6B
# K6gqEb9IWkYEKXY2Z9/r6Gydaf7Nq5bvOEBz/GlTPacFIYlBSqDykYeJKckiqie5
# KCaBE7JWggjJIX7h6aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTUyMzU3
# MjlaMC8GCSqGSIb3DQEJBDEiBCDIveCbiNvcboIxYVmlPjIV84241VJIb2M/tuXo
# 0KdWTzANBgkqhkiG9w0BAQEFAASCAgApWhhY3r5Y4YrWl8Ureu7DBOLdTq4dbQ71
# vcaInmkoNDnJKqfP7jnmkGGFX7kfUJil2wXs1MlwHUoXodFn/oJH9S7iDlstlFeD
# nXO1Pq/F6tye8pb/3L8RNO32XnhuxOZC53QWYVOAgm0dZ+AxvviIaDWl1Q2FTaf2
# ZJ86M8k7jC1v6xm9mvlSuemFZ8H9PynzV6b+4UP/ZeNnF50AJlU42LP21zHL3JqT
# O96b5IQhLtMfGB82UWZLFDEup72Pl2qV4qP9TieXLRuWc77B81nzlZHK3PzAfKRS
# T2Re0QKxrvaEwPsgtknUSwkQy5MIzw5A94/lOkEkxlwzDitXj+uaA+LJ9rgYaqJq
# truh/6j22HQrd7cBYgi4YExSWyfOooBfBJxQmJfH2KeivRCJ4w4Zv6ZXzPS8hdO4
# eacmESRMQcTjL7n8O11RXAZHWimaEOikNJJybJM7b6bechlqje+kn2haE+Jz3wEO
# MAhyQreweUDdpeuoSsrc0OGohGCuEDkqpuo3r74DJH288i7r94gEzfWaXmxtW6jl
# pPCAzZ5XUSK/WOsaTXVMXHxLLamMoTbXbnXTInZTNARN156ymZcjGOvf0a5B2Jsi
# SJ7Fmc7rmW4XID7S9kmFfLHf1WqInQyu+om+blgJhY7cZbV7/HN9KsTV6o+Y9876
# L0oG09NPtA==
# SIG # End signature block
