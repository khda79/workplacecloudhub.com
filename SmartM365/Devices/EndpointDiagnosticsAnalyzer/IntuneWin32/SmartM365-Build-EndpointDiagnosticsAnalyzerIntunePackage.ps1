#Requires -Version 5.1
<# .VERSION
1.0.2
#>

[CmdletBinding()]
param(
    [string]$IntuneWinAppUtilPath = 'C:\tmp\SmartM365Tools\IntuneWinAppUtil.exe',
    [string]$OutputRoot = (Join-Path ([IO.Path]::GetTempPath()) 'SmartM365\IntuneWin32\EndpointDiagnosticsAnalyzer'),
    [string]$GalleryBuildRoot = (Join-Path ([IO.Path]::GetTempPath()) 'SmartM365\PowerShellGallery'),
    [switch]$Force,
    [switch]$SkipAuthenticodeValidation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-WchDemoAssetPath {
    param([string]$Path)
    $name=$Path.Replace('\','/')
    return $name -match '(?i)(^|/)DemoAssets(/|$)|(^|/)[^/]*demo[-_ ]?assets[^/]*\.(zip|nupkg|7z|rar|tar|tgz|gz|intunewin)$'
}

function Assert-WchNoDemoAssets {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string[]]$Path)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $budget=@{Entries=0;Bytes=[long]0}
    function Read-WchReleaseArchive {
        param([IO.Stream]$Stream,[string]$Name,[int]$Depth)
        if($Depth -gt 8){throw "Release archive nesting limit exceeded: $Name"}
        $zip=[IO.Compression.ZipArchive]::new($Stream,[IO.Compression.ZipArchiveMode]::Read,$true)
        try {
            foreach($entry in $zip.Entries){
                $budget.Entries++;$budget.Bytes+=$entry.Length
                if($budget.Entries -gt 100000 -or $budget.Bytes -gt 536870912){throw "Release archive inspection limit exceeded: $Name"}
                if(Test-WchDemoAssetPath $entry.FullName){throw "DemoAssets are local-only: $Name!$($entry.FullName)"}
                if($entry.Name){
                    $content=$entry.Open();$memory=$null
                    try {
                        $head=New-Object byte[] 4;$read=$content.Read($head,0,4)
                        $magic=$read -eq 4 -and $head[0] -eq 80 -and $head[1] -eq 75 -and (($head[2] -eq 3 -and $head[3] -eq 4) -or ($head[2] -eq 5 -and $head[3] -eq 6) -or ($head[2] -eq 7 -and $head[3] -eq 8))
                        if($magic -or $entry.FullName -match '(?i)\.(zip|nupkg)$'){
                            $memory=[IO.MemoryStream]::new();$memory.Write($head,0,$read);$content.CopyTo($memory);$memory.Position=0
                            Read-WchReleaseArchive $memory ($Name+'!'+$entry.FullName) ($Depth+1)
                        }
                    }finally{$content.Dispose();if($memory){$memory.Dispose()}}
                }
            }
        }finally{$zip.Dispose()}
    }
    foreach($inputPath in $Path){
        if(Test-WchDemoAssetPath $inputPath){throw "DemoAssets are local-only: $inputPath"}
        $item=Get-Item -LiteralPath $inputPath -Force -ErrorAction Stop
        $pending=[Collections.Generic.Stack[IO.FileSystemInfo]]::new();$pending.Push($item)
        while($pending.Count){
            $file=$pending.Pop()
            if(Test-WchDemoAssetPath $file.FullName){throw "DemoAssets are local-only: $($file.FullName)"}
            if($file.Attributes -band [IO.FileAttributes]::ReparsePoint){throw "Release input must not be a reparse point: $($file.FullName)"}
            if($file.PSIsContainer){foreach($child in @(Get-ChildItem -LiteralPath $file.FullName -Force -ErrorAction Stop)){$pending.Push($child)}}else{
                $stream=[IO.File]::OpenRead($file.FullName)
                try {
                    $head=New-Object byte[] 4;$read=$stream.Read($head,0,4);$stream.Position=0
                    $magic=$read -eq 4 -and $head[0] -eq 80 -and $head[1] -eq 75 -and (($head[2] -eq 3 -and $head[3] -eq 4) -or ($head[2] -eq 5 -and $head[3] -eq 6) -or ($head[2] -eq 7 -and $head[3] -eq 8))
                    if($magic -or $file.Extension -match '(?i)^\.(zip|nupkg)$'){Read-WchReleaseArchive $stream $file.FullName 0}
                }finally{$stream.Dispose()}
            }
        }
    }
}


$applicationRoot = Split-Path -Parent $PSScriptRoot
$versionPath = Join-Path $applicationRoot 'SmartM365-EndpointDiagnosticsAnalyzer.version.json.txt'
$version = Get-Content -LiteralPath $versionPath -Raw | ConvertFrom-Json
$packageVersion = [string]$version.PackageVersion
if ([string]::IsNullOrWhiteSpace($packageVersion)) {
    throw 'PackageVersion is missing from the version manifest.'
}
if (-not (Test-Path -LiteralPath $IntuneWinAppUtilPath -PathType Leaf)) {
    throw "IntuneWinAppUtil.exe not found: $IntuneWinAppUtilPath"
}

$buildScript = Join-Path $applicationRoot 'PowerShellGallery\SmartM365-Build-EndpointDiagnosticsAnalyzerGalleryPackage.ps1'
$buildParameters = @{
    OutputRoot = $GalleryBuildRoot
    Force = $Force
    SkipAuthenticodeValidation = $SkipAuthenticodeValidation
}
$galleryBuild = & $buildScript @buildParameters
$sourcePath = [string]$galleryBuild.IntuneSourcePath
$setupFile = 'Deploy\SmartM365-EndpointDiagnosticsAnalyzer-Install.ps1'
if (-not (Test-Path -LiteralPath (Join-Path $sourcePath $setupFile) -PathType Leaf)) {
    throw "Intune setup file is missing from the staging source: $setupFile"
}

$resolvedOutputRoot = [IO.Path]::GetFullPath($OutputRoot)
$outputPath = Join-Path $resolvedOutputRoot $packageVersion
$normalizedRoot = $resolvedOutputRoot.TrimEnd('\') + '\'
if (-not ([IO.Path]::GetFullPath($outputPath).StartsWith($normalizedRoot, [StringComparison]::OrdinalIgnoreCase))) {
    throw "Unsafe Intune output path: $outputPath"
}
if (Test-Path -LiteralPath $outputPath) {
    if (-not $Force) { throw "Intune build target already exists: $outputPath" }
    Remove-Item -LiteralPath $outputPath -Recurse -Force
}
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null

Assert-WchNoDemoAssets -Path $sourcePath
& $IntuneWinAppUtilPath -c $sourcePath -s $setupFile -o $outputPath -q
if ($LASTEXITCODE -ne 0) {
    throw "IntuneWinAppUtil failed with exit code $LASTEXITCODE."
}
$generatedName = ([IO.Path]::GetFileNameWithoutExtension($setupFile) + '.intunewin')
$generatedPath = Join-Path $outputPath $generatedName
if (-not (Test-Path -LiteralPath $generatedPath -PathType Leaf)) {
    $generatedPath = Get-ChildItem -LiteralPath $outputPath -Filter '*.intunewin' -File |
        Select-Object -First 1 -ExpandProperty FullName
}
if ([string]::IsNullOrWhiteSpace([string]$generatedPath) -or -not (Test-Path -LiteralPath $generatedPath -PathType Leaf)) {
    throw 'IntuneWinAppUtil did not produce an .intunewin file.'
}
$finalPath = Join-Path $outputPath ("SmartM365-EndpointDiagnosticsAnalyzer-{0}.intunewin" -f $packageVersion)
if (-not [IO.Path]::GetFullPath($generatedPath).Equals([IO.Path]::GetFullPath($finalPath), [StringComparison]::OrdinalIgnoreCase)) {
    Move-Item -LiteralPath $generatedPath -Destination $finalPath -Force
}

[pscustomobject]@{
    Result = 'PASS'
    PackageVersion = $packageVersion
    SourcePath = $sourcePath
    SetupFile = $setupFile
    IntuneWinPath = $finalPath
    Sha256 = (Get-FileHash -LiteralPath $finalPath -Algorithm SHA256).Hash
    AutomaticUpdateEnabled = $false
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDPCE8xgZRPtcOP
# l+/iBKfBmzvjQaiuJRjNzXutLgRbVqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDrhZSZBngv33xBfKk5wKqqZa0HqIYOPrvKZ7R2AE5lPMA0GCSqG
# SIb3DQEBAQUABIIBgELyAw0r51Ip2b/MTFRMOYlpn1RzSydoVo6fI7oMQUcDFFkf
# g3Jp4wl47x1UELsvNu2E0dXm/7rDmLXOPx7DEP42d7nW6Wiyhdm++W0FQIR6izY9
# rtwyQj1JXxumisSBRkwObckoqWAdeMTAZI0JPokSlh/3PNReiLTJ5SZHv/47iA9j
# MaBXi0HEH4kBK+DF0cGjtCUL85gYaiwkmiGNIHw1RiYM3dtqxauOX7WaQsr1ywdD
# ZsSYm3aTZpEj/3VfonecalR3D7yQAlurMbzowJJkwBCujWeKVcpaj7oIxPUCkzpR
# KGHs+YLME+fYVcZGTk9sRjczl6x/kRt3rFJgdFnBsqhmf2jG9loGJi/vtZfcn0Ex
# hafu/mq4VkOIxGf5/AJl0ZzLlUiqQvfLrgs5fkPWEkDY1kWQqA2nDMeIB5abHPmi
# 8WXQ3u1PwSX7rfbwjpUT3UXCpiBKy7H8xyP6VI9p543ctKzbd4Q+cwHXWu5yYZxd
# 47xIxZjVEvvjUinxJ6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTMxMDE1
# NTJaMC8GCSqGSIb3DQEJBDEiBCAvPyejgmXgg89uLU1C+w8f1E5xuy4svuahLDHr
# TDkxdjANBgkqhkiG9w0BAQEFAASCAgBxW8xXtkmZVKJ8+6ma/hzBWXVKEGZZuMwP
# gKzyN4D4W0CO5zb6UI6tNpgmSOR748pgeUKZD/54GFwxI4OdntBpoV3homo+NJp2
# 3Qpnlti9JC7j9F4h84hvtvqs4a81vB7SMtbvdIWMR+Xc741dJUz0EZBp26jLbyfK
# rXqKzJt2g1li/6Msf2U8BhWxvBLoyq9a0O/veYchurZHDUtoIQ2FpUbJDJ+Kpcjw
# opPofHju6HjhyLPULbgFETeQiVzU4feffXG47XQR7RwrqQ85TlgodeaxJPoB+JSi
# CwSjZxlrd169kTcUWjzmBEC8a4Mza6+wb0I/AqOJEuUXckg2gdutIMv2F9Msh7Tp
# 2Xyuue8WNUAMucRG85IFDB1pe73/Hwy5s2d5vp5MQ6+WJu8vqQjDHYfuX2JoqUcu
# gXy34QvLyAQiDL48sgzCyXmBM+f1yx2cRvT0XZcOzin7RyGgnrHNpJJBu3Z3fMi2
# MQyzC25HINr7Ut7JbJ2ngQydBfxBGvfMIJRbHzSrItmhqwVFyX+fDa8QLCiP6BEw
# UtlBEP3oMoDW/QFYqLK1USWj+2wNC3g9JVuXCGBn4osqd3rPubgN4S+CCKwlorat
# wEoCWV9djuV4yM/++ZPFjlEIUfiDc3i8ipXDSgFjdULODZkjivNluxf5iRPoRp0P
# 1oYYiyGzNg==
# SIG # End signature block
