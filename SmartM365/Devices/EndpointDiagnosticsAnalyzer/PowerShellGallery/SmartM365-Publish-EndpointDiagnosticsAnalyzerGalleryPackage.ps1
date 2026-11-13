#Requires -Version 5.1
<# .VERSION
1.0.2
#>
[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
param(
    [string]$PackagePath='',
    [string]$OutputRoot=(Join-Path ([IO.Path]::GetTempPath()) 'SmartM365\PowerShellGallery'),
    [string]$Repository='PSGallery',
    [string]$ApiKeyEnvironmentVariable='PSGALLERY_API_KEY',
    [switch]$Execute,
    [switch]$AllowPrereleasePublication,
    [switch]$ForceBuild
)
Set-StrictMode -Version Latest;$ErrorActionPreference='Stop'

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

if(-not$PackagePath){$build=& (Join-Path $PSScriptRoot 'SmartM365-Build-EndpointDiagnosticsAnalyzerGalleryPackage.ps1') -OutputRoot $OutputRoot -Force:$ForceBuild;$PackagePath=[string]$build.PackagePath}
Assert-WchNoDemoAssets -Path $PackagePath
$runtime=Join-Path $PackagePath 'Runtime'
$version=Get-Content -LiteralPath (Join-Path $runtime 'SmartM365-EndpointDiagnosticsAnalyzer.version.json.txt') -Raw|ConvertFrom-Json
foreach($name in @('SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1','HardwareReadiness.ps1','SmartM365.GuiSplash.ps1','Start-SmartM365-EndpointDiagnosticsAnalyzer-GUI.cmd','WorkplaceCloudHub-lockup-WPF.png','WorkplaceCloudHub.ico')){
    $hash=$version.RuntimeHashes.PSObject.Properties[$name]
    if(-not$hash -or (Get-FileHash -LiteralPath (Join-Path $runtime $name)).Hash -ne [string]$hash.Value){throw "Packaged runtime integrity mismatch: $name"}
}
foreach($file in Get-ChildItem -LiteralPath $PackagePath -File -Recurse | Where-Object Extension -in @('.ps1','.psm1','.psd1')){
    $signature=Get-AuthenticodeSignature -LiteralPath $file.FullName
    if($signature.Status -ne 'Valid' -or -not$signature.SignerCertificate -or $signature.SignerCertificate.Thumbprint -ne 'D70ECB7B00377EBFB76B304C08DFC6620584E114'){throw "Package signature verification failed: $($file.Name)"}
}
$manifestPath=Join-Path $PackagePath 'SmartM365.EndpointDiagnosticsAnalyzer.psd1';if(-not(Test-Path $manifestPath -PathType Leaf)){throw "Manifest not found: $manifestPath"};$manifest=Test-ModuleManifest $manifestPath;$pre=[string]$manifest.PrivateData.PSData['Prerelease'];$key=[Environment]::GetEnvironmentVariable($ApiKeyEnvironmentVariable);$preview=[pscustomobject]@{Mode=if($Execute){'Execute'}else{'Preview'};Repository=$Repository;ModuleName=$manifest.Name;Version=[string]$manifest.Version;Prerelease=$pre;PackagePath=(Resolve-Path $PackagePath).Path;ApiKeyEnvironmentVariable=$ApiKeyEnvironmentVariable;ApiKeyPresent=(-not[string]::IsNullOrWhiteSpace($key));PublicMetadataComplete=(-not[string]::IsNullOrWhiteSpace([string]$manifest.LicenseUri));PublicationAttempted=$false}
if(-not$Execute){return $preview};if($pre-and-not$AllowPrereleasePublication){throw "Prerelease approval is required: $pre"};if(-not$preview.PublicMetadataComplete){throw 'LicenseUri is required.'};if(-not$key){throw "API key environment variable is empty: $ApiKeyEnvironmentVariable"};if(-not(Get-Command Publish-PSResource -ErrorAction SilentlyContinue)){throw 'Publish-PSResource is unavailable.'};if($PSCmdlet.ShouldProcess("$($manifest.Name) $($manifest.Version) to $Repository",'Publish public PowerShell resource')){Publish-PSResource -Path $PackagePath -ApiKey $key -Repository $Repository -ErrorAction Stop;$preview.PublicationAttempted=$true};$preview

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDBuy8skV8B3d5J
# VsJZk6VQaI9JIaoAnuf8yx8vvWCHzqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIARKQHFFaRDM6Bv8rP71ZV+hVPJveo0NXrXq/2QCfbn3MA0GCSqG
# SIb3DQEBAQUABIIBgGvjC5Nhv1JYEHZU/Nxd5PARvshLLYsS3YTPOaWDJ5a2zv3h
# p1i0qu2+/C2Da9RqlZPmPGdS6LgmujIuEDmLfkvpfxUxZtEu0B2Mokr6EpoYzJtf
# VivpuebuEETRxSH8MTJKTdTZCJgRF7vUqW4H21PoRLsKluyiIaliEQzHP1+emqwe
# qNb6LfncMz15FpR+pOGURRAO6we0dYSzxEKJchKY2C2P+RPgqrG6OHcPr18cx41W
# lLS01/iHZg1VMogLODI0sL8ppZ2tEPhbJB9i5aG8kI+kZ+nUCc2EwiHdU+4ACf/v
# QVe9MIuN+FiFTwEE4KzN9mBL02whRjWhUG9X9fooSBBFucD+fFYHVd/uFpPmrAnq
# 6PeJMsDoJFR71pjSoXovG7kSkCi63Jko1xMraRJ8RwCuRl8xzTXs1MO6Q/ToQL5i
# adVFReYkDWupC8cUapqnjuNFRLOJhmTd8wbFaY0YMea2z7zYpwLknjhAbMwVLAIb
# 03a9jVCb+9T6z/JHPqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTMxMDE1
# NTFaMC8GCSqGSIb3DQEJBDEiBCCqj/mmYscriIEN13Jj4ZFsgiZWCYtTG0yyU/g/
# wNwyeTANBgkqhkiG9w0BAQEFAASCAgBdEcTTMZ2KFHGKVUr8A0HlCH7FBFziZEnK
# NUzhYPpJkbB9q/Ebd2FsnDTHvEyoknjinx3F/c4w8pOZyjz60jU3DeyEeserAblY
# VXfz18rasy6djLaFtXLICjSD7nz6syt82WnnoDUUMUiPj+1/humIEnvK5nq3Muci
# 3iUfAICddKCiGBLx/isGAnQk904L2ew6OnWy12Z772miILenj9MfpGar7yiPMoxC
# 9mK6o0j5v9xSUeAJFi6PlS75cjC+ndRjlmonx2wgyJRijmp3EXCjcBF8BYb3PnbA
# N5UBOv1r50P+BTI7lteuY0jXZ0XQyVMxRJ6D9Ng1D9/aS2VG853f1PM7gGPXsCJc
# wOJvQkXLmvk45bnPse7CGu6TYxobVOnl0EoW/usYmqhf81CkU6jtrC5m53WTBOyG
# hob9Z+xxBrD8fHl9FczeIxRJhAbpEi0IfFAkH2W9bqOSQeGCbz4uUmLuyBwwSyX7
# /qsgV6+VBU4ZBCxMt6I3bV+nP7hRPcKm8/PQz8vfaFyrw4Cc5oqQoxyTSOzOh7LP
# /X/2qe+ugKDK60gQxdyJdZHB5HGoJifKzWyJGG7JeuURd3kui7w+LzK8xC2N84s2
# Wcf+sKlLTfZEQwqxbTEOTHgte6wefG40J4dnbc1k8JI0YqRz59alXbYldv7XtsA4
# xxbcffk9iA==
# SIG # End signature block
