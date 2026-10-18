<#
.SYNOPSIS
Tests removal of misplaced Authenticode signature blocks by the WorkplaceCloudHub signing script.

.DESCRIPTION
Version 1.0.0

Runs Sign-WorkplaceCloudHubRepositoryPowerShellScripts.ps1 on throwaway repositories and checks that:
- code after an old signature block is kept and the script ends up Valid and runnable under AllSigned;
- a script with two signature blocks keeps all its code and ends up with a single Valid block;
- a UTF-8 BOM and CRLF line endings are preserved;
- a script without block, and a normally signed script edited before its block, are signed exactly as
  Set-AuthenticodeSignature alone would sign them (no extra line);
- unbalanced markers fail with an explicit error and leave the file byte-identical.

Requires the workplacecloudhub.com code-signing certificate in Cert:\CurrentUser\My, trusted in
CurrentUser\Root and CurrentUser\TrustedPublisher. No timestamp server is used.
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent),
    [string]$Thumbprint = 'D70ECB7B00377EBFB76B304C08DFC6620584E114',
    [switch]$KeepTempRepository
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:Failures = 0
$crlf = [string][char]13 + [char]10
# Built by concatenation so this file never contains a marker line itself.
$beginMarker = '# SIG # Begin' + ' signature block'
$endMarker = '# SIG # End' + ' signature block'

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if ($Condition) { Write-Host "[PASS] $Message" }
    else { Write-Host "[FAIL] $Message" -ForegroundColor Red; $script:Failures++ }
}

function Initialize-TempRepository {
    param([string]$Name)
    $path = Join-Path -Path $script:TempRoot -ChildPath $Name
    New-Item -ItemType Directory -Path $path -Force | Out-Null
    & git init --quiet $path
    if ($LASTEXITCODE -ne 0) { throw 'git init failed.' }
    & git -C $path config core.autocrlf false
    return $path
}

function Add-TrackedFile {
    param([string]$Repository, [string]$Name)
    & git -C $Repository add -- $Name
    if ($LASTEXITCODE -ne 0) { throw "git add failed for $Name." }
}

function Invoke-Signer {
    param([string]$Repository)
    $command = "& '{0}' -RepoRoot '{1}' -Thumbprint '{2}' -TimestampServer '' -SkipTrustInstall; exit `$LASTEXITCODE" -f $script:Signer, $Repository, $Thumbprint
    $output = & pwsh -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
    $code = $LASTEXITCODE
    $text = ($output | ForEach-Object { [string]$_ }) -join [Environment]::NewLine
    Write-Host $text
    return [pscustomobject]@{ ExitCode = $code; Output = $text }
}

function Invoke-AllSigned {
    param([string]$Path)
    $output = & pwsh -NoProfile -ExecutionPolicy AllSigned -File $Path 2>&1
    return [pscustomobject]@{ ExitCode = $LASTEXITCODE; Lines = @($output | ForEach-Object { [string]$_ }) }
}

function Get-Text { param([string]$Path) return [System.IO.File]::ReadAllText($Path) }

function Get-MarkerCount {
    param([string]$Path, [string]$Marker)
    return @([System.IO.File]::ReadAllLines($Path) | Where-Object { $_ -ceq $Marker }).Count
}

function Get-CodeBeforeBlock {
    param([string]$Path)
    $text = Get-Text $Path
    $index = $text.IndexOf($beginMarker, [System.StringComparison]::Ordinal)
    if ($index -lt 0) { return $text }
    return $text.Substring(0, $index)
}

function Test-OnlyCrlf {
    param([string]$Path)
    $text = Get-Text $Path
    return -not ([regex]::IsMatch($text, '\r(?!\n)|(?<!\r)\n'))
}

function Write-SignedFixture {
    param([string]$Path, [string]$Code)
    [System.IO.File]::WriteAllText($Path, $Code, [System.Text.UTF8Encoding]::new($false))
    $result = Set-AuthenticodeSignature -FilePath $Path -Certificate $script:Cert -HashAlgorithm SHA256
    if ([string]$result.Status -ne 'Valid') { throw "Fixture signing failed for ${Path}: $($result.Status)" }
}

$script:Signer = Join-Path -Path $RepoRoot -ChildPath 'Sign-WorkplaceCloudHubRepositoryPowerShellScripts.ps1'
if (-not (Test-Path -LiteralPath $script:Signer -PathType Leaf)) { throw "Signing script not found: $script:Signer" }
$script:Cert = Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert | Where-Object { $_.Thumbprint -eq $Thumbprint } | Select-Object -First 1
if (-not $script:Cert) { throw "Code-signing certificate not found: $Thumbprint" }

$script:TempRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ('wch-sigblock-test-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $script:TempRoot -Force | Out-Null

try {
    # Run 1: repairable layouts and regressions of the normal path.
    $repo = Initialize-TempRepository -Name 'repair'

    $afterBlock = Join-Path $repo 'AfterBlock.ps1'
    Write-SignedFixture -Path $afterBlock -Code ("Write-Output 'before'" + $crlf)
    [System.IO.File]::AppendAllText($afterBlock, "Write-Output 'after'" + $crlf)

    $twoBlocks = Join-Path $repo 'TwoBlocks.ps1'
    Write-SignedFixture -Path $twoBlocks -Code ("Write-Output 'first'" + $crlf)
    $signedOnce = Get-Text $twoBlocks
    [System.IO.File]::WriteAllText($twoBlocks, $signedOnce + "Write-Output 'middle'" + $crlf + $signedOnce.Replace("'first'", "'last'"), [System.Text.UTF8Encoding]::new($false))

    $bom = Join-Path $repo 'Bom.ps1'
    Write-SignedFixture -Path $bom -Code ("Write-Output 'bom'" + $crlf)
    [System.IO.File]::WriteAllText($bom, (Get-Text $bom) + "Write-Output 'bom after'" + $crlf, [System.Text.UTF8Encoding]::new($true))

    $noBlockCode = "Write-Output 'plain'" + $crlf
    $noBlock = Join-Path $repo 'NoBlock.ps1'
    [System.IO.File]::WriteAllText($noBlock, $noBlockCode, [System.Text.UTF8Encoding]::new($false))
    $noBlockReference = Join-Path $script:TempRoot 'NoBlock.reference.ps1'
    Write-SignedFixture -Path $noBlockReference -Code $noBlockCode

    $edited = Join-Path $repo 'Edited.ps1'
    Write-SignedFixture -Path $edited -Code ("Write-Output 'v1'" + $crlf)
    [System.IO.File]::WriteAllText($edited, (Get-Text $edited).Replace("'v1'", "'v2'"), [System.Text.UTF8Encoding]::new($false))
    $editedReference = Join-Path $script:TempRoot 'Edited.reference.ps1'
    Copy-Item -LiteralPath $edited -Destination $editedReference
    [void](Set-AuthenticodeSignature -FilePath $editedReference -Certificate $script:Cert -HashAlgorithm SHA256)

    foreach ($name in @('AfterBlock.ps1', 'TwoBlocks.ps1', 'Bom.ps1', 'NoBlock.ps1', 'Edited.ps1')) { Add-TrackedFile -Repository $repo -Name $name }
    Assert-True ([string](Get-AuthenticodeSignature -FilePath $afterBlock).Status -ne 'Valid') 'Fixture: code after a block invalidates the signature'

    $run = Invoke-Signer -Repository $repo
    Assert-True ($run.ExitCode -eq 0) 'Signer succeeds on repairable files'

    foreach ($file in @($afterBlock, $twoBlocks, $bom, $noBlock, $edited)) {
        $name = Split-Path -Path $file -Leaf
        Assert-True ([string](Get-AuthenticodeSignature -FilePath $file).Status -eq 'Valid') "$name is Valid"
        Assert-True ((Get-MarkerCount $file $beginMarker) -eq 1 -and (Get-MarkerCount $file $endMarker) -eq 1) "$name has exactly one signature block"
        Assert-True (Test-OnlyCrlf $file) "$name uses CRLF only"
    }

    $exec = Invoke-AllSigned -Path $afterBlock
    Assert-True ($exec.ExitCode -eq 0 -and ($exec.Lines -join '|') -eq 'before|after') 'Code after the old block is kept and runs under AllSigned'
    $exec = Invoke-AllSigned -Path $twoBlocks
    Assert-True ($exec.ExitCode -eq 0 -and ($exec.Lines -join '|') -eq 'first|middle|last') 'Code around two old blocks is kept and runs under AllSigned'
    $exec = Invoke-AllSigned -Path $bom
    Assert-True ($exec.ExitCode -eq 0 -and ($exec.Lines -join '|') -eq 'bom|bom after') 'BOM script runs under AllSigned'

    $bomBytes = [System.IO.File]::ReadAllBytes($bom)
    Assert-True ($bomBytes.Length -ge 3 -and $bomBytes[0] -eq 0xEF -and $bomBytes[1] -eq 0xBB -and $bomBytes[2] -eq 0xBF) 'UTF-8 BOM is preserved'
    Assert-True ((Get-CodeBeforeBlock $noBlock) -ceq (Get-CodeBeforeBlock $noBlockReference)) 'File without block is signed exactly like Set-AuthenticodeSignature alone'
    Assert-True ((Get-CodeBeforeBlock $noBlock).StartsWith($noBlockCode, [System.StringComparison]::Ordinal)) 'File without block keeps its code unchanged'
    Assert-True ((Get-CodeBeforeBlock $edited) -ceq (Get-CodeBeforeBlock $editedReference)) 'Normal re-sign adds no extra line'

    # Run 2: unbalanced markers fail and leave files byte-identical.
    $repo = Initialize-TempRepository -Name 'unbalanced'
    $layouts = [ordered]@{
        'BeginOnly.ps1' = "Write-Output 'a'" + $crlf + $beginMarker + $crlf + '# AAAA' + $crlf
        'EndOnly.ps1'   = "Write-Output 'a'" + $crlf + $endMarker + $crlf
        'Nested.ps1'    = $beginMarker + $crlf + $beginMarker + $crlf + $endMarker + $crlf
        # Blank line 3 is tolerated; code on line 4 is not.
        'CodeInBlock.ps1' = "Write-Output 'a'" + $crlf + $beginMarker + $crlf + $crlf + "Write-Output 'b'" + $crlf + '# AAAA' + $crlf + $endMarker + $crlf
    }
    $hashes = @{}
    foreach ($name in $layouts.Keys) {
        $path = Join-Path $repo $name
        [System.IO.File]::WriteAllText($path, $layouts[$name], [System.Text.UTF8Encoding]::new($false))
        Add-TrackedFile -Repository $repo -Name $name
        $hashes[$name] = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    }
    $run = Invoke-Signer -Repository $repo
    Assert-True ($run.ExitCode -ne 0) 'Signer fails on unbalanced markers or code inside a block'
    Assert-True (([regex]::Matches($run.Output, 'Unbalanced signature markers')).Count -eq 3) 'Each unbalanced file reports an explicit error'
    Assert-True ($run.Output -match 'CodeInBlock\.ps1 :: Non-signature line inside a signature block at line 4; file left unchanged\.') 'Code inside a block reports an explicit error with its line'
    foreach ($name in $layouts.Keys) {
        Assert-True ((Get-FileHash -LiteralPath (Join-Path $repo $name) -Algorithm SHA256).Hash -eq $hashes[$name]) "$name is byte-identical"
    }
}
finally {
    if ($KeepTempRepository) { Write-Host "Temporary repositories kept: $script:TempRoot" }
    else { Remove-Item -LiteralPath $script:TempRoot -Recurse -Force -ErrorAction SilentlyContinue }
}

if ($script:Failures -gt 0) {
    Write-Host "$($script:Failures) assertion(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host 'All signature block assertions passed.'
exit 0

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCABGXxg8rhIVMDS
# Jq8qWDHnjE7IyrI03ytwIGdUHrvnXaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIETZUtKvFY4v+UNgQn/xaEN0whvyrjxnhZ1zhFvg7N6pMA0GCSqG
# SIb3DQEBAQUABIIBgIWmA5Gcw4gfqQMYXlIxnvSqkBNqPO9pU5ttrdQac6jdpnsr
# aetrRCui1+YVuQXN2CRtTiQhEmfPpOPrDZl2TdZOIqLTrOS6GcsgO/lcyUKM7uTI
# MuJRhFp5yP0aPZy9chmOmD69jedoFl1kLUYTdhoi0muf4G0WrR51aSK1cV/We7eO
# +T/lR+neOYWIIDtB5xxODJ6NgzC4G12YKRDNqTRHsTcRBkh6qOsIdwGWD6DzLS5N
# OmjwJPvAbicWtOHni+oezD5jpWLnweosuC6z4zbVmGxdluByyrQm1O8QbICh8lal
# H+18BaJyLTDM5njkGkv7RBJjEeoAX1mGTODfErEBfqGopg0N9knblSNo1Qf/P2OL
# FPdEvWCKFAlJVDENrlCpKZ4QGWKnEM+I8266SnrqFPmtw1GmiflcS0+ZKG7PtOIR
# 2F8zmUTXhjW6uPGGsXOzfjDfqbTg3CAxqQqgZxyD+dttnUV8qDVorpFp3UI1HeCy
# UnuRIV+THGtm0kFhgqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgwMzU3
# MzRaMC8GCSqGSIb3DQEJBDEiBCB9PDZK02aKmFbONjDMGDEJcMXj42rTH/HA8xIH
# yramPzANBgkqhkiG9w0BAQEFAASCAgBBFsVRcr0jycQyqy8GVIWrEwjaUkWVDi2J
# g5qM5Cuy7JnAC4ZiEZZjqv5ZmrAoFsUbK+zr1B/XCFLmYc33oZRWz3Xmen+klwKO
# d77U4K7XFxFLYsH9XK+XU8vgjlrQa738yv2wmsNizRYa+Tg6+u1GIrN5/HZWGxc1
# Z7IW6g1sXQfHs7oC1jDpNdP1ZgmGYXZr5oNGghUa5hFGToeyfeT21vxlywqAQkSP
# ++dR/xpVZjCxEGXWqM5DnSCCJAnyL2RhthZwAzb9VxJ00zv+L9rkT6qOCdqhPlTk
# tEj4FFfH282yiov1og7UpzqYGUeJ2hyQWRzljEsWJMqmuuNs/mdcNgr+CMMekZ3x
# NqR7O9x8rahMit755UKDr/ks62g3ecvyQoh/8RxOMQsrxSFqbhDkIm736uoEJDJ+
# hq330K65SxMzZ+3kWpWNwFl5nBDUUxKF71uQbSitR3A8XCBuP5tjNGZbrib98hT5
# n2wOodZvZOB+lyI0KesfMiwJ4zs1O7P5kNHVfhoLhTHcyG8dCjd63jV/ky1qjndq
# khpjNpe+FM6C4YjSZYwmaLbbBDjymMRl1UgvG6biZvYrtx+/+NkZjJzzLzt1L38H
# TCv5dpJ+K0AE3T3hSxib4E5Nl4L2OThX7RjH0OL/pr/d2AVRcsnzdOI2RuB7tWH7
# k8zILorhmA==
# SIG # End signature block
