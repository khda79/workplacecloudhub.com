<#
.SYNOPSIS
End-to-end test of the WorkplaceCloudHub pre-push signing hook.

.DESCRIPTION
Version 1.1.0

Builds a throwaway repository with a bare remote, installs the current pre-push hook and signing
script, then pushes through the real hook. Checks that:
- committed unsigned scripts are signed and the push is blocked until the signatures are committed;
- untracked scripts, and tracked scripts with uncommitted changes outside the pushed commits, stay
  byte-identical;
- a pushed script whose committed version is unsigned and whose working tree has uncommitted
  changes is left untouched and the push is refused;
- a push without PowerShell changes skips signing;
- a vendored Microsoft-signed script under a ThirdParty folder is pushed byte for byte with its
  Microsoft signature, and a tampered copy is refused and left untouched.

Requires the workplacecloudhub.com code-signing certificate in Cert:\CurrentUser\My and access to
the timestamp server.
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent),
    [string]$ThirdPartySamplePath = (Join-Path -Path (Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent) -ChildPath 'SmartM365\SmartInventory\ExchangeInventory\OnPremises\ServersAndStorage\ThirdParty\CSS-Exchange\HealthChecker\HealthChecker.ps1'),
    [switch]$KeepTempRepository
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:Failures = 0

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if ($Condition) { Write-Host "[PASS] $Message" }
    else { Write-Host "[FAIL] $Message" -ForegroundColor Red; $script:Failures++ }
}

function Invoke-Git {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)
    $output = & git -C $script:Work @Arguments 2>&1
    $code = $LASTEXITCODE
    $text = ($output | ForEach-Object { [string]$_ }) -join [Environment]::NewLine
    if ($text) { Write-Host $text }
    return [pscustomobject]@{ ExitCode = $code; Output = $text }
}

function Invoke-GitChecked {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)
    $result = Invoke-Git -Arguments $Arguments
    if ($result.ExitCode -ne 0) { throw "git $($Arguments -join ' ') failed with exit code $($result.ExitCode)." }
}

function Write-LfFile {
    param([string]$RelativePath, [string[]]$Lines)
    $path = Join-Path -Path $script:Work -ChildPath $RelativePath
    $folder = Split-Path -Path $path -Parent
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    [System.IO.File]::WriteAllText($path, (($Lines -join "`n") + "`n"), [System.Text.UTF8Encoding]::new($false))
}

function Get-FileHashValue {
    param([string]$RelativePath)
    return (Get-FileHash -LiteralPath (Join-Path -Path $script:Work -ChildPath $RelativePath) -Algorithm SHA256).Hash
}

function Get-SignatureStatus {
    param([string]$RelativePath)
    return [string](Get-AuthenticodeSignature -FilePath (Join-Path -Path $script:Work -ChildPath $RelativePath)).Status
}

function Get-RemoteMain {
    $value = & git -C $script:Remote rev-parse --verify --quiet refs/heads/main
    if ($LASTEXITCODE -ne 0) { return '' }
    return [string]$value
}

$hookSource = Join-Path -Path $RepoRoot -ChildPath '.githooks\pre-push'
$signerSource = Join-Path -Path $RepoRoot -ChildPath 'Sign-WorkplaceCloudHubRepositoryPowerShellScripts.ps1'
foreach ($source in @($hookSource, $signerSource, $ThirdPartySamplePath)) {
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Missing source file: $source" }
}

$tempRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ('wch-prepush-test-' + [guid]::NewGuid().ToString('N'))
$script:Work = Join-Path -Path $tempRoot -ChildPath 'work'
$script:Remote = Join-Path -Path $tempRoot -ChildPath 'remote.git'
New-Item -ItemType Directory -Path $script:Work -Force | Out-Null

try {
    & git init --quiet --bare $script:Remote
    if ($LASTEXITCODE -ne 0) { throw 'git init --bare failed.' }
    & git init --quiet $script:Work
    if ($LASTEXITCODE -ne 0) { throw 'git init failed.' }
    Invoke-GitChecked @('symbolic-ref', 'HEAD', 'refs/heads/main')
    Invoke-GitChecked @('config', 'user.name', 'Pre-push test')
    Invoke-GitChecked @('config', 'user.email', 'pre-push-test@example.invalid')
    Invoke-GitChecked @('config', 'core.autocrlf', 'false')
    Invoke-GitChecked @('config', 'core.hooksPath', '.githooks')
    Invoke-GitChecked @('remote', 'add', 'origin', $script:Remote)

    New-Item -ItemType Directory -Path (Join-Path -Path $script:Work -ChildPath '.githooks') -Force | Out-Null
    Copy-Item -LiteralPath $hookSource -Destination (Join-Path -Path $script:Work -ChildPath '.githooks\pre-push')
    Copy-Item -LiteralPath $signerSource -Destination (Join-Path -Path $script:Work -ChildPath 'Sign-WorkplaceCloudHubRepositoryPowerShellScripts.ps1')
    Write-LfFile -RelativePath '.gitattributes' -Lines @('* text=auto', '*.ps1 text eol=crlf', '.githooks/* text eol=lf', '**/ThirdParty/** -text -eol -whitespace')
    Write-LfFile -RelativePath 'Tracked.ps1' -Lines @("Write-Output 'tracked'")

    # 1. Initial push: committed unsigned scripts are signed and the push is blocked.
    Invoke-GitChecked @('add', '-A')
    Invoke-GitChecked @('commit', '--quiet', '-m', 'initial')
    $push = Invoke-Git @('push', 'origin', 'main')
    Assert-True ($push.ExitCode -ne 0) 'Initial push with an unsigned committed script is blocked'
    Assert-True ($push.Output -match 'signatures were updated') 'Hook asks to commit the signature changes'
    Assert-True ((Get-SignatureStatus 'Tracked.ps1') -eq 'Valid') 'Committed unsigned script was signed'
    Assert-True ((Get-RemoteMain) -eq '') 'Remote is unchanged after the blocked push'

    Invoke-GitChecked @('commit', '--quiet', '-a', '-m', 'chore(signing): sign initial scripts')
    $push = Invoke-Git @('push', 'origin', 'main')
    Assert-True ($push.ExitCode -eq 0) 'Push succeeds once signatures are committed'

    # 2. Foreign work in progress: untracked scripts and uncommitted changes outside the pushed commits.
    Write-LfFile -RelativePath 'Foreign\IntunePackager\Untracked.ps1' -Lines @("Write-Output 'foreign untracked'")
    Write-LfFile -RelativePath 'Untracked.psm1' -Lines @('function Get-Foreign { 1 }')
    Add-Content -LiteralPath (Join-Path -Path $script:Work -ChildPath 'Tracked.ps1') -Value "Write-Output 'foreign edit'"
    $foreignFiles = @('Foreign\IntunePackager\Untracked.ps1', 'Untracked.psm1', 'Tracked.ps1')
    $foreignHashes = @{}
    foreach ($file in $foreignFiles) { $foreignHashes[$file] = Get-FileHashValue $file }

    Write-LfFile -RelativePath 'New.ps1' -Lines @("Write-Output 'new'")
    Invoke-GitChecked @('add', '--', 'New.ps1')
    Invoke-GitChecked @('commit', '--quiet', '-m', 'add New.ps1')
    $remoteBefore = Get-RemoteMain
    $push = Invoke-Git @('push', 'origin', 'main')
    Assert-True ($push.ExitCode -ne 0) 'Push of a new unsigned committed script is blocked'
    Assert-True ((Get-SignatureStatus 'New.ps1') -eq 'Valid') 'New committed script was signed'
    Assert-True ((Get-RemoteMain) -eq $remoteBefore) 'Remote is unchanged after the blocked push'
    foreach ($file in $foreignFiles) {
        Assert-True ((Get-FileHashValue $file) -eq $foreignHashes[$file]) "Foreign file is byte-identical after the blocked push: $file"
    }

    Invoke-GitChecked @('add', '--', 'New.ps1')
    Invoke-GitChecked @('commit', '--quiet', '-m', 'chore(signing): sign New.ps1')
    $push = Invoke-Git @('push', 'origin', 'main')
    Assert-True ($push.ExitCode -eq 0) 'Push succeeds with foreign work in progress present'
    foreach ($file in $foreignFiles) {
        Assert-True ((Get-FileHashValue $file) -eq $foreignHashes[$file]) "Foreign file is byte-identical after the successful push: $file"
    }

    # 3. A push without PowerShell changes skips signing.
    Write-LfFile -RelativePath 'Notes.txt' -Lines @('notes')
    Invoke-GitChecked @('add', '--', 'Notes.txt')
    Invoke-GitChecked @('commit', '--quiet', '-m', 'add notes')
    $push = Invoke-Git @('push', 'origin', 'main')
    Assert-True ($push.ExitCode -eq 0) 'Push without PowerShell changes succeeds'
    Assert-True ($push.Output -match 'signing skipped') 'Signing is skipped when no PowerShell file is pushed'

    # 4. Pushed script with an unsigned committed version and uncommitted changes: refused, untouched.
    Write-LfFile -RelativePath 'Shared.ps1' -Lines @("Write-Output 'shared'")
    Invoke-GitChecked @('add', '--', 'Shared.ps1')
    Invoke-GitChecked @('commit', '--quiet', '-m', 'add Shared.ps1')
    Add-Content -LiteralPath (Join-Path -Path $script:Work -ChildPath 'Shared.ps1') -Value "Write-Output 'uncommitted'"
    $sharedHash = Get-FileHashValue 'Shared.ps1'
    $remoteBefore = Get-RemoteMain
    $push = Invoke-Git @('push', 'origin', 'main')
    Assert-True ($push.ExitCode -ne 0) 'Push is refused when a pushed unsigned script has uncommitted changes'
    Assert-True ((Get-FileHashValue 'Shared.ps1') -eq $sharedHash) 'Script with uncommitted changes is byte-identical'
    Assert-True ((Get-RemoteMain) -eq $remoteBefore) 'Remote is unchanged after the refused push'
    foreach ($file in $foreignFiles) {
        Assert-True ((Get-FileHashValue $file) -eq $foreignHashes[$file]) "Foreign file is byte-identical after the refused push: $file"
    }

    # 5. Pushed script with a signed committed version and uncommitted changes: pushed, untouched.
    Invoke-GitChecked @('checkout', '--', 'Shared.ps1')
    $push = Invoke-Git @('push', 'origin', 'main')
    Assert-True ($push.ExitCode -ne 0 -and (Get-SignatureStatus 'Shared.ps1') -eq 'Valid') 'Clean pushed script is signed and the push is blocked'
    Invoke-GitChecked @('commit', '--quiet', '-m', 'chore(signing): sign Shared.ps1', '--', 'Shared.ps1')
    Add-Content -LiteralPath (Join-Path -Path $script:Work -ChildPath 'Shared.ps1') -Value "Write-Output 'uncommitted again'"
    $sharedHash = Get-FileHashValue 'Shared.ps1'
    $push = Invoke-Git @('push', 'origin', 'main')
    Assert-True ($push.ExitCode -eq 0) 'Push succeeds when the committed version is signed despite uncommitted changes'
    Assert-True ((Get-FileHashValue 'Shared.ps1') -eq $sharedHash) 'Script with uncommitted changes is byte-identical after the push'
    $headCommit = [string](& git -C $script:Work rev-parse HEAD)
    Assert-True ((Get-RemoteMain) -eq $headCommit) 'Remote main matches the local commit'

    # 6. Vendored Microsoft-signed script: pushed as is, never re-signed; a tampered copy is refused.
    Invoke-GitChecked @('checkout', '--', 'Shared.ps1')
    $vendorRelative = 'Tools\ThirdParty\Vendor\HealthChecker.ps1'
    $vendorPath = Join-Path -Path $script:Work -ChildPath $vendorRelative
    New-Item -ItemType Directory -Path (Split-Path -Path $vendorPath -Parent) -Force | Out-Null
    [System.IO.File]::WriteAllBytes($vendorPath, [System.IO.File]::ReadAllBytes($ThirdPartySamplePath))
    $vendorHash = Get-FileHashValue $vendorRelative
    Invoke-GitChecked @('add', '--', $vendorRelative)
    Invoke-GitChecked @('commit', '--quiet', '-m', 'add vendored script')
    $push = Invoke-Git @('push', 'origin', 'main')
    $vendorSignature = Get-AuthenticodeSignature -FilePath $vendorPath
    Assert-True ($push.ExitCode -eq 0) 'Push of a vendored Microsoft-signed script succeeds without a signing commit'
    Assert-True ($push.Output -match 'Third-party file left untouched') 'Hook reports the vendored script as third-party'
    Assert-True ((Get-FileHashValue $vendorRelative) -eq $vendorHash) 'Vendored script is byte-identical after the push'
    Assert-True ([string]$vendorSignature.Status -eq 'Valid' -and [string]$vendorSignature.SignerCertificate.Subject -match 'CN=Microsoft Corporation') 'Vendored script keeps its Microsoft signature'
    $vendorBlob = [string](& git -C $script:Work rev-parse "HEAD:$($vendorRelative.Replace('\', '/'))")
    Assert-True ([string](& git -C $script:Remote cat-file -s $vendorBlob) -eq [string](Get-Item -LiteralPath $ThirdPartySamplePath).Length) 'Remote stores the vendored script with its published size (no line ending conversion)'

    $tampered = [System.IO.File]::ReadAllBytes($vendorPath)
    $tampered[2000] = if ($tampered[2000] -eq 0x41) { [byte]0x42 } else { [byte]0x41 }
    [System.IO.File]::WriteAllBytes($vendorPath, $tampered)
    $tamperedHash = Get-FileHashValue $vendorRelative
    Invoke-GitChecked @('commit', '--quiet', '-m', 'tamper vendored script', '--', $vendorRelative)
    $remoteBefore = Get-RemoteMain
    $push = Invoke-Git @('push', 'origin', 'main')
    Assert-True ($push.ExitCode -ne 0) 'Push of a tampered vendored script is refused'
    Assert-True ($push.Output -match 'no valid publisher signature') 'Refusal names the invalid publisher signature'
    Assert-True ((Get-FileHashValue $vendorRelative) -eq $tamperedHash) 'Tampered vendored script is left untouched (not re-signed)'
    Assert-True ((Get-RemoteMain) -eq $remoteBefore) 'Remote is unchanged after the refused vendored push'
}
finally {
    if ($KeepTempRepository) { Write-Host "Temporary repository kept: $tempRoot" }
    else { Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue }
}

if ($script:Failures -gt 0) {
    Write-Host "$($script:Failures) assertion(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host 'All pre-push signing assertions passed.'
exit 0

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAbsUxf0ABGjDdg
# cKovgxBdXynWl1/peA4fT4ywll+Fh6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGe9eI2H8V1Qvj83QDoRVyX6W/+XCVmrqi4enDcDvedWMA0GCSqG
# SIb3DQEBAQUABIIBgJ+PoDgi1gwYMfdhhBxXDx0orjbU8fKhQdgPU2GTsW6Xcd2K
# pyI6Qj0nv+JDhjanb+VEEEn2QeiwZqXr09YZqOIFrUIcbG4ZUMJyj/uYC4EufMwf
# bC82Wbs3w26KX1qxEyiGMWO6G343a3J4Nei79sfISNtxP7tmvqCUmyEEeyVwAwCy
# 1WxaMmZ6sLGlZqldhkYCURyKZf6EA8nHBuW10dk2tr11aKNFadgIdDp2IGklySeE
# zMQhQPyEC+AR3m/CSw2uU2NOC88Xw4Apt0v6qZF/F88vITyduO761VeDV9RefSpT
# QkLzXjn8C+uoa+kpuk8odxMdFkx8SBKxWKmjsXaja9J69VwvCORPOSFY1swsUssx
# DPaYszHm9If28rtIb5J+1/OeT/n5hKLoOi6ILnfWmJeoPKwCXEQjvz42sujvGs97
# beev0NqMubpkhDf5H2/qzVrs/1YeOnIh3RB8y1iFaJnThR64QmMNNmL798fd7G+W
# JeqT+3Tc4LOI5aHaIKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgwMjQ1
# MjlaMC8GCSqGSIb3DQEJBDEiBCDbTYqphu8a7KQMytssrqYcDm1r/Um4oE3AjcRK
# DhjIojANBgkqhkiG9w0BAQEFAASCAgC02HdhbPi4OnFrJRBE7CMcsW0tjo2UMlQj
# MN+l15bWtEepEJqT+C5AjPmdY52v6hW5G2RY0uBABKqFColuVLWeOvWfxE1/3Sg1
# 9pl6PXOMocpUPurs+x124lr+IhNMiA7Xsgd1e1XajfzHKtVLvT6fJkGiRIMMZH5i
# n9JP9g6z6wmSp+aMebDb/M09lcBRUqeiAVxs0fmQgFBgdJU44l9TeYPGjLUUGXB5
# YomwbS+ma/ZhfvcY8rCZFFgRyv7Qr6LwyxESnkHI+39JagOhxEK71n+CDUCk1C+F
# c8qvSFl0VasheMotfVxgoBPR5vfaL5HHokiWjIFDbYXhrPOT/B3oWWkF9shOrGmo
# 4mgtffGSLhCbhD8NZJbtKysy0Tfn5UU4JoSabmVHOd09xVy81D7irsNKUWQZUIC9
# CFKelVsYdZqUvvpoLUsk/UCYg0yxIijHYtmQ/VPzg44DERxC/xbtXZVzfmnrmf4M
# lVq0KC9vGjmZItK5nqMsNaIEc3FKEhy+GP8t1liqKYtSU/dTNjmU6ap7jEkGtwaI
# otK9Szz/SR8kRvfeBvKAL1udLZDUiskwUkgkgn1X9WXlFtEBJ+pKM+yfmuz0es7i
# vESPyok4RxE4kdt5ANf94AsrijIL1OI8agj+1b9PbfzH073xn7NizaikUm9/nIMl
# EISflHxAsA==
# SIG # End signature block
