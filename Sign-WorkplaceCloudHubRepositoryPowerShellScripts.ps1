<#
.SYNOPSIS
Signs repository PowerShell files with the workplacecloudhub.com Authenticode code-signing certificate.

.DESCRIPTION
Signs PowerShell scripts/modules/manifests tracked by Git, with an option to include
non-ignored untracked files. Intended for manual use and for the repository pre-push hook.

With -PathListFile (used by the pre-push hook), only the listed files are inspected. Each line is
"<commit><TAB><repo-relative path>". Untracked files are never touched. A listed file whose working
tree differs from <commit> (uncommitted work, possibly from another session) is left untouched: the
committed version must already carry a valid signature, otherwise the run fails. In this mode the
script exits with code 3 when it changed at least one file, so the hook can block the push.

Signature blocks that are not the single trailing block (code after a block, several blocks) are
removed before signing, keeping all other lines; unbalanced markers fail and leave the file unchanged.

Vendored third-party files (repository-relative paths matching -ExcludedPathPatterns, by default any
ThirdParty folder) are never modified: neither line endings nor signature. Their publisher signature
must be Valid and must not be the workplacecloudhub.com certificate, otherwise the run fails so the
published file can be restored.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$RepoRoot = $PSScriptRoot,
    [string]$Thumbprint = 'D70ECB7B00377EBFB76B304C08DFC6620584E114',
    [string]$ExpectedSignerEmail = 'contact@workplacecloudhub.com',
    [string]$TimestampServer = 'http://timestamp.digicert.com',
    [string[]]$Extensions = @('.ps1', '.psm1', '.psd1', '.ps1xml'),
    [string[]]$ExcludedPathPatterns = @('ThirdParty/*', '*/ThirdParty/*'),
    [string]$PathListFile,
    [switch]$IncludeUntracked,
    [switch]$Force,
    [switch]$SkipTrustInstall,
    [switch]$SkipLineEndingNormalization
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Info {
    param([string]$Message)
    Write-Host "[Sign-WorkplaceCloudHub] $Message"
}

function Normalize-Thumbprint {
    param([string]$Value)
    return ([string]$Value).Replace(' ', '').ToUpperInvariant()
}

function Test-MicrosoftAuthenticodeSignature {
    param([AllowNull()]$Signature)

    if (-not $Signature -or [string]$Signature.Status -ne 'Valid' -or -not $Signature.SignerCertificate) {
        return $false
    }

    $subject = [string]$Signature.SignerCertificate.Subject
    $issuer = [string]$Signature.SignerCertificate.Issuer
    return (($subject -match '(^|,\s*)O=Microsoft Corporation(,|$)' -or $subject -match '(^|,\s*)CN=Microsoft Corporation(,|$)') -and $issuer -match 'Microsoft')
}

function Get-RepositoryRelativePath {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$FullPath
    )

    $rootWithSeparator = $Root.TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
    if (-not $FullPath.StartsWith($rootWithSeparator, [System.StringComparison]::OrdinalIgnoreCase)) { return $FullPath.Replace('\', '/') }
    return $FullPath.Substring($rootWithSeparator.Length).Replace('\', '/')
}

function Test-ExcludedRepositoryPath {
    param(
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [AllowEmptyCollection()][string[]]$Patterns = @()
    )

    $normalizedPath = $RelativePath.Replace('\', '/')
    foreach ($pattern in $Patterns) {
        if (-not [string]::IsNullOrWhiteSpace($pattern) -and $normalizedPath -like $pattern.Replace('\', '/')) { return $true }
    }
    return $false
}
function Install-CertificateInCurrentUserStore {
    param(
        [Parameter(Mandatory = $true)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [Parameter(Mandatory = $true)][string]$StoreName
    )

    $thumbprint = Normalize-Thumbprint $Certificate.Thumbprint
    $store = [System.Security.Cryptography.X509Certificates.X509Store]::new($StoreName, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    try {
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        $existing = @($store.Certificates | Where-Object { (Normalize-Thumbprint $_.Thumbprint) -eq $thumbprint })
        if ($existing.Count -gt 0) {
            return $false
        }

        $publicCert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($Certificate.RawData)
        $store.Add($publicCert)
        return $true
    }
    finally {
        if ($store) { $store.Close() }
    }
}

function Get-GitPowerShellFiles {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string[]]$Extensions,
        [switch]$IncludeUntracked
    )

    $patterns = @($Extensions | ForEach-Object { '*{0}' -f $_ })
    $args = @('-C', $Root, 'ls-files')
    if ($IncludeUntracked) { $args += @('--cached', '--others', '--exclude-standard') }
    $args += '--'
    $args += $patterns

    $relativePaths = & git @args
    if ($LASTEXITCODE -ne 0) { throw 'git ls-files failed.' }

    return @($relativePaths | Where-Object { $_ } | Sort-Object -Unique | ForEach-Object {
        [System.IO.Path]::GetFullPath((Join-Path -Path $Root -ChildPath $_))
    })
}

function Get-PathListEntry {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$ListPath,
        [Parameter(Mandatory = $true)][string[]]$Extensions
    )

    $tracked = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    $trackedPaths = & git -C $Root -c core.quotePath=false ls-files --cached
    if ($LASTEXITCODE -ne 0) { throw 'git ls-files failed.' }
    foreach ($trackedPath in $trackedPaths) { if ($trackedPath) { [void]$tracked.Add($trackedPath) } }

    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($line in [System.IO.File]::ReadAllLines($ListPath)) {
        $parts = $line.Split([char]9, 2)
        if ($parts.Count -ne 2 -or [string]::IsNullOrWhiteSpace($parts[0]) -or [string]::IsNullOrWhiteSpace($parts[1])) { continue }
        $relativePath = $parts[1]
        if ($Extensions -notcontains [System.IO.Path]::GetExtension($relativePath)) { continue }
        if (-not $seen.Add($relativePath)) { continue }
        if (-not $tracked.Contains($relativePath)) {
            Write-Info "Not tracked, left untouched: $relativePath"
            continue
        }
        $entries.Add([pscustomobject]@{
            Commit = $parts[0]
            RelativePath = $relativePath
            FullPath = [System.IO.Path]::GetFullPath((Join-Path -Path $Root -ChildPath $relativePath))
        })
    }
    return $entries.ToArray()
}

function Test-WorkingTreeMatchesCommit {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Commit,
        [Parameter(Mandatory = $true)][string]$RelativePath
    )

    & git -C $Root diff --quiet $Commit -- $RelativePath
    if ($LASTEXITCODE -eq 0) { return $true }
    if ($LASTEXITCODE -eq 1) { return $false }
    throw "git diff failed for $RelativePath at $Commit."
}

function Get-CommittedFileSignature {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Commit,
        [Parameter(Mandatory = $true)][string]$RelativePath
    )

    # Checkout filters (eol=crlf) are applied so the bytes match what a checkout would write.
    $tempPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ('wch-sign-{0}{1}' -f [guid]::NewGuid().ToString('N'), [System.IO.Path]::GetExtension($RelativePath))
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = 'git'
    $startInfo.Arguments = '-C "{0}" cat-file --filters "{1}:{2}"' -f $Root, $Commit, $RelativePath
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.CreateNoWindow = $true
    try {
        $process = [System.Diagnostics.Process]::Start($startInfo)
        $stream = [System.IO.File]::Create($tempPath)
        try { $process.StandardOutput.BaseStream.CopyTo($stream) } finally { $stream.Dispose() }
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) { throw "git cat-file failed for ${Commit}:$RelativePath." }
        return Get-AuthenticodeSignature -FilePath $tempPath -ErrorAction Stop
    }
    finally {
        Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
    }
}

function Read-WorkplaceCloudHubText {
    param([Parameter(Mandatory = $true)][string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $preamble = [byte[]]@()
    $encoding = $null
    $offset = 0

    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $encoding = [System.Text.UTF8Encoding]::new($false, $true)
        $preamble = [byte[]]@(0xEF, 0xBB, 0xBF)
        $offset = 3
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        $encoding = [System.Text.UnicodeEncoding]::new($false, $false, $true)
        $preamble = [byte[]]@(0xFF, 0xFE)
        $offset = 2
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        $encoding = [System.Text.UnicodeEncoding]::new($true, $false, $true)
        $preamble = [byte[]]@(0xFE, 0xFF)
        $offset = 2
    }
    else {
        $encoding = [System.Text.UTF8Encoding]::new($false, $true)
    }

    return [pscustomobject]@{
        Encoding = $encoding
        Preamble = $preamble
        Text = $encoding.GetString($bytes, $offset, $bytes.Length - $offset)
    }
}

function Write-WorkplaceCloudHubText {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Content,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text
    )

    $contentBytes = $Content.Encoding.GetBytes($Text)
    $bytes = if ($Content.Preamble.Length -gt 0) { [byte[]]($Content.Preamble + $contentBytes) } else { $contentBytes }
    [System.IO.File]::WriteAllBytes($Path, $bytes)
}

function ConvertTo-WorkplaceCloudHubCrlf {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory = $true)][string]$Path)

    $content = Read-WorkplaceCloudHubText -Path $Path
    $text = $content.Text
    $crlf = [string][char]13 + [char]10
    $lineEndingPattern = [regex]::Escape([string][char]13 + [char]10) + '|' + [regex]::Escape([string][char]13) + '|' + [regex]::Escape([string][char]10)
    $normalizedText = [regex]::Replace($text, $lineEndingPattern, $crlf)
    if ($normalizedText -ceq $text) { return $false }

    if ($PSCmdlet.ShouldProcess($Path, 'Normalize line endings to CRLF before Authenticode signing')) {
        Write-WorkplaceCloudHubText -Path $Path -Content $content -Text $normalizedText
    }
    return $true
}

function Get-SignatureBlockLayout {
    param([Parameter(Mandatory = $true)][string]$Path)

    # Lines are kept with their line endings so removing a block never touches the surrounding text.
    $content = Read-WorkplaceCloudHubText -Path $Path
    $lines = @([regex]::Matches($content.Text, '[^\r\n]*(?:\r\n|\r|\n)|[^\r\n]+$') | ForEach-Object { $_.Value })
    $kept = New-Object System.Text.StringBuilder
    $blockCount = 0
    $inBlock = $false
    $lastEndIndex = -1
    for ($index = 0; $index -lt $lines.Count; $index++) {
        $marker = $lines[$index].TrimEnd()
        if ($marker -ceq '# SIG # Begin signature block') {
            if ($inBlock) { throw "Unbalanced signature markers: nested Begin marker at line $($index + 1); file left unchanged." }
            $inBlock = $true
            continue
        }
        if ($marker -ceq '# SIG # End signature block') {
            if (-not $inBlock) { throw "Unbalanced signature markers: End marker without Begin at line $($index + 1); file left unchanged." }
            $inBlock = $false
            $blockCount++
            $lastEndIndex = $index
            continue
        }
        if ($inBlock) {
            if ($marker -and -not $marker.StartsWith('#')) { throw "Non-signature line inside a signature block at line $($index + 1); file left unchanged." }
            continue
        }
        [void]$kept.Append($lines[$index])
    }
    if ($inBlock) { throw 'Unbalanced signature markers: Begin marker without End; file left unchanged.' }

    # A single block followed only by blank lines is the normal layout that Set-AuthenticodeSignature replaces itself.
    $trailingOnly = $true
    for ($index = $lastEndIndex + 1; $index -lt $lines.Count; $index++) {
        if ($lines[$index].Trim()) { $trailingOnly = $false; break }
    }

    return [pscustomobject]@{
        Content = $content
        BlockCount = $blockCount
        NeedsRemoval = ($blockCount -gt 1 -or ($blockCount -eq 1 -and -not $trailingOnly))
        TextWithoutBlocks = $kept.ToString()
    }
}
$RepoRoot = [System.IO.Path]::GetFullPath($RepoRoot)
if (-not (Test-Path -LiteralPath (Join-Path -Path $RepoRoot -ChildPath '.git'))) {
    throw "RepoRoot is not a Git repository root: $RepoRoot"
}

$normalizedThumbprint = Normalize-Thumbprint $Thumbprint
$cert = Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert |
    Where-Object { (Normalize-Thumbprint $_.Thumbprint) -eq $normalizedThumbprint } |
    Select-Object -First 1

if (-not $cert) { throw "Code-signing certificate not found in Cert:\CurrentUser\My: $normalizedThumbprint" }
if (-not $cert.HasPrivateKey) { throw "Code-signing certificate has no private key: $normalizedThumbprint" }
$certificateEmail = [string]$cert.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::EmailName, $false)
if (-not [string]::IsNullOrWhiteSpace($ExpectedSignerEmail) -and $certificateEmail -ine $ExpectedSignerEmail) {
    throw "Code-signing certificate email mismatch. Expected '$ExpectedSignerEmail', got '$certificateEmail'."
}

if (-not $SkipTrustInstall) {
    foreach ($storeName in @('Root', 'TrustedPublisher')) {
        $installed = Install-CertificateInCurrentUserStore -Certificate $cert -StoreName $storeName
        if ($installed) { Write-Info "Installed public certificate in CurrentUser\$storeName ($normalizedThumbprint)." }
    }
}

$signed = 0
$skipped = 0
$excluded = 0
$normalized = 0
$failed = 0
$failures = New-Object System.Collections.Generic.List[string]
$changedFiles = New-Object System.Collections.Generic.List[string]

if ($PathListFile) {
    if ($IncludeUntracked) { throw '-IncludeUntracked cannot be combined with -PathListFile.' }
    if (-not (Test-Path -LiteralPath $PathListFile -PathType Leaf)) { throw "Path list file not found: $PathListFile" }

    $files = New-Object System.Collections.Generic.List[string]
    foreach ($entry in @(Get-PathListEntry -Root $RepoRoot -ListPath $PathListFile -Extensions $Extensions)) {
        if (-not (Test-Path -LiteralPath $entry.FullPath -PathType Leaf)) { continue }
        if (Test-ExcludedRepositoryPath -RelativePath $entry.RelativePath -Patterns $ExcludedPathPatterns) {
            # Third-party files are only verified, never modified, so uncommitted work does not matter here.
            $files.Add($entry.FullPath)
            continue
        }
        try {
            if (Test-WorkingTreeMatchesCommit -Root $RepoRoot -Commit $entry.Commit -RelativePath $entry.RelativePath) {
                $files.Add($entry.FullPath)
                continue
            }

            # Uncommitted changes (possibly another session's work): never modify the file.
            $committedSignature = Get-CommittedFileSignature -Root $RepoRoot -Commit $entry.Commit -RelativePath $entry.RelativePath
            $committedThumbprint = ''
            if ($committedSignature.SignerCertificate) { $committedThumbprint = Normalize-Thumbprint $committedSignature.SignerCertificate.Thumbprint }
            if ([string]$committedSignature.Status -eq 'Valid' -and ($committedThumbprint -eq $normalizedThumbprint -or (Test-MicrosoftAuthenticodeSignature -Signature $committedSignature))) {
                Write-Info "Uncommitted changes, left untouched (committed version is signed): $($entry.RelativePath)"
                $skipped++
                continue
            }
            throw 'the committed version is not validly signed and the working tree has uncommitted changes, so the file was left untouched. Commit or discard those changes, then push again.'
        }
        catch {
            $failed++
            $failures.Add("$($entry.FullPath) :: $($_.Exception.Message)") | Out-Null
        }
    }
    $files = $files.ToArray()
}
else {
    $files = @(Get-GitPowerShellFiles -Root $RepoRoot -Extensions $Extensions -IncludeUntracked:$IncludeUntracked)
}
Write-Info ("Files to inspect: {0}" -f $files.Count)

foreach ($file in $files) {
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { continue }

    try {
        $signature = Get-AuthenticodeSignature -FilePath $file -ErrorAction Stop
        $signerThumbprint = ''
        if ($signature.SignerCertificate) { $signerThumbprint = Normalize-Thumbprint $signature.SignerCertificate.Thumbprint }

        $relativePath = Get-RepositoryRelativePath -Root $RepoRoot -FullPath $file
        if (Test-ExcludedRepositoryPath -RelativePath $relativePath -Patterns $ExcludedPathPatterns) {
            if ([string]$signature.Status -ne 'Valid' -or -not $signature.SignerCertificate -or $signerThumbprint -eq $normalizedThumbprint) {
                throw "Third-party file excluded from signing has no valid publisher signature (status $($signature.Status)); it was left untouched. Restore the file exactly as published."
            }
            Write-Info "Third-party file left untouched (signed by $($signature.SignerCertificate.Subject)): $relativePath"
            $excluded++
            continue
        }

        if (-not $Force -and [string]$signature.Status -eq 'Valid' -and (Test-MicrosoftAuthenticodeSignature -Signature $signature)) {
            $skipped++
            continue
        }

        # Validates the markers before any change so an unbalanced file is never modified.
        $blockLayout = Get-SignatureBlockLayout -Path $file

        $lineEndingsChanged = $false
        if (-not $SkipLineEndingNormalization) {
            $lineEndingsChanged = ConvertTo-WorkplaceCloudHubCrlf -Path $file
            if ($lineEndingsChanged) {
                $normalized++
                $changedFiles.Add($file)
            }
        }

        if (-not $Force -and -not $lineEndingsChanged -and [string]$signature.Status -eq 'Valid' -and $signerThumbprint -eq $normalizedThumbprint) {
            $skipped++
            continue
        }

        # Old blocks followed by code (or several blocks) make Set-AuthenticodeSignature fail with HashMismatch:
        # remove every block, keeping all code before, between and after them, then sign.
        if ($blockLayout.NeedsRemoval -and $PSCmdlet.ShouldProcess($file, "Remove $($blockLayout.BlockCount) misplaced signature block(s)")) {
            $blockLayout = Get-SignatureBlockLayout -Path $file
            Write-WorkplaceCloudHubText -Path $file -Content $blockLayout.Content -Text $blockLayout.TextWithoutBlocks
            Write-Info "Removed $($blockLayout.BlockCount) misplaced signature block(s): $file"
            if (-not $changedFiles.Contains($file)) { $changedFiles.Add($file) }
        }

        if ($PSCmdlet.ShouldProcess($file, 'Set Authenticode signature')) {
            $signingParams = @{
                FilePath = $file
                Certificate = $cert
                HashAlgorithm = 'SHA256'
                ErrorAction = 'Stop'
            }
            if (-not [string]::IsNullOrWhiteSpace($TimestampServer)) {
                $signingParams['TimestampServer'] = $TimestampServer
            }
            $result = Set-AuthenticodeSignature @signingParams
            if ([string]$result.Status -ne 'Valid') {
                throw "Signature status is $($result.Status): $($result.StatusMessage)"
            }
            $verifiedSignature = Get-AuthenticodeSignature -FilePath $file -ErrorAction Stop
            if ([string]$verifiedSignature.Status -ne 'Valid') {
                throw "Signature verification status is $($verifiedSignature.Status): $($verifiedSignature.StatusMessage)"
            }
            if (-not [string]::IsNullOrWhiteSpace($TimestampServer) -and -not $verifiedSignature.TimeStamperCertificate) {
                throw "The signature is valid but no timestamp countersignature was returned by $TimestampServer."
            }
            $signed++
            if (-not $changedFiles.Contains($file)) { $changedFiles.Add($file) }
        }
    }
    catch {
        $failed++
        $failures.Add("$file :: $($_.Exception.Message)") | Out-Null
    }
}

Write-Info ("Normalized={0}; Signed={1}; Skipped={2}; ThirdParty={3}; Failed={4}" -f $normalized, $signed, $skipped, $excluded, $failed)
if ($failed -gt 0) {
    $failures | ForEach-Object { Write-Host $_ -ForegroundColor Red }
    exit 1
}
if ($PathListFile -and $changedFiles.Count -gt 0) {
    $changedFiles | ForEach-Object { Write-Info "Updated: $_" }
    exit 3
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD09bJC+tZKpRhh
# QNohKo9eMvQVj/QyLzLe/u2i4glozaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJUi3Vrsinr/tRGcwtRwd3fX9XrqNRThd/E+HCltlfAuMA0GCSqG
# SIb3DQEBAQUABIIBgHaZZ5/7kGcZT1k0PyzsWT2FxnOaC6fYWIwSU7t/o0L8iQ+Z
# BUqT0PSIhMR74uC0ZadVDpxYUgqY6JOD6RydOschWrcLjg9VE4zQ8OHOkJI1EX3T
# zUkifwTEVoxLqSTt4+u/MHnw1w8E9QBcn9Mf6vxqBGq6MfFcWqWzfBcATuZ08VZC
# A0nIndqu8p2qnf6dRcG5+t64E0twmZJyjdLhMuL13YXc8lz6z1yBMFaxGyUcNTzr
# nM8evLYjTa55u4q/HUAiNghrMWYg8HlS48CYiTp+g2yHTmFVZzK+UQ9jQdZj0Cb7
# jiQfFcUSit/zi5bZuVhy8hi3crAQ3BUR+NCMzPshA4t1npPSoL8XsHmw9ns3TcSx
# dLaCUeUBxui70Ews1VFY70VoYEjZdr2vwEMgJ3uDoZNfpTycU+8IDlgA8FguDeO5
# dPih3rSJiU0/VGRGmhMwcK5x2KuAoKwArgrY+Bn5rANvLB3YY/EM+dFIb3mggKT/
# 67yLchqpIarRRM2qC6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjExNTMy
# MzZaMC8GCSqGSIb3DQEJBDEiBCAiTBN7daGljCH0Z9/Mfkf6ABhFjmtzVd9V1pl3
# SEHyoTANBgkqhkiG9w0BAQEFAASCAgAvEE23P/R51skBnS+25uhXdK4IBivMWl3K
# bp9waey/EbRBPUPL5A8adSPkzVBwLO35TVFYLTYIJXmJyl8jR2eW3FF53NWBeymV
# UoiUMPCwk7f2R7642T2/8jEoow2vtNj3M0WCXvbEh9KgSOj6stko/zX7WCjJP19D
# 18ygJWd3zKyR7jNvsc6Vh+SYfNC0ejrqZKgC+OA3M9MeKueOqZU7Lq5rgHjAcD6u
# XgPreClJ7FxLBUaO/6z2aEMn2exj3PfzzMH3OvvOI8ysK1ItIEvFkwBm4xIHjjZL
# +hhgVwWxKs9wVUBiD2T/z5hV08f4aLlIr0SeFDwrXm+TrEo5qzmhDdetPu7fRZaV
# Yo8l3OUGz0o6J0sU0j13Oj0sc2vuOh3AAVcicElDSts6PpGGwrdRGV7UF45ITIRR
# uOZ4gCrS7aT/CRpd4HjF8Djv2kRTOR6KrQ663YLIIPDB9Zb0XZ2yyYLlbnQzmuHt
# khOL1QW+2+q9Mr0fRMM6NGKLOPmnTlm2fgXpyvb5PgvtgyF/78JBwvnGXdk2r3mJ
# 5/klHOGDh1CHA2UocQSSEsmKXvI7d4I0rd8dlGRPJej1Gg8M9HaKC0PMqzUEU15l
# zB8bFEwP1qyeFn5eDnM+JxtPWibWugCpziA1Z8o0JO33wOrOZ1sC/RmEArazptTw
# 3D1B/Hq7PQ==
# SIG # End signature block
