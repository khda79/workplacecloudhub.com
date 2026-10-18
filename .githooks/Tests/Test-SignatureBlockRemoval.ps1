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
