<#
.SYNOPSIS
End-to-end test of the WorkplaceCloudHub pre-push signing hook.

.DESCRIPTION
Version 1.0.0

Builds a throwaway repository with a bare remote, installs the current pre-push hook and signing
script, then pushes through the real hook. Checks that:
- committed unsigned scripts are signed and the push is blocked until the signatures are committed;
- untracked scripts, and tracked scripts with uncommitted changes outside the pushed commits, stay
  byte-identical;
- a pushed script whose committed version is unsigned and whose working tree has uncommitted
  changes is left untouched and the push is refused;
- a push without PowerShell changes skips signing.

Requires the workplacecloudhub.com code-signing certificate in Cert:\CurrentUser\My and access to
the timestamp server.
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent),
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
foreach ($source in @($hookSource, $signerSource)) {
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
    Write-LfFile -RelativePath '.gitattributes' -Lines @('* text=auto', '*.ps1 text eol=crlf', '.githooks/* text eol=lf')
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
