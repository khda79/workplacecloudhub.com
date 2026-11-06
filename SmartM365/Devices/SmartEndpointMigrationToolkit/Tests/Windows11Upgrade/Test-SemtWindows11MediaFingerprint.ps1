<#
.SYNOPSIS
    Offline tests of the setup media manifest value used to compare two Windows 11 setup media.
.DESCRIPTION
    Extracts the real Get-SetupMediaIntegrityManifestHash function of the Windows11Upgrade
    endpoint script and runs it on synthetic manifests. Checks that two manifests describing the
    same files give the same value whatever their row order, encoding, line ending, path case
    or separator, and that any change of path, length or hash, or an added file, gives another
    value. Also checks the fallback for a manifest that cannot be read as rows, and that the
    fingerprint comparison still uses the manifest value.
    Prints one PASS/FAIL line per assertion and "Summary: Passed=N; Failed=M". Exits 1 on failure.
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

$ErrorActionPreference = 'Stop'
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$w11Path = Join-Path $root 'Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'
$passed = 0
$failed = 0
function Assert-Test([string]$Name, [bool]$Condition) {
    if ($Condition) { $script:passed++; Write-Host "PASS: $Name" }
    else { $script:failed++; Write-Host "FAIL: $Name" }
}

$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($w11Path, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ('Windows 11 upgrade script does not parse: {0}' -f $parseErrors[0].Message) }
$functionText = @{}
foreach ($name in 'Get-SetupMediaIntegrityManifestHash','Get-SetupMediaIntegrityManifestPath','Get-SmartFileSha256','Test-SetupFingerprintMatch') {
    $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    if ($null -eq $definition) { throw "Function not found in the Windows 11 upgrade script: $name" }
    $functionText[$name] = $definition.Extent.Text
    . ([scriptblock]::Create($definition.Extent.Text))
}
function ConvertTo-LongLiteralPath { param([string]$Path) return $Path }
$script:SetupMediaManifestFileName = 'SmartM365-SetupMediaManifest.sha256.csv'

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('SemtMediaFingerprint-{0}' -f [guid]::NewGuid().ToString('N'))
$rows = @(
    [pscustomobject]@{ RelativePath = 'setup.exe'; Length = '102400'; SHA256 = 'AAAA0000AAAA0000AAAA0000AAAA0000AAAA0000AAAA0000AAAA0000AAAA0000' },
    [pscustomobject]@{ RelativePath = 'sources\install.wim'; Length = '5000000000'; SHA256 = 'BBBB1111BBBB1111BBBB1111BBBB1111BBBB1111BBBB1111BBBB1111BBBB1111' },
    [pscustomobject]@{ RelativePath = 'sources\lang.ini'; Length = '120'; SHA256 = 'CCCC2222CCCC2222CCCC2222CCCC2222CCCC2222CCCC2222CCCC2222CCCC2222' }
)

function New-ManifestMedia {
    param([string]$Name, [object[]]$Rows, [string]$Encoding = 'UTF8', [string]$NewLine = "`r`n", [switch]$Raw, [string]$RawText)
    $media = Join-Path $tempRoot $Name
    New-Item -ItemType Directory -Path $media -Force | Out-Null
    $path = Join-Path $media 'SmartM365-SetupMediaManifest.sha256.csv'
    if ($Raw) { [System.IO.File]::WriteAllText($path, $RawText, (New-Object System.Text.UTF8Encoding($false))); return $media }
    $lines = @($Rows | ConvertTo-Csv -NoTypeInformation)
    $text = ($lines -join $NewLine) + $NewLine
    $encodingObject = switch ($Encoding) {
        'UTF8' { New-Object System.Text.UTF8Encoding($true) }
        'UTF8NoBom' { New-Object System.Text.UTF8Encoding($false) }
        'ASCII' { [System.Text.Encoding]::ASCII }
        'Unicode' { [System.Text.Encoding]::Unicode }
    }
    [System.IO.File]::WriteAllText($path, $text, $encodingObject)
    return $media
}

try {
    $reference = Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'reference' -Rows $rows)
    Assert-Test 'manifest value is a content value' ($reference -match '^content:[0-9A-F]{64}$')

    $reordered = Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'reordered' -Rows @($rows[2],$rows[0],$rows[1]))
    Assert-Test 'other row order: same value' ($reordered -eq $reference)

    $asciiLf = Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'ascii-lf' -Rows $rows -Encoding ASCII -NewLine "`n")
    Assert-Test 'ASCII with LF line endings: same value' ($asciiLf -eq $reference)

    $unicode = Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'unicode' -Rows $rows -Encoding Unicode)
    Assert-Test 'UTF-16 encoding: same value' ($unicode -eq $reference)

    $caseRows = @($rows | ForEach-Object { [pscustomobject]@{ RelativePath = ('\' + ([string]$_.RelativePath).ToLowerInvariant().Replace('\','/')); Length = (' {0} ' -f $_.Length); SHA256 = ([string]$_.SHA256).ToLowerInvariant() } })
    $caseValue = Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'case' -Rows $caseRows -Encoding UTF8NoBom)
    Assert-Test 'path case, separator, leading slash, spaces and hash case: same value' ($caseValue -eq $reference)

    $changedHash = @($rows | ForEach-Object { $_.PSObject.Copy() }); $changedHash[1].SHA256 = 'DDDD3333DDDD3333DDDD3333DDDD3333DDDD3333DDDD3333DDDD3333DDDD3333'
    Assert-Test 'changed file hash: other value' ((Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'hash' -Rows $changedHash)) -ne $reference)

    $changedLength = @($rows | ForEach-Object { $_.PSObject.Copy() }); $changedLength[0].Length = '102401'
    Assert-Test 'changed file length: other value' ((Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'length' -Rows $changedLength)) -ne $reference)

    $changedPath = @($rows | ForEach-Object { $_.PSObject.Copy() }); $changedPath[2].RelativePath = 'sources\lang2.ini'
    Assert-Test 'changed file path: other value' ((Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'path' -Rows $changedPath)) -ne $reference)

    $extraRows = @($rows) + @([pscustomobject]@{ RelativePath = 'sources\extra.dll'; Length = '1'; SHA256 = 'EEEE4444EEEE4444EEEE4444EEEE4444EEEE4444EEEE4444EEEE4444EEEE4444' })
    Assert-Test 'added file: other value' ((Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'extra' -Rows $extraRows)) -ne $reference)

    $missingRows = @($rows[0],$rows[1])
    Assert-Test 'missing file: other value' ((Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'missing' -Rows $missingRows)) -ne $reference)

    $emptyMedia = Join-Path $tempRoot 'no-manifest'; New-Item -ItemType Directory -Path $emptyMedia -Force | Out-Null
    Assert-Test 'no manifest: empty value' ((Get-SetupMediaIntegrityManifestHash -MediaRoot $emptyMedia) -eq '')

    $badMedia = New-ManifestMedia -Name 'bad' -Raw -RawText "Path,Size`r`nsetup.exe,1`r`n"
    $badValue = Get-SetupMediaIntegrityManifestHash -MediaRoot $badMedia
    $badFileHash = (Get-FileHash -LiteralPath (Join-Path $badMedia 'SmartM365-SetupMediaManifest.sha256.csv') -Algorithm SHA256).Hash
    Assert-Test 'manifest without the expected columns: file hash fallback' ($badValue -eq ('file:' + $badFileHash))

    # The fingerprint comparison still includes the manifest value.
    $left = [pscustomobject]@{ MediaId = 'Win11'; ExpectedLanguage = 'fr-FR'; SetupExeLength = 1; SetupExeHash = 'A'; InstallImageName = 'install.wim'; InstallImageLength = 2; LangIniHash = 'B'; IntegrityManifestHash = $reference }
    $right = $left.PSObject.Copy(); $right.IntegrityManifestHash = $reordered
    Assert-Test 'fingerprints of reordered manifests match' (Test-SetupFingerprintMatch -Left $left -Right $right)
    $right.IntegrityManifestHash = 'content:' + ('0' * 64)
    Assert-Test 'fingerprints with another manifest content do not match' (-not (Test-SetupFingerprintMatch -Left $left -Right $right))
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $passed,$failed)
if ($failed -gt 0) { exit 1 }
exit 0
