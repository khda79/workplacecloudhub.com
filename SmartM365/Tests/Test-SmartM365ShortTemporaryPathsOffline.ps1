<#
.SYNOPSIS
Offline check that atomic writes stay under MAX_PATH (260) for deep targets, as on Windows PowerShell 5.1 /
.NET Framework without long path support (launcher cache copies of the repository): every file the JSON
transport, the Core and the Windows PowerShell 5.1 module create while writing a target whose full path
is 245 characters long must itself stay below 260 characters.
.VERSION
1.0.0
.NOTES
Runs in the current host and, from PowerShell 7, again in Windows PowerShell 5.1 with the 5.1 module.
Long paths may be enabled on the test machine, so file creation is observed rather than expected to fail.
Synthetic TEMP files only; no tenant, network or synchronized data access.
#>
[CmdletBinding()]
param([switch]$InnerWindowsPowerShell)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$coreFolder = Join-Path $root 'Modules/SmartM365.Core'
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }

Import-Module (Join-Path $coreFolder 'SmartM365.JsonTransport.psd1') -MinimumVersion '1.0.8' -Global -Force -ErrorAction Stop
$isWindowsPowerShell = $PSVersionTable.PSEdition -ne 'Core'
if ($isWindowsPowerShell) {
    Import-Module (Join-Path $coreFolder 'Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psd1') -MinimumVersion '1.0.63' -Global -Force -WarningAction SilentlyContinue -ErrorAction Stop
}
else {
    Import-Module (Join-Path $coreFolder 'SmartM365.Core.psd1') -MinimumVersion '1.0.102' -Global -Force -WarningAction SilentlyContinue -ErrorAction Stop
}

# A folder deep enough for a 245-character target named like the Exchange infrastructure local configuration.
$leaf = 'SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.local.json.txt'
$base = Join-Path ([IO.Path]::GetTempPath()) ('sm-maxpath-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$folder = $base
while (($folder.Length + 1 + $leaf.Length) -lt 230) { $folder = Join-Path $folder 'SmartInventory' }
$folder = Join-Path $folder ('p' * (245 - $folder.Length - 2 - $leaf.Length))
$target = Join-Path $folder $leaf
[void][IO.Directory]::CreateDirectory($folder)
Check ($target.Length -eq 245) "Synthetic target length is $($target.Length), not 245."

$sourceId = 'SmartM365ShortTemporaryPaths-' + [guid]::NewGuid().ToString('N')
$watcher = New-Object IO.FileSystemWatcher $folder
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true
$null = Register-ObjectEvent -InputObject $watcher -EventName Created -SourceIdentifier $sourceId
try {
    # Helper: same folder, shorter than the target whenever the target name is longer than the short name.
    $temporary = Get-SmartM365SiblingTemporaryPath -Path $target -Extension '.pending'
    Check ([IO.Path]::GetDirectoryName($temporary) -eq $folder -and $temporary.Length -lt $target.Length -and $temporary -like '*.pending') "Sibling temporary path is not short or not beside the target: $temporary"

    # JSON transport: create then replace (temporary and previous copies).
    $json = [Text.Encoding]::UTF8.GetBytes('{"Name":"synthetic","Version":1}')
    Write-SmartM365JsonBytesAtomically -Path $target -Bytes $json -Validate { param($document) $null = $document } -ExpectedSHA256 'ABSENT' | Out-Null
    $json2 = [Text.Encoding]::UTF8.GetBytes('{"Name":"synthetic","Version":2}')
    $hash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
    Write-SmartM365JsonBytesAtomically -Path $target -Bytes $json2 -Validate { param($document) $null = $document } -ExpectedSHA256 $hash | Out-Null
    Check ([IO.File]::ReadAllText($target) -like '*"Version":2*') 'The JSON transport did not replace the deep target.'

    # Core (PowerShell 7) or the Windows PowerShell 5.1 module: atomic copy and text write next to a deep target.
    $copyTarget = Join-Path $folder ($leaf -replace '\.local\.json\.txt$', '_Copy.csv')
    $source = Join-Path $base 'source.csv'
    [IO.File]::WriteAllText($source, "Id`r`n1`r`n")
    Copy-SmartM365FileAtomically -SourcePath $source -DestinationPath $copyTarget | Out-Null
    Write-SmartM365TextAtomically -Path $copyTarget -Content "Id`r`n2`r`n"
    Check ([IO.File]::ReadAllText($copyTarget) -like '*2*') 'The atomic text write did not replace the deep target.'

    Start-Sleep -Milliseconds 800
    $created = @(Get-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue | ForEach-Object { [string]$_.SourceEventArgs.FullPath })
    $tooLong = @($created | Where-Object { $_.Length -ge 260 })
    Check ($created.Count -ge 4) "Too few file creations observed ($($created.Count))."
    Check ($tooLong.Count -eq 0) ("Files created at or beyond MAX_PATH: " + (($tooLong | ForEach-Object { "$($_.Length) $([IO.Path]::GetFileName($_))" }) -join '; '))

    # Launcher case: the cache is rebuilt from the share on every run, so a script local configuration missing
    # on the share is created again from its template, then must not be rewritten while no key is missing.
    Get-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue | Remove-Event
    . (Join-Path $root 'Config/SmartM365-TenantContext.ps1') | Out-Null
    Remove-Item -LiteralPath $target -Force
    $templatePath = $target + '.template'
    [IO.File]::WriteAllText($templatePath, '{"Alpha":"__USE_GLOBAL__","HealthCheckerTimeoutMinutesPerServer":30,"ExcludedServers":[],"Nested":{"Value":1}}')
    $legacyPath = $target -replace '\.txt$', ''
    $config = Read-SmartM365JsonConfig -Path $legacyPath -Required 6>$null
    Check ((Test-Path -LiteralPath $target) -and $config['HealthCheckerTimeoutMinutesPerServer'] -eq 30) 'The deep local configuration was not created from its template.'
    $createdHash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
    $null = Read-SmartM365JsonConfig -Path $legacyPath -Required 6>$null
    Check ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -eq $createdHash) 'A complete local configuration was rewritten.'
    # Same keys, another order and other values (an edit made with ConvertTo-Json): still no rewrite.
    $edited = [Text.Encoding]::UTF8.GetBytes('{"Nested":{"Value":2},"ExcludedServers":["Server1"],"HealthCheckerTimeoutMinutesPerServer":45,"Alpha":"x"}')
    [IO.File]::WriteAllBytes($target, $edited)
    $editedHash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
    $config = Read-SmartM365JsonConfig -Path $legacyPath -Required 6>$null
    Check ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -eq $editedHash -and $config['HealthCheckerTimeoutMinutesPerServer'] -eq 45) 'A reordered local configuration with every template key was rewritten.'
    Start-Sleep -Milliseconds 800
    $created = @(Get-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue | ForEach-Object { [string]$_.SourceEventArgs.FullPath })
    $tooLong = @($created | Where-Object { $_.Length -ge 260 })
    Check ($tooLong.Count -eq 0) ("Configuration files created at or beyond MAX_PATH: " + (($tooLong | ForEach-Object { "$($_.Length) $([IO.Path]::GetFileName($_))" }) -join '; '))
}
finally {
    Unregister-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue
    Get-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue | Remove-Event
    $watcher.Dispose()
    if (Test-Path -LiteralPath $base) { cmd /c rd /s /q "$base" | Out-Null }
}

# Static: no temporary name built from the full target name plus a GUID in the Core files.
$files = @('SmartM365.JsonTransport.psm1', 'SmartM365.Core.psm1', 'SmartM365-CmdbReceipt.ps1', 'SmartM365-DailyCaptureRetention.ps1', 'Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1') | ForEach-Object { Join-Path $coreFolder $_ }
$files += Join-Path $root 'Config/SmartM365-TenantContext.ps1'
foreach ($file in $files) {
    $long = @(Select-String -LiteralPath $file -Pattern "\[guid\]::NewGuid\(\)[^\r\n]*\.(tmp|pending|previous|partial|download)\b" | Where-Object { $_.Line -notmatch 'GetTempPath|~sm' })
    Check ($long.Count -eq 0) ("Long temporary name in {0}: line {1}" -f [IO.Path]::GetFileName($file), (($long | ForEach-Object { $_.LineNumber }) -join ','))
}

if (-not $isWindowsPowerShell -and -not $InnerWindowsPowerShell) {
    # The production case: Windows PowerShell 5.1 with the 5.1 module.
    $windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path -LiteralPath $windowsPowerShell) {
        $output = & $windowsPowerShell -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath -InnerWindowsPowerShell 2>&1 | Out-String
        Check ($LASTEXITCODE -eq 0 -and $output -match 'PASS:') ("Windows PowerShell 5.1 run failed: " + $output.Trim())
        Write-Output $output.Trim()
    }
}
"PASS: $script:checks short temporary path checks; PowerShell $($PSVersionTable.PSVersion)."
