# JSON transport names are independent from the JSON payload. Import has no side effects.
Set-StrictMode -Version 2.0

function Get-SmartM365JsonNames {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $full = [IO.Path]::GetFullPath($Path)
    if ($full.EndsWith('.json.txt', [StringComparison]::OrdinalIgnoreCase)) {
        $legacy = $full.Substring(0, $full.Length - 4)
    } elseif ($full.EndsWith('.json', [StringComparison]::OrdinalIgnoreCase)) {
        $legacy = $full
    } else { throw "Not an owned JSON transport filename: $Path" }
    [pscustomobject]@{ Legacy = $legacy; Preferred = $legacy + '.txt'; Lock = $legacy + '.transport.lock'; Journal = $legacy + '.migration.log' }
}

function Assert-SmartM365JsonUnlinkedPath {
    param([Parameter(Mandatory)][string]$Path)
    $cursor = [IO.Path]::GetFullPath($Path)
    $volumeRoot = [IO.Path]::GetPathRoot($cursor).TrimEnd('\', '/')
    while ($cursor) {
        if (Test-Path -LiteralPath $cursor) {
            $item = Get-Item -LiteralPath $cursor -Force -ErrorAction Stop
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Linked path is not eligible for JSON migration: $cursor" }
        }
        if ($cursor.TrimEnd('\', '/') -eq $volumeRoot) { break }
        $parent = [IO.Path]::GetDirectoryName($cursor)
        if ($parent -eq $cursor) { break }
        $cursor = $parent
    }
}

function Get-SmartM365JsonReadPath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [switch]$Optional)
    $names = Get-SmartM365JsonNames $Path
    # Test existence, not just Leaf: a directory at the preferred name must not enable fallback.
    foreach ($candidate in @($names.Preferred, $names.Legacy)) {
        if (Test-Path -LiteralPath $candidate -ErrorAction Stop) {
            if (-not (Test-Path -LiteralPath $candidate -PathType Leaf -ErrorAction Stop)) { throw "JSON path is not a file: $candidate" }
            return $candidate
        }
    }
    if (-not $Optional) { throw "JSON file missing: $($names.Preferred) (legacy: $($names.Legacy))" }
}

function Get-SmartM365JsonHash {
    param([Parameter(Mandatory)][byte[]]$Bytes)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($Bytes))).Replace('-', '') }
    finally { $sha.Dispose() }
}

function ConvertFrom-SmartM365JsonBytes {
    param([Parameter(Mandatory)][AllowEmptyCollection()][byte[]]$Bytes, [scriptblock]$Validate)
    if ($Bytes.Length -eq 0) { throw 'Empty JSON payload.' }
    $memory = [IO.MemoryStream]::new($Bytes, $false)
    $reader = [IO.StreamReader]::new($memory, [Text.UTF8Encoding]::new($false, $true), $true)
    try { $content = $reader.ReadToEnd() } finally { $reader.Dispose(); $memory.Dispose() }
    if ([string]::IsNullOrWhiteSpace($content)) { throw 'Blank JSON payload.' }
    $document = ConvertFrom-Json -InputObject $content -ErrorAction Stop
    if ($null -eq $document -and $content.TrimStart().StartsWith('[')) { $document = @() }
    if ($null -eq $document) { throw 'Null JSON payload cannot represent a valid owned state.' }
    if ($Validate) { & $Validate $document | Out-Null }
    # Preserve array shape instead of PowerShell pipeline unrolling.
    return ,$document
}

function Read-SmartM365JsonDocument {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [scriptblock]$Validate)
    $selected = Get-SmartM365JsonReadPath $Path
    $bytes = [IO.File]::ReadAllBytes($selected)
    # No catch/fallback: an invalid preferred file must surface to the owner.
    $document = ConvertFrom-SmartM365JsonBytes -Bytes $bytes -Validate $Validate
    $names = Get-SmartM365JsonNames $Path
    $coexistence = 'Single'
    if ($selected -eq $names.Preferred -and (Test-Path -LiteralPath $names.Legacy)) {
        $oldBytes = [IO.File]::ReadAllBytes($names.Legacy)
        $coexistence = if ((Get-SmartM365JsonHash $oldBytes) -eq (Get-SmartM365JsonHash $bytes)) { 'Identical' } else { 'Conflict' }
        if ($coexistence -eq 'Conflict') {
            # A regenerated replacement may have committed just before interruption.
            # Only the durable, exact old/new hash pair can authorize that coexistence.
            $receipt = Get-SmartM365LastJsonMigrationEvent $names.Journal
            if ($receipt -and $receipt.Phase -in @('ReplacementPrepared','ReplacementPublished') -and $receipt.SHA256 -eq (Get-SmartM365JsonHash $bytes) -and $receipt.Detail -eq (Get-SmartM365JsonHash $oldBytes)) {
                $coexistence = 'PublishedReplacement'
            } else { throw "Conflicting JSON transport names: $($names.Legacy) and $selected" }
        }
    }
    [pscustomobject]@{ Path = $selected; Document = $document; SHA256 = (Get-SmartM365JsonHash $bytes); Coexistence = $coexistence }
}

function Enter-SmartM365JsonTransportLock {
    param([Parameter(Mandatory)][string]$Path, [int]$TimeoutSeconds = 10)
    $deadline = [datetime]::UtcNow.AddSeconds($TimeoutSeconds)
    do {
        try { return [IO.File]::Open($Path, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
        catch [IO.IOException] {
            if ([datetime]::UtcNow -ge $deadline) { throw "JSON transport lock unavailable: $Path" }
            Start-Sleep -Milliseconds 100
        }
    } while ($true)
}

function Write-SmartM365JsonMigrationEvent {
    param([string]$Journal, [string]$Owner, [string]$Phase, [string]$SHA256, [string]$Detail)
    $line = ([ordered]@{ Utc = [datetime]::UtcNow.ToString('o'); Owner = $Owner; Phase = $Phase; SHA256 = $SHA256; Detail = $Detail } | ConvertTo-Json -Compress) + [Environment]::NewLine
    if([IO.File]::Exists($Journal)){
        $check=[IO.File]::Open($Journal,'Open','Read','ReadWrite')
        try{if($check.Length -gt 0){$null=$check.Seek(-1,'End');if($check.ReadByte() -ne 10){$line=[Environment]::NewLine+$line}}}finally{$check.Dispose()}
    }
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes($line)
    $stream = [IO.File]::Open($Journal, 'Append', 'Write', 'Read')
    try { $stream.Write($bytes, 0, $bytes.Length); $stream.Flush($true) } finally { $stream.Dispose() }
}

function Get-SmartM365LastJsonMigrationEvent {
    param([Parameter(Mandatory)][string]$Path)
    if(-not [IO.File]::Exists($Path)){return $null}
    $lines=@(Get-Content -LiteralPath $Path -Tail 32 -ErrorAction Stop)
    for($i=$lines.Count-1;$i -ge 0;$i--){
        if([string]::IsNullOrWhiteSpace($lines[$i])){continue}
        try{
            $event=ConvertFrom-Json -InputObject $lines[$i] -ErrorAction Stop
            if(-not $event.PSObject.Properties['Phase'] -or -not $event.PSObject.Properties['Owner']){throw 'Incomplete journal event.'}
            return $event
        }catch{Write-Warning "Incomplete migration journal tail preserved; checking preceding durable event: $Path"}
    }
    throw "Migration journal has no recoverable event: $Path"
}

function Move-SmartM365OwnedJsonFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$RelativePath,
        [Parameter(Mandatory)][string]$Owner,
        [Parameter(Mandatory)][scriptblock]$Validate,
        [switch]$QualifiedUnc,
        [switch]$RemoveIdenticalLegacy,
        [int]$LockTimeoutSeconds = 10
    )
    if ([IO.Path]::IsPathRooted($RelativePath)) { throw 'Migration requires an owner-relative filename.' }
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    $names = Get-SmartM365JsonNames (Join-Path $Root $RelativePath)
    if (-not $names.Legacy.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) { throw 'Migration path escapes the owner root.' }
    if ($names.Legacy.StartsWith('\\') -and -not $QualifiedUnc) { throw 'UNC migration requires qualification on the target share before activation.' }
    foreach ($path in @($names.Legacy, $names.Preferred, $names.Lock, $names.Journal)) { Assert-SmartM365JsonUnlinkedPath $path }
    $selected = Get-SmartM365JsonReadPath $names.Legacy -Optional
    if (-not $selected) { return [pscustomobject]@{ Status = 'Absent'; Path = $names.Preferred; SHA256 = '' } }
    $lock = Enter-SmartM365JsonTransportLock $names.Lock $LockTimeoutSeconds
    try {
        $current = Read-SmartM365JsonDocument $names.Legacy -Validate $Validate
        $hash = $current.SHA256
        if ($current.Path -eq $names.Preferred) {
            if($current.Coexistence -eq 'PublishedReplacement'){throw 'Regenerated replacement must be completed by its producer before persistent migration.'}
            if ($current.Coexistence -eq 'Identical' -and $RemoveIdenticalLegacy) {
                Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'DeduplicatePrepared' $hash ''
                if ((Get-FileHash -LiteralPath $names.Legacy -Algorithm SHA256).Hash -ne $hash) { throw 'Legacy file changed during migration.' }
                [IO.File]::Delete($names.Legacy)
                Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'Completed' $hash 'Identical legacy removed'
            } elseif ($current.Coexistence -eq 'Identical') {
                Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'PendingLegacy' $hash 'Publication/compatibility confirmation required'
                return [pscustomobject]@{ Status = 'PendingLegacy'; Path = $names.Preferred; SHA256 = $hash }
            } else {
                $last=Get-SmartM365LastJsonMigrationEvent $names.Journal
                if($last -and $last.Phase -ne 'Completed'){
                    if($last.Owner -ne $Owner){throw 'Migration journal belongs to another owner.'}
                    Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'Completed' $hash 'Resumed after rename'
                }
            }
            return [pscustomobject]@{ Status = 'Completed'; Path = $names.Preferred; SHA256 = $hash }
        }
        Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'Prepared' $hash ''
        if ((Get-FileHash -LiteralPath $names.Legacy -Algorithm SHA256).Hash -ne $hash) { throw 'Legacy file changed before rename.' }
        # Same directory; never emulate failed rename with copy/delete, especially on UNC.
        [IO.File]::Move($names.Legacy, $names.Preferred)
        if ((Get-FileHash -LiteralPath $names.Preferred -Algorithm SHA256).Hash -ne $hash) { throw 'JSON integrity failure after rename.' }
        Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'Completed' $hash ''
        [pscustomobject]@{ Status = 'Completed'; Path = $names.Preferred; SHA256 = $hash }
    } catch {
        $failure = $_
        try { Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'Failed' '' $failure.Exception.Message } catch { Write-Warning 'Could not persist JSON migration failure journal.' }
        throw $failure
    } finally { $lock.Dispose() }
}

function Get-SmartM365JsonTransportPolicy {
    [CmdletBinding()]
    param()
    $path = Join-Path $PSScriptRoot '../../Config/SmartM365-JsonTransport.policy.psd1'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw 'JSON transport deployment policy missing. Incomplete deployment.' }
    $policy = Import-PowerShellDataFile -LiteralPath $path -ErrorAction Stop
    if ($policy.Mode -notin @('Readers', 'JsonText')) { throw 'Unsupported JSON transport deployment mode.' }
    return $policy
}

function Write-SmartM365JsonBytesAtomically {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][AllowEmptyCollection()][byte[]]$Bytes,
        [Parameter(Mandatory)][scriptblock]$Validate,
        [string]$ExpectedSHA256 = '',
        [switch]$RetireLegacyAfterPublication,
        [string]$Owner = '',
        [int]$LockTimeoutSeconds = 10
    )
    $names = Get-SmartM365JsonNames $Path
    # Caller chooses the deployment-phase path. Serialization is never performed here.
    $target = [IO.Path]::GetFullPath($Path)
    foreach ($candidate in @($target, $names.Lock)) { Assert-SmartM365JsonUnlinkedPath $candidate }
    $null = ConvertFrom-SmartM365JsonBytes -Bytes $Bytes -Validate $Validate
    $lock = Enter-SmartM365JsonTransportLock $names.Lock $LockTimeoutSeconds
    $temporary = $target + '.' + [guid]::NewGuid().ToString('N') + '.pending'
    $backup = $target + '.' + [guid]::NewGuid().ToString('N') + '.previous'
    $restoreAttributes = $null
    $published = $false
    try {
        $legacyHash = ''
        if ($RetireLegacyAfterPublication) {
            if ($target -ne $names.Preferred -or [string]::IsNullOrWhiteSpace($Owner)) { throw 'Replacement retirement requires a preferred target and an explicit owner.' }
            foreach ($candidate in @($names.Legacy, $names.Journal)) { Assert-SmartM365JsonUnlinkedPath $candidate }
            if (Get-SmartM365JsonReadPath $Path -Optional) {
                $prior = Read-SmartM365JsonDocument -Path $Path -Validate $Validate
                if ($prior.Coexistence -eq 'PublishedReplacement') {
                    # Finish the already verified publication before preparing another one.
                    $oldHash = (Get-FileHash -LiteralPath $names.Legacy -Algorithm SHA256).Hash
                    $receipt = Get-SmartM365LastJsonMigrationEvent $names.Journal
                    if ($receipt.Owner -ne $Owner -or $receipt.Detail -ne $oldHash -or $receipt.SHA256 -ne $prior.SHA256) { throw 'Replacement recovery receipt does not match this owner and file pair.' }
                    [IO.File]::Delete($names.Legacy)
                    Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'Completed' $prior.SHA256 'Recovered previously published replacement'
                }
            }
            if ([IO.File]::Exists($names.Legacy)) { $legacyHash = (Get-FileHash -LiteralPath $names.Legacy -Algorithm SHA256).Hash }
        }
        if ($ExpectedSHA256 -eq 'ABSENT') {
            if (Test-Path -LiteralPath $target) { throw [IO.IOException]::new('JSON target was created by another writer.') }
        } elseif ($ExpectedSHA256) {
            if (-not [IO.File]::Exists($target) -or (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -ne $ExpectedSHA256) {
                throw 'JSON target changed since it was read; refusing a lost update.'
            }
        }
        $stream = [IO.File]::Open($temporary, 'CreateNew', 'Write', 'None')
        try { $stream.Write($Bytes, 0, $Bytes.Length); $stream.Flush($true) } finally { $stream.Dispose() }
        $expected = Get-SmartM365JsonHash $Bytes
        if ((Get-FileHash -LiteralPath $temporary -Algorithm SHA256).Hash -ne $expected) { throw 'Staged JSON integrity check failed.' }
        if ($legacyHash) { Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'ReplacementPrepared' $expected $legacyHash }
        $publishDeadline = [datetime]::UtcNow.AddSeconds($LockTimeoutSeconds)
        if ([IO.File]::Exists($target)) {
            $attributes = [IO.File]::GetAttributes($target)
            if ($attributes -band [IO.FileAttributes]::ReadOnly) {
                $restoreAttributes = $attributes
                [IO.File]::SetAttributes($target, ($attributes -band (-bnot [IO.FileAttributes]::ReadOnly)))
            }
        }
        while ($true) {
            try {
                if ([IO.File]::Exists($target)) {
                    # No non-atomic emulation when the filesystem cannot replace atomically.
                    [IO.File]::Replace($temporary, $target, $backup)
                } else { [IO.File]::Move($temporary, $target) }
                $published = $true
                break
            } catch [IO.IOException] {
                if ([datetime]::UtcNow -ge $publishDeadline -or -not [IO.File]::Exists($temporary)) { throw [IO.IOException]::new("Atomic replacement failed for '$target': $($_.Exception.Message)", $_.Exception) }
            } catch [UnauthorizedAccessException] {
                if ([datetime]::UtcNow -ge $publishDeadline -or -not [IO.File]::Exists($temporary)) { throw [IO.IOException]::new("Atomic replacement failed for '$target': $($_.Exception.Message)", $_.Exception) }
            }
            Start-Sleep -Milliseconds 100
        }
        if ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -ne $expected) { throw "Published JSON integrity failure; previous version retained at $backup" }
        if ($legacyHash) {
            Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'ReplacementPublished' $expected $legacyHash
            if ((Get-FileHash -LiteralPath $names.Legacy -Algorithm SHA256).Hash -ne $legacyHash) { throw 'Legacy JSON changed during replacement; both files preserved.' }
            [IO.File]::Delete($names.Legacy)
            Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'Completed' $expected 'Regenerated replacement verified before legacy retirement'
        }
        if ([IO.File]::Exists($backup)) { [IO.File]::Delete($backup) }
        [pscustomobject]@{ Path = $target; SHA256 = $expected; Bytes = $Bytes.Length }
    } finally {
        # Keep an unconsumed staging file on failure for recovery; never touch another owner's file.
        try {
            if (-not $published -and $null -ne $restoreAttributes -and [IO.File]::Exists($target)) { [IO.File]::SetAttributes($target, $restoreAttributes) }
        } finally { $lock.Dispose() }
    }
}

function Complete-SmartM365JsonConsumption {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path,[Parameter(Mandatory)][string]$Owner,[Parameter(Mandatory)][string]$ExpectedSHA256)
    $names = Get-SmartM365JsonNames $Path
    foreach ($candidate in @($names.Legacy,$names.Preferred,$names.Lock,$names.Journal)) { Assert-SmartM365JsonUnlinkedPath $candidate }
    $lock = Enter-SmartM365JsonTransportLock $names.Lock 10
    try {
        $current = Read-SmartM365JsonDocument $Path
        if ($current.SHA256 -ne $ExpectedSHA256) { throw 'JSON request changed before consumption; newer request preserved.' }
        if ($current.Coexistence -notin @('Single','Identical')) { throw 'Unexpected request coexistence; files preserved.' }
        Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'ConsumptionPrepared' $ExpectedSHA256 ''
        if ($current.Coexistence -eq 'Identical') { [IO.File]::Delete($names.Legacy) }
        [IO.File]::Delete($current.Path)
        Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'Consumed' $ExpectedSHA256 ''
    } finally { $lock.Dispose() }
}

function Get-SmartM365JsonTemplateName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    return (Get-SmartM365JsonNames $Path).Legacy + '.template'
}

function Publish-SmartM365RegeneratedJsonBytes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Owner,
        [Parameter(Mandatory)][byte[]]$Bytes,
        [Parameter(Mandatory)][scriptblock]$Validate
    )
    $names = Get-SmartM365JsonNames $Path
    $policy = Get-SmartM365JsonTransportPolicy
    $target = Get-SmartM365JsonReadPath $Path -Optional
    if ($policy.Mode -eq 'JsonText') {
        if ($names.Preferred.StartsWith('\\')) {
            $qualified = @($policy.QualifiedUncRoots | Where-Object { $names.Preferred.StartsWith(([IO.Path]::GetFullPath($_).TrimEnd('\') + '\'), [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0
            if (-not $qualified) { throw 'UNC replacement requires qualification on the target share before activation.' }
        }
        $target = $names.Preferred
    } elseif (-not $target) { $target = $names.Legacy }
    $retire = $policy.Mode -eq 'JsonText'
    Write-SmartM365JsonBytesAtomically -Path $target -Bytes $Bytes -Validate $Validate -RetireLegacyAfterPublication:$retire -Owner $Owner
}

function Resolve-SmartM365OwnedJsonPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Owner,
        [Parameter(Mandatory)][scriptblock]$Validate
    )
    $names = Get-SmartM365JsonNames $Path
    $policy = Get-SmartM365JsonTransportPolicy
    if ($policy.Mode -eq 'JsonText' -and $names.Legacy.StartsWith('\\')) {
        $qualifiedRoot = @($policy.QualifiedUncRoots | Where-Object { $names.Legacy.StartsWith(([IO.Path]::GetFullPath($_).TrimEnd('\') + '\'), [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0
        if (-not $qualifiedRoot) { throw 'UNC JSON activation requires qualification on the target share.' }
    }
    $existing = Get-SmartM365JsonReadPath $Path -Optional
    if ($existing) {
        $null = Read-SmartM365JsonDocument $Path -Validate $Validate
        if ($policy.Mode -eq 'JsonText') {
            $qualified = $false
            foreach ($root in @($policy.QualifiedUncRoots)) {
                if ($names.Legacy.StartsWith(([IO.Path]::GetFullPath($root).TrimEnd('\') + '\'), [StringComparison]::OrdinalIgnoreCase)) { $qualified = $true }
            }
            $result = Move-SmartM365OwnedJsonFile -Root ([IO.Path]::GetDirectoryName($names.Legacy)) -RelativePath ([IO.Path]::GetFileName($names.Legacy)) -Owner $Owner -Validate $Validate -QualifiedUnc:$qualified -RemoveIdenticalLegacy
            return $result.Path
        }
        return $existing
    }
    if ($policy.Mode -eq 'JsonText') { return $names.Preferred }
    return $names.Legacy
}

function Resolve-SmartM365WeeklyManifestPaths {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$HistoryRootPath,
        [Parameter(Mandatory)][string]$HistoryLabel,
        [scriptblock]$ValidateOwner
    )
    # The owner supplies its exact history root. Never recurse into arbitrary JSON.
    $rootFull = [IO.Path]::GetFullPath($HistoryRootPath).TrimEnd('\', '/')
    $plans = @()
    foreach ($folder in @(Get-ChildItem -LiteralPath $HistoryRootPath -Directory -ErrorAction Stop | Where-Object { $_.Name -match '^\d{4}-W(?:0[1-9]|[1-4]\d|5[0-3])$' })) {
        $legacy = Join-Path $folder.FullName 'manifest.json'
        if (-not (Get-SmartM365JsonReadPath $legacy -Optional)) { continue }
        $week = $folder.Name
        $validator = {
            param($document)
            if ($document.Week -ne $week) { throw 'Weekly manifest week mismatch.' }
            if ($ValidateOwner) { & $ValidateOwner $document | Out-Null }
            else {
                if ($document.HistoryLabel -ne $HistoryLabel) { throw 'Weekly manifest owner mismatch.' }
                if ([IO.Path]::GetFullPath([string]$document.HistoryRootPath).TrimEnd('\', '/') -ne $rootFull) { throw 'Weekly manifest history root mismatch.' }
            }
            if (-not $document.PSObject.Properties['Files']) { throw 'Weekly manifest Files is missing.' }
            foreach ($name in @($document.Files)) {
                if ([string]::IsNullOrWhiteSpace([string]$name) -or [IO.Path]::GetFileName([string]$name) -ne $name -or $name -notmatch '(?i)\.csv$') { throw 'Weekly manifest contains an invalid CSV filename.' }
            }
        }.GetNewClosure()
        $before = Read-SmartM365JsonDocument $legacy -Validate $validator
        $plans += [pscustomobject]@{ Legacy=$legacy; Before=$before.Path; Validate=$validator }
    }
    # Preflight all existing weeks before mutating any; divergent pairs fail closed.
    foreach ($plan in $plans) {
        $resolved = Resolve-SmartM365OwnedJsonPath -Path $plan.Legacy -Owner ('WeeklyHistory:' + $HistoryLabel) -Validate $plan.Validate
        [pscustomobject]@{ Path=$resolved; Renamed=($resolved -ne $plan.Before) }
    }
}

function Resolve-SmartM365JsonConfigurationPath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $validate = { param($document) if ($document -isnot [System.Management.Automation.PSCustomObject]) { throw 'Configuration must be a JSON object.' } }
    Resolve-SmartM365OwnedJsonPath -Path $Path -Owner ('Configuration:' + [IO.Path]::GetFileName($Path)) -Validate $validate
}

Export-ModuleMember -Function Get-SmartM365JsonNames, Get-SmartM365JsonReadPath, Read-SmartM365JsonDocument, Move-SmartM365OwnedJsonFile, Resolve-SmartM365JsonConfigurationPath, Get-SmartM365JsonTransportPolicy, Get-SmartM365JsonTemplateName, Resolve-SmartM365OwnedJsonPath, Write-SmartM365JsonBytesAtomically, Resolve-SmartM365WeeklyManifestPaths, Publish-SmartM365RegeneratedJsonBytes, Complete-SmartM365JsonConsumption

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAIdLnWp8NcKuhn
# XJ/+lDdkaQBIIzZ/It48RQXcCuV2OqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEII/2RoJgSxr3bbU6PzHFTmlrvO0DOUE54+nPmQVJGPWZMA0GCSqG
# SIb3DQEBAQUABIIBgJQ+lgCek/gE51PeHolPySNBh7OYRMJX3sFEFnoyhZLDud6p
# X7D2RPphwwUa4f7n/4VTeLZSzfdwtwh+Z1/AT13/TlaZwuKeY1pqu8zXRbeG8DMs
# o227nZBwSRDXZV7OoxgCa6s7QfurFh6K9YGYmHWbANyMqr8sNbJm3sBSmHTWnFFy
# vH15B1obAImWurZjhMeOgJWFEJtaOeeAKzv6iyqU3R2GMPELw8CS6/h1jgWPgaJd
# L5XgHRFDuOMuF6p8EBcePqGs52iz/b9WaZQK4n3eXlP4x+5TqAXu8avNA7XtigoK
# 8bg/nW3c8+4PPCNLgPbw2HQDnzE45GkNdSzb2YKvRiSxHjC+GpveeIn8SzJ7kdEy
# 5Ucb8ASva+riFKNCBltFeQ31UlGCs0bvnwpE7IeZ2x9KIDzDGI7Fj6FrpyhSlM4A
# P5LsTHS/9k2xQ2VmsMLFbxxYukBrZ2XbplUEW3u6KH57EEDoZED8SEVlRt1kOFQS
# Qr2fO1sZIbZYtETxdqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjcxNjU3
# MzhaMC8GCSqGSIb3DQEJBDEiBCCN5Ur8tWppC2unmeblFXdoPlLzvbCPMeLtRiwz
# OLP1FTANBgkqhkiG9w0BAQEFAASCAgAqMGqOjO4I0v+4+Miqe7qtbIDx5lJPVLY8
# C5DP+b4BvF5bzoVpx6E7J58x2kUM0xnwp32BHPgVvh+5O0RNWfhJYPDGAqUZUoOA
# ACXRD23eUf54XegwyQyQfISGdgWSGRHYxrmyGxqhA4oG/vD77UP0DS3wTXGxbjI4
# RHuThmSOGLmYhGgi8CEClt3f8I28iFylEN8tkCHYGS8lT9PGQyFBDj8Cymb5aPP2
# 7uFOyDaeqG+XPIcKVeCcvi12A5iR+Czl3DYcrn1xpvPhjdV7OrID7vVH7YzmNRDo
# gewOZK+AYctKsL6g25psmwHSYa6ZW6aoAHW5ierp+tejZb3G3S1hFaK4EK3TveoR
# zjtl91Zj8fM+rNbQTKW94wttLz6B5zdZRXdTKiwve6lDlyp1fCzSbQy7quSn5KCL
# lCVnlhVtAzY5J+iwYeM9eUfkZEgtkyvbqvuFdBZPgrTJtByzW76oyYmElNjmu82B
# wO/TsKi/kC1+t37iUGmTtVfs4LsGn04H0JglQXHJsQpz3Mwg+A/g8BUf8Jh+KIz9
# Kvw29MNWBkP56Hc4bNN8Hob4DMQrKdVM/5z7pMrRigMvOpcx0WI66t2oo/dJgDWg
# yLnHxlBoBpo7n6OY7DOrU2GVXAjYarq0v0K+N6Ja+AdupCqfaKDUG+otZAYw4/Lz
# h+c7JlOcIQ==
# SIG # End signature block
