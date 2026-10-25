<#
.SYNOPSIS
    Wave computer list helpers.

.DESCRIPTION
    Parses Computers.txt (quotes trimmed, comments ignored, case-insensitive de-duplication),
    computes short keys and moves completed computers to archive lists under an exclusive
    file lock shared by every process and session of the technician computer.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

function ConvertTo-SemtComputerName {
    [CmdletBinding()]
    param([AllowNull()][string]$Line)

    $name = ([string]$Line).Trim().Trim([char]34).Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($name) -or $name.StartsWith('#')) { return '' }
    return $name
}

function Get-SemtComputerKey {
    [CmdletBinding()]
    param([AllowNull()][string]$ComputerName)

    $name = ConvertTo-SemtComputerName -Line $ComputerName
    if ([string]::IsNullOrWhiteSpace($name)) { return '' }
    return ($name.Split('.')[0]).ToUpperInvariant()
}

function Get-SemtComputerList {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Computer list not found: $Path" }
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $result = New-Object System.Collections.Generic.List[string]
    foreach ($line in @(Get-Content -LiteralPath $Path -ErrorAction Stop)) {
        $name = ConvertTo-SemtComputerName -Line $line
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ($seen.Add($name)) { $result.Add($name) }
    }
    return @($result.ToArray())
}

function Get-SemtComputerListStats {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Computer list not found: $Path" }
    $names = @(Get-Content -LiteralPath $Path -ErrorAction Stop | ForEach-Object { ConvertTo-SemtComputerName -Line $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $groups = @($names | Group-Object { $_.ToUpperInvariant() } | Where-Object { $_.Count -gt 1 } | Sort-Object -Property @{ Expression = 'Count'; Descending = $true },Name)
    $duplicateLines = 0
    foreach ($group in $groups) { $duplicateLines += ([int]$group.Count - 1) }
    return [pscustomobject]@{
        RawLines = [int]$names.Count
        Unique = [int](@($names | ForEach-Object { $_.ToUpperInvariant() } | Select-Object -Unique).Count)
        DuplicateGroups = [int]$groups.Count
        DuplicateLines = [int]$duplicateLines
        DuplicateSamples = [string](($groups | Select-Object -First 10 | ForEach-Object { $_.Group[0] }) -join ', ')
    }
}

function Test-SemtComputerListPresent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [ValidateRange(1, 20)][int]$Attempts = 5,
        [ValidateRange(0, 5000)][int]$DelayMilliseconds = 200
    )

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        if (Test-Path -LiteralPath $Path -PathType Leaf) { return $true }
        if ($attempt -lt $Attempts -and $DelayMilliseconds -gt 0) { Start-Sleep -Milliseconds $DelayMilliseconds }
    }
    return $false
}

function Get-SemtComputerListLockPath {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$ComputerListPath)

    $key = (Get-SemtStringSha256 -Text ([System.IO.Path]::GetFullPath($ComputerListPath).ToUpperInvariant())).Substring(0, 32)
    return (Join-Path -Path (Get-SemtLauncherStateRoot) -ChildPath ("ComputerListLocks\{0}.lock" -f $key))
}

function Invoke-SemtWithComputerListLock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ComputerListPath,
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock,
        [object[]]$ArgumentList = @()
    )

    return (Invoke-SemtWithFileLock -LockPath (Get-SemtComputerListLockPath -ComputerListPath $ComputerListPath) -TimeoutSeconds 60 -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList)
}

function Write-SemtTextLinesAtomic {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Lines
    )

    $temp = '{0}.tmp.{1}.txt' -f $Path,([guid]::NewGuid().ToString('N'))
    try {
        [System.IO.File]::WriteAllLines($temp, [string[]]$Lines, (New-Object System.Text.UTF8Encoding($false)))
        Move-Item -LiteralPath $temp -Destination $Path -Force -ErrorAction Stop
    }
    finally {
        if (Test-Path -LiteralPath $temp -PathType Leaf) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
    }
}

function Remove-SemtComputerListDuplicates {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    return (Invoke-SemtWithComputerListLock -ComputerListPath $Path -ArgumentList @($Path) -ScriptBlock {
        param($ListPath)
        $lines = @(Get-Content -LiteralPath $ListPath -ErrorAction Stop)
        $seen = @{}
        $output = New-Object System.Collections.Generic.List[string]
        $duplicates = 0
        $samples = New-Object System.Collections.Generic.List[string]
        foreach ($line in $lines) {
            $name = ConvertTo-SemtComputerName -Line $line
            if ([string]::IsNullOrWhiteSpace($name)) { $output.Add([string]$line); continue }
            $key = $name.ToUpperInvariant()
            if ($seen.ContainsKey($key)) {
                $duplicates++
                if ($samples.Count -lt 10 -and -not $samples.Contains($seen[$key])) { $samples.Add($seen[$key]) }
                continue
            }
            $seen[$key] = $name
            $output.Add([string]$line)
        }
        if ($duplicates -eq 0) { return [pscustomobject]@{ Changed = $false; DuplicateLines = 0; DuplicateSamples = ''; BackupPath = '' } }
        $backup = '{0}.dedup_{1}.bak' -f $ListPath,(Get-Date -Format 'yyyyMMdd-HHmmss')
        Copy-Item -LiteralPath $ListPath -Destination $backup -Force -ErrorAction Stop
        Write-SemtTextLinesAtomic -Path $ListPath -Lines @($output.ToArray())
        return [pscustomobject]@{ Changed = $true; DuplicateLines = $duplicates; DuplicateSamples = ($samples -join ', '); BackupPath = $backup }
    })
}

function Move-SemtComputersFromList {
    <#
    Moves the named computers from Computers.txt to an archive list in the same Wave folder.
    Both files are rewritten atomically under the shared computer-list lock.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ComputerListPath,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$ComputerNames,
        [Parameter(Mandatory = $true)][string]$ArchiveFileName,
        [Parameter(Mandatory = $true)][string]$Label
    )

    $moveKeys = @{}
    foreach ($computer in @($ComputerNames)) {
        $key = Get-SemtComputerKey -ComputerName $computer
        if (-not [string]::IsNullOrWhiteSpace($key) -and -not $moveKeys.ContainsKey($key)) { $moveKeys[$key] = $true }
    }
    if ($moveKeys.Count -eq 0) { return [pscustomobject]@{ Moved = 0; ArchivePath = ''; Detail = "No $Label computer to move." } }

    return (Invoke-SemtWithComputerListLock -ComputerListPath $ComputerListPath -ArgumentList @($ComputerListPath,$moveKeys,$ArchiveFileName,$Label) -ScriptBlock {
        param($ListPath, $Keys, $ArchiveName, $MoveLabel)
        if (-not (Test-Path -LiteralPath $ListPath -PathType Leaf)) { return [pscustomobject]@{ Moved = 0; ArchivePath = ''; Detail = 'Computers.txt is missing.' } }
        $remaining = New-Object System.Collections.Generic.List[string]
        $moved = New-Object System.Collections.Generic.List[string]
        foreach ($line in @(Get-Content -LiteralPath $ListPath -ErrorAction Stop)) {
            $name = ConvertTo-SemtComputerName -Line $line
            if (-not [string]::IsNullOrWhiteSpace($name) -and $Keys.ContainsKey((Get-SemtComputerKey -ComputerName $name))) { $moved.Add($name); continue }
            $remaining.Add([string]$line)
        }
        if ($moved.Count -eq 0) { return [pscustomobject]@{ Moved = 0; ArchivePath = ''; Detail = "$MoveLabel computers were no longer present in Computers.txt." } }
        $archivePath = Join-Path -Path (Split-Path -Path $ListPath -Parent) -ChildPath $ArchiveName
        $archiveLines = New-Object System.Collections.Generic.List[string]
        $existing = @{}
        if (Test-Path -LiteralPath $archivePath -PathType Leaf) {
            foreach ($line in @(Get-Content -LiteralPath $archivePath -ErrorAction Stop)) {
                $archiveLines.Add([string]$line)
                $existingKey = Get-SemtComputerKey -ComputerName $line
                if (-not [string]::IsNullOrWhiteSpace($existingKey)) { $existing[$existingKey] = $true }
            }
        }
        foreach ($name in $moved) {
            $key = Get-SemtComputerKey -ComputerName $name
            if (-not $existing.ContainsKey($key)) { $archiveLines.Add($name); $existing[$key] = $true }
        }
        Write-SemtTextLinesAtomic -Path $archivePath -Lines @($archiveLines.ToArray())
        Write-SemtTextLinesAtomic -Path $ListPath -Lines @($remaining.ToArray())
        return [pscustomobject]@{ Moved = $moved.Count; ArchivePath = $archivePath; Detail = ('Moved {0} {1} computer(s) to {2}.' -f $moved.Count,$MoveLabel,$ArchiveName) }
    })
}
