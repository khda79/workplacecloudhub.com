<#
.SYNOPSIS
    Compares two SharePoint file inventory scans for the same endpoint.

.DESCRIPTION
    By default, this script selects the latest file inventory CSV for the selected endpoint, then
    selects the previous file inventory CSV with the same inventory name prefix. The older scan
    is used as the baseline and the newer scan is used as the current inventory.

    RemovedFromNewScan means: present in the older scan, missing from the newer scan.

    AddedInNewScan means: absent from the older scan, present in the newer scan.

.VERSION
    1.0.0

.EXAMPLE
    .\SmartM365-SharePointSource-FileInventoryHistoryCompare.ps1

.EXAMPLE
    .\SmartM365-SharePointSource-FileInventoryHistoryCompare.ps1 -ScanDirectory ".\Migrations\MyMigration\scans\source\files"

.EXAMPLE
    .\SmartM365-SharePointSource-FileInventoryHistoryCompare.ps1 -OldCsv "C:\Scans\old.csv" -NewCsv "C:\Scans\new.csv"
#>

[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [string]$OldCsv,

    [ValidateNotNullOrEmpty()]
    [string]$NewCsv,

    [ValidateNotNullOrEmpty()]
    [string]$ScanDirectory,

    [ValidateNotNullOrEmpty()]
    [string]$InventoryNameFilter,

    [ValidateSet('Source', 'Target')]
    [string]$Side = 'Source',

    [ValidateSet('SP2016', 'SP2019', 'SPO')]
    [string]$EndpointType = 'SP2019',

    [ValidateNotNullOrEmpty()]
    [string]$WebUrlsFile,

    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory,

    [ValidateNotNullOrEmpty()]
    [string]$LogPath,

    [ValidateNotNullOrEmpty()]
    [string]$ComparisonName = 'InventoryHistory',

    [long]$SizeToleranceBytes = 10240,

    [ValidateRange(0, [double]::MaxValue)]
    [double]$ModifiedDateToleranceMinutes = 0,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

function Get-ConsoleTimestamp {
    return (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
}

function Write-Host {
    param(
        [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
        [object[]]$Object,

        [ConsoleColor]$ForegroundColor,

        [switch]$NoNewline
    )

    $message = if ($Object) { ($Object -join ' ') } else { '' }
    $line = if ($message -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} ') { $message } else { "{0} {1}" -f (Get-ConsoleTimestamp), $message }
    $parameters = @{ Object = $line }
    if ($PSBoundParameters.ContainsKey('ForegroundColor')) { $parameters.ForegroundColor = $ForegroundColor }
    if ($NoNewline) { $parameters.NoNewline = $true }
    Microsoft.PowerShell.Utility\Write-Host @parameters
}

function Add-TimestampToLogFile {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return
    }

    $timestampPattern = '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} '
    $content = Get-Content -LiteralPath $Path
    $updated = foreach ($line in $content) {
        if ($line -match $timestampPattern -or [string]::IsNullOrWhiteSpace($line)) {
            $line
        }
        else {
            "{0} {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $line
        }
    }

    Set-Content -LiteralPath $Path -Value $updated -Encoding UTF8 -WhatIf:$false
}

function Stop-TimestampedTranscript {
    param([string]$Path)

    try {
        Stop-Transcript | Out-Null
    }
    catch {
        return
    }

    Add-TimestampToLogFile -Path $Path
}

function Write-Warning {
    param(
        [Parameter(Position = 0)]
        [string]$Message
    )

    $line = if ($Message -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} ') { $Message } else { "{0} WARNING: {1}" -f (Get-ConsoleTimestamp), $Message }
    Microsoft.PowerShell.Utility\Write-Host $line -ForegroundColor Yellow
}

function Read-Host {
    param(
        [Parameter(Position = 0)]
        [string]$Prompt
    )

    Microsoft.PowerShell.Utility\Read-Host ("{0} {1}" -f (Get-ConsoleTimestamp), $Prompt)
}

function Get-ProjectRoot {
    if ($PSScriptRoot) {
        return (Resolve-Path -LiteralPath (Join-Path -Path $PSScriptRoot -ChildPath '..\..')).ProviderPath
    }

    return (Get-Location).Path
}

function Get-PythonCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot
    )

    $portablePython = Join-Path -Path $ProjectRoot -ChildPath 'Tools\Python\python.exe'
    if (Test-Path -LiteralPath $portablePython) {
        & $portablePython --version *> $null
        if ($LASTEXITCODE -eq 0) {
            return [pscustomobject]@{
                Executable = $portablePython
                Arguments  = @()
                Source     = 'Portable'
            }
        }
    }

    $python = Get-Command -Name python -ErrorAction SilentlyContinue
    if ($python) {
        & $python.Source --version *> $null
        if ($LASTEXITCODE -eq 0) {
            return [pscustomobject]@{
                Executable = $python.Source
                Arguments  = @()
                Source     = 'PATH'
            }
        }
    }

    $pyLauncher = Get-Command -Name py -ErrorAction SilentlyContinue
    if ($pyLauncher) {
        & $pyLauncher.Source -3 --version *> $null
        if ($LASTEXITCODE -eq 0) {
            return [pscustomobject]@{
                Executable = $pyLauncher.Source
                Arguments  = @('-3')
                Source     = 'PyLauncher'
            }
        }
    }

    throw "Python 3 is required. Expected portable Python at: $portablePython"
}

function Get-InventoryNamePrefix {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$CsvItem
    )

    if ($CsvItem.Name -match '^((?:SP2016|SP2019|SPO)-FileInventory-.+)-\d{8}-\d{6}\.csv$') {
        return $Matches[1]
    }

    return [System.IO.Path]::GetFileNameWithoutExtension($CsvItem.Name)
}

function Get-InventoryCsvs {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory,

        [string]$NameFilter,

        [ValidateSet('SP2016', 'SP2019', 'SPO')]
        [string]$EndpointType = 'SP2019'
    )

    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) {
        throw "$EndpointType scan directory not found: $Directory"
    }

    $items = Get-ChildItem -LiteralPath $Directory -Recurse -File -Filter ("{0}-FileInventory*.csv" -f $EndpointType) |
        Where-Object {
            $_.Name -notlike '*-Errors.csv' -and
            $_.Name -notlike '*.tmp' -and
            $_.Name -notlike '*PermissionInventory*'
        }

    if (-not [string]::IsNullOrWhiteSpace($NameFilter)) {
        $items = @($items | Where-Object { $_.Name -like $NameFilter -or $_.FullName -like $NameFilter })
    }

    return @($items | Sort-Object LastWriteTime -Descending)
}

function Resolve-DefaultCsvPair {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory,

        [string]$NameFilter,

        [ValidateSet('SP2016', 'SP2019', 'SPO')]
        [string]$EndpointType = 'SP2019'
    )

    $csvs = @(Get-InventoryCsvs -Directory $Directory -NameFilter $NameFilter -EndpointType $EndpointType)
    if ($csvs.Count -lt 2) {
        throw "At least two $EndpointType file inventory CSV files are required in $Directory."
    }

    $latest = $csvs[0]
    $latestPrefix = Get-InventoryNamePrefix -CsvItem $latest
    $previous = @($csvs | Where-Object {
            $_.FullName -ne $latest.FullName -and (Get-InventoryNamePrefix -CsvItem $_) -eq $latestPrefix
        } | Select-Object -First 1)

    if (-not $previous) {
        throw "No previous $EndpointType file inventory CSV found with the same prefix as '$($latest.Name)'. Use -OldCsv and -NewCsv explicitly, or use -InventoryNameFilter."
    }

    return [pscustomobject]@{
        Old = $previous[0]
        New = $latest
    }
}

function Confirm-InventoryHistoryComparison {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$OldItem,

        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$NewItem,

        [TimeSpan]$AgeDifference
    )

    Write-Host ""
    Write-Host ("{0} {1} scan history CSV selection" -f $Side, $EndpointType) -ForegroundColor Cyan
    Write-Host ("  Previous CSV      : {0}" -f $OldItem.FullName)
    Write-Host ("  Old timestamp     : {0}" -f $OldItem.LastWriteTime)
    Write-Host ("  Current CSV       : {0}" -f $NewItem.FullName)
    Write-Host ("  New timestamp     : {0}" -f $NewItem.LastWriteTime)
    Write-Host ("  Difference        : {0:n2}h" -f $AgeDifference.TotalHours)
    Write-Host ""

    $confirmation = Read-Host ("Type COMPARE to compare these {0} {1} scans" -f $Side, $EndpointType)
    if ($confirmation -ne 'COMPARE') {
        throw "Comparison cancelled. Confirmation text did not match COMPARE."
    }
}

$ProjectRoot = Get-ProjectRoot

if ([string]::IsNullOrWhiteSpace($ScanDirectory)) {
    $ScanDirectory = Join-Path -Path $ProjectRoot -ChildPath ("Scans\{0}" -f $EndpointType)
}

if ([string]::IsNullOrWhiteSpace($OldCsv) -and [string]::IsNullOrWhiteSpace($NewCsv)) {
    $pair = Resolve-DefaultCsvPair -Directory $ScanDirectory -NameFilter $InventoryNameFilter -EndpointType $EndpointType
    $OldCsv = $pair.Old.FullName
    $NewCsv = $pair.New.FullName
}
elseif ([string]::IsNullOrWhiteSpace($OldCsv) -or [string]::IsNullOrWhiteSpace($NewCsv)) {
    throw "Use both -OldCsv and -NewCsv, or omit both to compare the two latest matching endpoint scans."
}

if (-not (Test-Path -LiteralPath $OldCsv -PathType Leaf)) {
    throw "Old CSV not found: $OldCsv"
}

if (-not (Test-Path -LiteralPath $NewCsv -PathType Leaf)) {
    throw "New CSV not found: $NewCsv"
}

if (-not [string]::IsNullOrWhiteSpace($WebUrlsFile) -and -not (Test-Path -LiteralPath $WebUrlsFile -PathType Leaf)) {
    throw "Web URLs file not found: $WebUrlsFile"
}

$oldItem = Get-Item -LiteralPath $OldCsv
$newItem = Get-Item -LiteralPath $NewCsv

if ($oldItem.LastWriteTime -gt $newItem.LastWriteTime) {
    Write-Warning "OldCsv is newer than NewCsv. Swapping them so source=older and target=newer."
    $temporaryItem = $oldItem
    $oldItem = $newItem
    $newItem = $temporaryItem
    $OldCsv = $oldItem.FullName
    $NewCsv = $newItem.FullName
}

$ageDifference = ($newItem.LastWriteTime - $oldItem.LastWriteTime).Duration()
if (-not $Force) {
    Confirm-InventoryHistoryComparison -OldItem $oldItem -NewItem $newItem -AgeDifference $ageDifference
}

$comparisonTimestampText = Get-Date -Format 'yyyyMMdd-HHmmss'
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path -Path (Join-Path -Path $ProjectRoot -ChildPath 'ComparisonResults\ScanHistory') -ChildPath ("{0}-{1}-Changes-{2}" -f $Side, $EndpointType, $comparisonTimestampText)
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path -Path (Join-Path -Path $OutputDirectory -ChildPath 'logs') -ChildPath ("{0}-{1}-Run.log" -f $ComparisonName, $comparisonTimestampText)
}

$logDirectory = Split-Path -Path $LogPath -Parent
if ($logDirectory -and -not (Test-Path -LiteralPath $logDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
}

$script:TranscriptStarted = $false
trap {
    if ($script:TranscriptStarted) {
        Stop-TimestampedTranscript -Path $LogPath
    }
    throw $_
}

Start-Transcript -Path $LogPath -Force -WhatIf:$false | Out-Null
$script:TranscriptStarted = $true
Write-Host ("Run log: {0}" -f $LogPath) -ForegroundColor Cyan

$pythonScript = Join-Path -Path $ProjectRoot -ChildPath 'Scripts\Compare\compare_sp_source_target_file_inventories.py'
$comparisonExcelScript = Join-Path -Path $ProjectRoot -ChildPath 'Scripts\Export\export_comparison_to_excel.py'
$duplicateKeysExcelScript = Join-Path -Path $ProjectRoot -ChildPath 'Scripts\Export\export_duplicate_keys_to_excel.py'

foreach ($requiredScript in @($pythonScript, $comparisonExcelScript, $duplicateKeysExcelScript)) {
    if (-not (Test-Path -LiteralPath $requiredScript -PathType Leaf)) {
        throw "Required script not found: $requiredScript"
    }
}

$pythonCommand = Get-PythonCommand -ProjectRoot $ProjectRoot
$pythonExecutable = $pythonCommand.Executable
$pythonArguments = @($pythonCommand.Arguments)

Write-Host ("Scan age difference: {0:n2}h" -f $ageDifference.TotalHours)
Write-Host ("Using Python: {0} ({1})" -f $pythonExecutable, $pythonCommand.Source) -ForegroundColor Cyan
Write-Host ("Previous CSV: {0}" -f $OldCsv)
Write-Host ("Current CSV: {0}" -f $NewCsv)
Write-Host ("Output directory: {0}" -f $OutputDirectory)
Write-Host ("Modified date tolerance: {0:n2} minute(s)" -f $ModifiedDateToleranceMinutes)

function Invoke-LoggedPythonCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,

        [Parameter(Mandatory = $true)]
        [string]$FailureMessage
    )

    & $pythonExecutable @pythonArguments @Arguments 2>&1 | ForEach-Object {
        Write-Host ([string]$_)
    }

    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        throw ("{0} failed with exit code {1}." -f $FailureMessage, $exitCode)
    }
}

$compareArguments = @(
    $pythonScript,
    '--source-csv', $OldCsv,
    '--target-csv', $NewCsv,
    '--output-directory', $OutputDirectory,
    '--comparison-name', $ComparisonName,
    '--size-tolerance-bytes', $SizeToleranceBytes,
    '--modified-date-tolerance-minutes', $ModifiedDateToleranceMinutes
)

if (-not [string]::IsNullOrWhiteSpace($WebUrlsFile)) {
    $compareArguments += @('--source-web-urls-file', $WebUrlsFile, '--target-web-urls-file', $WebUrlsFile)
}

Invoke-LoggedPythonCommand -Arguments $compareArguments -FailureMessage 'Python comparison'

$comparisonXlsx = Join-Path -Path $OutputDirectory -ChildPath ("{0}-{1}-Changes-{2}.xlsx" -f $Side, $EndpointType, $comparisonTimestampText)
$comparisonExcelArguments = @(
    $comparisonExcelScript,
    '--comparison-directory', $OutputDirectory,
    '--output-xlsx', $comparisonXlsx,
    '--sheet-profile', 'sp2019-changes'
)
Invoke-LoggedPythonCommand -Arguments $comparisonExcelArguments -FailureMessage 'Excel export'

$duplicateKeysXlsx = Join-Path -Path $OutputDirectory -ChildPath ("{0}-{1}-DuplicateKeys-{2}.xlsx" -f $Side, $EndpointType, $comparisonTimestampText)
$duplicateKeysExcelArguments = @(
    $duplicateKeysExcelScript,
    '--comparison-directory', $OutputDirectory,
    '--output-xlsx', $duplicateKeysXlsx
)
Invoke-LoggedPythonCommand -Arguments $duplicateKeysExcelArguments -FailureMessage 'Duplicate keys Excel export'

Write-Host ("Scan history comparison completed: {0}" -f $comparisonXlsx) -ForegroundColor Green

if ($script:TranscriptStarted) {
    Stop-TimestampedTranscript -Path $LogPath
    $script:TranscriptStarted = $false
}









