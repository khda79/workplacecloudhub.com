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










# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCCcY1SY4wa8pn7
# KVj5fPRsuFjZZm3KEWhw/R0HGJyNCqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIIsye+fySolrxLsHRhEl+H2qmEC3v38xWcldW7B2bNKlMA0GCSqG
# SIb3DQEBAQUABIIBgEVvJ0sEkhgTiFNV130kzuDt6cSd6ztt0nVBGI74ajE5ELJK
# EEz2xPqe9KDN/CrBaDd82jYbbN41O68fbtAr80MsQeFwW0c4zP5N0M6mipUmIm/j
# r+QQ2UcS7NAx8dEfSp0ajfpzsmkFV4koZ0QdfgvcMc26otOJM2wNzvY293nSrIeY
# FPwNgE7bBu60EUA4dciUrECiUrHJnuhipTNkvFyTF7SnmArZqGunzz13Wd9wcvn3
# wbrSarmr55dQ4FhAl8rhUZg1T7t5l8O0cWDX+0EYjatWtBQ0cd7AF8V4qlsljzSy
# 7KG3nKEmtus8qjf7IiA5NckwSywzZyRuqb7OXiy9jI9Jl3UGSlT7eMO+vkpNz3BL
# CpnU6UZjcPyC2YkiAzXBajFqXfCAU5+lNHslC1Nh0WAavSpexGNpvEHcJ2fVHp2Q
# YLK3R8mOfC2G+/aSVFc+0s6ETsuDajR7jXzOrsc1Elg0+1SEQQNqRvGgrUiXbJFv
# 78znNcUjoUkamN2S9aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODUw
# MTJaMC8GCSqGSIb3DQEJBDEiBCD9bL33cKLIuAYZiINjrfnLrENrCQ14o7ScWXQn
# pYYFWTANBgkqhkiG9w0BAQEFAASCAgCdw53f8s7GniYbLj0JBEGF6VIvEjm+KIUN
# DAsI/KOPbSxUMOr/JvgvCHz/AYdCleFmXf8F1gjffdIFlPXWVtBXoTaMSmp7Xaoy
# UwO/kAZQ5jyXj65BtwMgaW7o48BpCpmtvJIuwvFRn1nKlML346QJJuyk0qS/NGvT
# c63SGa8lAE2VMuUqctGVvgTVvlnlAhkPeUsLNss+9dkbA1JVqGWw4/QCHssYyPNV
# /98byCpEzQAtVevb65ES4Fzs8e8ZzRnKr8LVetwPzzlzJwqlRso3d4naHvfRhpgs
# +R9XjE7r0iPG2DpSICRHI0o9KVUVFl+d87nF4h05Qlc5FA7KBrweQE1jhNYvv5SV
# MmoudocY4AheZZ5qTQf5sQBhXXgnmAro1cuVXYnZHLq1XWKksZhID81JbsGbKlTj
# RJjHHdnx+FxugFNFTkF8GaQJ9G2/+lPFM8RrhlmPcCMCu3vVyWN5ILOwEUxw212U
# wHQMBdmBUWweqQavKYhvGnla01j6kuySLj3NF1vKIfRrxh9qwdQ5LdM42BxZaNCj
# RWdCPE8Q7YamRkX87HKfS+v3UbUg6wXwKpgP5xddqnU8Zc0IFQzAyL9ToeFI3t3m
# cjdGLT9EgR+oSpyZsScjab76h6GXsvxP5sSzq5vlhL+VnIpLRc7c3fJBO/wl/d6F
# nV7craRnJg==
# SIG # End signature block
