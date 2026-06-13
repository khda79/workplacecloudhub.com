<#
    Name: SmartM365-SetupDiag-Analyze-Upgrade-Failures-Remediation.ps1
    Version: 1.0
    Description: Runs SetupDiag to analyze Windows setup or upgrade failures and writes enterprise-readable registry and JSON summaries.

    Intended use:
    - Microsoft Intune remediation script
    - Windows 10 / Windows 11
    - Run as 64-bit PowerShell
    - No forced reboot

    Capabilities:
    - Automatically downloads SetupDiag.exe from Microsoft if not found locally
    - Keeps local fallback discovery from Windows Setup folders
    - Forces TLS 1.2 for download compatibility
    - Captures SetupDiag exit code
    - Writes JSON and registry summary
    - Supports offline scan when logs exist, otherwise online scan

    Outputs:
    - SetupDiag XML: C:\ProgramData\SmartM365\IntuneRemediation\Output\SetupDiag\SetupDiagResults.xml
    - Remediation log: C:\ProgramData\SmartM365\IntuneRemediation\Logs\SetupDiag\Remediation.log
    - JSON summary: C:\ProgramData\SmartM365\IntuneRemediation\Output\SetupDiag\Summary.json
    - Registry summary: HKLM:\SOFTWARE\SmartM365\IntuneRemediation\SetupDiag
#>

[CmdletBinding()]
param(
    [bool]$ZipLogs = $false,
    [string]$OutputRoot = (Join-Path -Path $env:ProgramData -ChildPath 'SmartM365\IntuneRemediation\Output\SetupDiag')
)

$ErrorActionPreference = "Stop"

# Relaunch in 64-bit PowerShell if Intune starts the script in 32-bit PowerShell
if ($env:PROCESSOR_ARCHITEW6432) {
    & "$env:WINDIR\Sysnative\WindowsPowerShell\v1.0\powershell.exe" `
        -ExecutionPolicy Bypass `
        -NoProfile `
        -File "$PSCommandPath" `
        -ZipLogs:$ZipLogs `
        -OutputRoot "$OutputRoot"

    exit $LASTEXITCODE
}

$targetDirectory = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Tools\SetupDiag"
$targetExecutable = Join-Path -Path $targetDirectory -ChildPath "SetupDiag.exe"
$setupDiagDownloadUrl = "https://go.microsoft.com/fwlink/?linkid=870142"

# Use single-quoted child paths containing "$Windows.~BT" to prevent PowerShell variable expansion
$setupDiagCandidates = @(
    (Join-Path -Path $env:ProgramData -ChildPath 'SetupDiag\SetupDiag.exe'),
    (Join-Path -Path $env:SystemDrive -ChildPath '$Windows.~BT\Sources\SetupDiag.exe'),
    (Join-Path -Path $env:SystemDrive -ChildPath 'Windows.old\$Windows.~BT\Sources\SetupDiag.exe')
)

$enterpriseDirectory = "C:\ProgramData\SmartM365\IntuneRemediation\Output\SetupDiag"
$logDirectory = "C:\ProgramData\SmartM365\IntuneRemediation\Logs\SetupDiag"
$logPath = Join-Path -Path $logDirectory -ChildPath "Remediation.log"
$jsonPath = Join-Path -Path $enterpriseDirectory -ChildPath "Summary.json"
$enterpriseRegistryPath = "HKLM:\SOFTWARE\SmartM365\IntuneRemediation\SetupDiag"
$setupDiagResultPath = Join-Path -Path $OutputRoot -ChildPath "SetupDiagResults.xml"

function Write-SmartM365Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "[{0}] {1}" -f (Get-Date).ToString("s"), $Message
    Add-Content -Path $logPath -Value $line -Encoding UTF8
    Write-Output $Message
}

function ConvertTo-RegistryValue {
    param(
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return ""
    }

    return [string]$Value
}

function Get-SetupDiagErrorHint {
    param(
        [AllowNull()]
        [string]$ErrorCode
    )

    if ([string]::IsNullOrWhiteSpace($ErrorCode)) {
        return $null
    }

    switch -Regex ($ErrorCode) {
        "^0xC1900208$" { return "Compatibility blocker" }
        "^0xC1900101"  { return "Driver or OEM rollback issue" }
        "^0xC1900200$" { return "Hardware requirements not met" }
        "^0x80070070$" { return "Insufficient disk space" }
        "^0x800F0922$" { return "Servicing stack, reserved partition, or connectivity issue" }
        "^0x8007042B$" { return "Migration phase failure, often caused by driver, service, or application conflict" }
        "^0x8007001F$" { return "Device, driver, or general I/O failure" }
        default        { return $null }
    }
}

function Save-SetupDiag {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    try {
        Write-SmartM365Log "Attempting to download SetupDiag.exe from Microsoft."

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $temporaryDownloadPath = "$DestinationPath.download"

        if (Test-Path -LiteralPath $temporaryDownloadPath) {
            Remove-Item -LiteralPath $temporaryDownloadPath -Force -ErrorAction SilentlyContinue
        }

        Invoke-WebRequest -Uri $setupDiagDownloadUrl -OutFile $temporaryDownloadPath -UseBasicParsing -TimeoutSec 120 -ErrorAction Stop

        if (-not (Test-Path -LiteralPath $temporaryDownloadPath -PathType Leaf)) {
            Write-SmartM365Log "SetupDiag download failed because the temporary file was not created."
            return $false
        }

        $downloadedFile = Get-Item -LiteralPath $temporaryDownloadPath -ErrorAction Stop

        if ($downloadedFile.Length -lt 100KB) {
            Write-SmartM365Log "SetupDiag download failed because the downloaded file is unexpectedly small. SizeBytes=$($downloadedFile.Length)"
            Remove-Item -LiteralPath $temporaryDownloadPath -Force -ErrorAction SilentlyContinue
            return $false
        }

        Move-Item -LiteralPath $temporaryDownloadPath -Destination $DestinationPath -Force
        Write-SmartM365Log "SetupDiag.exe downloaded successfully. SizeBytes=$($downloadedFile.Length)"
        return $true
    }
    catch {
        Write-SmartM365Log "SetupDiag download failed: $($_.Exception.Message)"
        return $false
    }
}

try {
    if (-not (Test-Path -Path $enterpriseDirectory)) {
        New-Item -ItemType Directory -Path $enterpriseDirectory -Force | Out-Null
    }
    if (-not (Test-Path -Path $logDirectory)) {
        New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
    }

    Write-SmartM365Log "===== SetupDiag remediation started ====="

    if (-not (Test-Path -Path $targetDirectory)) {
        New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
        Write-SmartM365Log "Created SetupDiag target directory: $targetDirectory"
    }

    if (-not (Test-Path -LiteralPath $targetExecutable -PathType Leaf)) {
        Write-SmartM365Log "SetupDiag.exe is missing from target directory. Probing local Windows Setup folders first."

        $setupDiagAvailable = $false

        foreach ($candidate in $setupDiagCandidates) {
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                Copy-Item -LiteralPath $candidate -Destination $targetExecutable -Force
                Write-SmartM365Log "Copied SetupDiag.exe from local candidate: $candidate"
                $setupDiagAvailable = $true
                break
            }
        }

        if (-not $setupDiagAvailable) {
            $setupDiagAvailable = Save-SetupDiag -DestinationPath $targetExecutable
        }

        if (-not $setupDiagAvailable -or -not (Test-Path -LiteralPath $targetExecutable -PathType Leaf)) {
            Write-SmartM365Log "ERROR: SetupDiag.exe is not available locally and could not be downloaded."
            exit 1
        }
    }
    else {
        Write-SmartM365Log "SetupDiag.exe is present: $targetExecutable"
    }

    if (-not (Test-Path -Path $OutputRoot)) {
        New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
        Write-SmartM365Log "Created SetupDiag output directory: $OutputRoot"
    }

    $setupLogRoots = @(
        (Join-Path -Path $env:SystemDrive -ChildPath '$Windows.~BT\Sources\Rollback'),
        (Join-Path -Path $env:WINDIR -ChildPath "Panther"),
        (Join-Path -Path $env:SystemDrive -ChildPath '$Windows.~BT\Sources\Panther'),
        (Join-Path -Path $env:SystemDrive -ChildPath 'Windows.old\$Windows.~BT\Sources\Rollback')
    )

    $selectedLogRoot = $null
    foreach ($setupLogRoot in $setupLogRoots) {
        if (Test-Path -LiteralPath $setupLogRoot) {
            $selectedLogRoot = $setupLogRoot
            break
        }
    }

    $setupDiagArguments = @(
        "/Format:xml",
        "/ZipLogs:$ZipLogs",
        "/Output:$setupDiagResultPath",
        "/RegPath:HKEY_LOCAL_MACHINE\SYSTEM\Setup\SetupDiag\Results"
    )

    if (-not [string]::IsNullOrWhiteSpace($selectedLogRoot)) {
        Write-SmartM365Log "Setup logs found. Running SetupDiag in offline mode. LogsPath=$selectedLogRoot"
        $setupDiagArguments += "/LogsPath:$selectedLogRoot"
        $executionMode = "Offline"
    }
    else {
        Write-SmartM365Log "No setup logs found. Running SetupDiag in online mode."
        $executionMode = "Online"
    }

    $setupDiagProcess = Start-Process -FilePath $targetExecutable -ArgumentList $setupDiagArguments -Wait -PassThru -WindowStyle Hidden
    Write-SmartM365Log "SetupDiag completed with exit code: $($setupDiagProcess.ExitCode)"

    try {
        $binary = Get-Item -LiteralPath $targetExecutable -ErrorAction Stop
        $fileVersion = $binary.VersionInfo.FileVersion
        $fileHash = (Get-FileHash -LiteralPath $targetExecutable -Algorithm SHA256 -ErrorAction Stop).Hash
    }
    catch {
        $fileVersion = $null
        $fileHash = $null
    }

    $summary = [ordered]@{
        TimestampUtc      = (Get-Date).ToUniversalTime().ToString("s")
        LastRunUtc        = (Get-Date).ToUniversalTime().ToString("s")
        ExecutionMode     = $executionMode
        SetupDiagExe      = $targetExecutable
        SetupDiagVersion  = $fileVersion
        SetupDiagSha256   = $fileHash
        SetupDiagExitCode = $setupDiagProcess.ExitCode
        XmlPath           = $setupDiagResultPath
        MatchingProfile   = $null
        LastOperation     = $null
        ErrorCode         = $null
        ErrorHint         = $null
        ResultFound       = $false
    }

    if (Test-Path -LiteralPath $setupDiagResultPath -PathType Leaf) {
        try {
            $xmlRaw = Get-Content -LiteralPath $setupDiagResultPath -Raw -ErrorAction Stop

            if ($xmlRaw -match "Matching Profile found:\s*(.+)") {
                $summary.MatchingProfile = $matches[1].Trim()
            }

            if ($xmlRaw -match "Last Operation:\s*(.+)") {
                $summary.LastOperation = $matches[1].Trim()
            }

            if ($xmlRaw -match "Error:\s*(0x[0-9A-Fa-f]+)") {
                $summary.ErrorCode = $matches[1].ToUpperInvariant()
            }

            if ($summary.ErrorCode) {
                $summary.ErrorHint = Get-SetupDiagErrorHint -ErrorCode $summary.ErrorCode
            }

            $summary.ResultFound = $true
            Write-SmartM365Log "SetupDiag XML result parsed successfully."
        }
        catch {
            Write-SmartM365Log "WARNING: SetupDiag XML parsing failed: $($_.Exception.Message)"
        }
    }
    else {
        Write-SmartM365Log "WARNING: SetupDiagResults.xml was not generated."
    }

    New-Item -Path "HKLM:\SOFTWARE\SmartM365\IntuneRemediation" -Force | Out-Null
    New-Item -Path $enterpriseRegistryPath -Force | Out-Null

    foreach ($key in $summary.Keys) {
        New-ItemProperty -Path $enterpriseRegistryPath -Name $key -Value (ConvertTo-RegistryValue -Value $summary[$key]) -PropertyType String -Force | Out-Null
    }

    $summary | ConvertTo-Json -Depth 5 | Out-File -FilePath $jsonPath -Encoding UTF8 -Force
    Write-SmartM365Log "JSON summary written: $jsonPath"
    Write-SmartM365Log "Registry summary written: $enterpriseRegistryPath"

    if (-not $summary.ResultFound) {
        Write-SmartM365Log "SetupDiag completed but no usable XML result was found."
        exit 1
    }

    Write-SmartM365Log "===== SetupDiag remediation completed successfully ====="
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"

    try {
        if (-not (Test-Path -Path $enterpriseDirectory)) {
            New-Item -ItemType Directory -Path $enterpriseDirectory -Force | Out-Null
        }
        if (-not (Test-Path -Path $logDirectory)) {
            New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
        }

        Add-Content -Path $logPath -Value ("[{0}] ERROR: {1}" -f (Get-Date).ToString("s"), $_.Exception.Message) -Encoding UTF8
    }
    catch {
        Write-Output "LogWriteFailed=$($_.Exception.Message)"
    }

    exit 1
}

