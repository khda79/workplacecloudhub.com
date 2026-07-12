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

    $line = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
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

        Add-Content -Path $logPath -Value ("{0} ERROR: {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $_.Exception.Message) -Encoding UTF8
    }
    catch {
        Write-Output "LogWriteFailed=$($_.Exception.Message)"
    }

    exit 1
}


# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBxawoxI1LtX3hQ
# 4J2d81mUL4azcdSuJ4JZ2oQfkGB89qCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCDUIUhTDM0btOpSKjcED1DZoCjLt8MQI453HYZrqvjLAjANBgkqhkiG9w0B
# AQEFAASCAYBVuTxoT0WphOA7oDgTa/IkX9vJWtQSq3Q4uGGOtgKaSlQ4wE6mtFmi
# R7YoXmtt11T+362FZ+KE/tsEAOOw8VYEfjb45/p3JSoxBvjimM/ZgZTZ7JNhr9/l
# rqsXzBHP6HM+9p/hpnb0vMjssKhf5/4Rvo1tueStuPPgmXDeZEHz6lCp38YI4Ltn
# blShnoOEo74K/jGUfxKmy33FfmNCRdauSge5SrR6iJygzqC63qBYMUw9vzf/GoTx
# RuXlKadpOmXfwd3MeAiXlMtaRSW5GBGRddeiNbP0sb/gdovtFs74c3bjQ6WuBELs
# dzLATShSDdVYO42at/DS7Ausu8U4YsnewfSRU8BoJ1Q6uE/BjKEHpjIFPqLP3mjM
# 1jVi1N2uw2xsHdTWhWJ55uVxVjzvDxrFeQ/hH7+i4u1O3LSwW3ORqWpxvbLQ4DX0
# hcHLsE7H3+/X4Ya8/pZjkoqyPvwM03gtGHLibaRRoffzKGT/D5q1DkaTmnkm8Esy
# fgaAYMoeCIU=
# SIG # End signature block
