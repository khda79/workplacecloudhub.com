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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBxawoxI1LtX3hQ
# 4J2d81mUL4azcdSuJ4JZ2oQfkGB89qCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDUIUhTDM0btOpSKjcE
# D1DZoCjLt8MQI453HYZrqvjLAjANBgkqhkiG9w0BAQEFAASCAYCqp+TPzaLjSoiq
# KEmwy98XYAUNVZuYxrt89QN8Qux+UJro6mNaH1jCVszyc7I1A6C78j4LoeIrquBA
# 71SdaqdT+p6vxnQwdzZO6BVtzkph8Zr0+XbJ3Yq5lhgNUnZc/VuSzi1cVmJD2JLs
# x6ObM8iH5AaM/qd35RPgVQB93Hg3Zj53i0e7V+Ick2dJbNymzKDr1rzV4C7jry54
# u+pyN3QSO+h/ATW5pl3vm45lzxl906Ctv+mup/BSmvs07TeLySYx7+k0pw+KSsnl
# 7aKXiwtZqnN0aZxBNKgVHdLj+fERqF/jpks30xQSklkAbYZfJs+6MgRfzv1/CnbF
# bUgCYS1rNLLc8JMqWRaeKTgpQlkLtCgXcJldk63FimhGEzW8teeC2reEQtyeKnHR
# jJsO5x3qc7C5XyFpJDfjyZwAxnhhaijN/YqQvSpgpBt9rl8dA4Wy1EyPG4mpruFu
# 5hOBwlWAxzMSBGYXk+pXfhLO7c+CZl2OIlsUPHCRD5Tutf/mZ0w=
# SIG # End signature block
