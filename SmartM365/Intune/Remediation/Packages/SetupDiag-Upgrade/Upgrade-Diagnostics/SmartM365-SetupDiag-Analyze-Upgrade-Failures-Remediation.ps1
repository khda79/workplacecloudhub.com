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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC6oiNachiwSiLt
# v1UeH9PzlEKHj6tT+ccex04WB5HVvKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICtbwJ67XBbSnOguJLrKCP+6X8+W2+LSNqNC3zsI9VEDMA0GCSqG
# SIb3DQEBAQUABIIBgINPhvqhFb38W/4es2s6ufKtqHjukZn6cbKcumBh1odjPBVJ
# 8oCdndE9FA9Es5W3jl1BpG3G6OGrJfJX7kOmICNHQVfD7ZSHgHqUuzB/14JlQ9mI
# FLhXFAqgmIowmsnSx7pUJ4WEm+gcLHw7XM/2TCNr+EX1j8GnZcbPSPJxYF/n0yW+
# moeNIEhYHS7eBbdHcnzZ7FCK8ObyBtxv/jXpV5f1G4+virYz5XytJ/m6sIsGFfq6
# HbqQb/Evh2iOpZd/td1+h66FFDovK6uFQsWH21CEv7e6TbWWkvYRUG+ShV2mROjT
# 2yh91ZhvGZnOxwtts2G0xGnrSV7BaFJLZOujxWpnjOmWyk2hKOTQnpfMzGL2pRQ+
# V4HV1j36nE/UdExkiJ6KKnGxo0iOlnvoFe/gw2gd4CpKYFS/plhhTthAChOfriKS
# sblSGclfGrqNwL46kVQvzVW4pz0VZetqUJLWnzvX6Xf2XnTbGb1EoJFPsAAUVFaN
# VtA31Eun6s6GFq4VYKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# NTBaMC8GCSqGSIb3DQEJBDEiBCCo64Yr5op/m5UOfGfYGJBLbnvPYy6+12EMfnsn
# hBIL7DANBgkqhkiG9w0BAQEFAASCAgCWly3F/b0ZJMrl23Rlb/7RbkfnpcVPZ9iu
# ucNqYXNTbMPToaYepKQdg49hpwpNgp124rfPSkbVW8ZwamHzd8nGVyE4/H8tN1Zj
# 6KCIePiW/J40hTRyDTi3lzPOzDUZu3IJtQxOor1ZtyRgNeUEVgwTeBTosFSiLhtT
# JgCeTtogQvgw2faMrctFR/ZI0iT581XwSa2QIPRB0C8X35kg+/1U5IhjTkZgK15Q
# CglK2TzOMg+1HooM92p8qFagZlB0YbBBT5e/2jZ/jr0S/lqbI0veWz6rlWUWaaPi
# dHx6TnjJDSNGJNyr4XGCSquMBIvQ/kxP6PBpTWNYp91eDMMWxAmcC3FkEC/XxfpS
# aGpbVvX1MqVz+WsUK207TOGaBaaUepvrgieaPNukqfe6bcJbpJwfN7WTwyaTjccF
# gsJOKTh5j+ZGz4ZDUOzJQguoVGNZMHVaOc4QSBOop3MjyzUwj8hgolGix8ShtByq
# w17ALgNYhIjgq7vq5D79VE+KB1DaXUCiBmW8XBGGD6wVDAoRZlwfHjfJfPjyc61r
# NCXTkfm+ejPwC7zdzht9Zch32scs2DBg/UO1Crqr6qY5ICZSHvr72dN+80qg0xVR
# zhsEJHo+qgFVEiMal4/88zlN7smm0GJVy2Dr7elr+NF14ANSXq1iB8WU3zBiexzW
# qHkf/XCtcg==
# SIG # End signature block
