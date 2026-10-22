#Requires -Version 5.1
<#
.SYNOPSIS
    Runs the Microsoft CSS-Exchange Health Checker for the local Exchange server only.

.DESCRIPTION
    Deployed by the Health Checker scheduled-task installer (and kept current by the
    infrastructure collector) to <root>\Bin on every Exchange server, next to HealthChecker.ps1.
    The scheduled task runs it daily and the collector starts it on demand over CIM/DCOM.

    The runner is standalone: the Exchange servers have no access to the repository.
    On every run it:
    - takes a machine-wide lock so two runs never overlap;
    - removes the previous run output (only the last run is kept);
    - checks HealthChecker.ps1 against the expected SHA256 and a valid Microsoft Corporation
      Authenticode signature, and refuses to run it otherwise;
    - starts Health Checker in a Windows PowerShell 5.1 child process for the local server only
      (never with -Server), with -SkipVersionCheck and a time limit;
    - converts the Health Checker XML on the server into <root>\Output\result.json, which holds
      strings and numbers only: the collector never deserializes XML that another server produced;
    - writes <root>\Output\status.json: state, UTC times, exit code, result file name and versions.
      The status holds no host name, domain or secret.

    Raw output (XML, text and debug files) stays on the server under <root>\Output\Raw until the
    next run. The collector copies only result.json.

.PARAMETER TimeoutMinutes
    Time limit for Health Checker. The child process tree is stopped when it is exceeded.

.EXAMPLE
    powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\Program Files\SmartM365\HealthChecker\Bin\SmartM365-Exchange-OnPrem-HealthChecker-LocalRunner.ps1"

.VERSION
1.1.1

.REQUIREMENTS
    Windows PowerShell 5.1 on an Exchange server, run elevated (scheduled task with highest privileges).
    Run account: local administrator, Exchange View-Only Organization Management, Log on as a batch job.
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 240)]
    [int]$TimeoutMinutes = 30
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.1.1'
# Must match the collector constants (checked by the offline test).
$HealthCheckerVersion = '26.10.13.1644'
$HealthCheckerExpectedSha256 = 'BF8758E96E8760226E3094D45A84909AFFA9F91167DE11F739D8E3D9343229DC'
$RunnerMutexName = 'Global\SmartM365-Exchange-HealthChecker-LocalRunner'

function Repair-HealthCheckerModulePath {
    # A Windows PowerShell 5.1 process started from PowerShell 7 (the orchestrator, a pwsh console) inherits its
    # PSModulePath, and the PowerShell 7 builds of built-in modules such as Microsoft.PowerShell.Security
    # cannot load in 5.1 (Get-Acl, Get-AuthenticodeSignature). Windows PowerShell modules come first again.
    if ($PSVersionTable.PSEdition -ne 'Desktop') { return }
    $kept = New-Object System.Collections.Generic.List[string]
    $kept.Add((Join-Path -Path $PSHOME -ChildPath 'Modules'))
    foreach ($entry in @([string]$env:PSModulePath -split [IO.Path]::PathSeparator)) {
        $trimmed = $entry.Trim().TrimEnd('\')
        if (-not $trimmed -or $trimmed -match '(?i)(^|\\)PowerShell\\(7[^\\]*\\)?Modules$') { continue }
        if (-not ($kept -contains $trimmed)) { $kept.Add($trimmed) }
    }
    $env:PSModulePath = $kept -join [IO.Path]::PathSeparator
}

function Test-HealthCheckerScriptIntegrity {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$ExpectedSha256
    )

    $result = [pscustomobject]@{ IsValid = $false; Reason = $null; Sha256 = $null; SignatureStatus = $null; Signer = $null }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $result.Reason = "Health Checker script not found: $Path"
        return $result
    }

    $result.Sha256 = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    if ($result.Sha256 -ne $ExpectedSha256.ToUpperInvariant()) {
        $result.Reason = "SHA256 mismatch: expected $($ExpectedSha256.ToUpperInvariant()), found $($result.Sha256)."
        return $result
    }

    $signature = Get-AuthenticodeSignature -FilePath $Path
    $result.SignatureStatus = [string]$signature.Status
    if ($signature.SignerCertificate) { $result.Signer = [string]$signature.SignerCertificate.Subject }
    if ([string]$signature.Status -ne 'Valid' -or -not $signature.SignerCertificate) {
        $result.Reason = "Authenticode signature is not valid (status $($signature.Status))."
        return $result
    }
    $subject = [string]$signature.SignerCertificate.Subject
    $issuer = [string]$signature.SignerCertificate.Issuer
    if ($subject -notmatch '(^|,\s*)CN=Microsoft Corporation(,|$)' -or $subject -notmatch '(^|,\s*)O=Microsoft Corporation(,|$)' -or $issuer -notmatch '(^|,\s*)O=Microsoft Corporation(,|$)') {
        $result.Reason = "Authenticode signer is not Microsoft Corporation: $subject"
        return $result
    }

    $result.IsValid = $true
    return $result
}

function Get-HealthCheckerLineValue {
    param(
        [AllowNull()][object]$Line,
        [Parameter(Mandatory)][string]$Name
    )

    # Serialized entries can be plain strings when Export-Clixml ran out of depth.
    if ($null -eq $Line) { return $null }
    if ($Line -is [string]) { return $(if ($Name -eq 'DisplayValue') { $Line } else { $null }) }
    $property = $Line.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function ConvertTo-HealthCheckerResultLine {
    param([Parameter(Mandatory)][string]$Path)

    # Import-Clixml is only used where the XML was produced (the server itself); the result lines
    # hold strings and numbers only, so they can cross to the collector as JSON.
    $data = Import-Clixml -LiteralPath $Path
    $displayResults = Get-HealthCheckerLineValue -Line $data -Name 'DisplayResults'
    if ($displayResults -isnot [System.Collections.IDictionary] -or $displayResults.Count -eq 0) {
        throw 'The Health Checker XML has no DisplayResults section.'
    }

    $lines = New-Object 'System.Collections.Generic.List[object]'
    $groups = @(foreach ($entry in $displayResults.GetEnumerator()) {
        $order = 0
        [void][int]::TryParse([string](Get-HealthCheckerLineValue -Line $entry.Key -Name 'DisplayOrder'), [ref]$order)
        $categoryName = if ($entry.Key -is [string]) { $entry.Key } else { [string](Get-HealthCheckerLineValue -Line $entry.Key -Name 'Name') }
        [pscustomobject]@{ Category = $categoryName; Order = $order; Lines = @($entry.Value) }
    })
    foreach ($group in @($groups | Sort-Object Order, Category)) {
        $parentCheck = $null
        foreach ($line in $group.Lines) {
            if ($null -eq $line) { continue }
            $name = [string](Get-HealthCheckerLineValue -Line $line -Name 'Name')
            $tab = 0
            [void][int]::TryParse([string](Get-HealthCheckerLineValue -Line $line -Name 'TabNumber'), [ref]$tab)
            $testingName = [string](Get-HealthCheckerLineValue -Line $line -Name 'TestingName')
            $isSubLine = $tab -gt 1 -or [string]::IsNullOrWhiteSpace($name) -or $name -in @('Error', 'Warning')
            if (-not $isSubLine -or [string]::IsNullOrWhiteSpace($parentCheck)) {
                $parentCheck = if ([string]::IsNullOrWhiteSpace($name)) { $testingName } else { $name }
            }
            # An Error/Warning line with its own testing name (for example "Out of Date") is its own check.
            $check = if ($isSubLine -and -not [string]::IsNullOrWhiteSpace($testingName) -and $testingName -ne $name -and $testingName -notin @('Error', 'Warning')) { $testingName } else { $parentCheck }
            $writeType = [string](Get-HealthCheckerLineValue -Line $line -Name 'WriteType')
            $displayValue = Get-HealthCheckerLineValue -Line $line -Name 'DisplayValue'
            $displayText = if ($writeType -eq 'OutColumns' -or $null -eq $displayValue) { $null } elseif ($displayValue -is [string]) { $displayValue } else { (@($displayValue) | ForEach-Object { [string]$_ }) -join '; ' }
            $lines.Add([pscustomobject][ordered]@{
                Category      = [string]$group.Category
                CategoryOrder = [int]$group.Order
                Check         = [string]$check
                Item          = $(if ($isSubLine -and -not [string]::IsNullOrWhiteSpace($name)) { $name } else { $null })
                Name          = $name
                WriteType     = $writeType
                DisplayValue  = $displayText
                TestingName   = [string](Get-HealthCheckerLineValue -Line $line -Name 'TestingName')
                CustomName    = [string](Get-HealthCheckerLineValue -Line $line -Name 'CustomName')
                CustomValue   = [string](Get-HealthCheckerLineValue -Line $line -Name 'CustomValue')
            }) | Out-Null
        }
    }
    return $lines.ToArray()
}

function Test-LocalRunnerElevation {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    return (New-Object Security.Principal.WindowsPrincipal($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Clear-LocalRunnerOutput {
    param([Parameter(Mandatory)][string]$OutputFolder)

    # Only the last run is kept: everything from the previous run is removed first.
    if (-not (Test-Path -LiteralPath $OutputFolder -PathType Container)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
        return
    }
    foreach ($item in @(Get-ChildItem -LiteralPath $OutputFolder -Force)) {
        Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
    }
}

function Write-LocalRunnerJson {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][object]$Value
    )

    # Written to a temporary file then moved, so a reader never sees a partial file.
    $temporaryPath = '{0}.{1}.tmp' -f $Path, [guid]::NewGuid().ToString('N')
    [IO.File]::WriteAllText($temporaryPath, ($Value | ConvertTo-Json -Depth 4 -Compress), (New-Object System.Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $temporaryPath -Destination $Path -Force
}

function Write-LocalRunnerStatus {
    param(
        [Parameter(Mandatory)][string]$OutputFolder,
        [Parameter(Mandatory)][object]$Status
    )

    Write-LocalRunnerJson -Path (Join-Path -Path $OutputFolder -ChildPath 'status.json') -Value $Status
}

function New-LocalRunnerStatus {
    param([Parameter(Mandatory)][string]$RunToken)

    return [pscustomobject][ordered]@{
        SchemaVersion        = 1
        RunToken             = $RunToken
        State                = 'Running'
        Message              = $null
        RunnerVersion        = $ScriptVersion
        HealthCheckerVersion = $HealthCheckerVersion
        StartedAtUtc         = [DateTime]::UtcNow.ToString('o')
        FinishedAtUtc        = $null
        DurationSeconds      = $null
        ExitCode             = $null
        ResultFileName       = $null
    }
}

function Get-LocalHealthCheckerArgument {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][string]$RawFolder
    )

    # No -Server: Health Checker defaults to the local computer, so nothing is collected remotely.
    return @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $ScriptPath), '-OutputFilePath', ('"{0}"' -f $RawFolder), '-SkipVersionCheck')
}

function Invoke-LocalHealthCheckerProcess {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][string]$RawFolder,
        [Parameter(Mandatory)][int]$TimeoutSeconds,
        [string]$PowerShellPath = (Join-Path -Path $env:SystemRoot -ChildPath 'System32\WindowsPowerShell\v1.0\powershell.exe')
    )

    if (-not (Test-Path -LiteralPath $RawFolder -PathType Container)) { New-Item -ItemType Directory -Path $RawFolder -Force | Out-Null }
    $consolePath = Join-Path -Path $RawFolder -ChildPath 'HealthChecker-Console.txt'
    $errorPath = Join-Path -Path $RawFolder -ChildPath 'HealthChecker-Errors.txt'
    $arguments = Get-LocalHealthCheckerArgument -ScriptPath $ScriptPath -RawFolder $RawFolder
    $started = Get-Date
    $result = [pscustomobject]@{ State = $null; Message = $null; ExitCode = $null; DurationSeconds = $null; XmlFileName = $null }
    try {
        $process = Start-Process -FilePath $PowerShellPath -ArgumentList $arguments -PassThru -WindowStyle Hidden -RedirectStandardOutput $consolePath -RedirectStandardError $errorPath -ErrorAction Stop
        $null = $process.Handle
    }
    catch {
        $result.State = 'LaunchFailed'
        $result.Message = "Health Checker could not be started: $($_.Exception.Message)"
        $result.DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 0)
        return $result
    }

    if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
        & taskkill.exe /PID $process.Id /T /F 2>&1 | Out-Null
        try { [void]$process.WaitForExit(30000) } catch { Write-Verbose "Stopped Health Checker process already gone: $($_.Exception.Message)" }
        $result.State = 'Timeout'
        $result.Message = "Health Checker exceeded the $([math]::Round($TimeoutSeconds / 60, 1)) minute limit on the server and was stopped."
        $result.DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 0)
        return $result
    }

    $result.ExitCode = $process.ExitCode
    $result.DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 0)
    $xmlFiles = @(Get-ChildItem -LiteralPath $RawFolder -Filter 'HealthChecker-*.xml' -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike 'HealthChecker-ScriptDebugObject*' -and $_.Name -notlike 'HealthChecker-ExchangeDCCoreRatio-*' } | Sort-Object LastWriteTime -Descending)
    $consoleText = if (Test-Path -LiteralPath $consolePath -PathType Leaf) { [System.IO.File]::ReadAllText($consolePath) } else { '' }
    if ($xmlFiles.Count -gt 0) {
        $result.State = 'Completed'
        $result.XmlFileName = $xmlFiles[0].Name
    }
    elseif ($consoleText -match 'needs to be executed in elevated mode') {
        $result.State = 'NotElevated'
        $result.Message = 'Health Checker requires an elevated (administrator) session.'
    }
    elseif ($consoleText -match 'Failed to Load Exchange PowerShell Module|Something went wrong while loading the Exchange Management Shell') {
        $result.State = 'ExchangeShellFailed'
        $result.Message = 'Health Checker could not load the Exchange Management Shell with the task account.'
    }
    else {
        $result.State = 'NoOutput'
        $result.Message = "Health Checker produced no XML result (exit code $($result.ExitCode))."
    }
    return $result
}

function Invoke-LocalHealthCheckerRun {
    param(
        [Parameter(Mandatory)][string]$RootFolder,
        [Parameter(Mandatory)][int]$TimeoutSeconds,
        [string]$ExpectedSha256 = $HealthCheckerExpectedSha256
    )

    $outputFolder = Join-Path -Path $RootFolder -ChildPath 'Output'
    $rawFolder = Join-Path -Path $outputFolder -ChildPath 'Raw'
    $scriptPath = Join-Path -Path $RootFolder -ChildPath 'Bin\HealthChecker.ps1'

    Clear-LocalRunnerOutput -OutputFolder $outputFolder
    $status = New-LocalRunnerStatus -RunToken ([guid]::NewGuid().ToString())
    Write-LocalRunnerStatus -OutputFolder $outputFolder -Status $status
    $started = Get-Date

    $integrity = Test-HealthCheckerScriptIntegrity -Path $scriptPath -ExpectedSha256 $ExpectedSha256
    if (-not $integrity.IsValid) {
        $status.State = 'IntegrityFailed'
        $status.Message = "Health Checker was not run: $($integrity.Reason -replace [regex]::Escape($RootFolder), '<root>')"
    }
    elseif (-not (Test-LocalRunnerElevation)) {
        $status.State = 'NotElevated'
        $status.Message = 'Health Checker was not run: the task must run with highest privileges as a local administrator.'
    }
    else {
        $run = Invoke-LocalHealthCheckerProcess -ScriptPath $scriptPath -RawFolder $rawFolder -TimeoutSeconds $TimeoutSeconds
        $status.State = $run.State
        $status.Message = $run.Message
        $status.ExitCode = $run.ExitCode
        if ($run.State -eq 'Completed') {
            try {
                $lines = @(ConvertTo-HealthCheckerResultLine -Path (Join-Path -Path $rawFolder -ChildPath $run.XmlFileName))
                Write-LocalRunnerJson -Path (Join-Path -Path $outputFolder -ChildPath 'result.json') -Value ([pscustomobject][ordered]@{ SchemaVersion = 1; HealthCheckerVersion = $HealthCheckerVersion; Lines = $lines })
                $status.ResultFileName = 'result.json'
            }
            catch {
                $status.State = 'ConvertFailed'
                $status.Message = "The Health Checker XML could not be converted on the server: $($_.Exception.Message -replace [regex]::Escape($RootFolder), '<root>')"
            }
        }
    }

    $status.FinishedAtUtc = [DateTime]::UtcNow.ToString('o')
    $status.DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 0)
    Write-LocalRunnerStatus -OutputFolder $outputFolder -Status $status
    return $status
}

# Entry point. The functions above are also loaded by the offline test without running this part.
if ($MyInvocation.InvocationName -ne '.') {
    Repair-HealthCheckerModulePath
    $rootFolder = Split-Path -Path $PSScriptRoot -Parent
    $mutex = New-Object System.Threading.Mutex($false, $RunnerMutexName)
    $hasLock = $false
    try {
        try { $hasLock = $mutex.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $hasLock = $true }
        if (-not $hasLock) {
            # Another run is in progress: its status.json stays authoritative.
            Write-Output 'Health Checker local runner: another run is in progress.'
            exit 2
        }
        $final = Invoke-LocalHealthCheckerRun -RootFolder $rootFolder -TimeoutSeconds ($TimeoutMinutes * 60)
        Write-Output ('Health Checker local runner {0}: {1}. {2}' -f $ScriptVersion, $final.State, $final.Message)
        if ($final.State -eq 'Completed') { exit 0 }
        exit 1
    }
    finally {
        if ($hasLock) { $mutex.ReleaseMutex() }
        $mutex.Dispose()
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD+hUDoz9EX2X0d
# URfS5+zeDA+SC04gVrqhpZkB6jWS0aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJMbxqqNAmEz7TlYFE2dYWY51BtW2b65BT7Nyu5MlEhdMA0GCSqG
# SIb3DQEBAQUABIIBgKpBpGMmJ8/A+IRiLYrtLiwJzMiCtcPvSW8JTiKr4C199RSK
# ZJm+i8gBZiN1ymvY/dTsC/Z2IqVA8UrD/Vp3IbwLs5zrtCKHHReiqXwjr1UcF2/2
# +BBrAZCYa6Ex+IBfMUN16DiY6QrqiT9OTNH7zaWSI8LnOBdyftczdKQFD6Uqjdsw
# ky40SXH2Bm+yGbZdhbV9EjiwJMgpxJndYbXFQUuf0jP0nu/+5GpY1Yx/kQGxpMSg
# RS9R2PGpcLCDdzrAIhZm8NZQE74GLISpElRl/O1UzccLvZobucRireEV1JNW6JyN
# 0JnLvtTHaalxriD4ccZz3IHSAT1dFu7ptSWI0RtN2bjL8swj4gnX7BrCdLkZY5MG
# J/pnFYq5+1crGuZq6zGqXQ3dUjLya1hjg8+W99Kz95JJHxqDYPAeDhQ5qMR8UndG
# CZb+EEa0fLyMStWqsEHocN3mKMjOZcjJrqwPLmA5EbZJ2zmK+mPlq0flJlURqK4J
# o0pLykb2rvz2c+Z/GqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjIwNzQx
# MDVaMC8GCSqGSIb3DQEJBDEiBCDPBclNs07EULei51/WOcMhZw5Xo4jhTOQdsyzv
# xmEdPDANBgkqhkiG9w0BAQEFAASCAgBQGxuCHoeWfAQ/muaGF/KmuKf0l5DXpsbK
# A1tC8T+T3hUjV54q9NJbnjIh7QuYY4rob36cZZOj7AnyiYm6wbpEPMkdKOJoQ8hd
# 1WnGrWfW1eBhzjald9UNGyDJHgxzfVSq6m59JzsU2M+1ek6BjByD4i3wCzE+xRgN
# w3/lttErqXuXN/C83Y+ssSpmNqq4DeZb96ykYhxexOpCJjPPDZDVYbqHNfKWXPEv
# /7kYyiA9d6/2wqCoSIs8Em63fQxplxJCJSEaWv9rCH5aPQLK6U7dPcmmscUgK66G
# rQK76mw1O0VmKpe+dW0D5qDeFHZNjXLpLY7J33QPeU1IixhlVGmVCwZApfjiIb+k
# JYXlwiXZ0LDt+elC0briBxUQ47H2yp8nXTjMtDSCKugQQ8XPDGyByAqT0Kosh5fi
# II4tmKycSWaEizM0kSc7oIlFQq3Ekhi9dzXcn0DiqZFnGW5VsW3clcEPoeyofOwh
# wdZ4njWvqYutHNijQGQgx8oyLrxVbrvU8p4VBuI5GLA8B0jrC3ID2rFxvS2Jg/1p
# 0hiB6+88RzgmUuxADAd+RzmwGecZK47bnG888c0tUuZJYCtj18jNGyDC+0IBkcKb
# JRPM25cD33mwWw5ge3CSRcH0QASqiWFzNiXN/x6uJ/EbLp8e3T+KrGIaDuYohTjR
# MqQzEiQg4Q==
# SIG # End signature block
