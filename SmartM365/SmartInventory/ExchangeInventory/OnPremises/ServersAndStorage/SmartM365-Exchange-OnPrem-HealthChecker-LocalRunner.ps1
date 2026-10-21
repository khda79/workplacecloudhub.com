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
    - writes <root>\Output\status.json: state, UTC times, exit code, XML file name and versions.
      The status holds no host name, domain or secret.

    Raw output (XML, text and debug files) stays on the server under <root>\Output\Raw until the
    next run. The collector copies only the XML file.

.PARAMETER TimeoutMinutes
    Time limit for Health Checker. The child process tree is stopped when it is exceeded.

.EXAMPLE
    powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\Program Files\SmartM365\HealthChecker\Bin\SmartM365-Exchange-OnPrem-HealthChecker-LocalRunner.ps1"

.VERSION
1.0.0

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

$ScriptVersion = '1.0.0'
# Must match the collector constants (checked by the offline test).
$HealthCheckerVersion = '26.10.13.1644'
$HealthCheckerExpectedSha256 = 'BF8758E96E8760226E3094D45A84909AFFA9F91167DE11F739D8E3D9343229DC'
$RunnerMutexName = 'Global\SmartM365-Exchange-HealthChecker-LocalRunner'

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

function Write-LocalRunnerStatus {
    param(
        [Parameter(Mandatory)][string]$OutputFolder,
        [Parameter(Mandatory)][object]$Status
    )

    # Written to a temporary file then moved, so a reader never sees a partial status.
    $statusPath = Join-Path -Path $OutputFolder -ChildPath 'status.json'
    $temporaryPath = Join-Path -Path $OutputFolder -ChildPath ('status.{0}.tmp' -f [guid]::NewGuid().ToString('N'))
    [IO.File]::WriteAllText($temporaryPath, ($Status | ConvertTo-Json -Depth 3), (New-Object System.Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $temporaryPath -Destination $statusPath -Force
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
        XmlFileName          = $null
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
        if ($run.XmlFileName) { $status.XmlFileName = 'Raw\' + $run.XmlFileName }
    }

    $status.FinishedAtUtc = [DateTime]::UtcNow.ToString('o')
    $status.DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 0)
    Write-LocalRunnerStatus -OutputFolder $outputFolder -Status $status
    return $status
}

# Entry point. The functions above are also loaded by the offline test without running this part.
if ($MyInvocation.InvocationName -ne '.') {
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
