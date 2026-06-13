# Name: SmartM365-GroupPolicy-Stale-Remediation.ps1
# Version: 1.0
$ErrorActionPreference = "Stop"
$Scenario = "GroupPolicy-Stale"
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path -Path $LogRoot -ChildPath "$Scenario-Remediation.log"
$groupPolicyStatePath = "Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Extension-List\{00000000-0000-0000-0000-000000000000}"

function Write-SmartM365Log {
    param([string]$Message)
    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8
}

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-SmartM365Log "RemediationStarted"

    $process = Start-Process -FilePath (Join-Path $env:SystemRoot "System32\gpupdate.exe") -ArgumentList "/force" -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop
    Write-SmartM365Log "GpupdateExitCode=$($process.ExitCode)"

    if ($process.ExitCode -ne 0) {
        Write-SmartM365Log "Status=GpupdateFailed"
        exit 1
    }

    $state = Get-ItemProperty -Path $groupPolicyStatePath -ErrorAction Stop
    $fileTime = ([Int64]$state.startTimeHi -shl 32) -bor [UInt32]$state.startTimeLo
    $lastGPUpdateDate = [datetime]::FromFileTime($fileTime)
    $lastGPUpdateHours = (New-TimeSpan -Start $lastGPUpdateDate -End (Get-Date)).TotalHours

    if ($lastGPUpdateHours -le 24) {
        Write-SmartM365Log ("Status=Completed LastGPUpdateHours={0:N1}" -f $lastGPUpdateHours)
        exit 0
    }

    Write-SmartM365Log ("Status=CompletedButStillStale LastGPUpdateHours={0:N1}" -f $lastGPUpdateHours)
    exit 1
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    try { Write-SmartM365Log "Status=Error Message=$($_.Exception.Message)" } catch { Write-Output "LogWriteFailed=$($_.Exception.Message)" }
    exit 1
}
