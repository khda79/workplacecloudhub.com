# Name: SmartM365-GroupPolicy-Stale-Detection.ps1
# Version: 1.0
$ErrorActionPreference = "Stop"
$groupPolicyStatePath = "Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Extension-List\{00000000-0000-0000-0000-000000000000}"
$maxAgeDays = 7

try {
    $state = Get-ItemProperty -Path $groupPolicyStatePath -ErrorAction Stop

    if ($null -eq $state.startTimeHi -or $null -eq $state.startTimeLo) {
        Write-Output "Status=GroupPolicyTimestampMissing"
        exit 1
    }

    $fileTime = ([Int64]$state.startTimeHi -shl 32) -bor [UInt32]$state.startTimeLo
    $lastGPUpdateDate = [datetime]::FromFileTime($fileTime)
    $lastGPUpdateDays = (New-TimeSpan -Start $lastGPUpdateDate -End (Get-Date)).TotalDays

    if ($lastGPUpdateDays -gt $maxAgeDays) {
        Write-Output ("Status=GroupPolicyStale AgeDays={0:N1} MaxAgeDays={1}" -f $lastGPUpdateDays, $maxAgeDays)
        exit 1
    }

    Write-Output ("Status=GroupPolicyFresh AgeDays={0:N1}" -f $lastGPUpdateDays)
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 1
}
