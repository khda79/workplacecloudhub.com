# Name: SmartM365-GroupPolicy-Stale-Remediation.ps1
# Version: 1.0
# Define Variables

try {
    $compGPUpd = gpupdate /force
    $gpResult = [datetime]::FromFileTime(([Int64] ((Get-ItemProperty -Path "Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Extension-List\{00000000-0000-0000-0000-000000000000}").startTimeHi) -shl 32) -bor ((Get-ItemProperty -Path "Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Extension-List\{00000000-0000-0000-0000-000000000000}").startTimeLo))
    $lastGPUpdateDate = Get-Date ($gpResult[0])
    [int]$lastGPUpdateDays = (New-TimeSpan -Start $lastGPUpdateDate -End (Get-Date)).Days

    if ($lastGPUpdateDays -eq 0){        
        Write-Host "gpupdate completed successfully"            
        exit 0
    }
    else{
        Write-Host "gpupdate failed"
        }
}
catch{
    $errMsg = $_.Exception.Message
    return $errMsg
    exit 1
}
