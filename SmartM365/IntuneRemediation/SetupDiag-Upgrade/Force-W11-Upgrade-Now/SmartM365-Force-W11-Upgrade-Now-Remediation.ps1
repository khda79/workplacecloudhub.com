<#
  Name: SmartM365-Force-W11-Upgrade-Now-Remediation.ps1
    Version: 1.0
  Purpose: Force an immediate Windows Update scan/download/install (for assigned updates such as Windows 11 feature upgrade).
  Use case: Manual "Run Script" from Intune per device.
  Run as: SYSTEM (administrator)
#>

$ErrorActionPreference = 'Stop'

function Invoke-AssignedUpdatesInstall {
    $session  = New-Object -ComObject 'Microsoft.Update.Session'
    $searcher = $session.CreateUpdateSearcher()
    $criteria = "IsInstalled=0 and IsHidden=0 and IsAssigned=1"
    $result   = $searcher.Search($criteria)

    if ($result.Updates.Count -eq 0) {
        Write-Host "No applicable assigned updates found via COM API."
        return 0
    }

    $toInstall = New-Object -ComObject 'Microsoft.Update.UpdateColl'
    foreach ($u in $result.Updates) { [void]$toInstall.Add($u) }

    Write-Host "Downloading updates..."
    $downloader = $session.CreateUpdateDownloader()
    $downloader.Updates = $toInstall
    [void]$downloader.Download()

    Write-Host "Installing updates..."
    $installer = $session.CreateUpdateInstaller()
    $installer.Updates = $toInstall
    $installResult = $installer.Install()

    if ($installResult.RebootRequired) {
        Write-Host "Reboot required � scheduling reboot in 60 seconds."
        shutdown.exe /r /t 60 /c "Rebooting to complete Windows update"
        return 2
    }
    else {
        Write-Host "Installation completed. No reboot required."
        return 1
    }
}

try {
    $result = Invoke-AssignedUpdatesInstall
    if ($result -eq 0) {
        Write-Host "Fallback to USOClient triggers..."
        Start-Process "$env:SystemRoot\System32\UsoClient.exe" -ArgumentList "StartScan" -WindowStyle Hidden -Wait
        Start-Process "$env:SystemRoot\System32\UsoClient.exe" -ArgumentList "StartDownload" -WindowStyle Hidden -Wait
        Start-Process "$env:SystemRoot\System32\UsoClient.exe" -ArgumentList "StartInstall" -WindowStyle Hidden -Wait
        exit 0
    }
    exit $result
}
catch {
    Write-Error ("Update process failed: " + $_.Exception.Message)
    Start-Process "$env:SystemRoot\System32\UsoClient.exe" -ArgumentList "StartScan" -WindowStyle Hidden -Wait
    Start-Process "$env:SystemRoot\System32\UsoClient.exe" -ArgumentList "StartDownload" -WindowStyle Hidden -Wait
    Start-Process "$env:SystemRoot\System32\UsoClient.exe" -ArgumentList "StartInstall" -WindowStyle Hidden -Wait
    exit 0
}
