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
        Write-Output "No applicable assigned updates found via COM API."
        return "NoUpdates"
    }

    $toInstall = New-Object -ComObject 'Microsoft.Update.UpdateColl'
    foreach ($u in $result.Updates) { [void]$toInstall.Add($u) }

    Write-Output "Downloading updates..."
    $downloader = $session.CreateUpdateDownloader()
    $downloader.Updates = $toInstall
    [void]$downloader.Download()

    Write-Output "Installing updates..."
    $installer = $session.CreateUpdateInstaller()
    $installer.Updates = $toInstall
    $installResult = $installer.Install()

    if ($installResult.RebootRequired) {
        Write-Output "Reboot required; scheduling reboot in 60 seconds."
        shutdown.exe /r /t 60 /c "Rebooting to complete Windows update"
        return "RebootScheduled"
    }
    else {
        Write-Output "Installation completed. No reboot required."
        return "Installed"
    }
}

try {
    $result = Invoke-AssignedUpdatesInstall
    if ($result -eq "NoUpdates") {
        Write-Output "Fallback to USOClient triggers..."
        $usoClient = Join-Path $env:SystemRoot "System32\UsoClient.exe"
        if (-not (Test-Path -LiteralPath $usoClient -PathType Leaf)) {
            throw "UsoClient.exe was not found."
        }
        Start-Process $usoClient -ArgumentList "StartScan" -WindowStyle Hidden -Wait
        Start-Process $usoClient -ArgumentList "StartDownload" -WindowStyle Hidden -Wait
        Start-Process $usoClient -ArgumentList "StartInstall" -WindowStyle Hidden -Wait
        exit 0
    }
    Write-Output "Status=$result"
    exit 0
}
catch {
    Write-Output ("Status=Error Message=" + $_.Exception.Message)
    try {
        $usoClient = Join-Path $env:SystemRoot "System32\UsoClient.exe"
        if (Test-Path -LiteralPath $usoClient -PathType Leaf) {
            Write-Output "Attempting USOClient fallback after COM API failure."
            Start-Process $usoClient -ArgumentList "StartScan" -WindowStyle Hidden -Wait
            Start-Process $usoClient -ArgumentList "StartDownload" -WindowStyle Hidden -Wait
            Start-Process $usoClient -ArgumentList "StartInstall" -WindowStyle Hidden -Wait
            Write-Output "Status=FallbackTriggered"
            exit 0
        }
    }
    catch {
        Write-Output ("Fallback failed: " + $_.Exception.Message)
    }
    exit 1
}
