<#
    Name: SmartM365-Endpoint-Repair-WinRM-UWP-Store-Remediation.ps1
    Version: 1.0 
    Script de remédiation Intune (Windows ne se mets pas à niveau car MS Store et SoftwareDistribution sont corrompus ou indisponibles (erreur 403).
	
	- Active / répare WinRM
	- Répare Windows Update (SoftwareDistribution)
	- Lance DISM + SFC
	- Réenregistre toutes les apps UWP
	- Répare AppRepository
	- Réinstalle Microsoft Store
	- Réenregistre Frameworks UWP
	- Ne force aucun redémarrage
#>

# Script de remédiation Intune
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath 'SmartM365\IntuneRemediation\Logs\WinRM-UWP-Store'
$Log = Join-Path -Path $LogRoot -ChildPath 'Endpoint-Repair-WinRM-UWP-Store-Remediation.log'
if (-not (Test-Path -LiteralPath $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
$ErrorFound = $false

function Log { param($m) "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $m" | Out-File -FilePath $Log -Append }
function FlagError { param($m) Log $m; $global:ErrorFound = $true }

Log "===== Début remédiation ====="

# ---------------------------------------------------------
# 1. WinRM
# ---------------------------------------------------------
Log "Réparation WinRM"
try {
    Set-Service WinRM -StartupType Automatic -ErrorAction SilentlyContinue
    Start-Service WinRM -ErrorAction SilentlyContinue
    winrm quickconfig -force | Out-Null

    $listeners = winrm enumerate winrm/config/Listener 2>$null
    if ($listeners -notmatch "Transport = HTTP") {
        winrm delete winrm/config/Listener?Address=*+Transport=HTTP 2>$null | Out-Null
        winrm create winrm/config/Listener?Address=*+Transport=HTTP 2>$null | Out-Null
    }

    netsh advfirewall firewall add rule name="WinRM HTTP" dir=in action=allow protocol=TCP localport=5985 2>$null | Out-Null
}
catch { FlagError "Erreur WinRM : $($_.Exception.Message)" }

# ---------------------------------------------------------
# 2. Réparation Windows Update
# ---------------------------------------------------------
Log "Nettoyage SoftwareDistribution"
try {
    Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
    Stop-Service bits -Force -ErrorAction SilentlyContinue
    Stop-Service dosvc -Force -ErrorAction SilentlyContinue

    Remove-Item "C:\Windows\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue

    Start-Service wuauserv -ErrorAction SilentlyContinue
    Start-Service bits -ErrorAction SilentlyContinue
    Start-Service dosvc -ErrorAction SilentlyContinue
}
catch { FlagError "Erreur SoftwareDistribution : $($_.Exception.Message)" }

# ---------------------------------------------------------
# 3. DISM + SFC
# ---------------------------------------------------------
Log "DISM / SFC"
try {
    Start-Process dism.exe -ArgumentList "/Online","/Cleanup-Image","/RestoreHealth" -Wait -NoNewWindow
    Start-Process sfc.exe -ArgumentList "/scannow" -Wait -NoNewWindow
}
catch { FlagError "Erreur DISM/SFC : $($_.Exception.Message)" }

# ---------------------------------------------------------
# 4. Réparation services UWP critiques
# ---------------------------------------------------------
Log "Réparation services UWP"
$uwpServices = @("AppXSvc", "ClipSVC", "StateRepository", "WpnService")

foreach ($svc in $uwpServices) {
    try {
        Set-Service $svc -StartupType Automatic -ErrorAction SilentlyContinue
        Start-Service $svc -ErrorAction SilentlyContinue
        Log "Service réparé : $svc"
    }
    catch { FlagError "Erreur service $svc : $($_.Exception.Message)" }
}

# ---------------------------------------------------------
# 5. Réparation AppRepository (base EDB)
# ---------------------------------------------------------
Log "Réparation AppRepository"
try {
    $repo = "C:\ProgramData\Microsoft\Windows\AppRepository\StateRepository-Machine.srd"
    if (Test-Path $repo) {
        esentutl /p $repo | Out-Null
        Log "AppRepository réparé via esentutl"
    }
}
catch { FlagError "Erreur AppRepository : $($_.Exception.Message)" }

# ---------------------------------------------------------
# 6. Réparation WindowsApps (ACL)
# ---------------------------------------------------------
Log "Réparation ACL WindowsApps"
try {
    icacls "C:\Program Files\WindowsApps" /reset /t | Out-Null
}
catch { FlagError "Erreur ACL WindowsApps : $($_.Exception.Message)" }

# ---------------------------------------------------------
# 7. Réenregistrement UWP (gestion InstallLocation null)
# ---------------------------------------------------------
Log "Réenregistrement UWP"
try {
    Get-AppxPackage -AllUsers | ForEach-Object {
        try {
            if ($_.InstallLocation -and (Test-Path $_.InstallLocation)) {
                $manifest = Join-Path $_.InstallLocation "AppxManifest.xml"
                if (Test-Path $manifest) {
                    Add-AppxPackage -DisableDevelopmentMode -Register $manifest -ErrorAction SilentlyContinue
                }
            }
            else {
                FlagError "Package UWP corrompu détecté : $($_.Name)"
            }
        }
        catch { FlagError "Erreur UWP ($($_.Name)) : $($_.Exception.Message)" }
    }
}
catch { FlagError "Erreur globale UWP : $($_.Exception.Message)" }

# ---------------------------------------------------------
# 8. Réparation Microsoft Store
# ---------------------------------------------------------
Log "Réparation Microsoft Store"
try {
    Get-AppxPackage -AllUsers Microsoft.WindowsStore | ForEach-Object {
        try {
            if ($_.InstallLocation) {
                $manifest = Join-Path $_.InstallLocation "AppxManifest.xml"
                Add-AppxPackage -DisableDevelopmentMode -Register $manifest -ErrorAction SilentlyContinue
            }
            else {
                FlagError "Store sans InstallLocation (corrompu)"
            }
        }
        catch { FlagError "Erreur Store : $($_.Exception.Message)" }
    }
}
catch { FlagError "Erreur globale Store : $($_.Exception.Message)" }

# ---------------------------------------------------------
# 9. Frameworks UWP
# ---------------------------------------------------------
$frameworks = @("Microsoft.NET.Native*", "Microsoft.VCLibs*", "Microsoft.UI.Xaml*")

foreach ($fw in $frameworks) {
    Log "Réenregistrement Framework : $fw"
    try {
        Get-AppxPackage -AllUsers $fw | ForEach-Object {
            try {
                if ($_.InstallLocation) {
                    $manifest = Join-Path $_.InstallLocation "AppxManifest.xml"
                    Add-AppxPackage -DisableDevelopmentMode -Register $manifest -ErrorAction SilentlyContinue
                }
                else {
                    FlagError "Framework corrompu : $($_.Name)"
                }
            }
            catch { FlagError "Erreur framework ($($_.Name)) : $($_.Exception.Message)" }
        }
    }
    catch { FlagError "Erreur globale framework $fw : $($_.Exception.Message)" }
}

# ---------------------------------------------------------
# 10. Nettoyage final + sortie
# ---------------------------------------------------------
Log "===== Fin remédiation (redémarrage manuel requis) ====="

if ($ErrorFound) {
    Log "Sortie script : 1 (erreurs détectées pendant la remédiation)"
    exit 1
}
else {
    Log "Sortie script : 0 (remédiation exécutée sans erreur bloquante)"
    exit 0
}
