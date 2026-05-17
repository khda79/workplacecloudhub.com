<#
    Name: SmartM365-Endpoint-Repair-WinRM-UWP-Store-Detection.ps1
    Version: 1.0 
    Script de détection Intune (Windows ne se mets pas à niveau car MS Store et SoftwareDistribution sont corrompus ou indisponibles (erreur 403).
	
	Ce script vérifie les points critiques :

- WindowsApps :
dossier présent mais illisible
dossier présent mais trop petit
dossier présent mais trop peu de packages
permissions cassées

- AppRepository :
base EDB corrompue
fichiers manquants
dossier lisible mais incomplet

- Microsoft Store :
Store manquant
Store non enregistré

- Frameworks UWP :
.NET Native Runtime
.NET Native Framework
VCLibs
UI.Xaml

- Services critiques UWP :
AppXSvc
ClipSVC
StateRepository
WpnService

- Windows Update :
SoftwareDistribution vide (corruption)
erreurs AppX dans l’EventLog

- WinRM :
service arrêté ou absent
#>

# Détection avancée Intune – UWP / Store / AppRepository / WindowsApps / WinRM / WU
$ErrorFound = $false

function Flag($msg) {
    Write-Output "Issue: $msg"
    $global:ErrorFound = $true
}

# ---------------------------------------------------------
# 1. Vérification WindowsApps (taille + lisibilité + permissions)
# ---------------------------------------------------------
$waPath = "C:\Program Files\WindowsApps"

if (-not (Test-Path $waPath)) {
    Flag "WindowsApps absent"
}
else {
    try {
        $waItems = Get-ChildItem $waPath -ErrorAction Stop
        $waSize = ($waItems | Measure-Object -Property Length -Sum).Sum

        if ($waItems.Count -lt 50) {
            Flag "WindowsApps contient trop peu de packages (corruption probable)"
        }

        if ($waSize -lt 500MB) {
            Flag "WindowsApps taille anormale (<500MB)"
        }
    }
    catch {
        Flag "WindowsApps illisible ou permissions cassées"
    }
}

# ---------------------------------------------------------
# 2. Vérification AppRepository (taille + lisibilité)
# ---------------------------------------------------------
$arPath = "C:\ProgramData\Microsoft\Windows\AppRepository"

if (-not (Test-Path $arPath)) {
    Flag "AppRepository absent"
}
else {
    try {
        $arItems = Get-ChildItem $arPath -ErrorAction Stop
        if ($arItems.Count -lt 20) {
            Flag "AppRepository contient trop peu de fichiers (corruption probable)"
        }
    }
    catch {
        Flag "AppRepository illisible ou corrompu"
    }
}

# ---------------------------------------------------------
# 3. Vérification des services critiques UWP
# ---------------------------------------------------------
$services = @("AppXSvc", "ClipSVC", "StateRepository", "WpnService")

foreach ($svc in $services) {
    $s = Get-Service -Name $svc -ErrorAction SilentlyContinue
    if ($s -and $s.Status -ne "Running") {
        Flag "Service critique UWP non démarré : $svc"
    }
}

# ---------------------------------------------------------
# 4. Vérification Microsoft Store installé
# ---------------------------------------------------------
$store = Get-AppxPackage -AllUsers Microsoft.WindowsStore -ErrorAction SilentlyContinue
if (-not $store) {
    Flag "Microsoft Store non installé ou corrompu"
}

# ---------------------------------------------------------
# 5. Vérification Frameworks UWP essentiels
# ---------------------------------------------------------
$frameworks = @(
    "Microsoft.NET.Native.Runtime",
    "Microsoft.NET.Native.Framework",
    "Microsoft.VCLibs",
    "Microsoft.UI.Xaml"
)

foreach ($fw in $frameworks) {
    $pkg = Get-AppxPackage -AllUsers | Where-Object { $_.Name -like "$fw*" }
    if (-not $pkg) {
        Flag "Framework UWP manquant : $fw"
    }
}

# ---------------------------------------------------------
# 6. Vérification SoftwareDistribution (corruption)
# ---------------------------------------------------------
$sdPath = "C:\Windows\SoftwareDistribution\Download"
if (Test-Path $sdPath) {
    $sdSize = (Get-ChildItem $sdPath -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
    if ($sdSize -eq 0) {
        Flag "SoftwareDistribution vide (corruption probable)"
    }
}

# ---------------------------------------------------------
# 7. Vérification WinRM
# ---------------------------------------------------------
$winrm = Get-Service -Name WinRM -ErrorAction SilentlyContinue
if ($winrm.Status -ne "Running") {
    Flag "WinRM non fonctionnel"
}

# ---------------------------------------------------------
# 8. Vérification EventLog AppX / Store / UWP
# ---------------------------------------------------------
$events = Get-WinEvent -LogName "Microsoft-Windows-AppXDeploymentServer/Operational" -MaxEvents 50 -ErrorAction SilentlyContinue |
          Where-Object { $_.LevelDisplayName -eq "Error" }

if ($events.Count -gt 0) {
    Flag "Erreurs AppX détectées dans l’EventLog"
}

# ---------------------------------------------------------
# Résultat final
# ---------------------------------------------------------
if ($ErrorFound) {
    Write-Output "Issues detected"
    exit 1
}
else {
    Write-Output "OK"
    exit 0
}
