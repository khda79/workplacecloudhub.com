<#
    Name: SmartM365-DeviceOps-Get-WindowsUpdateLog-Remediation.ps1
    Version: 1.0 
# Get-WindowsUpdateLog
- Génère le log Windows Update, 
- Crée un nom de fichier propre
- Copie vers un share réseau
- Récupère les erreurs 
- Transfère le fichier vers C:\ProgramData\SmartM365\IntuneRemediation\Temp
#>
# ================================
# Variables
# ================================
$LocalFolder = "C:\ProgramData\SmartM365\IntuneRemediation\Temp"
$ComputerName = $env:COMPUTERNAME
$Date = (Get-Date -Format "yyyy-MM-dd_HH-mm-ss")
$OutputFile = "WindowsUpdate-$Date-$ComputerName.log"
$LocalPath = Join-Path -Path $LocalFolder -ChildPath $OutputFile

# ================================
# Vérification du dossier local
# ================================
if (!(Test-Path $LocalFolder)) {
    try {
        New-Item -Path $LocalFolder -ItemType Directory -Force | Out-Null
    }
    catch {
        Write-Output "Impossible de créer le dossier local : $LocalFolder"
        exit 1
    }
}

# ================================
# Génération du WindowsUpdate.log
# ================================
try {
    Write-Output "Génération du WindowsUpdate.log..."
    Get-WindowsUpdateLog -LogPath $LocalPath
}
catch {
    Write-Output "Erreur lors de la génération du WindowsUpdate.log : $_"
    exit 1
}

Write-Output "Log généré : $LocalPath"
exit 0
