<#
    Name: SmartM365-DO-Issues-Detection.ps1
    Version: 1.0
# Detection: Delivery Optimization / Dynamic Download Drift V2

ne considère PAS l’absence du dossier comme une erreur
détecte les vrais problèmes DO
détecte les erreurs 0x80D0xxxx
détecte les erreurs DynamicDownload
détecte les erreurs BITS
détecte les erreurs Windows Update liées à DO
#>


Write-Output "=== Delivery Optimization / Content Engine Diagnostic ==="
# Version 2
# ================================
# 1. Vérification du dossier DO
# ================================
$DOFolder = "C:\ProgramData\Microsoft\Windows\DeliveryOptimization"

$DOFolderStatus = "OK"

if (!(Test-Path $DOFolder)) {
    Write-Output "[INFO] Le dossier DeliveryOptimization n'existe pas → DO n'a encore rien téléchargé."
    $DOFolderStatus = "NotCreated"
}
else {
    $size = (Get-ChildItem $DOFolder -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum

    if ($size -lt 1024) {
        Write-Output "[ALERTE] Le cache DO est présent mais vide → DO ne télécharge rien."
        $DOFolderStatus = "Empty"
    }
    else {
        Write-Output "[OK] Dossier DO présent, taille : $([math]::Round($size/1MB,2)) MB"
    }
}

# ================================
# 2. Lecture du journal Windows Update
# ================================
$WU = Get-WinEvent -LogName "Microsoft-Windows-WindowsUpdateClient/Operational" -MaxEvents 200 -ErrorAction SilentlyContinue

$WU_Errors = $WU | Where-Object {
    $_.Message -match "DynamicDownload" -or
    $_.Message -match "Delivery Optimization" -or
    $_.Message -match "Content" -or
    $_.Message -match "0x80070002" -or
    $_.Message -match "8024000C"
}

# ================================
# 3. Lecture du journal Delivery Optimization
# ================================
$DO = Get-WinEvent -LogName "Microsoft-Windows-DeliveryOptimization/Operational" -MaxEvents 200 -ErrorAction SilentlyContinue

$DO_Errors = $DO | Where-Object {
    $_.Message -match "0x80D0" -or
    $_.Message -match "error" -or
    $_.Message -match "fail" -or
    $_.Message -match "timeout" -or
    $_.Message -match "proxy" -or
    $_.Message -match "download"
}

# ================================
# 4. Lecture du journal BITS
# ================================
$BITS = Get-WinEvent -LogName "Microsoft-Windows-Bits-Client/Operational" -MaxEvents 200 -ErrorAction SilentlyContinue

$BITS_Errors = $BITS | Where-Object {
    $_.Message -match "0x80D0" -or
    $_.Message -match "error" -or
    $_.Message -match "fail" -or
    $_.Message -match "timeout"
}

# ================================
# 5. Synthèse
# ================================
$HasErrors = $false

if ($WU_Errors) {
    Write-Output "`n[WU] Erreurs détectées :"
    $WU_Errors | Select-Object TimeCreated, Id, Message
    $HasErrors = $true
}

if ($DO_Errors) {
    Write-Output "`n[DO] Erreurs Delivery Optimization détectées :"
    $DO_Errors | Select-Object TimeCreated, Id, Message
    $HasErrors = $true
}

if ($BITS_Errors) {
    Write-Output "`n[BITS] Erreurs BITS liées à DO :"
    $BITS_Errors | Select-Object TimeCreated, Id, Message
    $HasErrors = $true
}

if ($DOFolderStatus -eq "Empty") {
    Write-Output "`n[DOSCAN] Cache DO vide → DO ne télécharge rien."
    $HasErrors = $true
}

# ================================
# 6. Résultat final
# ================================
if ($HasErrors) {
    Write-Output "`n=== RESULTAT : Delivery Optimization / Content Engine en ERREUR ==="
    exit 1
}
else {
    Write-Output "`n=== RESULTAT : Delivery Optimization OK ==="
    exit 0
}
