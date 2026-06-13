# Intune Remediation Scripts

Ce repertoire regroupe les scripts PowerShell Intune par domaine fonctionnel et par scenario. Il fait partie de `SmartIntuneRemediation`, avec le manager CLI/GUI qui permet de publier et gerer ces packages.

## Organisation

- `WindowsUpdate/` : scripts lies a Windows Update, scan, cache, services, politiques et proxy.
- `DeliveryOptimization/` : scripts lies a Delivery Optimization et au Content Engine.
- `WUfB/` : identite Windows Update for Business et correction du binding PolicyState.
- `SetupDiag-Upgrade/` : diagnostics SetupDiag, upgrade Windows 10/11 et prerequis de mise a niveau.
- `Intune-MDM/` : etat Intune, MDM, Hybrid Join, IME et Data Boundary.
- `Disk-Cleanup-Storage/` : espace disque, nettoyage et dossiers candidats de cleanup.
- `UWP-Store-AppRepository/` : sante AppRepository / Store / Windows Update.
- `Standalone/` : scripts utilitaires, actions ponctuelles ou diagnostics manuels qui ne doivent pas tourner en remediation recurrente.

## Convention

Les scripts actifs sont places dans les dossiers de scenario. Les noms de fichiers ne contiennent pas de numero de version; la version courante est indiquee dans l'en-tete du script et repart a `1.0`.

Quand un scenario contient une paire Intune, le dossier contient generalement :

- un script `SmartM365-*-Detection.ps1` ;
- un script `SmartM365-*-Remediation.ps1`.

## Scenarios consolides

- `WindowsUpdate/WindowsUpdate-Reset/` remplace les anciens scenarios Autopatch `0x80244007` et reset force Windows Update.
- `WindowsUpdate/Policy-Blockers/` regroupe les anciens scenarios WSUS/GPO, WSUS remnants, NoAutoUpdate, Policy-Blocking-Access et WUfB configuration blockers.
- `WindowsUpdate/Service-And-Scan-Health/` remplace `Service-Health`, `Service-Refresh` et `ForceWUScan`.
- `WindowsUpdate/Cache-Health/` remplace `Download-Failure`.
- `Intune-MDM/MDM-Enrollment-Repair/` regroupe Device-Stale-Join, Enrollment-State, Hybrid-Join et MDM-Tasks-Missing. La detection cible l'etat local stale/casse; le fait que le script s'execute via Intune prouve deja une joignabilite IME minimale.
- `SetupDiag-Upgrade/Upgrade-Staging-Health/` remplace Upgrade-Files-Missing et Upgrade-Residues.
- `Disk-Cleanup-Storage/Upgrade-Storage-Readiness/` regroupe les anciens scenarios free space, cleanup candidates et force disk cleanup.
- `DeliveryOptimization/ContentEngine-Health/` remplace `DeliveryOptimization/DO-Issues`.
- `UWP-Store-AppRepository/WU-Health/` remplace l'ancien scenario large `Endpoint-Repair/WinRM-UWP-Store`.
- `SetupDiag-Upgrade/Upgrade-Diagnostics/` regroupe `SetupDiag-Required` et `Upgrade-Blocking-Issues`.
- `Standalone/WindowsUpdateLog/`, `Standalone/Repair-DISM/`, `Standalone/Upgrade-Actions/`, `Standalone/Network-Diagnostics/`, `Standalone/WindowsPolicy/` et `Standalone/Device-Recovery/` conservent les actions ponctuelles ou diagnostics manuels deplaces hors des remediations recurrentes.

## Verification

Apres reorganisation, les archives ont ete supprimees et les scripts actifs restent dans leurs dossiers de scenario.
