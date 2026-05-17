# Intune Remediation Scripts

Ce repertoire regroupe les scripts PowerShell Intune par domaine fonctionnel et par scenario.

## Organisation

- `WindowsUpdate/` : scripts lies a Windows Update, scan, cache, services, politiques et journaux.
- `DeliveryOptimization/` : scripts lies a Delivery Optimization et au Content Engine.
- `WUfB/` : identite Windows Update for Business et correction du binding PolicyState.
- `SetupDiag-Upgrade/` : diagnostics SetupDiag, upgrade Windows 10/11 et prerequis de mise a niveau.
- `Intune-MDM/` : etat Intune, MDM, Hybrid Join, IME et Data Boundary.
- `Disk-Cleanup-Storage/` : espace disque, nettoyage et dossiers candidats de cleanup.
- `UWP-Store-AppRepository/` : sante AppRepository / Store / Windows Update.
- `Endpoint-Repair/` : reparations endpoint plus larges, notamment WinRM et UWP Store.
- `TimeZone-Recovery/` : scripts de timezone et recovery check.
- `Standalone/` : scripts utilitaires ou sans paire detection/remediation claire.

## Convention

Les scripts actifs sont places dans les dossiers de scenario. Les noms de fichiers ne contiennent pas de numero de version; la version courante est indiquee dans l'en-tete du script et repart a `1.0`.

Quand un scenario contient une paire Intune, le dossier contient generalement :

- un script `SmartM365-*-Detection.ps1` ;
- un script `SmartM365-*-Remediation.ps1`.

## Scenarios consolides

- `WindowsUpdate/WindowsUpdate-Reset/` remplace les anciens scenarios Autopatch `0x80244007` et reset force Windows Update.
- `WindowsUpdate/Policy-Blockers/` regroupe les anciens scenarios WSUS/GPO, WSUS remnants, NoAutoUpdate, Policy-Blocking-Access et WUfB configuration blockers.
- `WindowsUpdate/Service-Refresh/` remplace les anciens scenarios Connectivity et Scan-Health.
- `Intune-MDM/MDM-Enrollment-Repair/` regroupe Device-Stale-Join, Enrollment-State, Hybrid-Join et MDM-Tasks-Missing. La detection cible l'etat local stale/casse; le fait que le script s'execute via Intune prouve deja une joignabilite IME minimale.
- `SetupDiag-Upgrade/Upgrade-Staging-Health/` remplace Upgrade-Files-Missing et Upgrade-Residues.
- `Disk-Cleanup-Storage/Upgrade-Storage-Readiness/` regroupe les anciens scenarios free space, cleanup candidates et force disk cleanup.
- `SetupDiag-Upgrade/Upgrade-Diagnostics/` documente la chaine de diagnostic entre WindowsUpdateLog, SetupDiag-Required et Upgrade-Blocking-Issues; ces scripts restent separes car ils sont complementaires.

## Verification

Apres reorganisation, les archives ont ete supprimees et les scripts actifs restent dans leurs dossiers de scenario.
