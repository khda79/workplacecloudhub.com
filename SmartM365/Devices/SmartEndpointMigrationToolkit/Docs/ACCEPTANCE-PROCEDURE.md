# Procédure de recette finale - Smart Endpoint Migration Toolkit 1.0.0

À dérouler ensemble, dans l'ordre. Chaque étape indique le résultat attendu. Les étapes 4 et
suivantes agissent sur des postes réels : uniquement des postes de **laboratoire** désignés,
jamais de production, et avec accord explicite avant chaque lancement.

## 0. Prérequis

- [ ] Poste technicien Windows 10/11, Windows PowerShell 5.1, console administrateur, PsExec64.
- [ ] 2 à 4 postes de laboratoire Windows 10 joignables (SMB 445, partage `ADMIN$`), dont au
      moins un avec un ancien outil installé (IntuneHybridJoinToolkit ou Windows11UpgradeToolkit).
- [ ] Un partage UNC de médias Windows 11 de test lisible par les comptes ordinateur.
- [ ] Revue du diff Git de la livraison (nouveau dossier + 2 correctifs critiques).

## 1. Package

- [ ] `Release\SmartM365-EndpointMigration-Build-Release.ps1` (aperçu) : validation OK, aucune
      écriture. Sans `-AllowUnsigned`, refus attendu tant que les sources ne sont pas signées.
- [ ] Build de qualification : `-Build -AllowUnsigned`. Attendu : ZIP, `.sha256`, notes
      marquées `UNSIGNED QUALIFICATION BUILD`, aucun fichier privé (Waves, Runs, `*.local.json`,
      inventaires, médias, DemoAssets, mémoires).
- [ ] Extraire le ZIP dans un dossier neuf (chemin avec espace) et travailler **depuis ce dossier**.
- [ ] Après signature des sources (pre-push), reconstruire sans `-AllowUnsigned` : manifeste
      `RequireSignature=true`.

## 2. Qualification hors ligne (depuis le package extrait ou le dépôt)

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Core\Test-SemtCore.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Tests\Core\Test-SemtCore.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\GUI\Test-SemtGui.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Packaging\Test-SemtPackaging.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\HybridJoin\Test-SemtHybridJoinAutomaticWave.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Windows11Upgrade\Test-SemtWindows11AutomaticWave.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Windows11Upgrade\Test-SemtWindows11IntuneWin32.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Windows11Upgrade\Test-SemtWindows11MediaResolution.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Windows11Upgrade\Test-SemtWindows11MediaAdoption.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Core\Test-SemtPowerShellDates.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Tests\Core\Test-SemtPowerShellDates.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Core\Test-SemtProgramFolder.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Tests\Core\Test-SemtProgramFolder.ps1
```

- [ ] Toutes les suites : `Failed=0`, code de sortie 0.
- [ ] Moteurs : GUI et endpoints en Windows PowerShell 5.1 ; moteur et outils en 5.1 ou 7.2+. Lancer
      l'endpoint ou la GUI avec `pwsh` : refus explicite (`UNSUPPORTED_POWERSHELL_HOST`,
      `SEMT_UNSUPPORTED_GUI_HOST`) sans écriture d'état.

## 3. GUI et démo

- [ ] `Start-SmartEndpointMigrationToolkit-Demo.cmd` : 6 onglets, les deux Campaigns, options
      réelles, actions simulées, aucune exécution.
- [ ] `Start-SmartEndpointMigrationToolkit-GUI.cmd` : splash, en-tête, version 1.0.0, contrôle
      de mise à jour silencieux.
- [ ] Changer de Campaign : options, Waves et Runs propres à chaque Campaign ; les choix sont
      retrouvés à la réouverture.
- [ ] **New Wave** : création sans lancement ; refus d'un nom existant ou invalide.
- [ ] **Automatic Wave** : Preview n'écrit que des preuves ; Create grisé tant que les filtres
      diffèrent de la prévisualisation ; Create ne lance rien.

## 4. Hybrid Join (laboratoire)

- [ ] Single PC, **Dry run**, Once : `DRYRUN_*` cohérents, rapport HTML, `EffectiveOptions.json`.
- [ ] Single PC, **Audit only** : diagnostic distant sans modification ; CorrelationId identique
      dans la ligne de rapport et `State\Results\<CorrelationId>.json` sur le poste.
- [ ] Single PC, réparation avec *Allow dsregcmd /leave* **décoché** : le journal endpoint montre
      que `/leave` n'est pas autorisé.
- [ ] Wave de 2 postes, Loop : cycles, run guard, archivage `Computers.AlreadyEnrolled.txt`.

## 5. Windows 11 (laboratoire)

- [ ] Média de test + manifeste SHA256 ; Wave avec source UNC.
- [ ] Dry run puis Audit : préparation, matériel, disque, source.
- [ ] Réparation avec **DirectSetup activé et *Allow controlled reboots* désactivé** : aucun
      redémarrage demandé (l'endpoint ne reçoit pas `-AllowReboot`).
- [ ] Réparation normale sur un poste : copie par le poste, validation, setup lancé, marqueur
      `TargetPendingOperation.json` présent jusqu'au redémarrage.

## 6. Coexistence avec les anciens outils

- [ ] Poste avec ancien outil installé et inactif : `LEGACY_INSTALLED_IDLE`, non bloquant.
- [ ] Ancien outil en cours d'exécution sur le poste : statut `LEGACY_*` bloquant explicite.
- [ ] Média Windows 11 de l'ancien outil, non géré par Intune : déplacé sur le même volume,
      validé, journalisé (`Moved`), preuves dans `State\MediaAdoption`.
- [ ] Adoption interrompue (par exemple média invalide après déplacement, sur un poste de
      laboratoire) : retour à l'emplacement d'origine, permissions identiques (`RolledBack`).
- [ ] Média détenu par un package Intune de l'ancien outil : réutilisé sur place, ni déplacé ni
      modifié (`LegacyIntuneOwned`, `SetupMediaSource=LegacyExternal`) ; ACL et marqueurs inchangés.
- [ ] Média non réutilisable : `SETUP_MEDIA_DUPLICATION_BLOCKED` sans option de duplication.
- [ ] Média de l'ancien outil présent mais modifiable par les utilisateurs (permissions héritées de
      `ProgramData`, à constater avec `icacls` sans les modifier) : `SETUP_MEDIA_LEGACY_UNTRUSTED`,
      aucune copie, permissions du média inchangées. Avec `AllowSetupMediaDuplication` : seconde
      copie journalisée, média ancien toujours inchangé.
- [ ] Permissions du média ancien illisibles pour SYSTEM (cas de laboratoire uniquement, préparé
      puis restauré par un administrateur) : `SETUP_MEDIA_STATE_INDETERMINATE`, rien n'est modifié,
      même avec l'option de duplication.
- [ ] Aucune configuration, LOT, inventaire ou historique de l'ancien outil n'est importé.

## 6 bis. Scripts lancés en SYSTEM (Program Files)

- [ ] Après un passage PsExec : l'endpoint est sous
      `C:\Program Files\WCH\SmartM365\SmartEndpointMigrationToolkit\<Campaign>`, propriétaire
      Administrators, `icacls` : SYSTEM et Administrators (F), Users (RX) seulement ; aucun `.ps1`
      exécuté depuis `C:\ProgramData\WCH\...`.
- [ ] Reprise après redémarrage planifiée : la tâche `SmartM365-SEMT-<Campaign>-RetryAfterReboot`
      pointe vers `...\<Campaign>\RetryAfterReboot\RetryAfterRebootRunner.ps1` ; l'état et la copie
      de l'endpoint sont dans ce même dossier.
- [ ] Intune : la tâche `SmartM365 SEMT Windows 11 Upgrade - Intune` lance
      `...\Windows11Upgrade\Intune\SmartM365-EndpointMigration-Run-IntuneUpgrade.ps1`.
- [ ] Cas de laboratoire, préparé puis nettoyé par un administrateur : un dossier
      `...\SmartEndpointMigrationToolkit` créé à l'avance avec un droit de modification pour Users
      provoque `SEMT_PROGRAM_FOLDER_UNTRUSTED` (installation Intune, copie PsExec, reprise) ;
      rien n'est installé ni exécuté et le dossier n'est pas corrigé automatiquement.
- [ ] Le compte technicien a bien le droit de fixer le propriétaire Administrators via le partage
      `C$` (sinon la copie PsExec échoue avec `SEMT_PROGRAM_FOLDER_UNTRUSTED`, sans exécution).
- [ ] Données : après un passage (PsExec ou Intune), `...\SmartEndpointMigrationToolkit` et
      `...\<Campaign>` ont le propriétaire Administrators et seulement SYSTEM (F),
      Administrators (F), Users (RX), protégés. `C:\ProgramData\WCH` et `...\WCH\SmartM365` ainsi
      que leurs autres sous-dossiers (autres produits, anciens caches réutilisés en place) gardent
      exactement leur propriétaire, leurs ACL et leur contenu (comparer `icacls` avant/après).
- [ ] Cas de laboratoire, préparé puis nettoyé par un administrateur : dossier
      `...\SmartEndpointMigrationToolkit\<Campaign>` créé à l'avance par un utilisateur standard avec
      un fichier dedans : `SEMT_DATA_FOLDER_UNTRUSTED` / `DATA_FOLDER_UNTRUSTED`, contenu refusé,
      aucune modification d'ACL, rien d'exécuté ; même résultat avec un lien dur ou une jonction ;
      `C:\ProgramData\WCH\SmartM365` modifiable par Users (suppression) : refus, dossier partagé
      inchangé.
- [ ] Intune : chaîne de données rendue non fiable entre deux cycles : le runner ne lance rien,
      passe le package en `RepairRequired` (`SEMT_DATA_FOLDER_UNTRUSTED`) et retire sa tâche ; la
      réinstallation Intune durcit de nouveau la chaîne.

## 7. Concurrence

- [ ] Une Wave Hybrid Join et une Wave Windows 11 en parallèle : le plafond partagé
      (`GlobalWorkerLimit`) est respecté pour les deux.
- [ ] Le même poste dans les deux Waves : la seconde opération reçoit `TARGET_OPERATION_ACTIVE`
      ou `TARGET_PENDING_OPERATION`.
- [ ] Relancer une Wave déjà en cours : refus (verrou de Wave).
- [ ] Deux sessions Windows / comptes différents sur le poste technicien : plafond partagé commun.
- [ ] Launch All Waves : uniquement la Campaign sélectionnée, une Wave toutes les 5 secondes.

## 8. Arrêt

- [ ] Arrêt contrôlé : aucun nouveau poste, les workers en cours terminent.
- [ ] Seconde demande : arrêt local forcé des workers et de PsExec.
- [ ] Lignes `CANCELLED_BY_OPERATOR_REMOTE_*` pour les postes encore actifs ; le Run suivant lit
      le résultat de la tentative.

## 9. Intune Win32 (sans assignation)

- [ ] Build `-SkipIntuneWinBuild` : identité `SmartM365-SEMT-Windows11Upgrade-*`, détection
      générée, pas de collision avec l'ancien package.
- [ ] Ordre de bascule des assignations relu (README IntuneWin32). Aucune assignation réelle.
- [ ] Poste de laboratoire avec média ancien détenu par l'ancien package : le package
      WithCacheOnly s'installe sans copie ni téléchargement de média ; le runner passe
      `-SetupMediaPath` vers ce média ; aucune réparation Intune en boucle.
- [ ] Poste sans média réutilisable : WithCacheOnly échoue explicitement
      (`SEMT_SETUP_MEDIA_NOT_REUSABLE`) ; WithMedia copie ses médias (téléchargement Intune préalable).
- [ ] Poste avec média ancien non protégé : l'installation échoue avec
      `SEMT_SETUP_MEDIA_LEGACY_UNTRUSTED` (WithMedia et WithCacheOnly) ; seul un package WithMedia
      construit avec `-AllowSetupMediaDuplication` copie ses médias (`SEMT_SETUP_MEDIA_DUPLICATION_ALLOWED`).
      Le runner ne démarre pas l'endpoint et indique `SEMT_MEDIA_UNTRUSTED`.

## 10. Correctifs critiques des anciens outils

- [ ] Windows11UpgradeToolkit : `Tests\Test-SmartM365Windows11DirectSetupRebootCmd.ps1` OK ;
      diff limité au CMD + README.
- [ ] IntuneHybridJoinToolkit : `Tests\Test-SmartM365IntuneHybridJoinSinglePcOptions.ps1` OK ;
      GUI 1.21 ; DemoAssets 2026-10-24.
- [ ] Décision de publication d'une nouvelle build pour chacun des deux anciens outils.

## 11. Clôture

- [ ] Revue finale du diff, décision commit / push (instruction explicite requise).
- [ ] Tag `endpoint-migration-toolkit-v1.0.0` et publication : décision séparée.
