# Procédure de migration Windows 11 (Smart Endpoint Migration Toolkit)

Version anglaise : [Windows11-Procedure-EN.md](Windows11-Procedure-EN.md). Les deux versions
évoluent ensemble.

## 1. Préparer les médias d'installation

1. Extraire des médias Windows 11 sous licence, un dossier par langue, sur un partage lisible
   par les **comptes ordinateur** des postes cibles (le script endpoint tourne en SYSTEM), par
   exemple `\\files-paris\Windows11\24H2\FR-fr`.
2. Générer le manifeste SHA256 de chaque dossier média (depuis la racine du package) :
   `New-SetupMediaManifest.cmd` (dossier local `Campaigns\Windows11Upgrade\SetupSource`) ou
   `Campaigns\Windows11Upgrade\Scripts\SmartM365-EndpointMigration-New-SetupMediaManifest.ps1 -MediaRoot <dossier> -Force`.
   Le manifeste `SmartM365-SetupMediaManifest.sha256.csv` reste à la racine du dossier média.
3. Plusieurs sites : indiquer plusieurs sources UNC dans `Windows11Upgrade.SetupSourcePath`
   (séparées par `;`), ou mieux un CSV de correspondance dans `Windows11Upgrade.SetupSourceMapPath`
   (modèle `Campaigns\Windows11Upgrade\Templates\SetupSourceMap.csv.template` ; portées
   `Subnet`/`CIDR`, `IPPrefix`, `ComputerName`, `ComputerPrefix`, `Default`). Chaque poste
   choisit la source valide la plus proche.
4. Portes de copie facultatives : `SetupSourceConcurrencyLimit` avec
   `SetupSourceConcurrencyGateRoot`, et `SetupSubnetConcurrencyLimit` (1 copie par sous-réseau
   par défaut) avec `SetupSubnetConcurrencyGateRoot`. Les racines de porte sont des dossiers
   UNC où les comptes ordinateur ont le droit Modification (voir
   `Campaigns\Windows11Upgrade\SetupSourceGates\README.md`).

## 2. Configurer

Dans la GUI, sélectionner **Windows 11 upgrade**, puis l'onglet **Options** :

| Groupe | Options principales |
| --- | --- |
| Mode | Dry run, audit distant seul, run guard |
| Actions | Réparation des stratégies, reset WU, upgrade forcé WU, upgrade par setup, redémarrages, réparation de profils, nettoyage disque |
| Setup | Source ou CSV de correspondance, mode d'exécution (`LocalCache` par défaut, `Share`, `Auto`), identifiant de média, langue (`MatchSystem`), Dynamic Update, réutilisation des médias existants, duplication |
| Parallélisme / Sécurité | Workers par Wave, plafond partagé, délai PsExec, run guard, espace disque minimum (32 Go) |

Les valeurs par défaut peuvent aussi être posées une fois dans
`Config\SmartEndpointMigrationToolkit.local.json.txt`, et par Wave dans `Wave.local.json.txt` (par
exemple une source locale au site pour chaque Wave).

Redémarrages : *Allow controlled reboots* (activé par défaut) redémarre pour un redémarrage en
attente ou la fin du setup uniquement si aucun utilisateur n'est connecté, **sauf** si l'uptime
dépasse `ForceRequiredRebootWhenUptimeOverDays` (7 jours par défaut ; 0 = jamais). Avec
DirectSetup et *Allow controlled reboots* désactivé, aucun redémarrage n'est jamais demandé.

## 3. Préparer la Wave

- **Automatic Wave** : rafraîchir les inventaires partagés AD et Intune, choisir la source
  (`Both`, `AD`, `Intune`), le domaine AD (l'inventaire partagé couvre toute la forêt ;
  *(All domains)* par défaut), les préfixes, l'ancienneté de logon AD et, si besoin,
  *Exclude devices present in Intune*, cliquer **Preview**, contrôler la sélection et les
  exclusions, puis **Create Wave**.
- **New Wave** : coller les noms de postes, **Create Wave**.
- Les postes déjà en Windows 11 dans AD ou Intune sont ignorés à chaque cycle et archivés dans
  `Computers.AlreadyWindows11.txt` ; le matériel confirmé incompatible part dans
  `Computers.HardwareNotCapable.txt`.

## 4. Qualifier, puis lancer

1. **Dry run** (Once) : DNS, ping, SMB 445 et partage administratif pour chaque poste. Rien
   n'est copié.
2. **Audit distant seul** (Once) : préparation, matériel, disque, médias et accès aux sources,
   sans modification.
3. **Réparation** (Loop) : les actions activées s'exécutent ; le média est copié par le poste
   lui-même depuis la source choisie (pas par le poste technicien), validé, puis le setup
   Windows démarre.
4. Suivre l'onglet **Runs** : rapport HTML en direct, statut et prochaine action par poste,
   preuves dans `CentralLogs`.

## 5. Médias laissés par l'ancien outil Windows 11

Avant toute copie volumineuse, la même résolution est utilisée pour les Runs PsExec et pour le
package Intune :

- média valide dans le cache de ce produit : utilisé ;
- média de l'ancien outil détenu par un package Intune Win32 de l'ancien outil encore installé :
  réutilisé **sur place**, en lecture seule, externe à ce produit (jamais déplacé, nettoyé, ni
  modifié dans ses permissions ou ses marqueurs) ;
- média ancien valide, inactif et non géré : **déplacé** sur le même volume vers le nouveau cache ;
  en cas d'échec il revient à son emplacement avec ses permissions réappliquées et vérifiées
  (`RolledBack`), ou l'échec est signalé explicitement (`RollbackPartial`, `RollbackFailed`) sans
  perte de données ;
- média utilisé ou dont l'état est illisible (y compris permissions ou arborescence illisibles) :
  rien n'est modifié (`SETUP_MEDIA_IN_USE`, `SETUP_MEDIA_STATE_INDETERMINATE`) ;
- média ancien présent mais modifiable par un non-administrateur : jamais utilisé, et ses
  permissions ne sont jamais modifiées (`SETUP_MEDIA_LEGACY_UNTRUSTED`) ;
- sinon (absent, invalide, autre langue ou autre build) : copie normale. Une seconde copie à côté
  d'un média ancien restant (valide ou non fiable) exige `AllowSetupMediaDuplication` (sinon
  `SETUP_MEDIA_DUPLICATION_BLOCKED` ou `SETUP_MEDIA_LEGACY_UNTRUSTED`).

Les scripts lancés en SYSTEM (endpoint lancé par PsExec, tâche de reprise après redémarrage,
runner Intune) sont stockés sous `C:\Program Files\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade` ;
un dossier programme non protégé arrête l'opération (`SEMT_PROGRAM_FOLDER_UNTRUSTED`).

Preuves de chaque adoption : `<DataRoot>\State\MediaAdoption\`. Détails :
[LEGACY-COEXISTENCE.md](LEGACY-COEXISTENCE.md).

## 6. Arrêter

- **Arrêt contrôlé** (onglet Runs ou *Stop running Waves*) : aucun nouveau poste ne démarre.
- Une seconde demande **force l'arrêt local** des workers et de PsExec.
- Un poste peut encore exécuter le setup Windows : vérifier sa ligne
  (`CANCELLED_BY_OPERATOR_REMOTE_RUNNING` / `_REMOTE_STATE_UNCONFIRMED`) avant de relancer.

## 7. Alternative Intune Win32

Pour les postes hors de portée PsExec, construire les packages Intune Win32 par langue
(`Campaigns\Windows11Upgrade\IntuneWin32\README.md`). Utiliser le package WithCacheOnly pour les postes qui disposent déjà d'un média réutilisable (aucun média téléchargé) et le package WithMedia sinon (Intune télécharge ses médias avant l'exécution de l'installateur). Respecter l'ordre de bascule des
assignations décrit dans ce README pour remplacer une assignation du package de l'ancien
outil : les deux packages ne doivent jamais cibler les mêmes postes en même temps.

## 8. Statuts courants

| Statut | Signification / action |
| --- | --- |
| `ALREADY_WINDOWS11` | Ignoré, archivé |
| `WINDOWS11_HARDWARE_NOT_CAPABLE` | Matériel confirmé incompatible ; archivé |
| `SETUP_UPGRADE_STARTED`, `SETUP_UPGRADE_REBOOT_*` | Setup en cours ou en attente de redémarrage ; le poste bloque l'autre Campaign jusqu'au redémarrage (168 h maximum) |
| `SETUP_MEDIA_DUPLICATION_BLOCKED` | Média de l'ancien outil présent et non réutilisable ; désassigner l'app Intune de l'ancien outil, libérer le média ou autoriser la duplication |
| `SETUP_MEDIA_LEGACY_UNTRUSTED` | Média de l'ancien outil présent mais modifiable par un non-administrateur (ou contenant un point de réanalyse) ; non utilisé, permissions non modifiées. Le sécuriser ou le supprimer (décision d'administrateur), ou autoriser la duplication |
| `SETUP_MEDIA_STATE_INDETERMINATE` | Propriétaire, utilisation, permissions ou arborescence du média illisibles ; rien n'est modifié. Vérifier le poste puis relancer |
| `TARGET_OPERATION_ACTIVE` | Une autre opération de l'outil tourne sur le poste ; cycle suivant |
| `LEGACY_*` | Voir [LEGACY-COEXISTENCE.md](LEGACY-COEXISTENCE.md) |
