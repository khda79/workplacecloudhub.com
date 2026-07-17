# Smart Exchange Migration Readiness

Application autonome PowerShell 7 / WPF de prévalidation en lecture seule des batches de migration Exchange hybride vers Exchange Online.

Elle charge un CSV de boîtes aux lettres, exécute les contrôles de préparation et produit un verdict par boîte : `GO`, `GO-WARNING`, `NO-GO` ou `UNKNOWN`.

## Modes de données

Le mode est sélectionnable directement dans l’interface.

### Live — mode par défaut

- Exchange Online : connexion interactive déléguée lancée automatiquement par Run assessment.
- Microsoft Graph : connexion interactive déléguée standard lancée automatiquement après Exchange Online.
- Active Directory : tentative live automatique ; fallback sur `AD_Users_AllDomains.csv` si AD est indisponible.
- Exchange on-premises / Exchange 2016 : utilisation des cmdlets live lorsqu’elles sont disponibles ; fallback sur `Exchange_OnPrem_Mailboxes_AllDomains.csv` sinon.
- Santé Microsoft Entra Connect : lecture live des cmdlets ADSync locales lorsqu’elles sont disponibles ; fallback sur `M365_Entra_AzureADConnect_SyncHealth.csv` sinon.
- Endpoint de migration et santé Entra Connect : contrôlés automatiquement pendant l’évaluation.

Les inventaires AD, Exchange 2016 et Entra Connect sont préchargés comme sources de secours. Sur une machine sans accès on-premises, la bascule CSV est automatique et la source réellement utilisée apparaît dans les findings et dans l’onglet Activity. L’onglet Sources est informatif : il n’expose plus de boutons de connexion séparés.

### CacheOnly

Aucune connexion EXO, Graph, AD, Exchange on-premises ou Entra Connect n’est utilisée. Toutes les preuves disponibles sont lues dans :

```text
\\server\share\WORKPLACE\DATA\Tenants\<TenantKey>
```

Le chargeur accepte les CSV directement dans ce dossier ou dans son sous-dossier `DATA-LAST`.

Les contrôles qui nécessitent obligatoirement une interrogation live, notamment le test d’endpoint et la recherche des mailboxes soft-deleted/inactive si le cache ne les expose pas, sont signalés `UNKNOWN` avec une recommandation de validation finale en mode Live. Ils ne sont jamais convertis silencieusement en `PASS`.

## Autonomie

L’application possède son propre fichier de configuration, son moteur de contrôles et ses exports. Elle ne charge aucun script ni fichier de configuration du projet SmartM365.

Les CSV du cache peuvent être produits par SmartM365, mais l’application les consomme directement et reste exécutable indépendamment.

Le profil opérationnel est défini dans le JSON local ignoré par Git avec :

- `TenantProfile.TenantId` : garde-fou pour empêcher une authentification live dans le mauvais tenant ; ce paramètre est nécessaire.
- `TenantProfile.RemoteRoutingDomain` : domaine de routage hybride attendu.
- `Cache.RootPath` : racine des inventaires CSV.
- `Cache.MaximumAgeHours` : âge maximal accepté pour les données.

Il n’existe plus de bloc générique `Tenant`, de paramètre `MicrosoftGraph.UseDeviceCode`, de bloc `OnPremises`, ni de paramètre `EntraConnect.Server`.

Graph utilise `Connect-MgGraph` en authentification interactive standard avec le `TenantId` configuré. Aucun flux device code n’est demandé.

## Sécurité

L’application est strictement diagnostique. Elle ne crée pas de batch, ne modifie aucun destinataire, attribut AD, licence, hold, permission ou objet de migration.

Les identifiants et jetons ne sont pas enregistrés. Le fichier JSON local, les CSV d’entrée et les rapports opérationnels sont exclus de Git.

## Prérequis

Pour tous les modes :

- Windows.
- PowerShell 7 disponible dans `C:\Program Files\PowerShell\7\pwsh.exe` pour le lanceur fourni.
- Accès en lecture au chemin du cache lorsque des preuves mises en cache sont nécessaires.

Uniquement pour le mode Live :

- module `ExchangeOnlineManagement` 3.0 ou ultérieur ;
- modules `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users` et `Microsoft.Graph.Identity.DirectoryManagement` ;
- compte disposant des droits de lecture nécessaires dans Exchange Online et Microsoft Graph ;
- pour les sources on-premises live, modules/cmdlets `ActiveDirectory`, Exchange Management Shell et/ou `ADSync` disponibles sur la machine d’exécution. Leur absence n’est pas bloquante si les CSV de fallback sont accessibles et suffisamment récents.

Scopes Graph par défaut :

- `User.Read.All`
- `Directory.Read.All`
- `Organization.Read.All`

## Démarrage

```text
Start-SmartM365-ExchangeMigrationReadiness-GUI.cmd
```

Le lanceur CMD démarre PowerShell 7 dans un processus détaché et masqué, puis se ferme immédiatement.

Validation statique sans afficher l’interface :

```powershell
pwsh -NoLogo -NoProfile -STA -File .\SmartM365-ExchangeMigrationReadiness-GUI.ps1 -ValidateOnly
```

## Configuration

Modèle versionné :

```text
Config\SmartM365-ExchangeMigrationReadiness.local.json.template
```

Fichier local créé automatiquement au premier lancement et exclu de Git :

```text
Config\SmartM365-ExchangeMigrationReadiness.local.json
```

Le modèle contient uniquement des valeurs neutres. Renseignez les valeurs du tenant et du cache dans le fichier local ; elles ne doivent pas être versionnées.

Clés principales :

- `Mode` : `Live` par défaut ou `CacheOnly` ; le sélecteur GUI s’applique à l’exécution courante.
- `TenantProfile` : profil tenant autonome de l’application.
- `Cache.RootPath` et `Cache.MaximumAgeHours` : emplacement et fraîcheur des inventaires.
- `ExchangeOnline.UserPrincipalName` : UPN administrateur optionnel.
- `ExchangeOnline.DisableWam` : conserve le flux interactif EXO sans WAM lorsque requis sur ce poste.
- `Hybrid.MigrationEndpointName` : endpoint `ExchangeRemoteMove` explicite optionnel.
- `Hybrid.TargetDeliveryDomain` : domaine `tenant.mail.onmicrosoft.com` attendu.
- `DefaultTargetSku`, `TargetQuotaGbBySku` et `QuotaSafetyBufferPercent` : politique de quota cible.
- `OutputRoot` : dossier d’export, absolu ou relatif à l’application.

## Inventaires CSV attendus

Selon le mode et les contrôles disponibles :

- `AD_Users_AllDomains.csv`
- `Exchange_OnPrem_Mailboxes_AllDomains.csv`
- `M365_Entra_AzureADConnect_SyncHealth.csv`
- `M365_Users_Active.csv`
- `Exchange_EXO_Mailboxes_AllDomains.csv`
- `Exchange_EXO_MigrationJobs.csv`
- `M365_Licenses_Tenant.csv`

L’application vérifie la présence et l’âge des fichiers avant de les considérer utilisables.

## CSV de migration

La colonne canonique est `EmailAddress`. Sont également acceptées : `PrimarySmtp`, `PrimarySmtpAddress`, `UserPrincipalName`, `UPN` ou `Mailbox`. Le fichier est chargé et validé immédiatement après sa sélection avec Browse.

Colonnes optionnelles interprétées :

- `MailboxType`
- `TargetSku` ou `TargetSkuPartNumber`
- `BadItemLimit`
- `LargeItemLimit`

Les délimiteurs virgule, point-virgule et tabulation sont détectés automatiquement. L’import essaie d’abord UTF-8 strict, puis Windows-1252.

## Principaux contrôles

- intégrité du CSV, adresses vides, syntaxe SMTP et doublons ;
- existence/unicité du compte AD et statut du compte ;
- état UserMailbox / RemoteMailbox / MailUser ;
- cohérence Primary SMTP, proxyAddresses et targetAddress ;
- taille de mailbox contre quota cible avec marge de sécurité ;
- Litigation Hold et In-Place Hold quand les propriétés sont disponibles ;
- baseline Full Access, Send As et Send on Behalf ;
- état MailUser/mailbox Exchange Online et détection split-brain ;
- conflits soft-deleted/inactive ;
- migration user ou move request existant ;
- unicité et synchronisation de l’utilisateur Entra ;
- licence actuelle, UsageLocation et capacité du SKU cible ;
- santé de synchronisation Entra Connect depuis le cache ;
- endpoint `ExchangeRemoteMove` en Live ;
- avertissement documenté pour `CannotMoveEnhancedRestoreMailboxesCrossOrgPermanentException`.

Une source obligatoire absente reste bloquante. Une propriété non disponible dans le cache devient `UNKNOWN`, jamais un faux `PASS`.

## Rapports

Chaque exécution crée :

```text
Output\SEMR-yyyyMMdd-HHmmss\
  Summary.csv
  Findings.csv
  Permissions-Baseline.csv
  Evidence.csv
```

`Summary.csv` contient un verdict par mailbox. `Findings.csv` contient le détail de chaque contrôle : sévérité, résultat, caractère bloquant, valeur observée, valeur attendue, source, message et action recommandée.

Après migration, la comparaison de permissions en mode Live peut aussi générer :

```text
Permissions-Current-EXO.csv
Permissions-Comparison.csv
```