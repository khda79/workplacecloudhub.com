# Smart Exchange Migration Readiness

Application autonome PowerShell 7 / WPF de prévalidation en lecture seule des batches de migration Exchange hybride vers Exchange Online.

Elle charge un CSV de boîtes aux lettres, exécute les contrôles de préparation et produit un verdict par boîte : `GO`, `GO-WARNING`, `NO-GO` ou `UNKNOWN`.

## Modes de données

Le mode est sélectionnable directement dans l’interface.

La phase d’évaluation est également sélectionnable dans l’interface :

- `PreCreation` — phase par défaut, avant création du batch : un move actif est bloquant et l’absence de licence cible est attendue.
- `ExistingBatch` — contrôle d’un batch déjà créé ou démarré : un move actif est attendu, son absence devient un avertissement, et une licence Exchange déjà attribuée est attendue.

La phase sélectionnée est inscrite dans `Summary.csv` afin que le verdict reste interprétable après l’exécution.

### Live — mode par défaut

- Exchange Online : connexion interactive déléguée lancée automatiquement par Run assessment.
- Microsoft Graph : connexion interactive déléguée standard lancée automatiquement après Exchange Online dans un processus PowerShell 7 isolé, afin d'éviter les conflits de bibliothèques MSAL avec ExchangeOnlineManagement.
- Active Directory : tentative live automatique ; fallback sur `AD_Users_AllDomains.csv` si AD est indisponible.
- Exchange on-premises / Exchange 2016 : utilisation des cmdlets live lorsqu’elles sont disponibles ; la session applique `Set-ADServerSettings -ViewEntireForest $true` (ou `Set-OnPremADServerSettings` avec la session préfixée) avant toute collecte, puis utilise `Exchange_OnPrem_Mailboxes_AllDomains.csv` en fallback si cette portée forêt ne peut pas être activée.
- Santé de synchronisation Microsoft Entra : en Live, lecture autoritaire de `onPremisesSyncEnabled` et `onPremisesLastSyncDateTime` directement sur l’objet tenant Microsoft Graph. Le CSV `M365_Entra_AzureADConnect_SyncHealth.csv` reste contextuel uniquement et ne peut pas produire un faux PASS si Graph échoue.
- Endpoint de migration et fraîcheur de la synchronisation tenant : contrôlés automatiquement pendant l’évaluation.

Les inventaires AD et Exchange 2016 sont préchargés comme sources de secours. L’inventaire de santé Entra est préchargé pour le mode CacheOnly et comme contexte en Live. Sur une machine sans accès on-premises, la bascule CSV AD/Exchange est automatique et la source réellement utilisée apparaît dans les findings et dans l’onglet Activity. L’onglet Sources est informatif : il n’expose plus de boutons de connexion séparés.

### CacheOnly

Aucune connexion EXO, Graph, AD, Exchange on-premises ou Entra Connect n’est utilisée. Toutes les preuves disponibles sont lues dans :

```text
\\server\share\WORKPLACE\DATA\Tenants\<TenantKey>
```

Le chargeur accepte les CSV directement dans ce dossier ou dans son sous-dossier `DATA-LAST`. `Cache.AlternativeRootPath` permet de déclarer un second emplacement, par exemple une copie SharePoint synchronisée. Il sélectionne, fichier par fichier, la copie existante la plus récente parmi la racine, `DATA-LAST` et le chemin alternatif.

L'onglet `CSV sources` indique pour chaque fichier attendu son chemin sélectionné, sa présence, sa fraîcheur et son utilisation effective (`Used`, fallback disponible, non utilisé, absent ou périmé).

L'onglet `Options` permet de désactiver, pour l'exécution suivante, les contrôles optionnels. Les contrôles d'intégrité minimale du CSV et de disponibilité des quatre sources principales restent obligatoires. Les contrôles décochés ne produisent pas de finding et leurs collectes coûteuses sont évitées lorsqu'elles ne servent à aucun autre contrôle. La sélection appliquée est exportée dans `Check-Options.csv`.

Lorsqu'un fichier requis ou susceptible de servir de fallback dépasse `Cache.MaximumAgeHours`, le GUI demande explicitement si l'opérateur souhaite continuer. Une réponse positive accepte les CSV périmés uniquement pour l'exécution courante ; le JSON n'est pas modifié. Une réponse négative annule l'assessment avant les connexions et les collectes.

Les contrôles qui nécessitent obligatoirement une interrogation live, notamment le test d’endpoint et la recherche des mailboxes soft-deleted/inactive si le cache ne les expose pas, sont signalés `UNKNOWN` avec une recommandation de validation finale en mode Live. Ils ne sont jamais convertis silencieusement en `PASS`.

## Autonomie

L’application possède son propre fichier de configuration, son moteur de contrôles et ses exports. Elle ne charge aucun script SmartM365. Si son `TenantProfile.TenantId` est vide, elle lit uniquement le profil tenant central SmartM365 afin de reprendre en mémoire le même tenant.

Les CSV du cache peuvent être produits par SmartM365, mais l’application les consomme directement et reste exécutable indépendamment.

Le profil opérationnel est défini dans le JSON local ignoré par Git avec :

- `TenantProfile.TenantId` : garde-fou pour empêcher une authentification live dans le mauvais tenant ; s’il est vide, le profil `DefaultTenant` de SmartM365 est utilisé en mémoire.
- `TenantProfile.ProfileKey` : clé explicite du profil tenant, par exemple `prod` ; la valeur générique `tenant` reprend seulement le `DefaultTenant` central.
- `TenantProfile.RemoteRoutingDomain` : domaine de routage hybride attendu.
- `Cache.RootPath` : racine des inventaires CSV.
- `Cache.MaximumAgeHours` : âge maximal accepté pour les données.

Il n’existe plus de bloc générique `Tenant`, de paramètre `MicrosoftGraph.UseDeviceCode`, de bloc `OnPremises`, ni de paramètre `EntraConnect.Server`.

Exchange Online et Microsoft Graph utilisent exclusivement une authentification interactive déléguée standard. Aucun flux device code, secret applicatif ou certificat d’application n’est utilisé ni attendu dans le JSON.

## Sécurité

L’application est strictement diagnostique. Elle ne crée pas de batch, ne modifie aucun destinataire, attribut AD, licence, hold, permission ou objet de migration.

Les jetons ne sont pas enregistrés. Le fichier JSON local, les CSV d’entrée et les rapports opérationnels sont exclus de Git.

## Prérequis

Pour tous les modes :

- Windows.
- PowerShell 7 disponible dans `C:\Program Files\PowerShell\7\pwsh.exe` pour le lanceur fourni.
- Accès en lecture au chemin du cache lorsque des preuves mises en cache sont nécessaires.

Uniquement pour le mode Live :

- module `ExchangeOnlineManagement` 3.0 ou ultérieur ;
- modules `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users` et `Microsoft.Graph.Identity.DirectoryManagement` ; lorsqu’un module manque dans PowerShell 7, le GUI propose de l’installer depuis PowerShell Gallery avec `-Scope CurrentUser` avant l’authentification ;
- compte interactif disposant des droits de lecture nécessaires dans Exchange Online et Microsoft Graph ;
- pour les sources on-premises live, module `ActiveDirectory` et Exchange Management Shell disponibles sur la machine d’exécution. Leur absence n’est pas bloquante si les CSV de fallback sont accessibles et suffisamment récents. Le module ADSync n’est ni utilisé ni requis.

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
- `AssessmentPhase` : `PreCreation` par défaut ou `ExistingBatch` ; le sélecteur GUI s’applique à l’exécution courante.
- `DisabledChecks` : liste optionnelle d'identifiants de contrôles désactivés par défaut ; le GUI ne modifie pas automatiquement le JSON.
- `TenantProfile` : profil tenant autonome de l’application.
- `Cache.RootPath` et `Cache.MaximumAgeHours` : emplacement et fraîcheur des inventaires.
- `ExchangeOnline.UserPrincipalName` : UPN administrateur optionnel.
- `ExchangeOnline.DisableWam` : conserve le flux interactif EXO sans WAM lorsque requis sur ce poste.
- `Hybrid.MigrationEndpointName` : présélection facultative d’un endpoint `ExchangeRemoteMove` dans le JSON local ; le modèle reste vide. En mode Live, le GUI charge automatiquement les endpoints après l’authentification EXO, sélectionne l’unique endpoint ou demande explicitement lequel utiliser lorsqu’il y en a plusieurs. Le choix reste limité à la session et le JSON n’est jamais réécrit. En CacheOnly, la liste est désactivée et aucun test Live n’est exécuté.
- En mode Live, Active Directory est interrogé dans chaque domaine retourné par `Get-ADForest`. Si un domaine ne peut pas être interrogé, la couverture Live est considérée incomplète et l’application utilise le CSV AD de fallback au lieu de conclure à tort que l’identité est absente.
- `Hybrid.TargetDeliveryDomain` : domaine `tenant.mail.onmicrosoft.com` attendu.
- `Hybrid.ActiveMigrationWarningThreshold` : seuil consultatif du nombre de migrations actives/non terminales, `100` par défaut.
- `DefaultTargetSku`, `TargetQuotaGbBySku`, `MailboxIneligibleTargetSkus` et `QuotaSafetyBufferPercent` : politique explicite du SKU et du quota cible. Un SKU absent de la table reste `UNKNOWN` bloquant ; aucun quota générique de 100 Go n’est supposé.
- `EntraConnectHealth.MaximumLastSyncAgeMinutes` : ancienneté maximale de la dernière synchronisation tenant ; `120` minutes par défaut. En Live, un état désactivé, une date absente, une collecte Graph indisponible ou une synchronisation trop ancienne produit un résultat bloquant.
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
- `M365_Licenses_ServicePlans.csv`
- `AD_Users_DuplicateSMTP.csv`
- `AD_Users_DuplicateRemoteRoutingAddress.csv`
- `Exchange_OnPrem_ProxyAddresses_Check.csv`
- `Exchange_EXO_AcceptedDomains.csv`
- `Exchange_EXO_Mailboxes_AllDomains_Archive.csv`
- `Exchange_OnPrem_MigrationReadiness_Config.csv`

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
- unicité globale des proxy SMTP et adresses de routage, doublons internes, domaines acceptés et préservation X500/LegacyExchangeDN ;
- cohérence ExchangeGuid/ArchiveGuid et type de destinataire pris en charge ;
- taille de mailbox contre quota explicite du SKU cible avec marge de sécurité ; les shared mailboxes sans SKU explicite utilisent la limite non licenciée de 50 Go ;
- préparation de l'archive, saturation Recoverable Items, limites de dossiers, gros éléments et quotas source personnalisés ;
- Litigation Hold et In-Place Hold quand les propriétés sont disponibles ;
- baseline Full Access, Send As et Send on Behalf ;
- dépendances de délégation hors batch, forwarding mailbox, règles Inbox de transfert, modération et restrictions de remise ;
- santé de la base Exchange source ;
- état MailUser/mailbox Exchange Online et détection split-brain ;
- conflits soft-deleted/inactive ;
- migration user ou move request actif/non terminal, avec déduplication des deux représentations d’une même opération et interprétation adaptée à la phase `PreCreation` ou `ExistingBatch` ; l’erreur Exchange Online `No such request exists in specified index` est interprétée comme une absence de move, pas comme une collecte inconnue ;
- historique de moves échoués, suspendus ou arrêtés ;
- unicité et synchronisation de l’utilisateur Entra ;
- erreurs de provisioning et identity anchor Entra ; l’ancienneté de synchronisation par objet reste informative et ne remplace pas la date de dernière synchronisation tenant collectée sur l’objet `organization` ;
- licence actuelle, UsageLocation, capacité du SKU cible et présence d’un service plan Exchange mailbox activé ; l’attente de licence dépend de la phase et `SPE_F1` est non éligible ;
- santé et fraîcheur de la dernière synchronisation tenant en Live directement via Microsoft Graph ; le cache reste contextuel, et CacheOnly utilise exclusivement le CSV ;
- endpoint `ExchangeRemoteMove` en Live ; l’absence de `Test-MigrationServerAvailability` produit `UNKNOWN`, pas un faux échec de l’endpoint ;
- MRSProxy, certificat hybride, charge active des migrations et cohérence Autodiscover/OAuth ; le contrôle OAuth est consultatif pour un move distant et ne bloque plus à lui seul la migration ;
- avertissement documenté pour `CannotMoveEnhancedRestoreMailboxesCrossOrgPermanentException`.

Une source obligatoire absente reste bloquante. Une propriété non disponible dans le cache devient `UNKNOWN`, jamais un faux `PASS`. Chaque contrôle obligatoire produit désormais explicitement un finding `PASS` ou `UNKNOWN`, et `SourceTimestamp` correspond à l’horodatage réel de la source utilisée plutôt qu’à l’heure de l’évaluation.

Les contrôles tenant (endpoint, MRSProxy, certificat, capacité, OAuth et synchronisation Microsoft Entra) sont évalués une seule fois. Ils apparaissent dans l’onglet `Tenant checks` et dans `Global-Findings.csv`. Seuls leurs vrais blocages sont répercutés dans le verdict de chaque mailbox, sans dupliquer les findings.

## Rapports

Chaque exécution crée :

```text
Output\SEMR-yyyyMMdd-HHmmss\
  Summary.csv
  Findings.csv
  Global-Findings.csv
  Permissions-Baseline.csv
  Evidence.csv
  Csv-Sources.csv
  Check-Options.csv
  SmartM365-ExchangeMigrationReadiness-SEMR-yyyyMMdd-HHmmss.xlsx
  SmartM365-ExchangeMigrationReadiness-SEMR-yyyyMMdd-HHmmss.html
```

`Summary.csv` contient un verdict par mailbox avec les compteurs mailbox et tenant séparés. `Findings.csv` contient les contrôles propres aux mailboxes. `Global-Findings.csv` contient les contrôles tenant exécutés une seule fois. Chaque finding précise la sévérité, le résultat, le caractère bloquant, la valeur observée, la valeur attendue, la source, son horodatage, le message et l’action recommandée.

Le classeur Excel autonome regroupe tous les CSV générés dans des onglets formatés — les sept exports actuels et tout futur CSV du même dossier — avec filtres, première ligne figée et couleurs de verdict. Il ne nécessite ni Microsoft Excel ni le module ImportExcel.

Le rapport HTML UTF-8 est autonome et contient la synthèse GO / NO-GO, les contrôles tenant, la fraîcheur des sources CSV, les détails bloquants et un filtre mailbox. Il n’utilise aucune ressource externe.

Chaque lancement GUI crée également un journal de session horodaté sous `Output\Logs`. Le statut supérieur et l'onglet `Activity` décrivent les phases longues et indiquent quand l'opérateur doit patienter.

Browse et Run assessment ouvrent une petite fenêtre WPF sur un thread dédié. Elle reste réactive pendant les appels longs, affiche l’étape, le détail, la progression et le temps écoulé. À la fin d’un assessment, elle propose d’ouvrir directement le HTML, l’Excel ou le dossier de sortie ; en cas d’échec, elle propose le journal.

Après migration, la comparaison de permissions en mode Live peut aussi générer :

```text
Permissions-Current-EXO.csv
Permissions-Comparison.csv
```
