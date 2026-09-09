# Smart Exchange Migration Readiness

Application autonome PowerShell 7 / WPF de prévalidation en lecture seule des batches de migration Exchange hybride vers Exchange Online. Préversion **1.11.17**, publication approuvée le 9 septembre 2026. [Release GitHub et package portable](https://github.com/khda79/workplacecloudhub.com/releases/tag/exchange-migration-readiness-v1.11.17). La qualification Live sur un hôte Exchange équipé reste nécessaire ; les limites de validation sont détaillées dans `VALIDATION-1.11.17.md`.

Elle charge un CSV de boîtes aux lettres, interroge les sources autoritaires en Live et produit un verdict par boîte : `GO`, `GO-WARNING`, `NO-GO` ou `UNKNOWN`.

## Mode Live strict

L’application fonctionne exclusivement en **Live strict**. Elle ne propose plus de mode `CacheOnly` et ne charge aucun inventaire CSV de secours.

L'outil détecte le produit et le build du serveur local via `Get-ExchangeServer` : Exchange Server 2016 (15.1), Exchange Server 2019 (15.2 avant le build 2562) et Exchange Server Subscription Edition (15.2 build 2562 ou ultérieur). Cette reconnaissance technique n'est pas une certification de support Microsoft : Exchange 2016 et 2019 ont atteint leur fin de support le 14 octobre 2025 ([Microsoft Learn](https://learn.microsoft.com/en-us/troubleshoot/exchange/administration/exchange-2019-2016-end-of-support)).

La phase d’évaluation reste sélectionnable :

- `PreCreation` — avant création du batch : un move actif est bloquant et l’absence de licence cible est attendue.
- `ExistingBatch` — batch déjà créé ou démarré : un move actif et une licence Exchange sont attendus.

Sources obligatoires :

- Exchange Online : connexion interactive déléguée lancée par `Run assessment` ;
- Microsoft Graph : connexion interactive déléguée dans un processus PowerShell 7 isolé ;
- Active Directory : interrogation groupée de tous les domaines retournés par `Get-ADForest` ;
- Exchange Server 2016, 2019 ou Subscription Edition (SE) : worker Windows PowerShell 5.1 local qui charge directement le snap-in `Microsoft.Exchange.Management.PowerShell.SnapIn`, puis applique `Set-ADServerSettings -ViewEntireForest $true` ;
- santé Microsoft Entra Connect : `onPremisesSyncEnabled` et `onPremisesLastSyncDateTime` lus directement sur l’organisation Microsoft Graph.

Au premier assessment, le GUI inspecte les sessions EXO actives dans le processus et le contexte délégué Graph CurrentUser. Lorsqu'une session existe, une boîte affiche le compte, le tenant, l'organisation ou le type d'authentification disponibles et propose de la réutiliser, de forcer une nouvelle authentification ou d'annuler. Une session EXO préexistante réutilisée n'est pas fermée par l'application. Les données Graph sont recollectées pour chaque batch, même lorsque le contexte d'authentification est réutilisé ; un nouveau navigateur apparaît uniquement si le contexte est absent, incompatible, expiré, incomplet en scopes ou si l'utilisateur force la reconnexion.

Si une source obligatoire est indisponible, les contrôles possibles continuent, mais l’assessment est marqué `INCOMPLETE`. Une source manquante ne peut jamais produire un verdict `GO`.

Le worker Exchange on-premises exécute un self-test de sérialisation CLIXML sous Windows PowerShell 5.1 avant le preflight. Une erreur limitée à une mailbox ou à une commande est isolée et produit une évidence `UNKNOWN` pour le contrôle concerné sans interrompre la collecte du reste du batch. Les erreurs partielles et fatales sont comptabilisées séparément dans le journal du worker et remontées avec leur mailbox, contrôle et commande dans le journal principal du GUI.

Le contrôle d’unicité SMTP regroupe par défaut les adresses du batch par lots de 50 et recherche leurs propriétaires dans toute la forêt Exchange. Chaque lot est exécuté dans un processus Windows PowerShell 5.1 distinct avec un timeout de 60 secondes : le processus enfant est tué à l’expiration et le lot produit une évidence `UNKNOWN` sans figer le worker principal. Un journal par processus enfant est conservé sous `Output\Logs\ExchangeOnPremChildren\<RunId>`. La taille des lots et le timeout sont configurables par `ExchangeOnPremises.SmtpUniquenessBatchSize` (1 à 50) et `ExchangeOnPremises.SmtpUniquenessBatchTimeoutSeconds` (5 à 300). Le bouton d’annulation arrête le processus enfant et termine le worker après trois secondes s’il ne répond pas.

Le contrôle bloquant de connectivité hybride est le test fonctionnel `Test-MigrationServerAvailability` appliqué à l’endpoint `ExchangeRemoteMove` sélectionné. Son résultat, son message et sa durée sont journalisés. L’inventaire local de toutes les virtual directories EWS et de leur propriété `MRSProxyEnabled` est un diagnostic serveur optionnel, décoché par défaut et non bloquant ; il peut être activé dans l’onglet `Options` lors d’une investigation de topologie.

La session GUI écrit un journal structuré sous Output\Logs avec le PID, le thread, l'identifiant de session, le RunId, le composant, l'étape, la mailbox et les durées. Les diagnostics Graph sont conservés sous Output\Logs\GraphWorkers\<RunId>\MicrosoftGraph-Worker.log. Le worker Exchange on-premises produit Output\Logs\ExchangeOnPremChildren\<RunId>\ExchangeOnPrem-Worker.log en plus des journaux SMTP par lot. Les exceptions incluent leur type, identifiant PowerShell, commande, ligne, pile et exceptions internes, sans jeton ni mot de passe.

L’onglet `Sources`, `Live-Sources.csv`, la feuille Excel `Live Sources` et le rapport HTML exposent l’état et le détail de chaque source obligatoire.

## Autonomie et sécurité

L’application possède son propre JSON, son moteur et ses exports. Elle ne charge aucun autre script SmartM365. Si `TenantProfile.TenantId` est vide, elle peut reprendre en mémoire le tenant par défaut de la configuration centrale SmartM365.

Aucun inventaire SmartM365 n’est requis. Seuls le CSV du batch et les rapports générés sont utilisés comme fichiers CSV.

L’application est strictement diagnostique : elle ne crée pas de batch et ne modifie aucun destinataire, attribut AD, licence, hold, permission ou objet de migration. Les jetons ne sont pas enregistrés.

## Prérequis

- Windows avec WPF et une version PowerShell compatible avec le module EXO : PowerShell 7.4 minimum pour EXO 3.5 à 3.9.2, PowerShell 7.6 minimum pour EXO 3.10 et ultérieur ([matrice Microsoft](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps)) ;
- module `ExchangeOnlineManagement` 3.7.2 minimum avec le réglage par défaut `DisableWam=true` ;
- modules `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users` et `Microsoft.Graph.Identity.DirectoryManagement` à une **même version installée** ; le worker choisit la plus récente commune aux trois. Le GUI peut proposer leur installation alignée sous `CurrentUser` ;
- module `ActiveDirectory` et accès à tous les domaines de la forêt ;
- rôle/outils Exchange Management Shell 2016, 2019 ou Subscription Edition installés localement et snap-in `Microsoft.Exchange.Management.PowerShell.SnapIn` disponible sous Windows PowerShell 5.1 ;
- compte interactif disposant des droits de lecture EXO et Graph nécessaires.

Le module ADSync n’est ni utilisé ni requis. Aucune session PowerShell distante Exchange, aucun `ConnectionUri` et aucun mode de compatibilité de module PS7 ne sont utilisés : le GUI PowerShell 7 orchestre un processus local Windows PowerShell 5.1 avec l’identité Windows courante.

Scopes Graph par défaut :

- `User.Read.All`
- `Directory.Read.All`
- `Organization.Read.All`

## Démarrage

```text
Start-SmartM365-ExchangeMigrationReadiness-GUI.cmd
```

Validation statique sans ouvrir l’interface :

```powershell
pwsh -NoLogo -NoProfile -STA -File .\SmartM365-ExchangeMigrationReadiness-GUI.ps1 -ValidateOnly
```

## Configuration

Modèle versionné :

```text
Config\SmartM365-ExchangeMigrationReadiness.local.json.template
```

Fichier local ignoré par Git :

```text
Config\SmartM365-ExchangeMigrationReadiness.local.json
```

Clés principales :

- `AssessmentPhase` : `PreCreation` ou `ExistingBatch` ;
- `DisabledChecks` : contrôles optionnels désactivés par défaut ;
- `TenantProfile.TenantId` : garde-fou du tenant interactif ;
- `TenantProfile.ProfileKey` : clé du profil, par exemple `prod` ;
- `TenantProfile.RemoteRoutingDomain` : domaine de routage hybride attendu ;
- `ExchangeOnline.UserPrincipalName` et `ExchangeOnline.DisableWam` ;
- `MicrosoftGraph.Scopes` ;
- `ExchangeOnPremises.SmtpUniquenessBatchSize` et `ExchangeOnPremises.SmtpUniquenessBatchTimeoutSeconds` ;
- `Hybrid.MigrationEndpointName` : endpoint `ExchangeRemoteMove` présélectionné, facultatif ;
- `Hybrid.TargetDeliveryDomain` et `Hybrid.ActiveMigrationWarningThreshold` ;
- `DefaultTargetSku`, `TargetQuotaGbBySku`, `MailboxIneligibleTargetSkus` et `QuotaSafetyBufferPercent` ;
- `EntraConnectHealth.MaximumLastSyncAgeMinutes` ;
- `OutputRoot`.

Il n’existe aucun paramètre `Mode`, `Cache`, `UseDeviceCode`, secret applicatif ou certificat d’application.

## CSV de migration

La colonne canonique est `EmailAddress`. Sont aussi acceptées : `PrimarySmtp`, `PrimarySmtpAddress`, `UserPrincipalName`, `UPN` ou `Mailbox`.

Colonnes optionnelles :

- `MailboxType`
- `TargetSku` ou `TargetSkuPartNumber`
- `BadItemLimit`
- `LargeItemLimit`

Les délimiteurs virgule, point-virgule et tabulation sont détectés automatiquement. L’import essaie UTF-8 strict puis Windows-1252.

## Contrôles principaux

- format, valeurs vides, syntaxe SMTP, colonnes et doublons du batch ;
- existence, unicité et état du compte AD dans toute la forêt ;
- état UserMailbox, RemoteMailbox et MailUser Exchange on-premises ;
- cohérence Primary SMTP, proxyAddresses, targetAddress et domaine de routage ;
- unicité globale Live des adresses SMTP et targetAddress via l’annuaire destinataires Exchange en portée forêt ;
- domaines SMTP acceptés dans Exchange Online et domaine UPN vérifié dans Microsoft Entra ;
- présence du LegacyExchangeDN, audit du proxy X500 correspondant (avertissement non bloquant pour un remote move hybride classique) et cohérence ExchangeGuid/ArchiveGuid ;
- taille de mailbox contre le quota du SKU cible et de la licence actuellement attribuée ;
- archives, Recoverable Items, limites de dossiers et quotas personnalisés ; le contrôle `MAILBOX-LARGE-ITEMS`, non collectable par les cmdlets live disponibles, est désactivé par défaut et reste activable dans `Options` pour une investigation ciblée ;
- Litigation Hold et In-Place Hold ;
- baseline Full Access, Send As et Send on Behalf, après exclusion de SELF et des groupes techniques Exchange ; les délégués sont rapprochés du batch par UPN, adresse, DN, `DOMAINE\sAMAccountName` et SID ;
- délégations hors batch, forwarding, règles Inbox, modération et restrictions de remise ;
- disponibilité de la base Exchange source ;
- état Exchange Online, conflits soft-deleted/inactive et objets de migration existants ;
- synchronisation Entra, provisioning errors, identity anchor, licences, UsageLocation et capacité du SKU ;
- endpoint `ExchangeRemoteMove` testé fonctionnellement, certificat TLS, charge et backlog ; diagnostic local EWS/MRSProxy optionnel ; OAuth hybride informatif et non bloquant pour le remote move ;
- erreur `CannotMoveEnhancedRestoreMailboxesCrossOrgPermanentException`.

Une source obligatoire absente produit des findings bloquants `UNKNOWN` et un assessment `INCOMPLETE`. Un blocage confirmé produit `NO-GO`. Les alertes tenant non bloquantes restent dans `Tenant checks` et ne transforment pas artificiellement toutes les mailboxes en `GO-WARNING`.

## Rapports

Chaque exécution crée :

```text
Output\SEMR-yyyyMMdd-HHmmss\
  Summary.csv
  Findings.csv
  Global-Findings.csv
  Permissions-Baseline.csv
  Evidence.csv
  Live-Sources.csv
  Check-Coverage.csv
  Check-Options.csv
  SmartM365-ExchangeMigrationReadiness-SEMR-yyyyMMdd-HHmmss.xlsx
  SmartM365-ExchangeMigrationReadiness-SEMR-yyyyMMdd-HHmmss.html
```

`Summary.csv` contient `AssessmentStatus`, le verdict, la taille, le SKU cible, les licences attribuées et les compteurs mailbox/tenant séparés. `DataCoverage` représente uniquement la couverture d’exécution et des sources (`Complete`, `Partial` ou `Incomplete`) ; un résultat métier `UNKNOWN` reste compté séparément et ne dégrade plus artificiellement la couverture. `Check-Coverage.csv` vérifie que chaque contrôle activé a produit exactement un résultat par mailbox ou par tenant, sans doublon ; les branches manquantes sont matérialisées en `UNKNOWN`. Le classeur ajoute les feuilles `Action Plan` et `Check Coverage`. Le HTML affiche le statut `COMPLETE/INCOMPLETE`, les sources Live, les décisions mailbox et les contrôles tenant.

## Permissions après migration

Après un assessment, l’onglet `Permissions baseline` permet de comparer la baseline enregistrée aux permissions Exchange Online actuelles. Cette comparaison est elle aussi en lecture seule.

La comparaison s'arrête explicitement si une commande Full Access, Send As ou Send on Behalf est absente, échoue ou ne fournit pas la propriété attendue. Un échec de lecture ne signifie jamais « aucun droit ». Les ACE `Deny` ne sont pas exportées comme des autorisations ; les SID de délégués sont conservés. Les identités sont comparées textuellement sans distinction de casse : un même délégué représenté par SID, DN et UPN peut encore apparaître comme différent et doit être rapproché manuellement. Une baseline initiale partielle ne certifie pas la conservation de tous les droits. Ces exports décrivent les droits explicites, pas le calcul complet des accès effectifs via groupes et héritage.

## Installation autonome et permissions opérateur

Copier le dossier complet, y compris les deux workers, le module, les fichiers splash/update-check, les images, `Config` et `Samples`. Aucun MSI, déploiement Intune ou module PowerShell Gallery du produit n'est fourni. Les modules Microsoft restent des prérequis externes. Ne copier ni `*.local.json`, ni `Output`, ni les CSV réels dans une distribution publique.

Sous PowerShell 7 compatible, installer les prérequis Microsoft si nécessaire :

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Repository PSGallery
$graphVersion = (Find-Module Microsoft.Graph.Authentication -Repository PSGallery).Version
Install-Module Microsoft.Graph.Authentication,Microsoft.Graph.Users,Microsoft.Graph.Identity.DirectoryManagement -RequiredVersion $graphVersion -Scope CurrentUser -Repository PSGallery
```

Préparer ensuite une copie du modèle `Config\SmartM365-ExchangeMigrationReadiness.local.json.template` nommée `SmartM365-ExchangeMigrationReadiness.local.json`. Renseigner `TenantProfile.TenantId` avec le GUID du tenant prévu, `RemoteRoutingDomain` et `Hybrid.TargetDeliveryDomain` avec le domaine hybride réel ; ajuster le SKU et le répertoire de sortie. Le chargement ajoute les clés manquantes du modèle sans écraser les valeurs locales. En copie autonome, aucune configuration centrale n'est disponible : le TenantId doit être renseigné explicitement. Une simple réussite de `Get-EXOMailbox` ne suffit plus à identifier le tenant ; des métadonnées EXO absentes ou plusieurs sessions imposent une nouvelle connexion vérifiable.

| Source | Permissions et contrôles à préparer |
| --- | --- |
| Microsoft Graph délégué | Consentement administrateur pour les scopes ci-dessus ; compte avec un rôle Entra pris en charge pour la lecture des licences, par exemple Directory Readers ou Global Reader. Les scopes seuls ne garantissent pas les droits du compte. Voir [licenseDetails](https://learn.microsoft.com/en-us/graph/api/user-list-licensedetails?view=graph-rest-1.0). |
| Exchange Online délégué | RBAC autorisant la lecture des destinataires, MailUser, mailboxes actives/soft-deleted/inactive, domaines, permissions, migrations/move requests et endpoints, ainsi que `Test-MigrationServerAvailability`. Vérifier les cmdlets réellement disponibles dans la session ; un simple rôle de lecture des destinataires ne couvre pas nécessairement les tests de migration. Voir [recherche des permissions de cmdlets](https://learn.microsoft.com/en-us/powershell/exchange/find-exchange-cmdlet-permissions?view=exchange-ps). |
| AD et Exchange local | Identité Windows courante avec accès de lecture à tous les domaines de la forêt et aux objets Exchange, statistiques, permissions, bases et configuration hybride. Aucun compte local alternatif ni accès distant Exchange n'est utilisé. |
| Poste opérateur | Répertoire de sortie inscriptible, accès réseau aux domaines AD, Microsoft Graph, EXO et à l'endpoint hybride. Le worker doit reconnaître un serveur Exchange local via `Get-ExchangeServer` : un poste RSAT seul ne suffit pas. |

Ne pas exécuter le bootstrap app-only SmartM365 pour cette application autonome. Elle ne demande ni `Mail.Send` ni permission Graph d'écriture ; elle ne transmet pas automatiquement ses rapports à SharePoint ou Teams. Le test fonctionnel d'endpoint est diagnostique et ne crée pas de migration. Le contrôle de mise à jour au lancement consulte GitHub ; `-ValidateOnly` n'ouvre pas l'interface et ne lance aucune connexion tenant.

## Validation et limites des rapports

```powershell
pwsh -NoProfile -STA -File .\SmartM365-ExchangeMigrationReadiness-GUI.ps1 -ValidateOnly
pwsh -NoProfile -File .\SmartM365-ExchangeMigrationReadiness-GraphWorker.ps1 -ValidateOnly
pwsh -NoProfile -File .\Tests\SmartM365-Test-ExchangeMigrationReadiness.ps1
```

Le premier contrôle valide le XAML, le JSON, la sérialisation PS5.1 et les auto-tests du moteur. Le second importe les modules Graph sans authentification. La suite de régression utilise exclusivement des fixtures locales, vérifie les erreurs de permissions, les sessions EXO ambiguës, les payloads Graph absents, les CSV et l'export réel d'un assessment sans sources. Ces tests ne certifient ni les permissions effectives d'un tenant ni la topologie hybride réelle.

Le classeur autonome contient neuf feuilles de base : Summary, Mailbox Findings, Tenant Checks, Permissions, Evidence, Live Sources, Check Coverage, Action Plan et Check Options. Excel et ImportExcel ne sont pas nécessaires à sa génération. Le HTML et le GUI restent en anglais ; la présentation WorkplaceCloudHub est prévue en EN/FR/IT/ES/DE/AR. Les CSV, classeurs, HTML et logs contiennent des identités, adresses, délégations et détails d'infrastructure : ils ne sont pas anonymisés. Protéger leur accès. Dans Excel, ouvrir les CSV non fiables par import en colonnes texte ; le classeur généré encode les chaînes comme du texte.

Un `GO` porte uniquement sur les contrôles activés et les preuves collectées à l'instant donné. Examiner aussi Check Options, Live Sources, Check Coverage et les `UNKNOWN`. Une information Graph absente du payload est une erreur de collecte, jamais la preuve qu'un utilisateur n'existe pas. La disponibilité des sources et le verdict métier sont deux informations distinctes.

## Changements de la préversion 1.11.17

- Sélection cohérente des versions Graph et diagnostic visible en cas d'échec de `-ValidateOnly`.
- Refus d'une identité tenant EXO non vérifiable ou de plusieurs connexions ambiguës.
- Distinction entre payload Graph absent et recherche d'utilisateur vide.
- Une licence cible déjà attribuée ne réclame plus un second siège disponible ; la capacité globale du batch continue de compter les destinataires non licenciés.
- Comparaison des permissions interrompue en cas d'évidence incomplète ; conservation des SID et exclusion des ACE Full Access `Deny` côté local et EXO.
- Tests hors connexion, guide d'installation et limites de validation explicités.
