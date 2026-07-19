# Smart Exchange Migration Readiness

Application autonome PowerShell 7 / WPF de prévalidation en lecture seule des batches de migration Exchange hybride vers Exchange Online.

Elle charge un CSV de boîtes aux lettres, interroge les sources autoritaires en Live et produit un verdict par boîte : `GO`, `GO-WARNING`, `NO-GO` ou `UNKNOWN`.

## Mode Live strict

L’application fonctionne exclusivement en **Live strict**. Elle ne propose plus de mode `CacheOnly` et ne charge aucun inventaire CSV de secours.

La phase d’évaluation reste sélectionnable :

- `PreCreation` — avant création du batch : un move actif est bloquant et l’absence de licence cible est attendue.
- `ExistingBatch` — batch déjà créé ou démarré : un move actif et une licence Exchange sont attendus.

Sources obligatoires :

- Exchange Online : connexion interactive déléguée lancée par `Run assessment` ;
- Microsoft Graph : connexion interactive déléguée dans un processus PowerShell 7 isolé ;
- Active Directory : interrogation groupée de tous les domaines retournés par `Get-ADForest` ;
- Exchange 2016 : worker Windows PowerShell 5.1 local qui charge directement le snap-in `Microsoft.Exchange.Management.PowerShell.SnapIn`, puis applique `Set-ADServerSettings -ViewEntireForest $true` ;
- santé Microsoft Entra Connect : `onPremisesSyncEnabled` et `onPremisesLastSyncDateTime` lus directement sur l’organisation Microsoft Graph.

Si une source obligatoire est indisponible, les contrôles possibles continuent, mais l’assessment est marqué `INCOMPLETE`. Une source manquante ne peut jamais produire un verdict `GO`.

Le worker Exchange 2016 exécute un self-test de sérialisation CLIXML sous Windows PowerShell 5.1 avant le preflight. Une erreur limitée à une mailbox est isolée et produit une évidence `UNKNOWN` pour cette mailbox sans interrompre la collecte du reste du batch.

Le contrôle d’unicité SMTP regroupe par défaut les adresses du batch par lots de 25 et recherche leurs propriétaires dans toute la forêt Exchange. Chaque lot est exécuté dans un processus Windows PowerShell 5.1 distinct avec un timeout de 60 secondes : le processus enfant est tué à l’expiration et le lot produit une évidence `UNKNOWN` sans figer le worker principal. Un journal par processus enfant est conservé sous `Output\Logs\Exchange2016Children\<RunId>`. Le bouton d’annulation arrête le processus enfant et termine le worker après trois secondes s’il ne répond pas.

L’onglet `Sources`, `Live-Sources.csv`, la feuille Excel `Live Sources` et le rapport HTML exposent l’état et le détail de chaque source obligatoire.

## Autonomie et sécurité

L’application possède son propre JSON, son moteur et ses exports. Elle ne charge aucun autre script SmartM365. Si `TenantProfile.TenantId` est vide, elle peut reprendre en mémoire le tenant par défaut de la configuration centrale SmartM365.

Aucun inventaire SmartM365 n’est requis. Seuls le CSV du batch et les rapports générés sont utilisés comme fichiers CSV.

L’application est strictement diagnostique : elle ne crée pas de batch et ne modifie aucun destinataire, attribut AD, licence, hold, permission ou objet de migration. Les jetons ne sont pas enregistrés.

## Prérequis

- Windows et PowerShell 7 ;
- module `ExchangeOnlineManagement` 3.0 ou ultérieur ;
- modules `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users` et `Microsoft.Graph.Identity.DirectoryManagement` ; le GUI peut proposer leur installation sous `CurrentUser` ;
- module `ActiveDirectory` et accès à tous les domaines de la forêt ;
- rôle/outils Exchange Management Shell 2016 installés localement et snap-in `Microsoft.Exchange.Management.PowerShell.SnapIn` disponible sous Windows PowerShell 5.1 ;
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
- état UserMailbox, RemoteMailbox et MailUser Exchange 2016 ;
- cohérence Primary SMTP, proxyAddresses, targetAddress et domaine de routage ;
- unicité globale Live des adresses SMTP et targetAddress via l’annuaire destinataires Exchange en portée forêt ;
- domaines SMTP acceptés dans Exchange Online et domaine UPN vérifié dans Microsoft Entra ;
- préservation X500/LegacyExchangeDN et cohérence ExchangeGuid/ArchiveGuid ;
- taille de mailbox contre le quota du SKU cible et de la licence actuellement attribuée ;
- archives, Recoverable Items, limites de dossiers, gros éléments et quotas personnalisés ;
- Litigation Hold et In-Place Hold ;
- baseline Full Access, Send As et Send on Behalf ;
- délégations hors batch, forwarding, règles Inbox, modération et restrictions de remise ;
- disponibilité de la base Exchange source ;
- état Exchange Online, conflits soft-deleted/inactive et objets de migration existants ;
- synchronisation Entra, provisioning errors, identity anchor, licences, UsageLocation et capacité du SKU ;
- endpoint `ExchangeRemoteMove`, MRSProxy, certificat TLS, charge, backlog et OAuth hybride ;
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

`Summary.csv` contient `AssessmentStatus`, le verdict, la taille, le SKU cible, les licences attribuées et les compteurs mailbox/tenant séparés. `Check-Coverage.csv` vérifie que chaque contrôle activé a produit exactement un résultat par mailbox ou par tenant, sans doublon ; les branches manquantes sont matérialisées en `UNKNOWN`. Le classeur ajoute les feuilles `Action Plan` et `Check Coverage`. Le HTML affiche le statut `COMPLETE/INCOMPLETE`, les sources Live, les décisions mailbox et les contrôles tenant.

## Permissions après migration

Après un assessment, l’onglet `Permissions baseline` permet de comparer la baseline enregistrée aux permissions Exchange Online actuelles. Cette comparaison est elle aussi en lecture seule.
