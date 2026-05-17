# SmartM365 App Registration Permissions

Ce document explique les permissions ajoutees par `SmartM365-Create-AppRegistration.ps1`, pourquoi elles sont necessaires, et quels scripts les utilisent actuellement.

## Notes importantes

- Les permissions ci-dessous sont des permissions **Application** Microsoft Graph pour l'app registration SmartM365.
- `SmartM365-Create-AppRegistration.ps1` s'execute uniquement en authentification interactive deleguee: un administrateur se connecte avec les droits necessaires pour creer ou mettre a jour l'app registration.
- `SmartM365-Create-AppRegistration.ps1` est a part concernant les permissions: ses droits interactifs de setup ne sont pas la reference de privilege pour les scripts runtime/inventaire.
- En mode multi-tenant, executer le bootstrap separement par tenant avec `-Tenant <TenantKey>`. Les valeurs d'app, certificat, SharePoint, mail et Teams restent dans `Config/Tenants/<TenantKey>.local.json`, jamais dans Git.
- Les scripts app-only utilisent `https://graph.microsoft.com/.default` ou `Connect-MgGraph -ClientId ...`; les permissions effectives sont donc celles accordees a l'application dans Entra ID.
- Les scripts avec `-InteractiveAuth` peuvent demander des scopes delegues equivalents, mais l'app registration doit surtout porter les permissions application pour les executions sans surveillance.
- Les uploads SharePoint sont centralises par `Modules/SmartM365.Core` et `Modules/SmartM365.SharePoint`; ils peuvent etre utilises par plusieurs scripts lorsque `EnableSharePointUpload` est active.
- Les notifications Teams utilisent `TeamsAlertsWebhookUrl`, `TeamsInfosWebhookUrl` et `Send-SmartM365TeamsNotification` avec des URLs Teams Workflows / Power Automate, pas un ancien Office 365 Connector. Elles n'ajoutent pas de permission Graph Teams a l'app registration. Ne pas utiliser `Teamwork.Migrate.All` pour des notifications operationnelles normales. En cas d'erreur terminale, chaque script d'inventaire/rapport doit envoyer une notification Teams dans le canal `Alerts` avec le contexte de diagnostic et un lien d'aide IA construit depuis le message d'erreur. En fin d'execution sans erreur, chaque script doit envoyer une notification dans le canal `Infos` avec un champ `Result summary` qui resume le resultat du script.

## Configuration utilisateur des notifications Teams

`SmartM365-Create-AppRegistration.ps1` cree ou reutilise l'equipe `SMART-M365` et les canaux `Alerts` / `Infos`, mais il ne cree pas les URLs Teams Workflows. Ces URLs appartiennent au contexte Teams / Power Automate de l'utilisateur et doivent etre creees manuellement dans Microsoft Teams.

Procedure attendue:

1. Ouvrir Microsoft Teams.
2. Aller dans l'equipe `SMART-M365`, canal `Alerts`.
3. Ouvrir `Workflows` depuis le menu du canal.
4. Rechercher `webhook`.
5. Creer un workflow de type `Send webhook alerts to a channel` ou base sur `When a Teams webhook request is received`.
6. Selectionner l'equipe `SMART-M365` et le canal `Alerts`, puis copier l'URL HTTP POST generee.
7. Refaire la meme operation dans le canal `Infos`.
8. Enregistrer et tester les URLs avec:

```powershell
.\SmartM365-Set-TeamsWebhook.ps1 -Channel Alerts -WebhookUrl "<Alerts workflow URL>"
.\SmartM365-Set-TeamsWebhook.ps1 -Channel Infos  -WebhookUrl "<Infos workflow URL>"
```

Les URLs sont stockees uniquement dans `Config/Tenants/<TenantKey>.local.json` avec les cles `TeamsAlertsWebhookUrl` et `TeamsInfosWebhookUrl`. Ce fichier est local et ignore par Git. Les URLs de webhook doivent etre traitees comme des secrets: ne pas les publier dans un commit, une issue, une documentation ou un canal public.

Cette approche evite d'ajouter des permissions Microsoft Graph Teams larges a l'app SmartM365. Les notifications operationnelles passent par Power Automate; les seules permissions Teams demandees par le bootstrap sont des scopes delegues de setup pour creer/verifier les canaux `Alerts` et `Infos`.

## Permissions Microsoft Graph

| Permission | Pourquoi elle est necessaire | Scripts / modules utilisateurs |
| --- | --- | --- |
| `Directory.Read.All` | Lecture large de l'annuaire Entra ID: organisation, utilisateurs, groupes, licences, domaines, devices, attributs de synchronisation et enrichissements directory. Plusieurs endpoints Graph repondent mieux ou uniquement avec ce droit dans les inventaires tenant. | `M365Inventory/Users/SmartM365-ActiveUsers-Inventory.ps1`; `M365Inventory/Licensing/SmartM365-Licences-Inventory.ps1`; `M365Inventory/Domains/SmartM365-VerifiedDomains-Inventory.ps1`; `M365Inventory/Devices/SmartM365-EntraDevices-Inventory.ps1`; `IntuneInventory/RBAC/SmartM365-Intune-RBAC-GroupMembers.ps1`; `IntuneInventory/Devices/SmartM365-Devices-Compliance-Inventory.ps1`; `ExchangeInventory/Mailboxes/SmartM365-EXO-Mailboxes-Inventory.ps1`; `ExchangeInventory/CalendarPermissions/SmartM365-EXO-Mailboxes-CalPerm_Inventory.ps1`. |
| `User.Read.All` | Lecture des utilisateurs et de leurs proprietes pour les exports M365 et les enrichissements Exchange Online. | `M365Inventory/Users/SmartM365-ActiveUsers-Inventory.ps1`; `M365Inventory/Licensing/SmartM365-Licences-Inventory.ps1`; `ExchangeInventory/Mailboxes/SmartM365-EXO-Mailboxes-Inventory.ps1`; `ExchangeInventory/BackupProtection/SmartM365-EXO-BackupProtection-Comparison.ps1`. |
| `AuditLog.Read.All` | Lecture des informations d'activite de connexion exposees par Graph, notamment `signInActivity` sur les utilisateurs pour remplir les colonnes de derniere connexion. | `M365Inventory/Users/SmartM365-ActiveUsers-Inventory.ps1`. |
| `Device.Read.All` | Lecture des objets devices Entra ID, notamment pour lier les donnees Intune/Autopilot aux objets Entra et recuperer des attributs comme deviceId, trustType ou sign-in approximatif. | `M365Inventory/Devices/SmartM365-EntraDevices-Inventory.ps1`; `IntuneInventory/Devices/SmartM365-Devices-Inventory.ps1`; `IntuneInventory/Devices/SmartM365-Devices-Compliance-Inventory.ps1`; `IntuneInventory/Autopilot/SmartM365-WindowsAutopilot-Inventory.ps1`. |
| `GroupMember.Read.All` | Lecture des membres de groupes Entra ID, utile pour les exports RBAC et les comparaisons basees sur des groupes. | `IntuneInventory/RBAC/SmartM365-Intune-RBAC-GroupMembers.ps1`; `ExchangeInventory/BackupProtection/SmartM365-EXO-BackupProtection-Comparison.ps1`. |
| `DeviceManagementApps.Read.All` | Lecture des applications detectees / applications Intune via `/deviceManagement/detectedApps` et relations associees. | `IntuneInventory/Applications/SmartM365-Intune-DiscoveredApps-Inventory.ps1`. |
| `DeviceManagementConfiguration.Read.All` | Lecture des configurations, politiques et rapports Intune, notamment compliance policies, Windows Update for Business, Feature Update / Quality Update profiles et report export jobs. | `IntuneInventory/Devices/SmartM365-Device-System-Inventory.ps1`; `IntuneInventory/Devices/SmartM365-Devices-Compliance-Inventory.ps1`; `IntuneInventory/WindowsUpdate/SmartM365-WinUpdate_Status_From_Intune.ps1`; `IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Get-IntuneAutopatchAlerts.ps1`. |
| `DeviceManagementManagedDevices.Read.All` | Lecture des devices geres par Intune et de leurs proprietes: inventaires devices, BIOS, compliance, systeme, upgrade eligibility et Endpoint Analytics. | `IntuneInventory/Devices/SmartM365-Devices-Inventory.ps1`; `IntuneInventory/Devices/SmartM365-Devices-BIOS-Inventory.ps1`; `IntuneInventory/Devices/SmartM365-Devices-Compliance-Inventory.ps1`; `IntuneInventory/Devices/SmartM365-Device-System-Inventory.ps1`; `IntuneInventory/Devices/SmartM365-Devices-UpgradeEligibility.ps1`; `IntuneInventory/Applications/SmartM365-Intune-DiscoveredApps-Inventory.ps1`; `IntuneInventory/WindowsUpdate/SmartM365-WinUpdate_Status_From_Intune.ps1`; `IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Get-IntuneAutopatchAlerts.ps1`. |
| `DeviceManagementScripts.Read.All` | Lecture/export des remediation scripts Intune exposes comme `deviceHealthScripts`, incluant details et assignments. | `IntuneInventory/SmartM365-Export-IntuneRemediations.ps1`. |
| `DeviceManagementServiceConfig.Read.All` | Lecture de la configuration du service Intune, en particulier les identites Windows Autopilot. | `IntuneInventory/Autopilot/SmartM365-WindowsAutopilot-Inventory.ps1`. |
| `Mail.Send` | Envoi des notifications d'erreur et rapports HTML via Microsoft Graph lorsque `SmtpServer` est vide. L'adresse d'expedition est resolue depuis `From` dans le profil tenant ou le `*.local.json` du script. Le bootstrap cree aussi une Application Access Policy Exchange Online pour limiter ce droit au groupe `MailSendAccessPolicyGroup`. | `Modules/SmartM365.Core/SmartM365.Core.psm1`; scripts qui appellent `SendEmailHtmlReport`, `Send-SmartM365Mail` ou `SendFileListEmailReport`. |
| `Sites.Selected` | Upload et remplacement de fichiers CSV uniquement sur le site SharePoint SmartM365. Le bootstrap attribue ensuite le role `write` a l'app sur le site cree/reutilise, puis retire les anciens grants larges `Files.ReadWrite.All` et `Sites.ReadWrite.All` lorsqu'ils existent. | `Modules/SmartM365.Core/SmartM365.Core.psm1`; `Modules/SmartM365.SharePoint/SmartM365.SharePoint.psm1`; tous les scripts d'inventaire/export qui appellent l'upload SharePoint quand `EnableSharePointUpload` est active. |

## Permissions Intune ReadWrite incluses par defaut

Ces permissions sont ajoutees par defaut car un script actuel les demande encore en authentification interactive. Elles sont a considerer comme transitoires: lancer `SmartM365-Create-AppRegistration.ps1` avec `-SkipBroadIntuneReadWritePermissions` des que les scripts concernes sont durcis en lecture seule.

| Permission | Pourquoi elle est la aujourd'hui | Scripts / modules utilisateurs |
| --- | --- | --- |
| `DeviceManagementApps.ReadWrite.All` | Permission large demandee par le script Autopatch actuel pour acceder aux donnees Intune/reporting. A reduire si les appels restent en lecture seule. | `IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Get-IntuneAutopatchAlerts.ps1`. |
| `DeviceManagementConfiguration.ReadWrite.All` | Permission large demandee par le script Autopatch actuel pour les profils Windows Update et export jobs. A remplacer par `DeviceManagementConfiguration.Read.All` si aucun POST/operation d'ecriture privilegiee n'est necessaire. | `IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Get-IntuneAutopatchAlerts.ps1`. |
| `DeviceManagementManagedDevices.ReadWrite.All` | Permission large demandee par le script Autopatch actuel pour les donnees managed devices/reporting. A remplacer par `DeviceManagementManagedDevices.Read.All` si aucun changement de device n'est effectue. | `IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Get-IntuneAutopatchAlerts.ps1`. |

## Permission non-Graph ajoutee par le script

| Permission | Pourquoi elle est necessaire | Scripts / modules utilisateurs |
| --- | --- | --- |
| `Exchange.ManageAsApp` sur l'API `Office 365 Exchange Online` | Autorise l'authentification app-only par certificat avec `Connect-ExchangeOnline`. Ce n'est pas une permission Microsoft Graph. Elle ne suffit pas seule: le token EXO app-only doit aussi contenir un role Entra supporte pour que le RBAC Exchange soit construit. | Scripts Exchange Online sous `ExchangeInventory`, par exemple `AcceptedDomains`, `BackupProtection`, `CalendarPermissions`, `Mailboxes` et `Migration`, lorsqu'ils utilisent l'authentification app-only. |

## Role Entra ID attribue au service principal SmartM365

| Role | Pourquoi il est necessaire | Scripts / modules utilisateurs |
| --- | --- | --- |
| `Global Reader` | Baseline privilegiee pour les scripts runtime Exchange Online en lecture seule. Ce role Entra est supporte par Exchange Online PowerShell app-only et suffit normalement pour les inventaires et rapports qui n'appellent que des cmdlets `Get-*`. | Scripts Exchange Online sous `ExchangeInventory`, par exemple `AcceptedDomains`, `BackupProtection`, `CalendarPermissions`, `Mailboxes` et `Migration`, tant qu'ils restent en lecture seule. |
| `Exchange Administrator` | Role de setup, pas baseline runtime. Il peut etre necessaire au compte administrateur interactif qui execute `SmartM365-Create-AppRegistration.ps1`, car le bootstrap cree/modifie une shared mailbox, un groupe mail-enabled, une Application Access Policy et l'attribution du role au service principal. | `SmartM365-Create-AppRegistration.ps1` uniquement, ou futur script Exchange Online qui ferait explicitement des modifications. |

## Configuration Exchange Online creee par le bootstrap

| Objet | Pourquoi il est necessaire | Scripts / modules utilisateurs |
| --- | --- | --- |
| Shared mailbox `smartm365-reports@<domaine>` | Fournit une boite dediee comme expediteur des rapports et notifications SmartM365. Le bootstrap ecrit cette adresse dans `From`. | `Modules/SmartM365.Core/SmartM365.Core.psm1`; scripts qui envoient des rapports ou erreurs par Graph. |
| Mail-enabled security group `SMART-M365-MailSend-Allowed` | Liste les boites autorisees pour l'application SmartM365 avec `Mail.Send`. Le bootstrap ecrit son adresse dans `MailSendAccessPolicyGroup`. | Restriction Exchange Online appliquee a `Mail.Send`. |
| Application Access Policy Exchange Online | Limite l'application SmartM365 aux boites membres de `MailSendAccessPolicyGroup` pour les permissions Outlook Graph comme `Mail.Send`. Sans cette restriction, `Mail.Send` application est tenant-wide. | Envois Graph faits par `Modules/SmartM365.Core/SmartM365.Core.psm1`. |

## Permissions utilisees uniquement pour executer le bootstrap

Ces scopes sont demandes a l'administrateur qui lance `SmartM365-Create-AppRegistration.ps1`. Ils ne sont pas ajoutes a l'app SmartM365 comme permissions metier; ils servent a creer/modifier l'app registration et a accorder le consentement.

`SmartM365-Create-AppRegistration.ps1` est un script interactif de bootstrap: il est volontairement plus privilegie que les scripts d'inventaire runtime. Les roles requis par ce setup ne doivent pas etre recopies comme prerequis des scripts Exchange Online read-only.

| Scope de connexion | Pourquoi il est demande |
| --- | --- |
| `Application.ReadWrite.All` | Creer ou mettre a jour l'app registration, ses API permissions et ses certificats publics. |
| `AppRoleAssignment.ReadWrite.All` | Accorder le consentement admin sous forme d'app role assignments. `SmartM365-Create-AppRegistration.ps1` le fait par defaut; utiliser `-DisableGrantAdminConsent` uniquement pour preparer l'app sans consentement immediat. |
| `Channel.Create` | Creer les canaux Teams standard `Alerts` et `Infos` dans l'equipe `SMART-M365` lorsque le bootstrap les initialise. |
| `Channel.ReadBasic.All` | Verifier si les canaux Teams standard `Alerts` et `Infos` existent deja avant de les creer. |
| `Directory.Read.All` | Lire les service principals des APIs Microsoft Graph et Exchange Online, verifier le contexte tenant, et retrouver l'utilisateur administrateur connecte. |
| `Group.ReadWrite.All` | Creer ou reutiliser le groupe Microsoft 365 qui porte l'equipe Teams `SMART-M365`, puis convertir ce groupe en team. |
| `RoleManagement.ReadWrite.Directory` | Attribuer au service principal SmartM365 le role Entra `Global Reader` pour Exchange Online app-only, et retirer l'ancien role `Exchange Administrator` du service principal s'il avait ete ajoute par une version precedente du bootstrap. Ce scope est utilise uniquement par l'administrateur qui execute le bootstrap et n'est pas ajoute a l'app SmartM365. |
| `Sites.FullControl.All` | Attribuer le role `write` au service principal SmartM365 sur le site SharePoint cible avec `Sites.Selected`. Ce scope est utilise uniquement par l'administrateur qui execute le bootstrap et n'est pas ajoute a l'app SmartM365. |

## Points a revoir plus tard

- Remplacer les permissions Intune `ReadWrite` par des permissions `Read` lorsque `SmartM365-Get-IntuneAutopatchAlerts.ps1` est confirme en lecture seule.
- Verifier que tous les uploads SharePoint restent compatibles avec `Sites.Selected`; ne reintroduire `Files.ReadWrite.All` ou `Sites.ReadWrite.All` qu'en dernier recours documente.
- Revoir plus tard si le role Entra `Global Reader` peut etre remplace par une attribution Exchange RBAC encore plus fine lorsque les besoins exacts des scripts EXO sont stabilises.
- Si le site SharePoint de l'equipe Teams n'est pas disponible immediatement, relancer le bootstrap plus tard: la provision SharePoint d'une equipe Teams est asynchrone cote Microsoft 365.
