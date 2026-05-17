# SmartM365 App Registration Permissions

Ce document explique les permissions ajoutees par `SmartM365-Create-AppRegistration.ps1`, pourquoi elles sont necessaires, et quels scripts les utilisent actuellement.

## Notes importantes

- Les permissions ci-dessous sont des permissions **Application** Microsoft Graph pour l'app registration SmartM365.
- `SmartM365-Create-AppRegistration.ps1` s'execute uniquement en authentification interactive deleguee: un administrateur se connecte avec les droits necessaires pour creer ou mettre a jour l'app registration.
- Les scripts app-only utilisent `https://graph.microsoft.com/.default` ou `Connect-MgGraph -ClientId ...`; les permissions effectives sont donc celles accordees a l'application dans Entra ID.
- Les scripts avec `-InteractiveAuth` peuvent demander des scopes delegues equivalents, mais l'app registration doit surtout porter les permissions application pour les executions sans surveillance.
- Les uploads SharePoint sont centralises par `Modules/SmartM365.Core` et `Modules/SmartM365.SharePoint`; ils peuvent etre utilises par plusieurs scripts lorsque `EnableSharePointUpload` est active.

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
| `Files.ReadWrite.All` | Upload et remplacement de fichiers CSV dans une bibliotheque SharePoint via Microsoft Graph. Necessaire avec l'implementation actuelle qui ecrit dans un drive/document library. | `Modules/SmartM365.Core/SmartM365.Core.psm1`; `Modules/SmartM365.SharePoint/SmartM365.SharePoint.psm1`; tous les scripts d'inventaire/export qui appellent l'upload SharePoint quand `EnableSharePointUpload` est active. |
| `Mail.Send` | Envoi des notifications d'erreur et rapports HTML via Microsoft Graph lorsque `SmtpServer` est vide. L'adresse d'expedition est resolue depuis `From` dans `SmartM365.global.local.json` ou le `*.local.json` du script. | `Modules/SmartM365.Core/SmartM365.Core.psm1`; scripts qui appellent `SendEmailHtmlReport`, `Send-SmartM365Mail` ou `SendFileListEmailReport`. |
| `Sites.ReadWrite.All` | Resolution du site SharePoint, lecture des drives/libraries et ecriture dans le site cible pour les exports CSV. | `Modules/SmartM365.Core/SmartM365.Core.psm1`; `Modules/SmartM365.SharePoint/SmartM365.SharePoint.psm1`; tous les scripts d'inventaire/export qui appellent l'upload SharePoint quand `EnableSharePointUpload` est active. |

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
| `Exchange.ManageAsApp` sur l'API `Office 365 Exchange Online` | Autorise l'authentification app-only par certificat avec `Connect-ExchangeOnline`. Ce n'est pas une permission Microsoft Graph. Elle ne suffit pas seule: Exchange Online doit aussi recevoir le RBAC adapte pour le service principal. | Scripts Exchange Online sous `ExchangeInventory`, par exemple `AcceptedDomains`, `BackupProtection`, `CalendarPermissions`, `Mailboxes` et `Migration`, lorsqu'ils utilisent l'authentification app-only. |

## Permissions utilisees uniquement pour executer le bootstrap

Ces scopes sont demandes a l'administrateur qui lance `SmartM365-Create-AppRegistration.ps1`. Ils ne sont pas ajoutes a l'app SmartM365 comme permissions metier; ils servent a creer/modifier l'app registration et a accorder le consentement.

| Scope de connexion | Pourquoi il est demande |
| --- | --- |
| `Application.ReadWrite.All` | Creer ou mettre a jour l'app registration, ses API permissions et ses certificats publics. |
| `AppRoleAssignment.ReadWrite.All` | Accorder le consentement admin sous forme d'app role assignments. `SmartM365-Create-AppRegistration.ps1` le fait par defaut; utiliser `-DisableGrantAdminConsent` uniquement pour preparer l'app sans consentement immediat. |
| `Directory.Read.All` | Lire les service principals des APIs Microsoft Graph et Exchange Online, verifier le contexte tenant, et retrouver l'utilisateur administrateur connecte. |
| `Group.ReadWrite.All` | Creer ou reutiliser le groupe Microsoft 365 qui porte l'equipe Teams `SMART-M365`, puis convertir ce groupe en team. |
| `Sites.Read.All` | Lire le site SharePoint associe a l'equipe Teams `SMART-M365` afin de renseigner `SharePointSiteHostname` et `SharePointSitePath` dans `SmartM365.global.local.json`. |

## Points a revoir plus tard

- Remplacer les permissions Intune `ReadWrite` par des permissions `Read` lorsque `SmartM365-Get-IntuneAutopatchAlerts.ps1` est confirme en lecture seule.
- Evaluer `Sites.Selected` pour SharePoint afin de limiter l'upload au site cible au lieu de `Sites.ReadWrite.All`, mais cela demande une implementation et une attribution supplementaires.
- Les scripts Exchange Online ont besoin d'un RBAC Exchange explicite pour le service principal, meme avec `Exchange.ManageAsApp`.
- Si le site SharePoint de l'equipe Teams n'est pas disponible immediatement, relancer le bootstrap plus tard: la provision SharePoint d'une equipe Teams est asynchrone cote Microsoft 365.
