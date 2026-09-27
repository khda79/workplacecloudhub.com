# SmartInventory : transition des noms JSON

27 septembre 2026. Candidat local fondé sur `b1ef27654e0fff2e9deb498fa733be4e264cc9c6`, préparé pour publication Git sur main ; aucun déploiement ni migration réelle. Le checkout d'origine et ses changements concurrents sont préservés. L'ancien candidat `smartinventory-json-txt` est obsolète ; seul `json-txt-current` porte cette implémentation.

L'implémentation et les tests synthétiques sont réalisés. La qualification UNC, SharePoint, OneDrive et Power BI reste à effectuer. Les sources PowerShell du lot sont signées avant publication et vérifiées sur une extraction Git. À la demande explicite de Khaled, la politique est désormais `JsonText`. Les listes publiques restent vides : les destinations effectivement qualifiées sont déclarées dans la configuration locale privée décrite ci-dessous. Aucun script propriétaire ni migration réelle exécuté par Codex. Après récupération Git, les propriétaires exécutés peuvent migrer leurs fichiers locaux admissibles.

## Règles communes

Qualification opérateur reçue : les deux emplacements UNC ont passé 14 contrôles sur chacun des trois serveurs ; le test SharePoint a confirmé octets, identifiants et versions préservés, sans suppression. La copie OneDrive du fichier et de son archive a été vérifiée localement avec le SHA-256 attendu. Cela ne constitue pas un lancement réel des collecteurs ni une mesure de leur fraîcheur.

Les destinations qualifiées se configurent maintenant dans `Config/SmartM365-JsonTransport.policy.local.json.txt`, ignoré par Git, pour préserver les chemins et identifiants privés. Ce fichier accepte uniquement `QualifiedUncRoots` et `QualifiedSharePointDrives` et ne peut pas changer le mode public. Il est lu sans migration ni réécriture ; nouveau invalide ou doublons divergents restent bloquants. Aucun chemin réseau n'est accédé pour charger cette liste.

- `Readers` reconnaît les deux noms ; `JsonText` active les nouvelles écritures et la transition intégrée à chaque propriétaire. Aucune conversion manuelle préalable imposée.
- `.json.txt` prioritaire ; repli uniquement en son absence. Nouveau fichier invalide, inaccessible, répertoire à sa place ou doublon divergent : erreur explicite, sans reprise silencieuse de l'ancien.
- Persistants : déplacement des octets sans re-sérialisation, validation du propriétaire, SHA-256, verrou commun aux deux noms, journal durable et reprise. Doublons identiques retirés sous contrôle du propriétaire ; divergents conservés et bloquants.
- Régénérés : publication atomique du nouveau fichier avant traitement de l'ancien. Reçu avec les deux empreintes pour reprendre une interruption après publication. Une nouvelle collecte peut produire un nouveau contenu ; la migration seule ne le recalcule jamais.
- Écritures atomiques avec empreinte attendue ; preuves `.pending`, `.previous`, `.lock`, `.log` préservées en cas d'échec. Leur extension n'est pas `.json`.
- Des noms anciens restent dans le code pour la compatibilité et dans les contenus historiques pour préserver leurs octets. Les lecteurs résolvent les deux noms.
- Un verrou local n'est pas un verrou distribué OneDrive. Les chemins liés/reparse sont refusés : d'éventuels placeholders doivent être qualifiés sans contourner ce garde-fou.

## Inventaire par famille

Chemins relatifs aux racines configurées ; `.json` désigne le nom historique, cible avec suffixe `.txt`. Aucune configuration privée incluse.

| Famille / producteur | Emplacement et rôle | Consommateurs | Transfert et rétention | Décision |
|---|---|---|---|---|
| TenantContext, Core, collecteurs, Setup | Global/tenant `Config/*.local.json`, configurations adjacentes | Tous les lecteurs concernés, Setup, Exchange Notifications et Migration Readiness | Peut résider dans OneDrive ; aucun upload ni purge ajouté | `.json.txt`, migration propriétaire, création atomique et enrichissement avec contrôle de concurrence ; aucune exclusion générale des configurations |
| Templates | `*.json.template`, initialisation | Configurations et manifestes | Sources versionnées | Conserver : extension réelle `.template`, pas `.json` ; résolution adaptée |
| Core et compatibilité PS5 | `WeeklyHistory/<année>/<semaine>/manifest.json`, CSV associés | Collecteurs, publication, lecteurs historiques | Upload des manifestes et CSV ; reprise des manifestes convertis ; rétention existante après publication réussie | `.json.txt`, prévalidation puis migration de toutes les semaines reconnues, sans recalcul |
| AD | Manifestes de snapshots hebdomadaires | Collecteur AD et publication | Origine et CSV requis vérifiés ; publication avant rétention | `.json.txt`, schéma propriétaire, octets conservés |
| Licensing | Manifestes hebdomadaires, liste des CSV | Collecteur et nettoyage de doublons CSV | Migration/validation avant nettoyage ; capture conservée | `.json.txt`, modification métier ultérieure distincte de la migration |
| HybridIdentity et Windows11 Issues | Manifestes spécialisés ScriptName/Tenant | Collecteurs et lecteurs de snapshots | Reprise d'upload même si snapshot complet ; rétention conservée | `.json.txt`, historique migré sous contrôle du propriétaire |
| DiscoveredApps | Checkpoint, cache JSON, CSV partiel | Reprise du collecteur | Compteurs et empreintes conservés ; checkpoint incompatible archivé par SHA, CSV conservé | `.json.txt`, checkpoint persistant déplacé, cache régénéré publié avant retrait ancien |
| Intune Remediations | Exports horodatés : ExportInfo, IntuneRemediations, Scripts/<id>/Metadata et Assignments | Exporteur et consommateurs de restauration | Timestamp, tenant, chemins et GUID vérifiés ; export non qualifiable exclu de la rétention | `.json.txt`, prévalidation de l'export entier ; chemins embarqués historiques inchangés, lecteurs externes à qualifier |
| Orchestrator Management | Jobs, cluster, archives avant/après, PublicationFailed | Orchestrator, GUI, Pipeline, launchers indirectement | Archives exactes, aucune purge nouvelle | `.json.txt`, publications atomiques et archives reconnues migrées |
| Orchestrator persistant | State et SharePointMirrorState sous DATA-ALL/Orchestrator | Reprise, GUI, miroir | État invalide jamais remplacé par un état vide | `.json.txt`, validation de schéma, identifiants et empreintes conservés |
| Orchestrator régénéré | Heartbeat, Capabilities, ElectionPlan | Pairs, GUI, élection | Miroir opérationnel, rétention existante | Nouveau `.json.txt` publié avant retrait de l'ancien |
| Orchestrator commandes | StopRequested, RebalanceRequest | Processus propriétaire, GUI/launchers | Consommation verrouillée avec empreinte et journal | `.json.txt`, arrêt gracieux inchangé, aucune suppression générique |
| Distributed et Pipeline | Election/Claims/<job>/<slot>, leases, états de lots/jobs, demandes | Coordination, reprise, miroir | Verrous communs ; archive/retrait de leases contrôlés | `.json.txt`, historiques migrés, mises à jour avec empreinte attendue ; anciens écrivains interdits après activation |
| Core PS7/PS5, miroir SharePoint | JSON publiés et index | Upload, download, OneDrive | Filtres adaptés, doublon identique envoyé une fois ; ancien index conservé jusqu'au remplacement vérifié | Renommage par item ID/eTag, SHA et versions vérifiés ; aucune suppression globale |
| WorkplaceEvidence | Pointeur courant, manifestes, validations, reçus publication/échec/retrait, marqueurs | Pipeline, validation, transfert, Power Query | Verrou de publication ; pointeur après lots ; rétention captures et TransferOnly/ExpectedBatchId conservés | Migration propriétaire intégrée, SHA/tenant/batch vérifiés, chaînes historiques conservées |
| Audits WorkplaceEvidence | transfers, workforce-diagnostics, DATA-REPAIR-BACKUPS | Transfert, diagnostic, réparation | Racines et reçus vérifiés, pas de réparation ni diagnostic réel imposé | `.json.txt`, anciens audits reconnus migrés lors de l'exécution concernée |
| Contrats WorkplaceEvidence | config/prepared-source-contract et prepared-evidence-contract | Pipeline, validation, réparation | Sources versionnées ; aucune donnée privée ni rétention | Deux sources renommées `.json.txt`, octets inchangés ; retirer les deux anciens chemins seulement après contrôle du paquet |
| Power Query | GetPreparedEvidenceFile.pq | Modèle consommant les données préparées | Aucun rapport publié | Lecture compatible avec détection des conflits ; qualification Power BI encore requise |
| Power BI natif | Noms imposés version.json, report.json, page.json, visual.json | Desktop/PBIR | Aucun changement de politique OneDrive | **Exception validée par Khaled le 27/09/2026** pour les fichiers natifs uniquement ; aucun renommage |
| Fixtures et temporaires | Nouveaux dossiers GUID sous TEMP ou qualification dédiée | Tests uniquement | Ancien format créé volontairement ; fixtures conservées hors DATA réel | Ce ne sont pas des sorties de collecte ni des exceptions pour les données opérationnelles |

Les autres produits autonomes du dépôt ne sont pas renommés globalement. Les dépendances identifiées ci-dessus sont incluses. [Tableau des 45 collecteurs](JsonTransport-Collectors.md).

## SharePoint et cas exigeant une décision

Un renommage Graph ne fournit pas les garanties NTFS/SMB. Le code utilise `PATCH name` sur le même item avec `If-Match`, puis vérifie ID, octets et présence des versions antérieures. Doublons identiques : ancien item renommé `*.legacy-<empreinte-ID>.json.txt` pour conserver aussi ses propres versions. Doublons différents : blocage sans fusion ni suppression. Sources : [update driveItem](https://learn.microsoft.com/en-us/graph/api/driveitem-update?view=graph-rest-1.0), [list versions](https://learn.microsoft.com/en-us/graph/api/driveitem-list-versions?view=graph-rest-1.0).

Les situations suivantes restent des **arrêts de sécurité, pas des exceptions approuvées** :

1. Copie distante sans fichier local : miroir conserve et signale. Établir une liste fermée depuis les reçus/index du propriétaire, vérifier ID/eTag/SHA/versions et soumettre cette liste avant traitement. Aucune suppression globale.
2. Export sans ExportInfo, lot préparé sans reçu d'appartenance, schéma historique inconnu ou audit déplacé : propriété non démontrable. Conserver et signaler, établir une règle spécifique sur copie synthétique, puis soumettre l'exception ou la récupération à Khaled. Ne pas inventer le propriétaire ni recalculer les données.
3. Conflit de contenu, journal sans événement durable, modification concurrente, verrou indisponible ou filesystem non qualifié : opération arrêtée, preuves conservées. La décision du contenu faisant autorité reste explicite.

Le dépôt ne permet pas de déterminer si ces situations existent sur les serveurs. Aucune inspection des données privées n'a été faite ; aucune disparition complète des anciens JSON distants n'est revendiquée.

La seule exception technique approuvée concerne les fichiers natifs Power BI : [documentation Microsoft](https://learn.microsoft.com/power-bi/developer/projects/projects-report). Elle ne couvre pas données, contrats ou audits.

## Ordre de déploiement proposé initialement

L'activation de `JsonText` a depuis été approuvée explicitement. Cela ne confirme ni la qualification des destinations distantes ni le lancement des collectes. Les étapes ci-dessous restent la référence de qualification ; le mode du dépôt est maintenant actif.

1. Conserver les tâches désactivées et les trois machines sans écrivain Orchestrator/collecteur, y compris lancements manuels et consommateurs externes.
2. Examiner la liste de fichiers et empreintes ; préparer signatures et paquet isolé approuvé, sans configurations privées. Ne pas déployer l'ancien candidat.
3. Qualifier le paquet signé sur les trois machines, puis SMB, SharePoint dédié et OneDrive : [commandes](JsonTransport-Qualification.md). Ne pas utiliser l'entrée Orchestrator, même DryRun : en phase active, l'initialisation peut migrer des états.
4. Déployer lecteurs avant producteurs actifs : transport partagé/Core/PS5, TenantContext et lecteurs de configuration, Orchestrator/GUI/Pipeline/Distributed, WorkplaceEvidence/Power Query et lecteurs externes. Garder `Readers`. Relancer les sessions seulement selon le plan approuvé pour charger les nouveaux modules.
5. Vérifier de nouveau l'arrêt coordonné et l'absence d'ancien écrivain. La compatibilité de lecture n'autorise pas un parc mélangeant anciens et nouveaux producteurs ; des sessions en mémoire peuvent conserver d'anciens chemins.
6. Après approbation active, définir `JsonText` et uniquement les UNC/drives qualifiés. Le lancement approuvé de chaque propriétaire migre ses historiques/états et produit les nouveaux fichiers. Pas de conversion manuelle séparée obligatoire.
7. Contrôler reçus, SHA, complétude, fraîcheur, historique et synchronisation avant approbation de remise en service. Retour arrière : arrêter les écrivains et restaurer un lecteur compatible, jamais reprendre silencieusement un ancien `.json`.

## Preuves et limites

Correctif après premier démarrage réel : la libération puis réutilisation d'une même clé de concurrence utilisait deux libellés de propriétaire de journal. Le propriétaire des leases est unifié, avec reconnaissance explicite du seul ancien libellé après validation du schéma et du chemin ; les journaux étrangers restent refusés et inchangés, même après plusieurs essais. Les anciens journaux d'échec correspondants sont repris automatiquement, sans suppression manuelle.

La pagination des versions SharePoint ne compare plus le préfixe textuel encodé de l'URL. Elle vérifie HTTPS, hôte Graph, drive et item exacts après décodage, accepte les routes équivalentes avec segments ou clés OData et conserve le lien de pagination intact. Les boucles et changements de destination restent refusés. Cette correction cible le refus observé sur les fichiers possédant un historique paginé ; sa confirmation réelle nécessite le chargement du correctif et une nouvelle publication réussie de l'état et du heartbeat.

Validation ciblée du correctif : Distributed 17, SharePoint simulé 38 sous PS7 et PS5, transport PS5 34, lifecycle 22 et recovery 17, plus régression du miroir. Core 1.0.59, compatibilité PS5 1.0.43, JsonTransport 1.0.2, SharePointJsonTransition 1.0.1, Distributed 1.1.7. Arrêter proprement les processus avant mise à jour des sources partagées ; ne pas effacer leurs fichiers de coordination.

- Campagne consolidée : **35 suites / 35 réussies**, dont 7 sous Windows PowerShell 5.1. Les résultats détaillés sont conservés localement, hors Git.
- Recontrôles après corrections : lecteurs collecteurs **93**, SharePoint simulé **26 par moteur**, lifecycle Orchestrator **22/22**, recovery **17/17**. Le journal tronqué et la réponse Graph perdue sont couverts.
- Filesystem **sur disque local seulement** : **14 contrôles par moteur**. Cela ne qualifie pas SMB, OneDrive ni un lecteur distant concurrent.
- Analyse statique : **313 fichiers PowerShell sans erreur de syntaxe**, 6 composants ciblés sans diagnostic PSScriptAnalyzer de sévérité Error, contrôle des espaces Git sans erreur. Les 6 blocs de commandes de qualification ont aussi été analysés sans erreur de syntaxe.
- Deux suites complètes préexistantes échouent également sur la base inchangée : Pipeline (dépendance Full vers Backup désactivé/manual), Distributed (groupe sans affectation dans la topologie simulée). Exclues des 35 vertes, à traiter séparément avec la configuration réelle. Aucun job activé pour masquer ces échecs.
- Aucune connexion Graph réelle, collecte, migration privée, observation de synchronisation OneDrive ou exécution Power Query dans Power BI. Aucun accès au serveur ni au partage de production.
- Les tests synthétiques ne prouvent pas la couverture ni la fraîcheur des collectes réelles.

[Liste fermée et empreintes du candidat du lot Git signé](JsonTransport-CandidateFiles.csv) : 78 modifications, 29 ajouts, 2 anciens chemins de contrats retirés (contenu conservé dans les deux nouveaux noms). Le manifeste lui-même est exclu de ses propres empreintes. Le manifeste décrit les octets attendus après extraction Git ; aucun fichier privé inclus.

La revue automatique avait refusé une tentative de rétablir la suppression du CSV partiel DiscoveredApps en cas de checkpoint incompatible, pour risque de perte irréversible. La solution appliquée conserve le CSV et archive le checkpoint par empreinte ; aucune autorisation de suppression n'est demandée.
