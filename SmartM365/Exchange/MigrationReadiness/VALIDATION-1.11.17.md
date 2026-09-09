# MigrationReadiness 1.11.17 — validation candidate

Audit local du 8 septembre 2026. Statut : **publication de la préversion approuvée par l’utilisateur le 9 septembre 2026**. Cet accord ne constitue pas un résultat de test Live.

## Défauts démontrés et corrections

| Cas observé en 1.11.16 | Correction et preuve 1.11.17 |
| --- | --- |
| Graph Authentication 2.38.1 chargé avec Users/DirectoryManagement 2.37.0 : FileLoadException, assembly déjà chargé. | Choix de la version commune 2.37.0 sur cet hôte. Import des trois modules réussi sans authentification ; installation proposée alignée. |
| Un payload Graph manquant retourne une collection vide sans erreur ni horodatage réel. | QueryError explicite et timestamp absent ; la recherche vide réellement réussie reste distinguée. |
| Une erreur RBAC Full Access est avalée et retourne zéro permission. | Les erreurs et cmdlets/propriétés manquantes arrêtent la comparaison, sans générer de faux droits manquants. |
| Tous les SID et les noms contenant SELF sont filtrés côté EXO. | Conservation des SID utilisateurs et des noms ordinaires ; filtre exact des identités techniques. |
| Une ACE Full Access Deny est exportée comme autorisation locale/EXO. | Exclusion des refus des listes d'autorisations ; test du bloc réel du worker local extrait par AST. |
| Un probe EXO réussi sans métadonnées peut valider une session de tenant inconnu ; plusieurs sessions sont ambiguës. | Identité tenant vérifiable et session unique exigées ; reconnexion sinon. Aucun appel cloud réalisé pour ce test. |
| LICENSE-CAPACITY exige un siège supplémentaire même lorsque le SKU cible est déjà attribué. | Aucun double comptage ; une affectation absente sans siège libre reste bloquante et une capacité inconnue reste UNKNOWN. |

L'import CSV UTF-8 avec BOM a été testé et fonctionne déjà ; aucune modification de cet import.

## Tests effectués

- PowerShell 7.6.5 : **31 assertions de régression réussies**, incluant les entrées CSV, les cas ci-dessus et le pipeline réel d'assessment/export avec sources indisponibles.
- GUI `-ValidateOnly` : XAML et JSON valides, sérialisation CLIXML Windows PowerShell 5.1.19041.6456 réussie, 4 cas de verdict et 10 cas de readiness intégrés réussis.
- Les tests et l'entrée GUI signés passent avec `-ExecutionPolicy AllSigned`. Le self-test enfant PS5.1 utilise le mécanisme interne existant `Bypass` : ce n'est pas une certification AllSigned de tous les processus enfants.
- GraphWorker `-ValidateOnly` : import réel Authentication/Users/DirectoryManagement 2.37.0 réussi avec la politique locale habituelle, sans connexion tenant. L'essai AllSigned de ces dépendances Microsoft échoue sur le fichier installé `Microsoft.Graph.Authentication.format.ps1xml` (AuthorizationManager). Aucun module tiers ni politique de sécurité n'a été modifié pour contourner cette limite.
- Rapport synthétique : INCOMPLETE et verdict non positif en absence de sources ; huit CSV relus, neuf feuilles Excel et XML OpenXML valides, statut et identité de test présents dans le HTML, échappement HTML vérifié.
- Le rendu illustratif provient du XAML réel avec données synthétiques. Aucun assessment Live ni comparaison de permissions sur tenant réel n'a été exécuté.

## Prérequis et limites

EXO 3.9.2 est installé localement. ActiveDirectory et le snap-in Exchange ne sont pas présents sur cet hôte : droits AD, RBAC EXO, endpoint hybride, topologie Exchange et synchronisation tenant restent **non validés en Live**. `-ValidateOnly` ne remplace pas ces vérifications opérateur sur un hôte adapté.

La comparaison des permissions reste textuelle : SID/DN/UPN différents pour une même identité peuvent nécessiter un rapprochement manuel. Elle ne calcule pas les droits effectifs issus des groupes/héritages. Les rapports et logs ne sont pas anonymes. La compatibilité de détection Exchange 2016/2019 ne signifie pas que Microsoft prend encore ces versions en charge.

## Site et publication préparés

- Guide dédié en EN/FR/IT/ES/DE/AR : prérequis, installation, permissions, rapports, limites et illustration synthétique. Traduction stricte réussie pour cette page dans les cinq langues secondaires.
- Contrôle navigateur des six versions à largeur mobile : aucun débordement horizontal ni image cassée ; arabe RTL. Revue visuelle française et arabe.
- Delta préparé de **14 fichiers** : six nouvelles pages produit, six liens de catalogue modifiés uniquement pour ce produit, une image, un sitemap avec six nouvelles URL et les 168 URL existantes préservées.
- Les six catalogues publics ont répondu HTTP 200 ; les six nouvelles routes étaient absentes (HTTP 404). Empreintes du contenu public et des candidats enregistrées séparément. Le reste des traductions du site n'est pas certifié par cet audit.
- Les modifications web sont préparées dans une copie isolée et un patch ; les fichiers du projet web partagé conservent les travaux concurrents. Avant transfert, vérifier à nouveau les empreintes sources et distantes ; ne pas publier un build global non audité.

## Accord de publication et qualification restante

Le 9 septembre 2026, l’utilisateur a validé le périmètre Git, la première release et le delta web préparés. La [préversion 1.11.17](https://github.com/khda79/workplacecloudhub.com/releases/tag/exchange-migration-readiness-v1.11.17) conserve les limites de test ci-dessus. Tester le fonctionnement diagnostique sur l’hôte Exchange prévu avant toute utilisation opérationnelle. Aucun lancement de migration, changement tenant ou LinkedIn n’est autorisé par cet accord.

Le package portable, ses empreintes, le manifeste de publication, le patch web et le dossier de validation sont préparés localement. Les exports de test et journaux restent hors du périmètre Git.
