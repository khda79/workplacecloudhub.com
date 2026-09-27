# Commandes opérateur : qualification séparée du déploiement

Le lot est publié dans main à la demande de Khaled pour qualification. Récupérer le commit indiqué, dans une copie de test distincte des fichiers de production. Les essais distants restent à lancer par l'opérateur après son accord. Les commandes ci-dessous ne démarrent aucun collecteur ni Orchestrator et ne changent pas la politique de migration. Les essais UNC/SharePoint sont des opérations réelles sur des fixtures synthétiques dédiées ; ils ne sont pas équivalents aux tests locaux déjà réussis.

Les trois tâches restent désactivées. Aucun accès au serveur ni au partage n'a été fait par Codex. Les chemins et identifiants saisis restent locaux à votre session. Exécuter les blocs entiers pour que toute erreur arrête le bloc ; ne pas continuer ligne par ligne après une erreur.

## 1. Sur chacune des trois machines : contrôle et tests locaux

Ouvrir PowerShell 7 dans une session dédiée. Fournir la racine du checkout Git de qualification (contenant SmartM365 et SmartWorkplaceIntelligence). Les sources du lot sont signées avant publication ; le contrôle ci-dessous refuse une source PowerShell à signature invalide ou absente.

```powershell
& {
    $ErrorActionPreference = 'Stop'
    $package = Read-Host 'Racine locale du checkout Git de qualification'
    if ([string]::IsNullOrWhiteSpace($package)) { throw 'Chemin requis' }
    $package = (Resolve-Path -LiteralPath $package).Path
    if ($package.StartsWith('\\')) { throw 'Utiliser une copie locale du paquet de qualification' }
    $manifestPath = Join-Path $package 'SmartM365\Tests\JsonTransport-CandidateFiles.csv'
    $manifest = @(Import-Csv -LiteralPath $manifestPath)
    foreach ($entry in $manifest) {
        $file = Join-Path $package $entry.Path
        if ($entry.Action -eq 'Remove') { continue }
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Fichier absent : $($entry.Path)" }
        if ((Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash -ne $entry.SHA256) {
            throw "Empreinte differente : $($entry.Path)"
        }
        if ([IO.Path]::GetExtension($file) -in @('.ps1','.psm1','.psd1')) {
            if ((Get-AuthenticodeSignature -LiteralPath $file).Status -ne 'Valid') {
                throw "Signature non valide : $($entry.Path)"
            }
        }
    }
    $task = Get-ScheduledTask -TaskPath '\WCH\' -TaskName 'SmartM365 Inventory Orchestrator - prod' -ErrorAction Stop
    if ($task.Settings.Enabled -or $task.State -eq 'Running') { throw 'Tache encore active' }
    $running = @(Get-CimInstance Win32_Process | Where-Object {
        $_.Name -in @('pwsh.exe','powershell.exe') -and
        $_.CommandLine -match 'SmartM365-Inventory-Orchestrator\.ps1|SmartInventory[\\/]'
    })
    if ($running.Count) { $running | Select-Object ProcessId,ParentProcessId,CommandLine; throw 'Processus a examiner avant qualification' }
    $policy = Import-PowerShellDataFile (Join-Path $package 'SmartM365\Config\SmartM365-JsonTransport.policy.psd1')
    if ($policy.Mode -notin @('Readers','JsonText')) { throw 'Mode de politique inconnu' }
    Write-Host ('Mode du checkout : ' + $policy.Mode + ' ; tests synthetiques uniquement')
    & (Join-Path $package 'SmartM365\Tests\Invoke-SmartM365JsonSyntheticQualification.ps1') -IncludeWindowsPowerShell
    if ($LASTEXITCODE -ne 0) { throw 'Echec de qualification synthetique' }
}
```

Le CSV décrit les fichiers signés et les octets attendus dans un checkout Git respectant .gitattributes. Vérifier le commit récupéré ; le manifeste ne s'auto-authentifie pas. Conserver le contrôle des signatures.

## 2. Sur chaque machine : qualification UNC dans un dossier dédié

Après approbation de ces écritures de test, le bloc crée uniquement le parent dédié s'il manque, puis le script crée un sous-dossier GUID et le conserve. Il crée volontairement des anciens `.json` synthétiques afin de tester leur conversion. Aucun fichier DATA existant n'est migré.

```powershell
& {
    $ErrorActionPreference = 'Stop'
    $package = Read-Host 'Racine locale du checkout Git de qualification'
    if ([string]::IsNullOrWhiteSpace($package)) { throw 'Chemin requis' }
    $probeParent = Read-Host 'Chemin UNC du dossier dedie SmartM365-Qualification'
    if (-not (Test-Path -LiteralPath $probeParent -PathType Container)) {
        New-Item -ItemType Directory -Path $probeParent -ErrorAction Stop | Out-Null
    }
    $test = Join-Path $package 'SmartM365\Tests\Test-SmartM365JsonFileSystemQualification.ps1'
    & $test -ProbeParentPath $probeParent -AllowUnc
    & powershell.exe -NoProfile -ExecutionPolicy AllSigned -File $test -ProbeParentPath $probeParent -AllowUnc
    if ($LASTEXITCODE -ne 0) { throw 'Echec de qualification UNC PS5' }
}
```

Résultat attendu : 14 contrôles par moteur, CSV et fixtures conservés. Ces essais exercent deux processus sur le partage depuis une machine. Ils ne prouvent pas à eux seuls la contention entre deux machines, une coupure SMB ni le comportement d'un antivirus/agent de synchronisation en charge. Avant activation, compléter avec une session tenant un verrou et une seconde machine tentant l'accès au même fichier synthétique, puis reprise après libération. La campagne locale couvre les interruptions de processus ; une coupure réseau réelle doit être faite seulement dans un environnement de test approuvé.

### Verrou entre deux machines

Sur A, exécuter ce bloc et conserver la fenêtre en attente. Communiquer le chemin affiché à B. Tout se passe dans un nouveau sous-dossier du parent de qualification.

```powershell
& {
    $ErrorActionPreference = 'Stop'
    $root = Read-Host 'Chemin UNC du dossier dedie SmartM365-Qualification'
    $folder = Join-Path $root ('CrossHost-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $folder | Out-Null
    $path = Join-Path $folder 'cross-host.lock'
    $lock = [IO.File]::Open($path,'OpenOrCreate','ReadWrite','None')
    try { Write-Host "Fichier a tester depuis B : $path"; $null = Read-Host 'Appuyer sur Entree APRES le premier test de B pour liberer le verrou' }
    finally { $lock.Dispose() }
}
```

Sur B, exécuter une première fois pendant le verrou, puis une deuxième fois après libération. Attendus : « Accès refusé » puis « Accès obtenu ». Un refus peut aussi provenir d'une permission ; seule la réussite après libération confirme le scénario complet.

```powershell
& {
    $ErrorActionPreference = 'Stop'
    $path = Read-Host 'Chemin exact cross-host.lock affiche par A'
    if ([IO.Path]::GetFileName($path) -ne 'cross-host.lock') { throw 'Chemin de fixture inattendu' }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw 'Fixture absente' }
    try { $lock = [IO.File]::Open($path,'Open','ReadWrite','None') }
    catch [IO.IOException] { Write-Host 'Acces refuse pendant le verrou ; refaire apres liberation'; return }
    try { Write-Host 'Acces obtenu : attendu uniquement apres liberation par A' }
    finally { $lock.Dispose() }
}
```

## 3. SharePoint réel, données synthétiques seulement

Préparer un dossier dédié **SmartM365-Qualification** dans le drive concerné avec vos outils habituels. Employer une connexion Graph déjà approuvée avec les droits existants ; le script ne se connecte pas et ne modifie aucune permission. Fournir les IDs exacts, pas une URL. Il garde les items et les journaux pour inspection et ne fait aucun DELETE.

```powershell
& {
    $ErrorActionPreference = 'Stop'
    $package = Read-Host 'Racine locale du checkout Git de qualification'
    $driveId = Read-Host 'DriveId de la bibliotheque a qualifier'
    $folderId = Read-Host 'ItemId du dossier SmartM365-Qualification'
    if ([string]::IsNullOrWhiteSpace($package) -or [string]::IsNullOrWhiteSpace($driveId) -or [string]::IsNullOrWhiteSpace($folderId)) { throw 'Parametres requis' }
    & (Join-Path $package 'SmartM365\Tests\Test-SmartM365JsonSharePointQualification.ps1') -DriveId $driveId -TestParentFolderId $folderId -AllowSyntheticRemoteWrites
}
```

Vérifier le CSV `qualification.csv` affiché : ID original conservé, versions antérieures vérifiées, octets identiques, aucun DELETE, politique non activée. Inspecter l'historique des deux items depuis SharePoint. L'absence d'erreur ne prouve pas la propagation OneDrive.

## 4. OneDrive et Power BI

Sur une machine qui synchronise le dossier SharePoint synthétique, attendre la fin de synchronisation puis comparer le fichier `state.json.txt` au SHA256 du rapport précédent. Répéter sur un second client si utilisé. Ne pas lancer le test filesystem dans l'ensemble du dossier OneDrive.

```powershell
& {
    $ErrorActionPreference = 'Stop'
    $file = Read-Host 'Chemin local du state.json.txt synchronise dans le dossier JsonTransport-GUID'
    $expected = (Read-Host 'ExpectedSHA256 du rapport SharePoint').Trim()
    if ([IO.Path]::GetFileName($file) -ne 'state.json.txt' -or $expected -notmatch '^[A-Fa-f0-9]{64}$') { throw 'Parametres invalides' }
    $actual = (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash
    if ($actual -ne $expected) { throw 'Octets synchronises differents' }
    Get-Item -LiteralPath $file | Select-Object FullName,Length,LastWriteTimeUtc
    Write-Host 'Empreinte conforme ; verifier aussi absence de conflit ou attente dans OneDrive'
}
```

Dans une copie de qualification Power BI : intégrer le lecteur `.pq` adapté, charger séparément un lot ancien et le même lot renommé sans changer ses octets, comparer nombre de lignes, colonnes, identifiants, batch et empreintes de sources. Vérifier que nouveau invalide et doublon divergent produisent une erreur au lieu d'un repli. Les fichiers PBIR natifs restent en `.json` conformément à l'exception approuvée. Aucun refresh de production ni publication de rapport autorisé par ce document.

## 5. Critères de feu vert et éléments à retourner

Après réussite, enregistrer les seuls chemins et drives effectivement qualifiés dans `Config/SmartM365-JsonTransport.policy.local.json.txt` à côté de la politique publique signée. Le module transport 1.0.1 fusionne ces listes privées avec la politique de déploiement ; ne pas modifier le fichier `.psd1` signé. Structure :

```json
{
  "QualifiedUncRoots": [],
  "QualifiedSharePointDrives": []
}
```

Remplacer les tableaux vides par les destinations vérifiées sur le site concerné. Ce fichier reste hors Git, comme toutes les configurations `.local.json.txt`. La lecture n'effectue aucune écriture ni accès aux destinations déclarées. Contrôle en session neuve : importer `SmartM365.JsonTransport.psd1`, puis appeler `Get-SmartM365JsonTransportPolicy`.

Conserver résultats CSV/logs, versions PowerShell et OS, empreintes du paquet signé, preuves des signatures, trois états de tâches et processus, rapports UNC/SharePoint, contrôles de propagation OneDrive et qualification Power Query. Ne pas transmettre de secrets de configuration.

Le feu vert demande aussi la validation des lecteurs externes de Remediations, des éventuelles copies distantes sans propriétaire local et des historiques incomplets. Aucune perte d'historique, baisse de couverture ou reprise de données périmées ne peut être acceptée pour faire passer les tests.

Ces commandes ne signent, ne déploient, ne démarrent de collecte et ne réactivent de tâche. `JsonText` est maintenant activé dans Git sur demande explicite. Les lancements réels, qualifications distantes et réactivation des tâches restent des opérations séparées.
