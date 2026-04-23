# Déploiement BitLocker TPM + PIN avec sauvegarde de la clé de récupération dans AD DS (UFCV)

---

## Vue d’ensemble

Ce dépôt contient un unique script PowerShell, [`BitLocker-Enable-TPM-PIN-Recovery_UFCV.ps1`](./BitLocker-Enable-TPM-PIN-Recovery_UFCV.ps1), qui met en œuvre un flux d’activation BitLocker pour un poste Windows joint à un domaine Active Directory.

Le script :

- vérifie que la stratégie BitLocker attendue est présente dans le registre ;
- valide que la machine appartient au domaine `ufcvfr.lan` et qu’un contrôleur de domaine de ce domaine est joignable ;
- refuse de démarrer si BitLocker est déjà en cours de chiffrement ou de déchiffrement sur `C:`, ou si le volume est déjà protégé ;
- affiche une interface WPF pour saisir et confirmer un PIN numérique ;
- crée ou réutilise un protecteur de mot de passe de récupération ;
- sauvegarde ce protecteur dans AD DS ;
- active BitLocker sur `C:` avec `TPM + PIN`, `XtsAes256` et `UsedSpaceOnly` ;
- gère un compteur de reports persistant ;
- affiche l’avancement dans une vue de progression pendant le provisioning.

> [!IMPORTANT]
> Le dépôt est manifestement destiné à un contexte interne UFCV. Cette spécialisation est visible dans le code par le domaine codé en dur, les valeurs de stratégie attendues, les chemins de stockage locaux et le texte des messages utilisateur.

---

## Ce que fait réellement le script

Le workflow réel, tel qu’implémenté, est le suivant :

1. Le script force l’encodage UTF-8 BOM et la culture `fr-FR`.
2. Il émet un avertissement si l’exécution n’est pas effectuée sous le SID `S-1-5-18` (`LocalSystem`).
3. Il charge les assemblies WPF nécessaires.
4. Il lit `HKLM:\SOFTWARE\Policies\Microsoft\FVE` et compare un ensemble de valeurs attendues.
5. Si au moins une valeur manque, diffère ou a un type inattendu, le script affiche un message d’inéligibilité et s’arrête.
6. Il vérifie ensuite le domaine courant via `GetCurrentDomain()` et tente de trouver un contrôleur de domaine joignable.
7. Il lit un compteur de reports persistant dans `C:\ProgramData\BitLockerActivation\PostponeCount.txt`.
8. Il vérifie l’état BitLocker du volume `C:`.
9. Il charge une fenêtre WPF avec deux zones de saisie, une vue de progression et des boutons de validation, report et fermeture.
10. Le PIN est validé côté UI et côté logique :
    - uniquement des chiffres ;
    - entre 6 et 20 caractères ;
    - non strictement croissant ;
    - non strictement décroissant ;
    - identique dans les deux champs.
11. En cas de validation, le provisioning est exécuté dans un runspace asynchrone.
12. Le runspace vérifie ou crée un `RecoveryPassword`, le sauvegarde dans AD DS, supprime d’éventuels protecteurs `TpmPin` existants, puis appelle `Enable-BitLocker`.
13. Si `Enable-BitLocker` renvoie l’erreur `0x80310060`, le script crée un fichier indicateur `PendingReboot.flag` et demande un redémarrage avant de relancer.
14. En cas de succès, le compteur de report est réinitialisé en supprimant `PostponeCount.txt`.
15. L’interface passe en mode progression pendant l’exécution, puis affiche un bouton de fermeture en fin de traitement.

> [!NOTE]
> Le script ne propose pas de mode silencieux, pas de paramètre CLI et pas de configuration externe. Toute la logique est codée dans le `.ps1`.

---

## Structure du dépôt

À la racine du dépôt, on trouve les fichiers suivants :

| Fichier | Rôle |
| --- | --- |
| [`BitLocker-Enable-TPM-PIN-Recovery_UFCV.ps1`](./BitLocker-Enable-TPM-PIN-Recovery_UFCV.ps1) | Script principal. Contient la logique de validation, l’UI WPF, le compteur de reports et le provisioning BitLocker. |
| [`README.md`](./README.md) | Documentation du dépôt. |
| [`.gitattributes`](./.gitattributes) | Normalisation LF des fichiers texte. |
| [`.gitignore`](./.gitignore) | Fichier d’exclusion Git pour les artefacts courants et pour `.repoignore`. |
| [`.repoignore`](./.repoignore) | Liste d’exclusion locale utilisée par l’outillage du dépôt ; elle ignore notamment `.git`, les répertoires de build, les caches, les fichiers temporaires et divers artefacts courants. |

Il n’y a pas de dossier de modules, pas de XAML séparé, pas de test automatisé et pas de manifeste d’installation.

---

## Flux d’exécution

### 1. Initialisation de l’environnement

Le script commence par :

- forcer `OutputEncoding` en UTF-8 BOM ;
- définir `fr-FR` comme culture, locale système et langue d’interface ;
- charger `PresentationFramework`, `PresentationCore` et `WindowsBase` ;
- avertir si l’identité Windows courante n’est pas `LocalSystem`.

Cette phase n’est pas entourée de blocs `try/catch` spécifiques. Si l’environnement refuse l’un de ces réglages, le script peut s’arrêter avant l’UI.

### 2. Contrôle de conformité BitLocker dans le registre

Le script ouvre `HKLM:\SOFTWARE\Policies\Microsoft\FVE` en lecture seule et compare la clé à une liste de valeurs attendues.

- Si la clé n’existe pas, toutes les valeurs sont considérées comme manquantes.
- Les chaînes sont comparées en insensible à la casse et acceptent `String` ou `ExpandString`.
- Les entiers sont comparés comme `DWord`.
- Toute différence de valeur, de type ou d’absence bloque le traitement.

Quand la comparaison échoue, une boîte de dialogue informe que le poste n’est pas éligible au déploiement BitLocker et le script se termine avec `exit 1`.

### 3. Validation du domaine et du contrôleur de domaine

Le domaine attendu est codé en dur à `ufcvfr.lan`.

Le script :

- récupère le domaine courant via `System.DirectoryServices.ActiveDirectory.Domain.GetCurrentDomain()`;
- récupère un contrôleur de domaine via `FindDomainController()`;
- vérifie que le domaine détecté correspond exactement à `ufcvfr.lan`;
- vérifie que le nom du contrôleur de domaine se termine bien par `.ufcvfr.lan`.

En cas d’échec, une boîte de dialogue indique que le poste n’est pas connecté au réseau UFCV ou n’est pas sur le bon domaine, puis le script s’arrête.

### 4. Lecture du compteur de reports

Le script lit `C:\ProgramData\BitLockerActivation\PostponeCount.txt`.

- Si le dossier n’existe pas, il est créé.
- Si le fichier n’existe pas, le compteur démarre à `0`.
- Le nombre maximal de reports est fixé à `99`.

Si la limite est atteinte, le script affiche un avertissement et désactive le report dans l’UI.

### 5. Vérification de l’état BitLocker sur `C:`

Avant d’afficher la fenêtre, le script appelle `Get-BitLockerVolume -MountPoint "C:"`.

Il interrompt l’exécution si :

- `VolumeStatus` vaut `EncryptionInProgress` ;
- `VolumeStatus` vaut `DecryptionInProgress` ;
- `VolumeStatus` vaut `FullyEncrypted` et `ProtectionStatus` vaut `On`.

### 6. Affichage de l’interface WPF

La fenêtre principale est créée à partir d’un XAML embarqué dans le script. Il n’existe pas de fichier XAML externe.

L’interface initiale affiche :

- deux champs `PasswordBox` pour le PIN et sa confirmation ;
- un compteur de reports restants ;
- un bouton `Valider` ;
- un bouton `Plus tard` ;
- un bouton de fermeture `×`.

La vue de progression est masquée au départ.

### 7. Validation du PIN

Le PIN est filtré et validé à deux niveaux.

Au niveau de la saisie :

- les caractères non numériques sont bloqués à la frappe ;
- les deux champs ont `MaxLength = 20`.

Au niveau de la logique :

- le PIN doit contenir uniquement des chiffres ;
- sa longueur doit être comprise entre 6 et 20 caractères ;
- il ne doit pas être strictement croissant ;
- il ne doit pas être strictement décroissant ;
- les deux champs doivent être identiques pour activer `Valider`.

Si la validation échoue au clic, le script affiche une boîte de dialogue explicite et ne lance pas le provisioning.

### 8. Provisioning asynchrone

Quand l’utilisateur valide le PIN, le script lance `Start-BitLockerProvisioningAsync`.

Le traitement de provisioning est exécuté dans un runspace séparé afin de ne pas bloquer l’interface. Un `DispatcherTimer` lit ensuite les messages du runspace et met à jour la vue de progression.

### 9. Finalisation

À la fin :

- le bouton `Fermer` devient visible ;
- la fermeture est réautorisée ;
- si le provisioning a réussi, le compteur de reports est supprimé ;
- si l’utilisateur a reporté, le compteur est incrémenté ;
- si l’erreur correspond à un PIN non encore autorisé par la stratégie, un fichier `PendingReboot.flag` est créé pour marquer l’état.

---

## Prérequis

> [!NOTE]
> Le script ne vérifie pas tous les prérequis possibles, mais son code suppose au minimum :

- un hôte Windows capable de charger WPF ;
- des cmdlets et assemblies disponibles pour `Set-Culture`, `Set-WinSystemLocale`, `Set-WinUILanguageOverride`, `Get-BitLockerVolume` et les cmdlets BitLocker ;
- un volume système sur `C:` ;
- une machine jointe au domaine `ufcvfr.lan` ;
- une configuration de stratégie BitLocker déjà appliquée dans `HKLM:\SOFTWARE\Policies\Microsoft\FVE` ;
- un accès à un contrôleur de domaine joignable ;
- un contexte d’exécution suffisamment privilégié pour lire la stratégie, modifier BitLocker et, si nécessaire, appliquer les réglages de locale ;
- un environnement compatible avec `TPM + PIN`, puisque `Enable-BitLocker` est appelé avec `-TpmAndPinProtector`.

Le code avertit si l’exécution n’a pas lieu en `LocalSystem`, mais il ne bloque pas pour autant.

---

## Hypothèses d’environnement codées en dur

Les valeurs suivantes sont fixes dans le script et doivent être adaptées avant réutilisation hors de l’environnement UFCV :

| Hypothèse | Valeur codée en dur | Impact |
| --- | --- | --- |
| Domaine AD attendu | `ufcvfr.lan` | Le script refuse de continuer hors de ce domaine. |
| Volume cible | `C:` | Le provisioning ne vise que le volume système. |
| Clé registre BitLocker | `HKLM:\SOFTWARE\Policies\Microsoft\FVE` | Toute la validation de conformité repose sur cette branche. |
| Dossier de stockage | `C:\ProgramData\BitLockerActivation` | Dossier utilisé pour le compteur de reports et le drapeau de redémarrage. |
| Fichier de report | `C:\ProgramData\BitLockerActivation\PostponeCount.txt` | Contient le nombre de reports déjà consommés. |
| Fichier drapeau | `C:\ProgramData\BitLockerActivation\PendingReboot.flag` | Créé quand la stratégie BitLocker n’autorise pas encore le PIN. |
| Nombre maximal de reports | `99` | Au-delà, le report est désactivé et la fermeture est bloquée tant que l’activation n’est pas lancée. |
| Méthode de chiffrement | `XtsAes256` | Méthode passée à `Enable-BitLocker`. |
| Culture et locale | `fr-FR` | Le script force les formats de langue et de culture français. |
| DLL Network Unlock | `C:\Windows\System32\nkpprov.dll` | Valeur attendue dans la stratégie `NetworkUnlockProvider`. |

Le script ne prend aucun de ces paramètres en ligne de commande.

---

## Validation des stratégies / du registre

> [!NOTE]
> La validation du registre est purement comparative. Le script ne corrige pas les valeurs, ne les crée pas et ne les écrit pas. Il s’en sert uniquement pour décider si le poste est éligible.

### Règles appliquées

- une valeur absente est signalée `MISSING` ;
- un type inattendu est signalé `TYPE_MISMATCH` ;
- une valeur différente est signalée `DIFF` ;
- seul l’état `OK` est accepté ;
- si au moins une valeur est `MISSING`, `DIFF` ou `TYPE_MISMATCH`, le script s’arrête avant l’UI BitLocker.

### Valeurs vérifiées

| Valeur de registre | Valeur attendue | Type attendu | Lecture dans le script |
| --- | --- | --- | --- |
| `NetworkUnlockProvider` | `C:\Windows\System32\nkpprov.dll` | `String` ou `ExpandString` | Chemin de la DLL Network Unlock attendu. |
| `OSManageNKP` | `1` | `DWord` | Valeur de stratégie comparée telle quelle. |
| `TPMAutoReseal` | `1` | `DWord` | Valeur de stratégie comparée telle quelle. |
| `EncryptionMethodWithXtsOs` | `7` | `DWord` | Méthode attendue pour le volume OS. |
| `EncryptionMethodWithXtsFdv` | `7` | `DWord` | Méthode attendue pour les lecteurs fixes. |
| `EncryptionMethodWithXtsRdv` | `4` | `DWord` | Méthode attendue pour les lecteurs amovibles. |
| `OSEnablePrebootInputProtectorsOnSlates` | `1` | `DWord` | Stratégie comparée telle quelle. |
| `OSEncryptionType` | `2` | `DWord` | Stratégie comparée telle quelle. |
| `OSRecovery` | `1` | `DWord` | Stratégie comparée telle quelle. |
| `OSManageDRA` | `1` | `DWord` | Stratégie comparée telle quelle. |
| `OSRecoveryPassword` | `2` | `DWord` | Stratégie comparée telle quelle. |
| `OSRecoveryKey` | `2` | `DWord` | Stratégie comparée telle quelle. |
| `OSHideRecoveryPage` | `1` | `DWord` | Stratégie comparée telle quelle. |
| `OSActiveDirectoryBackup` | `1` | `DWord` | Stratégie comparée telle quelle. |
| `OSActiveDirectoryInfoToStore` | `1` | `DWord` | Stratégie comparée telle quelle. |
| `OSRequireActiveDirectoryBackup` | `1` | `DWord` | Stratégie comparée telle quelle. |
| `ActiveDirectoryBackup` | `1` | `DWord` | Stratégie comparée telle quelle. |
| `RequireActiveDirectoryBackup` | `1` | `DWord` | Stratégie comparée telle quelle. |
| `ActiveDirectoryInfoToStore` | `1` | `DWord` | Stratégie comparée telle quelle. |
| `UseRecoveryPassword` | `1` | `DWord` | Stratégie comparée telle quelle. |
| `UseRecoveryDrive` | `1` | `DWord` | Stratégie comparée telle quelle. |
| `UseAdvancedStartup` | `1` | `DWord` | Stratégie comparée telle quelle. |
| `EnableBDEWithNoTPM` | `0` | `DWord` | Stratégie comparée telle quelle. |
| `UseTPM` | `0` | `DWord` | Stratégie comparée telle quelle. |
| `UseTPMPIN` | `1` | `DWord` | Stratégie comparée telle quelle. |
| `UseTPMKey` | `0` | `DWord` | Stratégie comparée telle quelle. |
| `UseTPMKeyPIN` | `0` | `DWord` | Stratégie comparée telle quelle. |

---

## Comportement de l’interface graphique

### Vue de saisie du PIN

La première vue est une fenêtre WPF sans bordure (`WindowStyle=None`), transparente, centrée, toujours au premier plan.

Elle contient :

- un bouton de fermeture `×` ;
- deux champs de saisie masquée ;
- un texte d’information ;
- un compteur de reports restants ;
- les boutons `Valider` et `Plus tard`.

### Règles de saisie

- seuls les chiffres sont acceptés à la frappe ;
- la validation est désactivée tant que les deux champs ne sont pas remplis ;
- la validation est activée uniquement si les deux PIN passent la règle de longueur et sont identiques ;
- le script compare les deux valeurs au moment du clic même si l’état visuel paraît correct.

### Indicateurs visuels

- bordure orange si le PIN est invalide ;
- bordure verte si le PIN est valide et que les deux champs concordent ;
- bordure rouge si les deux valeurs diffèrent ;
- bordure neutre si le champ est vide.

### Vue de progression

Lors du provisioning :

- la vue de saisie est masquée ;
- la vue de progression devient visible ;
- les boutons de fermeture, report et validation sont désactivés ;
- une barre de progression, un statut textuel et une liste de lignes de progression sont affichés ;
- le bouton `Fermer` n’apparaît qu’à la fin.

Le bouton `Fermer` est aussi activé après une erreur, ce qui permet de quitter l’écran de progression une fois l’état final affiché.

---

## Mécanisme de report

> [!NOTE]
> Le report n’est pas basé sur une durée, mais sur un compteur persistant.

### Stockage

- le compteur est lu et écrit dans `C:\ProgramData\BitLockerActivation\PostponeCount.txt` ;
- le dossier parent est créé si nécessaire ;
- le compteur est réinitialisé automatiquement après un provisioning réussi.

### Ce qui compte comme un report

Le compteur est incrémenté si :

- l’utilisateur clique sur `Plus tard` ;
- l’utilisateur ferme la fenêtre avant toute action explicite ;
- l’utilisateur ferme via le bouton `×` sans que le provisioning soit terminé.

### Ce qui ne compte pas comme un report

Le compteur n’est pas incrémenté si :

- le provisioning est terminé et l’utilisateur ferme ensuite la fenêtre ;
- l’utilisateur valide et le traitement se termine correctement ;
- l’utilisateur valide et le traitement termine sur une erreur autre qu’un report.

### Limite

- la limite est fixée à `99` ;
- si cette limite est atteinte, le bouton `Plus tard` est désactivé ;
- le bouton `×` est aussi désactivé ;
- la fermeture de la fenêtre est bloquée tant que l’utilisateur n’a pas lancé la configuration.

Le script considère alors l’activation BitLocker comme obligatoire.

---

## Logique de provisioning BitLocker

Le provisioning est implémenté dans le runspace `Start-BitLockerProvisioningAsync`.

### Vérification initiale

Le traitement revalide d’abord l’état du volume `C:` :

- chiffrement déjà en cours ;
- déchiffrement en cours ;
- volume déjà chiffré et protégé.

Dans ces cas, le runspace renvoie un état `already` et le script s’arrête sans modifier le volume.

### Protecteur de récupération

Le script cherche un protecteur de type `RecoveryPassword`.

- s’il en trouve un, il le réutilise ;
- s’il n’en trouve pas, il en crée un avec `Add-BitLockerKeyProtector -RecoveryPasswordProtector` ;
- il récupère ensuite l’ID du premier protecteur trouvé.

Le script ne supprime pas les autres protecteurs de récupération déjà existants.

### Sauvegarde AD DS

Le protecteur de récupération est ensuite sauvegardé avec `Backup-BitLockerKeyProtector`.

Cette étape est obligatoire dans le flux courant. Si elle échoue, le runspace remonte une erreur.

### Protecteur TPM + PIN

Avant l’activation finale :

- le PIN est converti en `SecureString` avec `ConvertTo-SecureString -AsPlainText -Force` ;
- tous les protecteurs existants de type `TpmPin` sont supprimés ;
- `Enable-BitLocker` est appelé avec :
  - `-EncryptionMethod XtsAes256`
  - `-UsedSpaceOnly`
  - `-TpmAndPinProtector`
  - `-Pin $UserPin`

Le script ne vérifie pas explicitement la présence du TPM avant cet appel.

### Succès

Un succès signifie que `Enable-BitLocker` a été lancé sans erreur dans ce flux.

Le script affiche ensuite qu’un redémarrage est requis pour finaliser et démarrer le chiffrement. Il ne suit pas l’état de chiffrement jusqu’à son achèvement complet.

### Cas particulier `0x80310060`

Si `Enable-BitLocker` échoue avec `0x80310060` ou avec le code HRESULT associé `-2144272384`, le script :

- crée `C:\ProgramData\BitLockerActivation\PendingReboot.flag` ;
- renvoie un état `policy_pending` ;
- demande à l’utilisateur de redémarrer puis de relancer le script.

Le drapeau est écrit par le script, mais il n’est pas relu ailleurs dans ce dépôt.

---

## Implémentation asynchrone / progression

L’UI ne bloque pas pendant le provisioning grâce à un runspace dédié.

### Mécanisme

- le runspace est créé avec `ApartmentState = "MTA"` ;
- le script principal injecte le code de provisioning sous forme de `ScriptBlock` ;
- les messages sont émis sous forme d’objets `pscustomobject` ;
- un `DispatcherTimer` WPF, déclenché toutes les `200 ms`, lit les objets reçus ;
- la barre de progression, le pourcentage et la liste des étapes sont mis à jour depuis le thread UI.

### Remontée de progression

Les messages de progression sont affichés sous trois formes principales :

- `⏳` pour les étapes en cours ;
- `✅` pour les étapes validées ;
- `⚠️` pour les avertissements.

L’état final est affiché via `Complete-Ui`, qui :

- affiche le texte final ;
- rend le bouton `Fermer` visible ;
- réautorise la fermeture ;
- marque le provisioning comme terminé.

---

## Gestion des erreurs et conditions bloquantes

Le script s’arrête ou bloque dans les cas suivants :

- stratégie BitLocker absente, différente ou incomplète dans le registre ;
- domaine différent de `ufcvfr.lan` ;
- aucun contrôleur de domaine joignable dans ce domaine ;
- chiffrement BitLocker déjà en cours sur `C:`;
- déchiffrement BitLocker déjà en cours sur `C:`;
- volume `C:` déjà chiffré et protégé ;
- échec de parsing du XAML ;
- contrôle XAML introuvable après chargement ;
- PIN invalide ;
- confirmation de PIN différente ;
- erreur de lancement du runspace ;
- erreur remontée dans le flux d’erreur PowerShell du runspace ;
- stratégie BitLocker n’autorisant pas encore le PIN (`0x80310060`) ;
- fermeture de la fenêtre pendant le provisioning ;
- fermeture de la fenêtre alors que la limite de reports est atteinte et que la configuration n’a pas été lancée.

Quelques points ne sont pas protégés par un traitement spécifique et peuvent donc faire échouer le script plus tôt que prévu :

- `Set-Culture`, `Set-WinSystemLocale` et `Set-WinUILanguageOverride` ;
- la conversion du compteur de reports en entier ;
- le chargement des assemblies WPF ;
- les cmdlets BitLocker elles-mêmes si elles ne sont pas disponibles.

---

## Notes de sécurité

Le code montre plusieurs comportements sensibles :

- le PIN est saisi dans des `PasswordBox`, mais il existe ensuite en mémoire dans les variables du script et dans l’argument passé au runspace ;
- la conversion en `SecureString` utilise `-AsPlainText -Force`, ce qui signifie que la valeur est d’abord manipulée en clair dans le processus ;
- le script ne persiste pas le PIN sur disque ;
- la sauvegarde du protecteur de récupération dans AD DS est obligatoire dans le flux normal ;
- la fermeture de la fenêtre est bloquée pendant le provisioning ;
- les protecteurs `TpmPin` existants sont supprimés avant de réenregistrer le protecteur TPM + PIN ;
- le script appelle `GC.Collect()` et `GC.WaitForPendingFinalizers()` à la fin, mais il ne procède pas à un effacement explicite du contenu des champs UI.

Le dépôt ne fournit pas de mécanisme de journalisation centralisée. Les traces visibles sont essentiellement des sorties console et des messages UI.

---

## Limites

> [!WARNING]
> Ce dépôt n’est pas un outil générique de déploiement BitLocker.

Ses principales limites sont les suivantes :

- domaine `ufcvfr.lan` codé en dur ;
- stratégie BitLocker attendue codée en dur ;
- volume cible limité à `C:`;
- compteur de reports limité à 99 ;
- absence de paramètres CLI ;
- absence de mode silencieux ;
- absence de configuration externe ;
- absence de vérification explicite du TPM avant `Enable-BitLocker` ;
- absence de suivi de fin de chiffrement au-delà du lancement correct de `Enable-BitLocker` ;
- absence de journalisation structurée ;
- absence de gestion explicite d’un fichier `PostponeCount.txt` corrompu ;
- fichier `PendingReboot.flag` écrit mais non consommé par le code ;
- dépendance forte à une interface WPF interactive.

En pratique, le script ressemble à un artefact de déploiement interne UFCV plutôt qu’à une bibliothèque réutilisable telle quelle.

---

## Utilisation

Le dépôt ne contient qu’un script d’exécution direct.

### Exemple de lancement

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\BitLocker-Enable-TPM-PIN-Recovery_UFCV.ps1
```

### Contexte d’exécution recommandé par le code

- environnement Windows ;
- session capable d’afficher la fenêtre WPF ;
- contexte très privilégié, idéalement `LocalSystem` selon l’intention du script ;
- machine déjà jointe au domaine UFCV et conforme aux stratégies attendues.

Le script ne propose aucun paramètre. Toutes les valeurs de fonctionnement sont internes au fichier `.ps1`.

---

## Dépannage

### Le script s’arrête immédiatement avec un message d’inéligibilité

Vérifier les valeurs sous `HKLM:\SOFTWARE\Policies\Microsoft\FVE`.

Le script affiche les écarts dans la console avant de quitter. Toute valeur manquante, différente ou de type inattendu bloque l’exécution.

### Le message indique que le poste n’est pas connecté au réseau UFCV

Vérifier :

- l’appartenance au domaine ;
- la connectivité LAN ou VPN ;
- la résolution du domaine `ufcvfr.lan` ;
- la joignabilité d’un contrôleur de domaine.

### BitLocker est déjà en cours ou déjà activé

Le script s’arrête volontairement dans ces cas :

- chiffrement en cours ;
- déchiffrement en cours ;
- volume déjà chiffré et protégé.

Il faut attendre la fin de l’opération existante ou constater l’état du volume avant de relancer.

### Le PIN est refusé

Le PIN doit :

- contenir uniquement des chiffres ;
- avoir entre 6 et 20 caractères ;
- ne pas former une suite strictement croissante ;
- ne pas former une suite strictement décroissante ;
- être identique dans les deux champs.

### Le message `0x80310060` apparaît

La stratégie BitLocker n’autorise pas encore le PIN au moment du lancement.

Le script crée `C:\ProgramData\BitLockerActivation\PendingReboot.flag`, affiche une demande de redémarrage et s’attend à être relancé ensuite.

### La fenêtre WPF ne s’affiche pas ou le parsing XAML échoue

Vérifier :

- l’intégrité du fichier `.ps1` ;
- l’encodage UTF-8 BOM ;
- la disponibilité des assemblies WPF ;
- le fait que l’hôte peut afficher une interface graphique.

Le script mentionne explicitement un possible problème avec `AllowsTransparency` si l’affichage échoue.

### Le compteur de reports semble incohérent

Vérifier `C:\ProgramData\BitLockerActivation\PostponeCount.txt`.

Le fichier doit contenir un entier valide. Si le contenu est corrompu, le cast en `[int]` peut échouer.

### Rien ne se passe après la validation

Vérifier que :

- les deux PIN sont identiques ;
- le bouton `Valider` est bien activé ;
- l’interface est passée en vue de progression ;
- aucun message d’erreur n’est remonté dans le runspace.

---

## Public visé / périmètre

Ce dépôt vise des administrateurs postes et des équipes de déploiement travaillant dans l’environnement UFCV.

Il ne s’agit pas d’un kit générique de déploiement BitLocker. Le code contient des hypothèses organisationnelles explicites et des valeurs codées en dur qui le rendent directement dépendant du contexte UFCV.

---

## Licence / note d’usage interne

Aucune licence open source n’est fournie dans ce dépôt.

Le projet doit être considéré comme un outil interne. Toute réutilisation, modification ou redistribution hors du cadre prévu doit être vérifiée par les responsables concernés.
