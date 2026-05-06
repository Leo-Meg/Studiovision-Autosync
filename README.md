# Studiovision Autosync

Studiovision Autosync est un routeur d’images pour StudioVision. Il surveille un dossier source alimenté par un appareil d’imagerie, identifie le patient actuellement ouvert dans StudioVision, déplace l’image vers le bon dossier patient, puis ajoute l’entrée correspondante dans la base Access afin que l’image apparaisse dans le dossier du patient.

Le projet est destiné à un environnement Windows avec StudioVision, Microsoft Access/ODBC et des postes d’acquisition d’images.

---

## Principe général

1. Le programme surveille un dossier source avec `watchdog`.
2. Lorsqu’une image arrive, elle est placée dans une file de traitement.
3. Le programme attend que l’écriture du fichier soit terminée.
4. Le patient actif est récupéré dans StudioVision via COM/Access.
5. Le dossier patient est résolu depuis `PUBLIC.MDB`.
6. L’image est déplacée dans le dossier patient.
7. Une ligne est ajoutée dans `PUBLIC.MDB` pour référencer l’image.
8. L’interface StudioVision est actualisée pour rendre l’image visible.
9. Si aucun patient valide n’est trouvé dans le délai prévu, l’image est déplacée dans le dossier des orphelins.

---

## Nommage normalisé des scripts

Tous les scripts destinés à produire un exécutable commencent maintenant par `studVMonitor_`.

| Script | Rôle |
|---|---|
| `src/studVMonitor_V1.py` | Version de base du routeur d’images. |
| `src/studVMonitor_V2.py` | Version améliorée avec actualisation groupée de l’interface et requery ciblé du sous-formulaire `SFDoc`. |
| `src/studVMonitor_V3.py` | Version V2 avec verrou mono-instance Windows empêchant plusieurs processus simultanés. |
| `src/studVMonitor_Windows7.py` | Variante compatible Python 3.9 / Windows 7 de la version de base. |
| `src/studVMonitor_Windows7_V2.py` | Variante Windows 7 avec les améliorations de la V2. |
| `src/studVMonitor_Box1_V3.py` | Variante dédiée BOX 1. |
| `src/studVMonitor_Box2.py` | Variante dédiée BOX 2 avec prise en charge d’une structure de fichiers issue d’un appareil Nidek. |
| `src/studVMonitor_Box2_V2.py` | Variante BOX 2 avec les améliorations de la V2. |

Les anciens noms de fichiers ont été retirés afin d’éviter les ambiguïtés lors de la génération des exécutables.

---

## Version V3 : verrou mono-instance

`studVMonitor_V3.py` ajoute un verrou système Windows basé sur un mutex global.

Objectif : empêcher définitivement la multiplication des processus si l’utilisateur lance deux fois le même exécutable, ou si un raccourci de démarrage automatique est déclenché alors qu’une instance est déjà active.

Comportement attendu :

- la première instance continue normalement ;
- toute instance supplémentaire quitte immédiatement ;
- une ligne d’erreur est écrite dans les logs :

```text
ERROR [MainThread] Another instance is already running. Exiting to prevent duplicate processing.
```

Cette protection réduit le risque de double traitement, de conflits sur les fichiers surveillés et d’insertions multiples dans la base Access.

---

## Configuration

Avant de lancer ou compiler un script, renseigner les chemins en haut du fichier choisi :

| Variable | Description |
|---|---|
| `SOURCE_DIR` | Dossier surveillé dans lequel l’appareil dépose les images. |
| `ORPHAN_DIR` | Dossier de sortie pour les images non associées à un patient. |
| `DEST_PHOTOS` | Racine des dossiers patients StudioVision. |
| `PUBLIC_MDB` | Chemin vers la base Access principale `PUBLIC.MDB`. |
| `DOCUM_MDB` | Chemin vers `DOCUM.MDB`, conservé pour compatibilité. |

Autres paramètres importants :

| Constante | Valeur par défaut | Rôle |
|---|---:|---|
| `FILE_LOCK_RETRY_DELAY` | `3` secondes | Pause entre deux tentatives lorsqu’un fichier est encore verrouillé. |
| `FILE_LOCK_MAX_ATTEMPTS` | `15` | Nombre maximal de tentatives avant abandon. |
| `PATIENT_POLL_INTERVAL` | `3` secondes | Fréquence de recherche du patient actif. |
| `PATIENT_WAIT_TIMEOUT` | `900` secondes | Délai maximal avant déplacement vers le dossier des orphelins. |
| `SFDOC_SUBFORM_NAME` | `"SFDoc"` | Nom du sous-formulaire StudioVision à actualiser. |

---

## Extensions surveillées

Extensions prises en charge par défaut :

```text
.jpg, .jpeg, .jfif, .png, .bmp, .tif, .tiff, .dcm
```

Correspondance utilisée pour la description enregistrée en base :

| Extension | Description |
|---|---|
| `.tif`, `.tiff` | `OCT` |
| `.dcm` | `DICOM` |
| autres extensions | `Image` |

---

## Prérequis

- Windows.
- Python adapté à la version du script utilisé.
- Pilote ODBC Microsoft Access installé.
- StudioVision et ses bases Access accessibles depuis le poste.
- Dépendances Python du fichier `requirements.txt`.

Installation des dépendances :

```bash
pip install -r requirements.txt
```

Dépendances principales :

| Package | Utilisation |
|---|---|
| `watchdog` | Surveillance du système de fichiers. |
| `pyodbc` | Connexion aux bases Access. |
| `pywin32` | Automatisation COM Windows et verrou mono-instance. |

---

## Lancement en Python

Exemple générique :

```bash
python src/studVMonitor_V3.py
```

Choisir le script correspondant au poste, à l’OS ou à la BOX concernée.

---

## Génération d’un exécutable

Les exécutables sont générés avec PyInstaller.

Exemple générique :

```bash
pyinstaller --onefile --noconsole --name studVMonitor_V3 src/studVMonitor_V3.py
```

Le nom passé à `--name` doit rester cohérent avec le script source afin de faciliter le diagnostic et la maintenance.

---

## Notebook de commandes shell

Le fichier texte historique `shell/Shell commands` a été remplacé par :

```text
shell/Shell_command.ipynb
```

Ce notebook explique les commandes utiles pour :

- générer un exécutable ;
- créer un raccourci de démarrage automatique Windows ;
- vérifier le dossier de démarrage ;
- arrêter un processus ;
- supprimer un raccourci de démarrage ;
- nettoyer un dossier de build.

Les exemples sont volontairement génériques afin de pouvoir être réutilisés proprement sur n’importe quel poste.

---

## Logs

Les logs sont écrits dans la console et/ou dans un fichier selon la variante utilisée.

Les variantes récentes utilisent de préférence un dossier utilisateur dédié :

```text
%USERPROFILE%\studiovision\image_router.log
```

Sur certaines variantes historiques, le fichier `image_router.log` peut être créé dans le répertoire courant.

---

## Fichiers orphelins

Une image est déplacée vers `ORPHAN_DIR` lorsqu’elle ne peut pas être associée automatiquement à un patient.

Causes fréquentes :

- aucun patient ouvert dans StudioVision ;
- délai d’attente dépassé ;
- dossier patient introuvable dans `PUBLIC.MDB` ;
- erreur d’accès au fichier ou à la base.

Ces fichiers doivent être contrôlés manuellement.

---

## Notes techniques

- `pythoncom.CoInitialize()` et `pythoncom.CoUninitialize()` sont utilisés dans le thread de travail, car les objets COM ne doivent pas être partagés directement entre threads.
- Les variantes V2 réduisent les blocages StudioVision lors de l’arrivée de plusieurs images en rafale.
- `studVMonitor_V3.py` doit être privilégié lorsqu’un poste risque de lancer plusieurs instances par erreur.
- Les variantes Windows 7 évitent certaines annotations Python modernes afin de rester compatibles avec Python 3.9.
