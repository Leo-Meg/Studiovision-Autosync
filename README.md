# Studiovision Autosync


Routeur automatique d’images vers les dossiers patients StudioVision.

Le programme surveille un dossier d’entrée. Quand une image arrive, il récupère le patient actif dans StudioVision / Access, déplace le fichier dans le dossier patient, puis ajoute l’entrée documentaire dans `PUBLIC.MDB`.

---

## Documentation complémentaire

| Fichier | Rôle |
|---|---|
| `README-version.md` | Compare les scripts Python et indique les différences concrètes entre versions. |
| `README-rendu-final.md` | Décrit l’organisation attendue des dossiers et fichiers sur le poste installé. |
| `README-lecture-log.md` | Explique comment lire et interpréter le fichier `image_router.log`. |
| `shell/Shell_command.ipynb` | Regroupe les commandes utiles pour compiler, installer, vérifier et désinstaller. |

---

## Fonctionnement

1. L’appareil dépose une image dans `SOURCE_DIR`.
2. Le programme détecte le fichier.
3. Il attend que l’écriture du fichier soit terminée.
4. Il lit le patient actif dans StudioVision via COM Access.
5. Il retrouve le dossier patient via `PUBLIC.MDB`.
6. Il déplace l’image dans le bon dossier.
7. Il insère une ligne documentaire dans `PUBLIC.MDB`.
8. Il actualise l’interface StudioVision.
9. Si aucun patient n’est trouvé, le fichier part dans `ORPHAN_DIR`.

---

## Structure du projet

```text
Studiovision-Autosync/
├── README.md
├── README-version.md
├── README-rendu-final.md
├── README-lecture-log.md
├── requirements.txt
├── src/
│   ├── studVMonitor_V1.py
│   ├── studVMonitor_V2.py
│   ├── studVMonitor_V3.py
│   ├── studVMonitor_Windows7.py
│   ├── studVMonitor_Windows7_V2.py
│   ├── studVMonitor_Box1_V3.py
│   ├── studVMonitor_Box2.py
│   └── studVMonitor_Box2_V2.py
└── Shell_command.ipynb
```

---

## Script recommandé

Pour une nouvelle installation standard :

```text
src/studVMonitor_V3.py
```

Pourquoi :

- reprend l’actualisation ciblée de `SFDoc` ;
- écrit les logs dans `%USERPROFILE%\studiovision\image_router.log` ;
- empêche le lancement simultané de deux instances.

Pour les différences entre scripts, lire :

```text
README-version.md
```

---

## Configuration à modifier

Dans le script choisi, adapter les constantes suivantes :

| Constante | Contenu attendu |
|---|---|
| `SOURCE_DIR` | Dossier surveillé. |
| `ORPHAN_DIR` | Dossier des fichiers non associés. |
| `DEST_PHOTOS` | Racine des dossiers photos patients. |
| `PUBLIC_MDB` | Base Access utilisée pour lire et insérer les documents. |
| `DOCUM_MDB` | Base conservée pour compatibilité selon installation. |

Exemple :

```python
SOURCE_DIR  = Path(r"C:\Chemin\Entree")
ORPHAN_DIR  = Path(r"C:\Chemin\Orphelins")
DEST_PHOTOS = Path(r"C:\Chemin\Photos")
PUBLIC_MDB  = Path(r"C:\Chemin\PUBLIC.MDB")
DOCUM_MDB   = Path(r"C:\Chemin\DOCUM.MDB")
```

---

## Extensions prises en charge

```text
.jpg, .jpeg, .jfif, .png, .bmp, .tif, .tiff, .dcm
```

Descriptions insérées dans la base :

| Extension | Description |
|---|---|
| `.tif`, `.tiff` | `OCT` |
| `.dcm` | `DICOM` |
| autres | `Image` |

---

## Prérequis

- Windows.
- StudioVision installé.
- Microsoft Access ou Runtime Access.
- Pilote ODBC Microsoft Access.
- Python adapté au script choisi.
- Droits d’écriture sur les dossiers source, destination, orphelins et logs.

Installation des dépendances :

```bash
pip install -r requirements.txt
```

Dépendances principales :

| Module | Rôle |
|---|---|
| `watchdog` | Surveillance du dossier source. |
| `pyodbc` | Accès aux bases `.MDB`. |
| `pywin32` | COM Access et fonctions Windows. |
| `pyinstaller` | Création de l’exécutable. |

---

## Lancer en Python

```bash
python src/studVMonitor_V3.py
```

Remplacer `studVMonitor_V3.py` par le script choisi.

---

## Créer l’exécutable

```bash
pyinstaller --onefile --noconsole --name studVMonitor_V3 src/studVMonitor_V3.py
```

Résultat :

```text
dist/studVMonitor_V3.exe
```

Le notebook `shell/Shell_command.ipynb` contient les commandes utiles pour :

- installer les dépendances ;
- compiler ;
- créer le raccourci de démarrage automatique ;
- vérifier le démarrage ;
- désinstaller le raccourci ;
- nettoyer `build/` et `dist/`.

---

## Démarrage automatique

Créer un raccourci `.lnk` dans :

```text
%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\
```

Le raccourci doit pointer vers :

```text
C:\Studiovision-Autosync\dist\studVMonitor_V3.exe
```

Le dossier de travail recommandé est :

```text
C:\Studiovision-Autosync\dist
```

Pour l’organisation finale des dossiers sur le poste installé, lire :

```text
README-rendu-final.md
```

---

## Logs

Emplacements possibles selon script :

```text
%USERPROFILE%\studiovision\image_router.log
image_router.log
C:\Studiovision-Autosync\logs\image_router.log
```

Pour savoir quoi vérifier dans `image_router.log`, lire :

```text
README-lecture-log.md
```

---

## Fichiers orphelins

Un fichier va dans `ORPHAN_DIR` si :

- aucun patient n’est ouvert ;
- le patient est ouvert trop tard ;
- le dossier patient est introuvable ;
- `PUBLIC.MDB` est inaccessible ;
- le fichier reste verrouillé trop longtemps.

Ces fichiers doivent être traités manuellement.

---

## Contrôle avant production

Avant installation définitive :

1. tester avec un dossier source temporaire ;
2. tester avec un patient de démonstration ;
3. vérifier le dossier destination ;
4. vérifier l’insertion dans StudioVision ;
5. vérifier les logs ;
6. vérifier le dossier orphelin ;
7. vérifier qu’une seule instance tourne.
