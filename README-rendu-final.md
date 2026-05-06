# README-rendu-final — installation sur un poste

Ce fichier décrit les dossiers et fichiers à placer sur l’ordinateur qui exécutera `studVMonitor`.

---

## Emplacement recommandé

Installer dans un dossier stable :

```text
C:\Studiovision-Autosync\
```

Éviter :

```text
Downloads
Desktop\Nouveau dossier
dossier temporaire
```

Le raccourci de démarrage automatique dépend du chemin de l’exécutable. Si le dossier est déplacé, le démarrage automatique casse.

---

## Arborescence recommandée

```text
C:\Studiovision-Autosync\
├── dist\
│   └── studVMonitor_<Version>.exe
├── logs\
│   └── image_router.log
├── config\
│   └── configuration_locale.txt
├── docs\
│   ├── README.md
│   ├── README-version.md
│   ├── README-rendu-final.md
│   └── Shell_command.ipynb
└── install\
    ├── install_startup.ps1
    └── uninstall_startup.ps1
```

---

## Minimum nécessaire en production

```text
C:\Studiovision-Autosync\
├── dist\
│   └── studVMonitor_V3.exe
└── logs\
    └── image_router.log
```

Et le raccourci :

```text
%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\studVMonitor_V3.lnk
```

---

## Dossier `dist\`

Contient l’exécutable final.

Exemple :

```text
C:\Studiovision-Autosync\dist\studVMonitor_V3.exe
```

Ce fichier vient de PyInstaller :

```bash
pyinstaller --onefile --noconsole --name studVMonitor_V3 src/studVMonitor_V3.py
```

Résultat :

```text
dist\studVMonitor_V3.exe
```

En mode `--onefile`, c’est normalement le seul fichier à lancer.

---

## Dossier `logs\`

Emplacement recommandé :

```text
C:\Studiovision-Autosync\logs\image_router.log
```

Ce log doit permettre de vérifier :

- le démarrage du programme ;
- le dossier surveillé ;
- les fichiers détectés ;
- le patient actif ;
- le dossier patient trouvé ;
- les fichiers déplacés ;
- les insertions dans `PUBLIC.MDB` ;
- les fichiers envoyés en orphelin ;
- les erreurs Access / ODBC ;
- les erreurs de fichier verrouillé ;
- les doubles lancements avec la V3.

Certaines versions écrivent ailleurs :

| Scripts | Emplacement de log |
|---|---|
| `studVMonitor_V2.py`, `studVMonitor_V3.py`, `studVMonitor_Box1_V3.py` | `%USERPROFILE%\studiovision\image_router.log` |
| `studVMonitor_V1.py`, `studVMonitor_Windows7.py`, `studVMonitor_Windows7_V2.py`, `studVMonitor_Box2.py`, `studVMonitor_Box2_V2.py` | `image_router.log` dans le dossier courant |

Pour une installation propre, modifier le script pour écrire explicitement dans :

```text
C:\Studiovision-Autosync\logs\image_router.log
```

---

## Dossier `config\`

Contient les informations locales du poste.

Exemple :

```text
C:\Studiovision-Autosync\config\configuration_locale.txt
```

Contenu conseillé :

```text
SOURCE_DIR=...
ORPHAN_DIR=...
DEST_PHOTOS=...
PUBLIC_MDB=...
DOCUM_MDB=...
LOG_FILE=C:\Studiovision-Autosync\logs\image_router.log
```

Même si les chemins sont codés dans le script Python, ce fichier documente l’installation réelle.

---

## Dossier `docs\`

Contient la documentation utile sur le poste :

```text
README.md
README-version.md
README-rendu-final.md
Shell_command.ipynb
```

---

## Dossier `install\`

Contient les scripts d’installation et de suppression du démarrage automatique.

Exemple :

```text
install_startup.ps1
uninstall_startup.ps1
```

---

## Raccourci Startup

Chemin du dossier Startup :

```text
%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\
```

Le raccourci doit pointer vers :

```text
C:\Studiovision-Autosync\dist\studVMonitor_V3.exe
```

Le dossier de travail doit être :

```text
C:\Studiovision-Autosync\dist
```

Commande PowerShell type :

```powershell
$exe = "C:\Studiovision-Autosync\dist\studVMonitor_V3.exe"
$startup = [System.Environment]::GetFolderPath("Startup")

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut("$startup\studVMonitor_V3.lnk")
$shortcut.TargetPath = $exe
$shortcut.WorkingDirectory = "C:\Studiovision-Autosync\dist"
$shortcut.Save()
```

---

## Dossier `build\`

`build\` est créé par PyInstaller pendant la compilation.

Exemple :

```text
build\
└── studVMonitor_V3\
    ├── Analysis-00.toc
    ├── PYZ-00.pyz
    ├── PYZ-00.toc
    ├── PKG-00.toc
    ├── EXE-00.toc
    ├── warn-studVMonitor_V3.txt
    └── xref-studVMonitor_V3.html
```

Rôle :

| Fichier | Rôle |
|---|---|
| `Analysis-00.toc` | Modules et dépendances détectés. |
| `PYZ-00.pyz` | Archive Python intermédiaire. |
| `PKG-00.toc` | Informations de packaging. |
| `EXE-00.toc` | Informations de génération de l’exécutable. |
| `warn-*.txt` | Avertissements PyInstaller. |
| `xref-*.html` | Rapport des dépendances. |

À retenir :

- `build\` sert à compiler ;
- `build\` ne sert pas à exécuter ;
- `build\` n’est pas à copier en production ;
- `build\` peut être supprimé après validation de l’exécutable ;
- le fichier utile en production est dans `dist\`.

---

## Fichier `.spec`

PyInstaller peut créer :

```text
studVMonitor_V3.spec
```

Ce fichier sert à reproduire ou personnaliser la compilation.

À conserver côté développement si nécessaire.

Pas obligatoire sur le poste final.

---

## Contrôle après installation

Vérifier :

```text
C:\Studiovision-Autosync\dist\studVMonitor_V3.exe
C:\Studiovision-Autosync\logs\
%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\studVMonitor_V3.lnk
```

Puis contrôler :

1. le raccourci pointe vers le bon `.exe` ;
2. le dossier de travail est `dist\` ;
3. le log est créé ;
4. StudioVision reçoit bien le document ;
5. les fichiers orphelins arrivent dans `ORPHAN_DIR` ;
6. une seule instance tourne avec la V3.
