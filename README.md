# Studiovision Autosync

**Studiovision Autosync** est un outil Windows de routage automatique d’images vers les dossiers patients StudioVision.

Le programme surveille un dossier d’entrée alimenté par un appareil d’imagerie. Lorsqu’une image arrive, il identifie le patient actuellement ouvert dans StudioVision, déplace le fichier vers le bon dossier patient, puis ajoute l’entrée correspondante dans la base Access afin que le document apparaisse dans l’interface StudioVision.

Ce dépôt contient plusieurs scripts Python correspondant à des environnements ou variantes de déploiement différents. Pour choisir le bon script, consulter aussi [`README-version.md`](README-version.md).

---

## Objectif du projet

Le projet vise à automatiser une tâche normalement manuelle : récupérer les images produites par un appareil externe et les rattacher au bon patient dans StudioVision.

Sans automatisation, l’utilisateur doit généralement :

1. récupérer l’image dans le dossier de sortie de l’appareil ;
2. retrouver le patient dans StudioVision ;
3. déplacer le fichier au bon emplacement ;
4. créer ou actualiser l’entrée documentaire dans StudioVision.

Studiovision Autosync automatise cette chaîne dès lors que le patient est ouvert dans StudioVision au moment du traitement.

---

## Fonctionnement général

Le fonctionnement est le même pour toutes les versions principales du programme.

1. **Surveillance du dossier source**  
   Le programme observe un dossier dans lequel l’appareil dépose les images.

2. **Détection d’un nouveau fichier**  
   Lorsqu’un fichier compatible est créé, il est ajouté à une file de traitement.

3. **Attente de fin d’écriture**  
   Le programme vérifie que le fichier n’est plus verrouillé par l’appareil avant de le déplacer.

4. **Lecture du patient actif**  
   Le programme interroge Microsoft Access / StudioVision via COM afin de récupérer les informations du patient actuellement ouvert.

5. **Résolution du dossier patient**  
   Le dossier cible est déterminé à partir des informations présentes dans `PUBLIC.MDB`.

6. **Déplacement de l’image**  
   L’image est déplacée dans le dossier patient. En cas de conflit de nom, un suffixe temporel peut être ajouté.

7. **Insertion en base Access**  
   Une ligne est ajoutée dans la table documentaire de `PUBLIC.MDB` pour référencer le nouveau fichier.

8. **Actualisation de l’interface StudioVision**  
   Selon la version du script, le formulaire ou le sous-formulaire documentaire est actualisé pour afficher le nouveau document.

9. **Gestion des fichiers non associés**  
   Si aucun patient valide n’est trouvé dans le délai prévu, le fichier est déplacé dans un dossier d’orphelins pour traitement manuel.

---

## Structure du dépôt

```text
Studiovision-Autosync/
├── README.md
├── README-version.md
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
└── shell/
    └── Shell_command.ipynb
```

---

## Convention de nommage des scripts

Tous les scripts Python destinés à générer un exécutable suivent la convention :

```text
studVMonitor_[version_ou_OS_ou_BOX].py
```

Exemples :

```text
studVMonitor_V3.py
studVMonitor_Windows7_V2.py
studVMonitor_Box1_V3.py
```

Cette convention permet d’identifier rapidement :

- la version fonctionnelle ;
- la compatibilité système ;
- la variante dédiée à un poste ou une BOX ;
- le nom attendu de l’exécutable généré.

---

## Choix recommandé

Pour une nouvelle installation standard, utiliser en priorité :

```text
src/studVMonitor_V3.py
```

Cette version reprend les améliorations de la V2 et ajoute un verrou mono-instance empêchant le lancement simultané de plusieurs processus.

Utiliser une variante spécifique seulement si l’environnement le justifie :

- poste Windows 7 ;
- BOX ou appareil avec une organisation de fichiers particulière ;
- besoin de conserver un comportement historique.

Le détail des différences est documenté dans [`README-version.md`](README-version.md).

---

## Configuration à adapter avant exécution

Chaque script contient une section de configuration en haut de fichier. Les chemins doivent être adaptés au poste cible avant lancement ou compilation.

| Variable | Rôle |
|---|---|
| `SOURCE_DIR` | Dossier surveillé, alimenté par l’appareil d’imagerie. |
| `ORPHAN_DIR` | Dossier où placer les fichiers non associés automatiquement à un patient. |
| `DEST_PHOTOS` | Racine des dossiers patients StudioVision. |
| `PUBLIC_MDB` | Base Access principale utilisée pour retrouver et enregistrer les documents. |
| `DOCUM_MDB` | Base Access conservée dans la configuration pour compatibilité avec certains environnements. |

Paramètres de traitement importants :

| Constante | Rôle |
|---|---|
| `WATCHED_EXTENSIONS` | Extensions de fichiers acceptées. |
| `FILE_LOCK_RETRY_DELAY` | Délai entre deux essais lorsqu’un fichier est encore verrouillé. |
| `FILE_LOCK_MAX_ATTEMPTS` | Nombre maximal d’essais avant abandon du fichier. |
| `PATIENT_POLL_INTERVAL` | Fréquence de recherche du patient actif. |
| `PATIENT_WAIT_TIMEOUT` | Délai maximal d’attente avant classement en orphelin. |
| `SFDOC_SUBFORM_NAME` | Nom du sous-formulaire documentaire StudioVision à actualiser. |

---

## Extensions prises en charge

Les extensions surveillées par défaut sont :

```text
.jpg, .jpeg, .jfif, .png, .bmp, .tif, .tiff, .dcm
```

La description insérée en base dépend de l’extension :

| Extension | Description insérée |
|---|---|
| `.tif`, `.tiff` | `OCT` |
| `.dcm` | `DICOM` |
| autres extensions surveillées | `Image` |

---

## Prérequis

Environnement attendu :

- Windows ;
- StudioVision installé et fonctionnel ;
- Microsoft Access ou runtime Access disponible ;
- pilote ODBC Microsoft Access installé ;
- Python compatible avec la version du script choisi ;
- droits suffisants sur les dossiers source, destination et orphelins.

Installer les dépendances Python :

```bash
pip install -r requirements.txt
```

Dépendances principales :

| Dépendance | Utilisation |
|---|---|
| `watchdog` | Surveillance du dossier source. |
| `pyodbc` | Connexion aux bases Access. |
| `pywin32` | Interaction COM avec Access / StudioVision et fonctionnalités Windows. |
| `pyinstaller` | Génération des exécutables Windows. |

---

## Lancement en mode Python

Depuis la racine du projet :

```bash
python src/studVMonitor_V3.py
```

Remplacer le nom du script par la version adaptée au poste cible.

---

## Génération d’un exécutable

Les exécutables sont générés avec PyInstaller.

Exemple générique :

```bash
pyinstaller --onefile --noconsole --name studVMonitor_V3 src/studVMonitor_V3.py
```

Résultat attendu :

```text
dist/studVMonitor_V3.exe
```

Le notebook [`shell/Shell_command.ipynb`](shell/Shell_command.ipynb) fournit une procédure générique pour :

- installer les dépendances ;
- générer un exécutable ;
- créer un raccourci de démarrage automatique ;
- vérifier le démarrage automatique ;
- arrêter proprement un processus ;
- supprimer un raccourci ou une ancienne tâche planifiée ;
- nettoyer les dossiers de build.

---

## Démarrage automatique Windows

Le déploiement courant consiste à placer un raccourci de l’exécutable dans le dossier de démarrage Windows de l’utilisateur.

Principe :

1. générer l’exécutable avec PyInstaller ;
2. créer un raccourci `.lnk` vers l’exécutable ;
3. placer ce raccourci dans le dossier Startup de Windows ;
4. redémarrer la session ou le poste ;
5. vérifier qu’une seule instance du programme est active.

La version V3 est recommandée pour ce mode de déploiement, car elle empêche les doubles lancements.

---

## Logs et diagnostic

Les scripts écrivent des logs avec :

- la date et l’heure ;
- le niveau de log ;
- le nom du thread ;
- le message de traitement ou d’erreur.

Selon la version, le fichier de log peut être écrit :

```text
%USERPROFILE%\studiovision\image_router.log
```

ou dans le répertoire courant sous le nom :

```text
image_router.log
```

À vérifier en priorité en cas de problème :

- le dossier source existe-t-il ?
- le fichier est-il encore verrouillé par l’appareil ?
- StudioVision est-il ouvert sur le bon patient ?
- `PUBLIC.MDB` est-il accessible ?
- le dossier patient existe-t-il sur disque ?
- l’image est-elle partie dans le dossier d’orphelins ?
- plusieurs instances du programme sont-elles lancées ?

---

## Fichiers orphelins

Un fichier est déplacé dans `ORPHAN_DIR` lorsqu’il ne peut pas être associé automatiquement.

Causes fréquentes :

- aucun patient ouvert dans StudioVision ;
- patient ouvert trop tardivement ;
- délai d’attente dépassé ;
- dossier patient introuvable ;
- erreur d’accès à la base Access ;
- fichier incomplet ou verrouillé trop longtemps.

Les fichiers orphelins doivent être contrôlés puis réintégrés manuellement si nécessaire.

---

## Sécurité d’exploitation

Avant toute mise en production :

1. tester sur un dossier source temporaire ;
2. vérifier le dossier cible réellement utilisé ;
3. tester avec un patient de démonstration ;
4. contrôler les logs ;
5. vérifier l’apparition du document dans StudioVision ;
6. vérifier le comportement en cas de patient absent ;
7. vérifier qu’une seule instance du programme tourne.

Pour une installation stable, privilégier `studVMonitor_V3.py` ou une variante dédiée reprenant le verrou mono-instance.
