# Lot de modifications Studiovision Autosync

Ce dossier contient uniquement les fichiers ajoutés ou modifiés, ainsi qu’un manifeste des fichiers à supprimer.

## Fichiers modifiés

- `README.md` : réécriture en français et mise à jour du nommage normalisé.

## Fichiers ajoutés

- `src/studVMonitor_V1.py`
- `src/studVMonitor_V2.py`
- `src/studVMonitor_V3.py`
- `src/studVMonitor_Windows7.py`
- `src/studVMonitor_Windows7_V2.py`
- `src/studVMonitor_Box1_V3.py`
- `src/studVMonitor_Box2.py`
- `src/studVMonitor_Box2_V2.py`
- `shell/Shell_command.ipynb`

## Fichiers à supprimer du dépôt

Voir `DELETED_FILES.txt`.

## Renommage appliqué

| Ancien fichier | Nouveau fichier |
|---|---|
| `src/studiovision_monitor.py` | `src/studVMonitor_V1.py` |
| `src/studiovision_monitorV2.py` | `src/studVMonitor_V2.py` |
| `src/windows7.py` | `src/studVMonitor_Windows7.py` |
| `src/windows7V2.py` | `src/studVMonitor_Windows7_V2.py` |
| `src/box1V3.py` | `src/studVMonitor_Box1_V3.py` |
| `src/box2.py` | `src/studVMonitor_Box2.py` |
| `src/box2V2.py` | `src/studVMonitor_Box2_V2.py` |
| `shell/Shell commands` | `shell/Shell_command.ipynb` |

## Ajout fonctionnel

`src/studVMonitor_V3.py` ajoute un verrou mono-instance Windows via mutex global. La deuxième instance quitte avec le code `2` et écrit le message d’erreur demandé dans les logs.
