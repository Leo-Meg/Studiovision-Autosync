# README-version — différences entre scripts

Ce fichier compare les scripts Python présents dans `src/`.

Convention :

```text
studVMonitor_[version_ou_OS_ou_Box].py
```

---

## Tableau factuel

| Script | Base fonctionnelle | Refresh StudioVision | Logs | Mono-instance | Particularités visibles |
|---|---|---|---|---|---|
| `studVMonitor_V1.py` | Version initiale | `Requery()` récursif sur les sous-formulaires puis formulaire parent + `MoveLast()` | `image_router.log` dans le dossier courant | Non | Sert de base historique. |
| `studVMonitor_V2.py` | Évolution V1 | Recherche `_find_sfdoc()` puis `Requery()` uniquement sur `SFDoc` + `MoveLast()` | `%USERPROFILE%\studiovision\image_router.log` | Non | Évite de modifier le pointeur du formulaire parent. |
| `studVMonitor_V3.py` | V2 | Identique V2 | `%USERPROFILE%\studiovision\image_router.log` | Oui | Ajoute `SingleInstanceGuard` avec mutex Windows global. |
| `studVMonitor_Windows7.py` | V1 adaptée | Identique V1 | `image_router.log` dans le dossier courant | Non | Utilise `typing.Optional` et évite la syntaxe `dict | None`. |
| `studVMonitor_Windows7_V2.py` | V2 adaptée | Identique V2 | `image_router.log` dans le dossier courant | Non | Compatible syntaxe Python ancienne : `Optional`, commentaires `# type:`. |
| `studVMonitor_Box1_V3.py` | Variante BOX 1 | `refresh_ui(expected_patient_code)` sur `SFDoc` | `%USERPROFILE%\studiovision\image_router.log` | Non | Vérifie que le patient affiché est encore le patient attendu avant refresh. |
| `studVMonitor_Box2.py` | Variante BOX 2 historique | Identique V1 | `image_router.log` dans le dossier courant | Non | Ajoute `_try_rmdir()` pour supprimer certains dossiers vides après déplacement. |
| `studVMonitor_Box2_V2.py` | Variante BOX 2 améliorée | Identique V2 | `image_router.log` dans le dossier courant | Non | Combine refresh ciblé `SFDoc` et `_try_rmdir()`. |

---

## Détails par script

### `studVMonitor_V1.py`

Version historique.

Caractéristiques :

- surveille `SOURCE_DIR` avec `watchdog` ;
- utilise une file `queue.Queue` ;
- lit le patient actif via COM Access ;
- résout le dossier patient via `PUBLIC.MDB` ;
- insère le document dans `PUBLIC.MDB` ;
- actualise largement l’interface avec `_requery_form()` ;
- déplace les fichiers non associés vers `ORPHAN_DIR` ;
- écrit `image_router.log` dans le dossier courant.

Limite principale : le refresh est large. Il peut perturber la position du formulaire StudioVision.

---

### `studVMonitor_V2.py`

Évolution de V1.

Différences par rapport à V1 :

- remplace `_requery_form()` et `_goto_last_record()` par `_find_sfdoc()` ;
- actualise seulement le sous-formulaire documentaire `SFDoc` ;
- ne touche pas le recordset du formulaire parent ;
- place les logs dans `%USERPROFILE%\studiovision\image_router.log`.

Usage : version standard si le verrou mono-instance n’est pas nécessaire.

---

### `studVMonitor_V3.py`

Évolution de V2.

Différences par rapport à V2 :

- importe `win32event`, `win32api`, `winerror` ;
- ajoute la classe `SingleInstanceGuard` ;
- crée le mutex `Global\StudiovisionAutosync_ImageRouter_V2` ;
- quitte avec `sys.exit(2)` si une autre instance existe déjà ;
- log attendu en cas de double lancement :

```text
Another instance is already running. Exiting to prevent duplicate processing.
```

Usage : version recommandée pour nouvelle installation.

---

### `studVMonitor_Windows7.py`

Variante de V1 pour environnement Python plus ancien.

Différences par rapport à V1 :

- utilise `from typing import Optional` ;
- remplace les annotations modernes de type `dict | None` par `Optional[...]` ;
- conserve le refresh large de V1 ;
- conserve le log local `image_router.log`.

Usage : poste Windows 7 ou Python ancien.

---

### `studVMonitor_Windows7_V2.py`

Variante de V2 pour environnement Python plus ancien.

Différences par rapport à `studVMonitor_Windows7.py` :

- utilise `_find_sfdoc()` ;
- actualise seulement `SFDoc` ;
- conserve les annotations compatibles Python ancien ;
- réduit fortement l’usage de f-strings ;
- conserve le log local `image_router.log`.

Usage : poste Windows 7 si la variante V2 fonctionne.

---

### `studVMonitor_Box1_V3.py`

Variante dédiée BOX 1.

Différences visibles par rapport à `studVMonitor_V3.py` :

- ne contient pas `SingleInstanceGuard` ;
- ne contient pas de mutex Windows ;
- `refresh_ui()` accepte `expected_patient_code` ;
- vérifie que le patient affiché n’a pas changé avant d’actualiser `SFDoc`;
- sauvegarde le formulaire parent s’il est en état `Dirty` avant `Requery()`;
- retente `Requery()` jusqu’à 3 fois avant fallback `Refresh()`;
- logs dans `%USERPROFILE%\studiovision\image_router.log`.

Usage : installation BOX 1 uniquement.

---

### `studVMonitor_Box2.py`

Variante BOX 2 historique.

Différences visibles par rapport à `studVMonitor_V1.py` :

- contient `_try_rmdir()` ;
- utilise cette logique pour tenter de supprimer certains dossiers vides après déplacement ;
- conserve le refresh large `_requery_form()` ;
- conserve le log local `image_router.log`.

Usage : BOX 2 si le comportement historique doit être conservé.

---

### `studVMonitor_Box2_V2.py`

Variante BOX 2 améliorée.

Différences visibles par rapport à `studVMonitor_Box2.py` :

- remplace le refresh large par `_find_sfdoc()` ;
- actualise seulement `SFDoc` ;
- conserve `_try_rmdir()` ;
- conserve le log local `image_router.log`.

Usage : BOX 2 si aucun besoin n’impose la version historique.

---

## Choix rapide

| Situation | Script |
|---|---|
| Nouvelle installation standard | `studVMonitor_V3.py` |
| Standard sans mono-instance | `studVMonitor_V2.py` |
| Référence historique | `studVMonitor_V1.py` |
| Windows 7 | `studVMonitor_Windows7_V2.py` |
| Windows 7 très contraint | `studVMonitor_Windows7.py` |
| BOX 1 | `studVMonitor_Box1_V3.py` |
| BOX 2 recommandé | `studVMonitor_Box2_V2.py` |
| BOX 2 historique | `studVMonitor_Box2.py` |

---

## Règle de maintenance

Pour tout nouveau script :

1. garder le préfixe `studVMonitor_` ;
2. indiquer version, OS ou BOX dans le nom ;
3. documenter les différences concrètes ici ;
4. générer un `.exe` avec le même nom.

Exemple :

```text
src/studVMonitor_Box3_V1.py
dist/studVMonitor_Box3_V1.exe
```
