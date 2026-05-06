# README-version — différences entre les scripts Python

Ce fichier explique le rôle de chaque script Python présent dans `src/`.

La convention de nommage est :

```text
studVMonitor_[version_ou_OS_ou_BOX].py
```

Elle permet de distinguer les versions fonctionnelles, les variantes de compatibilité système et les variantes propres à un poste ou une BOX.

---

## Vue d’ensemble

| Script | Usage recommandé | Particularité principale |
|---|---|---|
| `studVMonitor_V1.py` | Version historique de référence | Routage de base, actualisation large de l’interface. |
| `studVMonitor_V2.py` | Version standard améliorée | Actualisation plus ciblée du sous-formulaire documentaire `SFDoc`. |
| `studVMonitor_V3.py` | Version recommandée pour nouvelle installation | V2 + verrou mono-instance Windows. |
| `studVMonitor_Windows7.py` | Poste Windows 7 historique | Variante V1 adaptée à un environnement Python plus ancien. |
| `studVMonitor_Windows7_V2.py` | Poste Windows 7 avec amélioration V2 | Variante Windows 7 avec actualisation ciblée de `SFDoc`. |
| `studVMonitor_Box1_V3.py` | Variante dédiée BOX 1 | Variante spécialisée avec logique proche V3 / rafales / refresh sécurisé. |
| `studVMonitor_Box2.py` | Variante dédiée BOX 2 historique | Variante BOX 2 basée sur le comportement initial. |
| `studVMonitor_Box2_V2.py` | Variante dédiée BOX 2 améliorée | BOX 2 avec logique V2 d’actualisation ciblée. |

---

## `studVMonitor_V1.py`

Version de base du routeur d’images.

Elle assure :

- la surveillance du dossier source ;
- la mise en file des images détectées ;
- l’attente de fin d’écriture du fichier ;
- la lecture du patient actif dans StudioVision ;
- la résolution du dossier patient via `PUBLIC.MDB` ;
- le déplacement de l’image ;
- l’insertion documentaire dans `PUBLIC.MDB` ;
- l’actualisation de l’interface StudioVision.

À utiliser principalement comme version historique ou comme base de comparaison.

---

## `studVMonitor_V2.py`

Version améliorée de la V1.

Différences principales :

- meilleure gestion de l’actualisation de l’interface ;
- recherche récursive du sous-formulaire documentaire `SFDoc` ;
- actualisation ciblée du sous-formulaire au lieu de manipuler largement le formulaire parent ;
- réduction du risque de retour intempestif au premier enregistrement ;
- meilleure gestion des rafales d’images avec actualisation différée après stabilisation de la file.

Cette version est préférable à la V1 lorsque plusieurs images peuvent arriver rapidement ou lorsque l’interface StudioVision est sensible aux `Requery()` trop larges.

---

## `studVMonitor_V3.py`

Version recommandée pour les nouvelles installations standard.

Elle reprend les améliorations de la V2 et ajoute un verrou mono-instance Windows.

Objectif : empêcher deux processus identiques de fonctionner simultanément sur le même poste.

Comportement attendu :

- la première instance démarre normalement ;
- toute seconde instance détecte qu’une instance existe déjà ;
- la seconde instance écrit une erreur dans les logs ;
- la seconde instance quitte immédiatement.

Message typique :

```text
Another instance is already running. Exiting to prevent duplicate processing.
```

Cette version est particulièrement adaptée au démarrage automatique Windows, car elle limite le risque de double traitement, de double insertion ou de conflit sur les fichiers.

---

## `studVMonitor_Windows7.py`

Variante destinée aux environnements Windows 7 ou aux installations Python plus anciennes.

Différences principales :

- style de typage compatible avec les anciennes versions de Python ;
- syntaxe plus prudente pour éviter certaines incompatibilités ;
- comportement fonctionnel proche de la V1.

À utiliser uniquement si le poste cible impose Windows 7 ou une contrainte forte de compatibilité.

---

## `studVMonitor_Windows7_V2.py`

Variante Windows 7 intégrant les améliorations fonctionnelles de la V2.

Différences principales par rapport à `studVMonitor_Windows7.py` :

- actualisation ciblée du sous-formulaire `SFDoc` ;
- meilleure stabilité de l’interface StudioVision ;
- logique de traitement plus proche de `studVMonitor_V2.py` ;
- syntaxe conservant la compatibilité Windows 7 / Python ancien.

À privilégier sur Windows 7 lorsque l’environnement supporte cette variante.

---

## `studVMonitor_Box1_V3.py`

Variante dédiée à la BOX 1.

Elle reprend une logique proche des versions récentes :

- logs placés dans un dossier utilisateur dédié ;
- actualisation ciblée de `SFDoc` ;
- traitement plus robuste des rafales ;
- vérification du patient attendu avant certaines opérations d’interface.

À utiliser uniquement pour le poste ou l’appareil correspondant à la BOX 1.

---

## `studVMonitor_Box2.py`

Variante dédiée à la BOX 2 dans sa forme historique.

Elle reprend le fonctionnement général du routeur, avec des adaptations liées à l’environnement BOX 2.

À conserver si le poste BOX 2 dépend encore de ce comportement historique.

---

## `studVMonitor_Box2_V2.py`

Variante améliorée de `studVMonitor_Box2.py`.

Différences principales :

- actualisation ciblée du sous-formulaire `SFDoc` ;
- comportement plus proche de la V2 ;
- meilleure robustesse lors de l’arrivée de plusieurs images ;
- meilleure séparation entre logique de traitement et logique d’interface.

À privilégier pour BOX 2 si la version historique n’est pas explicitement requise.

---

## Recommandations de choix

| Situation | Script conseillé |
|---|---|
| Nouvelle installation standard | `studVMonitor_V3.py` |
| Installation standard sans verrou mono-instance | `studVMonitor_V2.py` |
| Audit ou comparaison avec le comportement initial | `studVMonitor_V1.py` |
| Poste Windows 7 | `studVMonitor_Windows7_V2.py` |
| Poste Windows 7 très contraint | `studVMonitor_Windows7.py` |
| BOX 1 | `studVMonitor_Box1_V3.py` |
| BOX 2 | `studVMonitor_Box2_V2.py` |
| BOX 2 avec contrainte historique | `studVMonitor_Box2.py` |

---

## Règle de maintenance

Lorsqu’une nouvelle variante est créée :

1. conserver le préfixe `studVMonitor_` ;
2. indiquer clairement la version, l’OS ou la BOX dans le nom ;
3. documenter la variante dans ce fichier ;
4. conserver un nom d’exécutable cohérent avec le nom du script source.

Exemple :

```text
src/studVMonitor_Box3_V1.py
```

peut produire :

```text
dist/studVMonitor_Box3_V1.exe
```
