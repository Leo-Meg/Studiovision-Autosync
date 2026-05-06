# README-lecture-log — lecture de `image_router.log`

Ce fichier explique comment lire le journal `image_router.log`.

Le log sert à vérifier le comportement réel du programme : démarrage, détection des images, patient actif, déplacement du fichier, insertion dans `PUBLIC.MDB`, actualisation StudioVision et erreurs.

---

## Emplacements possibles

Selon le script utilisé, le fichier peut être créé ici :

```text
%USERPROFILE%\studiovision\image_router.log
```

ou ici :

```text
image_router.log
```

ou, si le chemin a été rendu explicite dans le script :

```text
C:\Studiovision-Autosync\logs\image_router.log
```

Pour une installation stable, préférer :

```text
C:\Studiovision-Autosync\logs\image_router.log
```

---

## Format d’une ligne de log

Format courant :

```text
2026-05-06 10:15:30  INFO      [Worker]  Message
```

Signification :

| Partie | Signification |
|---|---|
| `2026-05-06 10:15:30` | Date et heure de l’événement. |
| `INFO` | Niveau du message. |
| `[Worker]` | Thread qui a écrit le message. |
| `Message` | Action réalisée ou erreur détectée. |

---

## Niveaux de log

| Niveau | Signification | Action attendue |
|---|---|---|
| `INFO` | Fonctionnement normal. | Lecture simple. |
| `WARNING` | Situation anormale mais non bloquante. | Vérifier le contexte. |
| `ERROR` | Échec d’une action importante. | Corriger le problème. |
| `DEBUG` | Détail technique. | Utile surtout en diagnostic. |

---

## Threads fréquents

| Thread | Signification |
|---|---|
| `MainThread` | Démarrage, arrêt, initialisation, verrou mono-instance. |
| `Worker` | Traitement des fichiers détectés. |
| Thread watchdog | Détection d’événements fichiers dans le dossier surveillé. |

Le nom exact peut varier selon la version du script.

---

## Démarrage du programme

Messages attendus :

```text
Watching: <SOURCE_DIR>
```

ou :

```text
Started image router
```

Signification :

- le programme est lancé ;
- le dossier source est surveillé ;
- le traitement peut commencer.

À vérifier :

- le chemin affiché correspond bien à `SOURCE_DIR` ;
- le dossier existe ;
- le compte Windows a le droit de lecture sur ce dossier.

---

## Détection d’un fichier

Message typique :

```text
Detected file: image001.jpg
```

ou :

```text
Queued: image001.jpg
```

Signification :

- un nouveau fichier compatible a été détecté ;
- il a été ajouté à la file de traitement.

Si rien n’apparaît :

- vérifier `SOURCE_DIR` ;
- vérifier l’extension du fichier ;
- vérifier que l’appareil écrit bien dans le dossier surveillé.

---

## Fichier encore verrouillé

Message typique :

```text
File locked (1/15), retrying...
```

Signification :

- le fichier existe ;
- l’appareil ou Windows écrit encore dedans ;
- le programme attend avant de le déplacer.

Si le message se répète puis échoue :

```text
File still locked after 15 attempts
```

Causes probables :

- l’appareil garde le fichier ouvert ;
- le fichier est incomplet ;
- un antivirus scanne le fichier ;
- un autre logiciel utilise le fichier.

Action :

- vérifier l’appareil ;
- vérifier les droits ;
- attendre ou tester avec un autre fichier.

---

## Patient actif trouvé

Message typique :

```text
Active patient: code=12345
```

Signification :

- StudioVision / Access est ouvert ;
- le formulaire actif contient les champs attendus ;
- le programme a récupéré le code patient.

Si aucun patient n’est trouvé :

```text
No active patient
```

ou :

```text
Waiting for active patient
```

Causes probables :

- StudioVision n’est pas ouvert ;
- aucun patient n’est affiché ;
- le formulaire actif n’est pas le bon ;
- les champs `Code patient`, `NOM`, `Prénom` ne sont pas accessibles ;
- Access ne répond pas via COM.

Action :

- ouvrir StudioVision ;
- afficher le bon patient ;
- vérifier que la fenêtre StudioVision / Access est active ;
- relancer le programme si COM Access ne répond plus.

---

## Dossier patient trouvé

Message typique :

```text
Patient folder resolved: C:\...\Photos\...
```

Signification :

- le programme a lu `PUBLIC.MDB` ;
- il a trouvé un champ `[Photo externe]` exploitable ;
- il a reconstruit le dossier patient ;
- le dossier existe sur disque.

Erreur possible :

```text
No existing document found for patient 12345.
```

Signification :

- le patient existe peut-être ;
- mais aucune ligne documentaire exploitable n’a été trouvée pour résoudre le dossier photo.

Erreur possible :

```text
Folder found in DB but missing on disk
```

Signification :

- la base contient un chemin ;
- le dossier correspondant n’existe pas sur disque.

Action :

- vérifier `PUBLIC_MDB` ;
- vérifier `DEST_PHOTOS` ;
- vérifier l’existence du dossier patient ;
- vérifier les droits réseau ou disque.

---

## Déplacement du fichier

Message typique :

```text
Moved: image001.jpg -> C:\...\patient\image001.jpg
```

Signification :

- le fichier a été déplacé vers le dossier patient.

Message possible :

```text
Name conflict, renamed to image001_1715000000.jpg
```

Signification :

- un fichier du même nom existait déjà ;
- le programme a ajouté un suffixe pour éviter l’écrasement.

Erreur possible :

```text
Move failed
```

Causes probables :

- droits insuffisants ;
- dossier cible inexistant ;
- fichier encore verrouillé ;
- chemin trop long ;
- disque ou partage réseau indisponible.

---

## Insertion dans `PUBLIC.MDB`

Message typique :

```text
Insert OK: patient=12345 path='...' db=PUBLIC.MDB
```

Signification :

- une ligne a été ajoutée dans la table `Documents` ;
- StudioVision peut référencer le fichier.

Erreur possible :

```text
DB insert failed
```

Causes probables :

- `PUBLIC.MDB` inaccessible ;
- pilote ODBC Access absent ;
- base verrouillée ;
- droits insuffisants ;
- champ attendu absent ou différent ;
- valeur patient invalide.

Action :

- vérifier le chemin `PUBLIC_MDB` ;
- vérifier le pilote ODBC Microsoft Access ;
- vérifier les droits d’écriture ;
- vérifier que la base n’est pas verrouillée exclusivement.

---

## Actualisation de StudioVision

Messages possibles :

```text
Requery() on 'SFDoc'
```

```text
Refresh() on 'SFDoc'
```

```text
MoveLast() on 'SFDoc'
```

Signification :

- le programme tente d’actualiser la liste documentaire ;
- `MoveLast()` place l’affichage sur le dernier document.

Avec les versions V2 / V3 :

- le refresh vise surtout le sous-formulaire `SFDoc` ;
- le formulaire parent est évité pour limiter les effets de bord.

Erreur possible :

```text
Subform 'SFDoc' not found
```

Signification :

- le formulaire actif ne contient pas le sous-formulaire attendu ;
- ou le nom du sous-formulaire diffère.

Action :

- vérifier que le bon écran StudioVision est ouvert ;
- vérifier la constante `SFDOC_SUBFORM_NAME`.

---

## Fichier envoyé en orphelin

Message typique :

```text
Moved to orphan
```

ou :

```text
No patient found before timeout
```

Signification :

- le programme n’a pas pu associer le fichier à un patient ;
- le fichier a été déplacé dans `ORPHAN_DIR`.

Causes fréquentes :

- aucun patient ouvert ;
- mauvais formulaire actif ;
- délai `PATIENT_WAIT_TIMEOUT` dépassé ;
- dossier patient introuvable ;
- base Access inaccessible.

Action :

- ouvrir le bon patient ;
- contrôler `ORPHAN_DIR` ;
- réintégrer manuellement le fichier si nécessaire.

---

## Deuxième instance bloquée avec V3

Message typique :

```text
Another instance is already running. Exiting to prevent duplicate processing.
```

Signification :

- une instance de `studVMonitor_V3` tourne déjà ;
- la nouvelle instance s’arrête ;
- le verrou mono-instance fonctionne.

Action :

- ne pas lancer une deuxième fois le programme ;
- vérifier le Gestionnaire des tâches ;
- vérifier les raccourcis Startup en double.

---

## Erreur COM Access

Message typique :

```text
COM error
```

ou :

```text
COM refresh failed
```

Signification :

- le programme n’a pas réussi à communiquer avec Access / StudioVision.

Causes fréquentes :

- StudioVision fermé ;
- Access ne répond pas ;
- mauvais formulaire actif ;
- session Windows verrouillée ou instable ;
- problème COM temporaire.

Action :

- ouvrir StudioVision ;
- afficher un patient ;
- relancer Access / StudioVision ;
- relancer `studVMonitor` si nécessaire.

---

## Erreur `pyodbc`

Message typique :

```text
pyodbc not available
```

Signification :

- le module Python `pyodbc` n’est pas disponible ;
- ou il n’a pas été embarqué correctement dans l’exécutable.

Action :

```bash
pip install pyodbc
```

Puis recompiler avec PyInstaller si nécessaire.

---

## Erreur pilote Access

Message typique :

```text
Microsoft Access Driver (*.mdb, *.accdb)
```

ou :

```text
Data source name not found
```

Signification :

- le pilote ODBC Access est absent ou incompatible.

Action :

- installer Microsoft Access Runtime ou le moteur ODBC Access adapté ;
- vérifier la compatibilité 32/64 bits entre Python, pyodbc, Access et le pilote.

---

## Lecture rapide d’un traitement réussi

Un traitement réussi ressemble à cette séquence :

```text
Detected file
File ready
Active patient
Patient folder resolved
Moved
Insert OK
Requery() on 'SFDoc'
MoveLast() on 'SFDoc'
```

Conclusion :

- l’image a été détectée ;
- le patient a été identifié ;
- le fichier a été déplacé ;
- la base a été mise à jour ;
- StudioVision a été actualisé.

---

## Lecture rapide d’un traitement échoué

Exemples de séquences à surveiller :

```text
File still locked after 15 attempts
```

Problème fichier.

```text
No active patient
Moved to orphan
```

Problème patient actif.

```text
PUBLIC.MDB not found
```

Problème chemin base Access.

```text
DB insert failed
```

Problème écriture base Access.

```text
Subform 'SFDoc' not found
```

Problème écran StudioVision ou nom du sous-formulaire.

```text
Another instance is already running
```

Deuxième lancement bloqué par la V3.

---

## Checklist de diagnostic

En cas de problème, vérifier dans cet ordre :

1. Le programme démarre-t-il ?
2. Le bon `SOURCE_DIR` est-il surveillé ?
3. Le fichier est-il détecté ?
4. Le fichier reste-t-il verrouillé ?
5. Le patient actif est-il trouvé ?
6. Le dossier patient est-il résolu ?
7. Le fichier est-il déplacé ?
8. L’insertion `PUBLIC.MDB` est-elle réussie ?
9. StudioVision est-il actualisé ?
10. Le fichier est-il parti en orphelin ?
11. Une autre instance tourne-t-elle déjà ?
