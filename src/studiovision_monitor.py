import queue
import shutil
import sys
import threading
import time
import logging
import unicodedata
from datetime import datetime
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

try:
    import pyodbc
    PYODBC_AVAILABLE = True
except ImportError:
    PYODBC_AVAILABLE = False


# Paths
SOURCE_DIR  = Path(r"??")
ORPHAN_DIR  = SOURCE_DIR.parent / "??"
DEST_PHOTOS = Path(r"??")
PUBLIC_MDB  = Path(r"??")

# File types to watch
WATCHED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".dcm"}

# How long to retry a locked file (15 attempts x 3s = 45s max)
FILE_LOCK_RETRY_DELAY  = 3
FILE_LOCK_MAX_ATTEMPTS = 15

# How long to wait for a patient to be opened in Access (poll every 3s, give up after 15min)
PATIENT_POLL_INTERVAL = 3
PATIENT_WAIT_TIMEOUT  = 900

# How many characters to take from last name and first name to build the folder pattern
NOM_CHARS    = 4
PRENOM_CHARS = 3

# Field names as they appear in the StudioVision Access form
ACCESS_FIELD_CODE   = "Code patient"
ACCESS_FIELD_NOM    = "NOM"
ACCESS_FIELD_PRENOM = "Prénom"

# Description written to the DB per file type
EXAM_DESCRIPTION = {
    ".jpg":  "Image",
    ".jpeg": "Image",
    ".png":  "Image",
    ".bmp":  "Image",
    ".tif":  "OCT",
    ".tiff": "OCT",
    ".dcm":  "DICOM",
}


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  [%(threadName)s]  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler("image_router.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("image_router")


def normaliser(texte: str) -> str:
    nfkd = unicodedata.normalize("NFKD", texte)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).lower().strip()


def construire_motif(patient: dict) -> str:
    # Build the patient folder name: code + 4 chars of last name + 3 chars of first name
    code   = str(patient["code"])
    nom    = normaliser(patient["nom"])[:NOM_CHARS]
    prenom = normaliser(patient["prenom"])[:PRENOM_CHARS]
    return f"{code}{nom}.{prenom}"


def get_patient_actif() -> dict | None:
    # Read the currently open patient from the StudioVision Access form via COM.
    # Returns a dict or None if Access is not open / no patient is displayed.
    # To test on Mac: uncomment the mock return below.
    if not WIN32_AVAILABLE:

        log.debug("win32com unavailable (non-Windows).")
        return None

    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            return None

        champs_cibles = {ACCESS_FIELD_CODE, ACCESS_FIELD_NOM, ACCESS_FIELD_PRENOM}
        data: dict = {}

        for i in range(form.Controls.Count):
            ctrl = form.Controls(i)
            try:
                if str(ctrl.Name) in champs_cibles:
                    data[ctrl.Name] = ctrl.Value
            except Exception:
                pass  # labels and buttons have no .Value

        if not champs_cibles.issubset(data.keys()):
            return None  # open form is not the patient form

        return {
            "code":   str(data[ACCESS_FIELD_CODE]),
            "nom":    str(data[ACCESS_FIELD_NOM]),
            "prenom": str(data[ACCESS_FIELD_PRENOM]),
        }

    except Exception as e:
        log.debug(f"COM Access unavailable: {e}")
        return None


def trouver_dossier_patient(patient: dict, motif: str) -> Path | None:
    # Fallback: scan all group folders if Fast Path misses (negative codes, edge cases).
    if not DEST_PHOTOS.exists():
        log.error(f"Network drive unreachable: {DEST_PHOTOS}")
        return None

    try:
        code_int      = int(patient["code"])
        numero_groupe = abs(code_int) // 1000
        groupe_nom    = f"{numero_groupe:02d}.000"
        candidat      = DEST_PHOTOS / groupe_nom / motif

        if candidat.is_dir():
            log.info(f"[Fast Path] Found -> {candidat}")
            return candidat

        log.debug(f"[Fast Path] '{candidat}' not found -> Fallback")

    except (ValueError, TypeError) as e:
        log.debug(f"[Fast Path] Non-numeric code '{patient['code']}' ({e}) -> Fallback")

    log.info(f"[Fallback] Scanning {DEST_PHOTOS} for '{motif}'...")
    for groupe in DEST_PHOTOS.iterdir():
        if not groupe.is_dir():
            continue
        candidat = groupe / motif
        if candidat.is_dir():
            log.info(f"[Fallback] Found -> {candidat}")
            return candidat

    log.warning(f"No folder found for '{motif}'")
    return None


def inserer_document(patient: dict, chemin_relatif: str, description: str) -> bool:
    # Insert one row in PUBLIC.MDB > Documents so StudioVision shows the image.
    if not PYODBC_AVAILABLE:
        log.warning("pyodbc unavailable — DB insert skipped.")
        return False

    if not PUBLIC_MDB.exists():
        log.error(f"PUBLIC.MDB not found: {PUBLIC_MDB}")
        return False

    try:
        conn   = pyodbc.connect(f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={PUBLIC_MDB};")
        cursor = conn.cursor()

        # Use MAX(NUMDOC) + 1 as the next unique ID
        cursor.execute("SELECT MAX(NUMDOC) FROM Documents")
        row    = cursor.fetchone()
        numdoc = int(row[0] if row[0] is not None else 0) + 1

        cursor.execute(
            """
            INSERT INTO Documents
                (NUMDOC, [code patient], Date, DESCRIPTIONS, TEXTE, [Photo externe], TypeVW, NumDocExterne)
            VALUES (?, ?, ?, ?, NULL, ?, 99, NULL)
            """,
            (numdoc, int(patient["code"]), datetime.now(), description, chemin_relatif)
        )
        conn.commit()
        conn.close()

        log.info(f"[DB] INSERT OK — NUMDOC={numdoc}  patient={patient['code']}  path='{chemin_relatif}'")
        return True

    except Exception as e:
        log.error(f"[DB] INSERT failed: {e}")
        return False


def attendre_disponibilite(fichier: Path) -> bool:
    # Wait until the file is no longer locked by the medical device.
    for tentative in range(1, FILE_LOCK_MAX_ATTEMPTS + 1):
        try:
            with fichier.open("rb"):
                return True
        except (PermissionError, OSError):
            log.debug(f"File locked ({tentative}/{FILE_LOCK_MAX_ATTEMPTS}), retrying in {FILE_LOCK_RETRY_DELAY}s...")
            time.sleep(FILE_LOCK_RETRY_DELAY)

    log.error(f"File still locked after {FILE_LOCK_MAX_ATTEMPTS} attempts: {fichier}")
    return False


def deplacer_vers(source: Path, dossier_dest: Path, label: str = "") -> Path | None:
    # Move source to dossier_dest. Adds a timestamp suffix if filename already exists.
    # Returns the final path on success, None on failure.
    dossier_dest.mkdir(parents=True, exist_ok=True)
    dest = dossier_dest / source.name

    if dest.exists():
        ts   = int(time.time())
        dest = dossier_dest / f"{source.stem}_{ts}{source.suffix}"
        log.info(f"Name conflict -> renamed to {dest.name}")

    try:
        shutil.move(str(source), str(dest))
        tag = f"[{label}]  " if label else ""
        log.info(f"✅  {tag}{source.name}  ->  {dest}")
        return dest
    except Exception as e:
        log.error(f"Move failed: {e}")
        return None


def orpheliner(fichier: Path) -> None:
    # No patient found after timeout — move file to the orphan folder.
    log.warning(f"Timeout reached — orphaning: {fichier.name}")
    deplacer_vers(fichier, ORPHAN_DIR, label="ORPHAN")


def worker(file_queue: queue.Queue) -> None:
    # Background thread that processes one image at a time from the queue.
    log.info("Worker started — waiting for images...")

    while True:
        try:
            fichier: Path = file_queue.get()
        except Exception as e:
            log.error(f"Queue read error: {e}")
            continue

        log.info(f"Processing: {fichier.name}  ({file_queue.qsize()} pending)")

        if not fichier.exists():
            log.warning(f"File gone before processing: {fichier}")
            file_queue.task_done()
            continue

        # Step 1: wait until the device has finished writing the file
        if not attendre_disponibilite(fichier):
            log.error(f"Aborting (persistent lock): {fichier.name}")
            file_queue.task_done()
            continue

        # Step 2: poll Access until a patient is open, or timeout after 15min
        patient       = None
        debut_attente = time.monotonic()
        premier_log   = True

        while True:
            patient = get_patient_actif()

            if patient:
                break

            temps_ecoule = time.monotonic() - debut_attente

            if temps_ecoule >= PATIENT_WAIT_TIMEOUT:
                orpheliner(fichier)
                file_queue.task_done()
                patient = None
                break

            if premier_log:
                log.info(f"No patient open — waiting (timeout in {PATIENT_WAIT_TIMEOUT // 60} min)")
                premier_log = False
            else:
                log.debug(f"Still waiting... ({(PATIENT_WAIT_TIMEOUT - temps_ecoule) / 60:.1f} min left)")

            time.sleep(PATIENT_POLL_INTERVAL)

        if patient is None:
            continue

        # Step 3: build the patient folder name pattern
        motif = construire_motif(patient)
        log.info(f"Patient: {patient['nom']} {patient['prenom']} (code {patient['code']}) -> '{motif}'")

        # Step 4: find the folder on disk
        dossier_patient = trouver_dossier_patient(patient, motif)
        if not dossier_patient:
            log.error(f"Folder '{motif}' not found. Orphaning.")
            orpheliner(fichier)
            file_queue.task_done()
            continue

        # Step 5: move the image into the patient folder
        dest = deplacer_vers(fichier, dossier_patient)
        if dest is None:
            file_queue.task_done()
            continue

        # Step 6: insert a row in PUBLIC.MDB so StudioVision can see the image
        groupe_reel    = dossier_patient.parent.name
        chemin_relatif = f"\\{groupe_reel}\\{motif}\\{dest.name}"
        description    = EXAM_DESCRIPTION.get(fichier.suffix.lower(), "Image")

        inserer_document(patient, chemin_relatif, description)

        file_queue.task_done()


class ImageProducer(FileSystemEventHandler):
    # Watchdog handler — only puts the file path in the queue, never blocks.

    def __init__(self, file_queue: queue.Queue) -> None:
        super().__init__()
        self._queue = file_queue

    def on_created(self, event) -> None:
        if event.is_directory:
            return

        fichier = Path(event.src_path)

        if fichier.suffix.lower() not in WATCHED_EXTENSIONS:
            log.debug(f"Ignored (extension not watched): {fichier.name}")
            return

        log.info(f"Enqueued: {fichier.name}  (queue -> {self._queue.qsize() + 1} item(s))")
        self._queue.put(fichier)


def main() -> None:
    if not SOURCE_DIR.exists():
        log.critical(f"Source folder not found: {SOURCE_DIR}")
        sys.exit(1)

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("Image Router v3.0 started")
    log.info(f"  Source   : {SOURCE_DIR}")
    log.info(f"  Dest     : {DEST_PHOTOS}")
    log.info(f"  Database : {PUBLIC_MDB}")
    log.info(f"  Orphans  : {ORPHAN_DIR}")
    log.info(f"  Timeout  : {PATIENT_WAIT_TIMEOUT // 60} min")
    log.info(f"  Ext      : {', '.join(sorted(WATCHED_EXTENSIONS))}")

    file_queue: queue.Queue = queue.Queue()

    worker_thread = threading.Thread(target=worker, args=(file_queue,), name="Worker", daemon=True)
    worker_thread.start()
    log.info("Worker thread started.")

    producer = ImageProducer(file_queue)
    observer = Observer()
    observer.schedule(producer, str(SOURCE_DIR), recursive=True)
    observer.start()
    log.info("Watchdog observer started. Press Ctrl+C to stop.")

    try:
        while observer.is_alive():
            time.sleep(1)
    except KeyboardInterrupt:
        log.info("Shutdown requested.")
    finally:
        observer.stop()
        observer.join()

        remaining = file_queue.qsize()
        if remaining:
            log.info(f"Waiting for {remaining} remaining image(s)...")
            file_queue.join()

        log.info("Image Router stopped.")


if __name__ == "__main__":
    main()