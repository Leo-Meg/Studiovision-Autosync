import pythoncom
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

# Retry settings for locked files (15 x 3s = 45s max)
FILE_LOCK_RETRY_DELAY  = 3
FILE_LOCK_MAX_ATTEMPTS = 15

# Patient polling settings (every 3s, give up after 15min)
PATIENT_POLL_INTERVAL = 3
PATIENT_WAIT_TIMEOUT  = 900

# Characters used to build the patient folder name
NOM_CHARS    = 4
PRENOM_CHARS = 3

# Field names in the StudioVision Access form
ACCESS_FIELD_CODE   = "Code patient"
ACCESS_FIELD_NOM    = "NOM"
ACCESS_FIELD_PRENOM = "Prénom"

# Description inserted in DB per file type
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


def normalise(text: str) -> str:
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).lower().strip()


def build_folder_pattern(patient: dict) -> str:
    code   = str(patient["code"])
    nom    = normalise(patient["nom"])[:NOM_CHARS]
    prenom = normalise(patient["prenom"])[:PRENOM_CHARS]
    return f"{code}{nom}.{prenom}"


def get_active_patient() -> dict | None:
    if not WIN32_AVAILABLE:
        log.debug("win32com unavailable.")
        return None

    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form = access.Screen.ActiveForm
        if form is None:
            return None

        target_fields = {ACCESS_FIELD_CODE, ACCESS_FIELD_NOM, ACCESS_FIELD_PRENOM}
        data: dict = {}

        for i in range(form.Controls.Count):
            ctrl = form.Controls(i)
            try:
                if str(ctrl.Name) in target_fields:
                    data[ctrl.Name] = ctrl.Value
            except Exception:
                pass

        if not target_fields.issubset(data.keys()):
            return None

        return {
            "code":   str(data[ACCESS_FIELD_CODE]),
            "nom":    str(data[ACCESS_FIELD_NOM]),
            "prenom": str(data[ACCESS_FIELD_PRENOM]),
        }

    except Exception as e:
        log.debug(f"COM error: {e}")
        return None


def find_folder_from_db(patient: dict) -> Path | None:
    # Look up an existing document for this patient to extract their folder path.
    if not PYODBC_AVAILABLE or not PUBLIC_MDB.exists():
        return None

    try:
        conn = pyodbc.connect(
            f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={PUBLIC_MDB};"
        )
        cursor = conn.cursor()
        cursor.execute(
            "SELECT TOP 1 [Photo externe] FROM Documents "
            "WHERE [code patient] = ? AND [Photo externe] IS NOT NULL",
            (int(patient["code"]),)
        )
        row = cursor.fetchone()
        conn.close()

        if not row or not row[0]:
            log.info(f"No existing document for patient {patient['code']}")
            return None

        # Photo externe format: \17.000\1758506693bon.eri\image.jpg
        parts = row[0].strip().strip("\\").split("\\")
        if len(parts) >= 2:
            folder = DEST_PHOTOS / parts[0] / parts[1]
            if folder.is_dir():
                log.info(f"Folder found via DB: {folder}")
                return folder
            log.warning(f"DB path found but folder missing on disk: {folder}")

        return None

    except Exception as e:
        log.error(f"DB folder lookup failed: {e}")
        return None


def find_folder_on_disk(patient: dict, pattern: str) -> Path | None:
    # Fallback: scan all group folders when DB lookup fails.
    if not DEST_PHOTOS.exists():
        log.error(f"Network drive unreachable: {DEST_PHOTOS}")
        return None

    try:
        code_int      = int(patient["code"])
        group_number  = abs(code_int) // 1000
        group_name    = f"{group_number:02d}.000"
        candidate     = DEST_PHOTOS / group_name / pattern

        if candidate.is_dir():
            log.info(f"Fast path found: {candidate}")
            return candidate

        log.debug(f"Fast path miss, scanning...")

    except (ValueError, TypeError):
        pass

    for group in DEST_PHOTOS.iterdir():
        if not group.is_dir():
            continue
        candidate = group / pattern
        if candidate.is_dir():
            log.info(f"Fallback found: {candidate}")
            return candidate

    log.warning(f"No folder found for '{pattern}'")
    return None


def insert_document(patient: dict, relative_path: str, description: str) -> bool:
    if not PYODBC_AVAILABLE:
        log.warning("pyodbc unavailable, DB insert skipped.")
        return False

    if not PUBLIC_MDB.exists():
        log.error(f"PUBLIC.MDB not found: {PUBLIC_MDB}")
        return False

    try:
        conn   = pyodbc.connect(
            f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={PUBLIC_MDB};"
        )
        cursor = conn.cursor()

        cursor.execute("SELECT MAX(NUMDOC) FROM Documents")
        row    = cursor.fetchone()
        numdoc = int(row[0] if row[0] is not None else 0) + 1

        cursor.execute(
            """
            INSERT INTO Documents
                (NUMDOC, [code patient], Date, DESCRIPTIONS, TEXTE, [Photo externe], TypeVW, NumDocExterne)
            VALUES (?, ?, ?, ?, NULL, ?, 99, NULL)
            """,
            (numdoc, int(patient["code"]), datetime.now(), description, relative_path)
        )
        conn.commit()
        conn.close()

        log.info(f"DB insert OK: NUMDOC={numdoc} patient={patient['code']} path='{relative_path}'")
        return True

    except Exception as e:
        log.error(f"DB insert failed: {e}")
        return False


def wait_for_file(file: Path) -> bool:
    for attempt in range(1, FILE_LOCK_MAX_ATTEMPTS + 1):
        try:
            with file.open("rb"):
                return True
        except (PermissionError, OSError):
            log.debug(f"File locked ({attempt}/{FILE_LOCK_MAX_ATTEMPTS}), retrying...")
            time.sleep(FILE_LOCK_RETRY_DELAY)

    log.error(f"File still locked after {FILE_LOCK_MAX_ATTEMPTS} attempts: {file}")
    return False


def move_file(source: Path, dest_folder: Path, label: str = "") -> Path | None:
    dest_folder.mkdir(parents=True, exist_ok=True)
    dest = dest_folder / source.name

    if dest.exists():
        ts   = int(time.time())
        dest = dest_folder / f"{source.stem}_{ts}{source.suffix}"
        log.info(f"Name conflict, renamed to {dest.name}")

    try:
        shutil.move(str(source), str(dest))
        tag = f"[{label}]  " if label else ""
        log.info(f"{tag}{source.name} -> {dest}")
        return dest
    except Exception as e:
        log.error(f"Move failed: {e}")
        return None


def orphan_file(file: Path) -> None:
    log.warning(f"Orphaning: {file.name}")
    move_file(file, ORPHAN_DIR, label="ORPHAN")


def worker(file_queue: queue.Queue) -> None:
    # Initialize COM for this thread.
    pythoncom.CoInitialize()
    log.info("Worker started.")

    try:
        while True:
            try:
                file: Path = file_queue.get()
            except Exception as e:
                log.error(f"Queue error: {e}")
                continue

            log.info(f"Processing: {file.name} ({file_queue.qsize()} pending)")

            if not file.exists():
                log.warning(f"File gone before processing: {file}")
                file_queue.task_done()
                continue

            # Wait until the file is fully written and unlocked.
            if not wait_for_file(file):
                log.error(f"Aborting, persistent lock: {file.name}")
                file_queue.task_done()
                continue

            # Poll Access until a patient is open or timeout.
            patient       = None
            start_time    = time.monotonic()
            first_log     = True

            while True:
                patient = get_active_patient()

                if patient:
                    break

                elapsed = time.monotonic() - start_time

                if elapsed >= PATIENT_WAIT_TIMEOUT:
                    orphan_file(file)
                    file_queue.task_done()
                    patient = None
                    break

                if first_log:
                    log.info(f"No patient open, waiting (timeout in {PATIENT_WAIT_TIMEOUT // 60} min)")
                    first_log = False

                time.sleep(PATIENT_POLL_INTERVAL)

            if patient is None:
                continue

            pattern = build_folder_pattern(patient)
            log.info(f"Patient: {patient['nom']} {patient['prenom']} (code {patient['code']}) -> '{pattern}'")

            # Find the patient folder: DB lookup first, disk scan as fallback.
            patient_folder = find_folder_from_db(patient)
            if not patient_folder:
                log.info("DB lookup failed, falling back to disk scan.")
                patient_folder = find_folder_on_disk(patient, pattern)

            if not patient_folder:
                log.error(f"Folder not found for '{pattern}'. Orphaning.")
                orphan_file(file)
                file_queue.task_done()
                continue

            # Move the image into the patient folder.
            dest = move_file(file, patient_folder)
            if dest is None:
                file_queue.task_done()
                continue

            # Insert a row in Documents so StudioVision can display the image.
            group_name    = patient_folder.parent.name
            relative_path = f"\\{group_name}\\{patient_folder.name}\\{dest.name}"
            description   = EXAM_DESCRIPTION.get(file.suffix.lower(), "Image")

            insert_document(patient, relative_path, description)

            file_queue.task_done()

    finally:
        pythoncom.CoUninitialize()


class ImageProducer(FileSystemEventHandler):

    def __init__(self, file_queue: queue.Queue) -> None:
        super().__init__()
        self._queue = file_queue

    def on_created(self, event) -> None:
        if event.is_directory:
            return

        file = Path(event.src_path)

        if file.suffix.lower() not in WATCHED_EXTENSIONS:
            return

        log.info(f"Enqueued: {file.name} (queue size: {self._queue.qsize() + 1})")
        self._queue.put(file)


def main() -> None:
    if not SOURCE_DIR.exists():
        log.critical(f"Source folder not found: {SOURCE_DIR}")
        sys.exit(1)

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("Image Router v3.1 started")
    log.info(f"  Source   : {SOURCE_DIR}")
    log.info(f"  Dest     : {DEST_PHOTOS}")
    log.info(f"  Database : {PUBLIC_MDB}")
    log.info(f"  Orphans  : {ORPHAN_DIR}")
    log.info(f"  Timeout  : {PATIENT_WAIT_TIMEOUT // 60} min")
    log.info(f"  Ext      : {', '.join(sorted(WATCHED_EXTENSIONS))}")

    file_queue: queue.Queue = queue.Queue()

    worker_thread = threading.Thread(target=worker, args=(file_queue,), name="Worker", daemon=True)
    worker_thread.start()

    producer = ImageProducer(file_queue)
    observer = Observer()
    observer.schedule(producer, str(SOURCE_DIR), recursive=True)
    observer.start()
    log.info("Watching for images. Press Ctrl+C to stop.")

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
            log.info(f"Waiting for {remaining} remaining file(s)...")
            file_queue.join()

        log.info("Image Router stopped.")


if __name__ == "__main__":
    main()