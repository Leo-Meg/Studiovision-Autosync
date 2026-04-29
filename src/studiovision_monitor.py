import pythoncom
import queue
import shutil
import sys
import threading
import time
import logging
from datetime import datetime
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# win32com is Windows-only; gracefully degrade on other platforms
try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# pyodbc requires the Microsoft Access ODBC driver to be installed
try:
    import pyodbc
    PYODBC_AVAILABLE = True
except ImportError:
    PYODBC_AVAILABLE = False


# Paths — must be configured before running
# Note: retype DOCUM_MDB manually, do not copy-paste (risk of invisible characters)
SOURCE_DIR  = Path(r"??")
ORPHAN_DIR  = Path(r"??")
DEST_PHOTOS = Path(r"??")
PUBLIC_MDB  = Path(r"??")
DOCUM_MDB   = Path(r"??")

WATCHED_EXTENSIONS     = {".jpg", ".jpeg", ".jfif", ".png", ".bmp", ".tif", ".tiff", ".dcm"}
FILE_LOCK_RETRY_DELAY  = 3    # seconds between lock-check retries
FILE_LOCK_MAX_ATTEMPTS = 15   # max retries before giving up on a locked file
PATIENT_POLL_INTERVAL  = 3    # seconds between Access polls
PATIENT_WAIT_TIMEOUT   = 900  # seconds before orphaning a file (15 min)

# Field names as they appear on the StudioVision patient form — do not translate
ACCESS_FIELD_CODE   = "Code patient"
ACCESS_FIELD_NOM    = "NOM"
ACCESS_FIELD_PRENOM = "Prénom"

# Name of the documents subform in Access — do not translate
SFDOC_SUBFORM_NAME = "SFDoc"

# Description inserted into the DB per file extension
EXAM_DESCRIPTION = {
    ".jpg":  "Image",
    ".jpeg": "Image",
    ".jfif": "Image",
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


def db_connect(mdb_path: Path):
    """Return a pyodbc connection to an Access MDB/ACCDB file."""
    return pyodbc.connect(
        f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_path};"
    )


def get_active_patient() -> dict | None:
    """Read the patient currently open in StudioVision via COM automation.

    Returns a dict with keys 'code', 'nom', 'prenom', or None if no patient
    form is active or COM is unavailable.
    """
    if not WIN32_AVAILABLE:
        return None
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            return None

        target = {ACCESS_FIELD_CODE, ACCESS_FIELD_NOM, ACCESS_FIELD_PRENOM}
        data: dict = {}

        for i in range(form.Controls.Count):
            ctrl = form.Controls(i)
            try:
                if str(ctrl.Name) in target:
                    data[ctrl.Name] = ctrl.Value
            except Exception:
                pass  # labels and buttons have no .Value

        if not target.issubset(data.keys()):
            return None  # active form is not the patient form

        return {
            "code":   str(data[ACCESS_FIELD_CODE]),
            "nom":    str(data[ACCESS_FIELD_NOM]),
            "prenom": str(data[ACCESS_FIELD_PRENOM]),
        }
    except Exception as e:
        log.debug(f"COM error: {e}")
        return None


def find_patient_folder(patient_code: str) -> Path | None:
    """Resolve the patient's photo folder from PUBLIC.MDB.

    Queries the most recent 'Photo externe' entry for the patient and extracts
    the group and patient folder names from the stored relative path.
    Returns the absolute Path if found and valid, otherwise None.
    """
    if not PYODBC_AVAILABLE:
        log.error("pyodbc not available.")
        return None
    if not PUBLIC_MDB.exists():
        log.error(f"PUBLIC.MDB not found: {PUBLIC_MDB}")
        return None
    try:
        conn   = db_connect(PUBLIC_MDB)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT TOP 1 [Photo externe] FROM Documents "
            "WHERE [code patient] = ? AND [Photo externe] IS NOT NULL",
            (int(patient_code),)
        )
        row = cursor.fetchone()
        conn.close()

        if not row or not row[0]:
            log.warning(f"No existing document found for patient {patient_code}.")
            return None

        # Expected format: \<group_folder>\<patient_folder>\filename.ext
        parts = row[0].strip().strip("\\").split("\\")
        if len(parts) < 2:
            log.error("Unexpected Photo externe format (value redacted).")
            return None

        folder = DEST_PHOTOS / parts[0] / parts[1]
        if not folder.is_dir():
            log.error(f"Folder found in DB but missing on disk: {folder}")
            return None

        log.info(f"Patient folder resolved: {folder}")
        return folder
    except Exception as e:
        log.error(f"DB folder lookup failed: {e}")
        return None


def insert_document(patient: dict, relative_path: str, description: str) -> bool:
    """Insert a document row exclusively into DOCUM_MDB. No fallback to PUBLIC_MDB is attempted.

    Returns False on any failure so the caller can orphan the image.
    """
    if not PYODBC_AVAILABLE:
        log.error("pyodbc not available — insert skipped.")
        return False

    if not DOCUM_MDB.exists():
        log.error(f"DOCUM_MDB not found or unreachable: {DOCUM_MDB} — insert aborted.")
        return False

    try:
        conn   = db_connect(DOCUM_MDB)
        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO Documents
                ([code patient], [Date], DESCRIPTIONS, TEXTE, [Photo externe], TypeVW, NumDocExterne)
            VALUES (?, ?, ?, ?, ?, 99, NULL)
            """,
            (int(patient["code"]), datetime.now(), description, relative_path, relative_path)
        )
        conn.commit()
        conn.close()
        log.info(f"Insert OK: patient={patient['code']} path='{relative_path}' db={DOCUM_MDB.name}")
        return True

    except Exception as e:
        # Covers network drops, locked files, driver errors, etc.
        log.error(f"Insert failed in {DOCUM_MDB.name}: {e} — no fallback attempted.")
        return False


# acSubform constant from the Access object model
_AC_SUBFORM = 112


def _find_and_requery_sfdoc(form) -> bool:
    """Recursively search for the SFDoc subform, requery it, and move to the last record.

    Returns True if the subform was found and processed.
    """
    for i in range(form.Controls.Count):
        ctrl = form.Controls(i)
        try:
            if ctrl.ControlType != _AC_SUBFORM:
                continue

            if ctrl.Name == SFDOC_SUBFORM_NAME:
                # Requery only this subform, leaving other forms untouched
                try:
                    ctrl.Form.Requery()
                    log.info(f"Requery() on '{SFDOC_SUBFORM_NAME}'")
                except Exception as e_req:
                    log.warning(f"Requery() unavailable on '{SFDOC_SUBFORM_NAME}' ({e_req}), trying Refresh()...")
                    try:
                        ctrl.Form.Refresh()
                        log.info(f"Refresh() on '{SFDOC_SUBFORM_NAME}'")
                    except Exception as e_ref:
                        log.warning(f"Refresh() also unavailable on '{SFDOC_SUBFORM_NAME}' ({e_ref})")

                # Move to the last record so the new image is already selected
                try:
                    ctrl.Form.Recordset.MoveLast()
                    log.info(f"MoveLast() on '{SFDOC_SUBFORM_NAME}'")
                except Exception as e:
                    log.debug(f"MoveLast() failed on '{SFDOC_SUBFORM_NAME}': {e}")

                return True

            if _find_and_requery_sfdoc(ctrl.Form):
                return True

        except Exception:
            pass

    return False


def refresh_ui() -> None:
    """Requery only the SFDoc subform to show the new record without touching other forms."""
    if not WIN32_AVAILABLE:
        return
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            log.warning("Refresh skipped: no active form in Access.")
            return

        if not _find_and_requery_sfdoc(form):
            log.warning(f"Subform '{SFDOC_SUBFORM_NAME}' not found — refresh skipped.")

    except Exception as e:
        log.warning(f"COM refresh failed (non-blocking): {e}")


def wait_for_file(file: Path) -> bool:
    """Block until the file is no longer locked by the imaging device.

    Returns True when the file can be opened, False after the maximum number of retries.
    """
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
    """Move source into dest_folder, appending a timestamp suffix on name conflict.

    Returns the final destination Path, or None on failure.
    """
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
    """Move an unmatched or failed file to the orphan folder."""
    log.warning(f"Orphaning: {file.name}")
    move_file(file, ORPHAN_DIR, label="ORPHAN")


def worker(file_queue: queue.Queue) -> None:
    """Background thread that processes image files one at a time.

    For each file:
      1. Wait for the device to finish writing (lock check).
      2. Poll Access until a patient form is open or the timeout expires.
      3. Resolve the patient folder from the database.
      4. Move the image into that folder.
      5. Insert a document record and refresh the Access UI.
         If the insert fails, orphan the already-moved file to avoid a
         ghost image with no database record.
    """
    # COM must be initialized per thread; it cannot be shared across threads
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

            if not wait_for_file(file):
                log.error(f"Aborting, persistent lock: {file.name}")
                file_queue.task_done()
                continue

            patient    = None
            start_time = time.monotonic()
            first_log  = True

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

            log.info(f"Patient matched (code {patient['code']})")

            patient_folder = find_patient_folder(patient["code"])
            if not patient_folder:
                log.error(f"Could not resolve folder for patient {patient['code']}. Orphaning.")
                orphan_file(file)
                file_queue.task_done()
                continue

            dest = move_file(file, patient_folder)
            if dest is None:
                file_queue.task_done()
                continue

            group_name    = patient_folder.parent.name
            relative_path = f"\\{group_name}\\{patient_folder.name}\\{dest.name}"
            description   = EXAM_DESCRIPTION.get(file.suffix.lower(), "Image")

            if insert_document(patient, relative_path, description):
                # Give Access time to see the committed record before requerying
                time.sleep(1.5)
                refresh_ui()
            else:
                # Orphan the already-moved file to avoid a ghost image with no DB record
                log.error("Insert failed — orphaning the moved file to avoid a record-less image.")
                orphan_file(dest)

            file_queue.task_done()

    finally:
        pythoncom.CoUninitialize()


class ImageProducer(FileSystemEventHandler):
    """Watchdog event handler — enqueues new image files for the worker thread."""

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

    log.info("Image Router v3.6 started")
    log.info(f"  Source     : {SOURCE_DIR}")
    log.info(f"  Dest       : {DEST_PHOTOS}")
    log.info(f"  PUBLIC.MDB : {PUBLIC_MDB}")
    log.info(f"  DOCUM.MDB  : {DOCUM_MDB}")
    log.info(f"  Orphans    : {ORPHAN_DIR}")
    log.info(f"  Timeout    : {PATIENT_WAIT_TIMEOUT // 60} min")
    log.info(f"  Ext        : {', '.join(sorted(WATCHED_EXTENSIONS))}")

    file_queue: queue.Queue = queue.Queue()

    # Single daemon worker processes files sequentially to avoid race conditions
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

        # Drain the queue before exiting so no image is lost
        remaining = file_queue.qsize()
        if remaining:
            log.info(f"Waiting for {remaining} remaining file(s)...")
            file_queue.join()

        log.info("Image Router stopped.")


if __name__ == "__main__":
    main()