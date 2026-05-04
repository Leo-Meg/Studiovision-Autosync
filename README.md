# Studiovision-Autosync

Automatic image routing script for [StudioVision](https://www.studiodentaire.com/) — a dental practice management software.  
When a medical imaging device saves a photo, the script detects it, identifies the open patient in StudioVision, moves the file to the correct patient folder on the network drive, and inserts a record in the Access database so the image appears immediately in the patient's file.

---

## Scripts

Three variants are provided in `src/`. They share the same core logic and configuration constants.

| File | Version | Description |
|---|---|---|
| `studiovision_monitor.py` | v3.5 | Base version. Watches a flat source folder for new images. |
| `windows7.py` | v3.5 | Same as above, using `typing.Optional` for Python 3.9 / Windows 7 compatibility. |
| `box2.py` | v3.6 | Extended version with **Nidek device support** (see below). |
| `studiovision_monitorV2.py` | v3.6 | Improved base version with **batched UI refresh** and **SFDoc-only requery** (see below). |

---

## How it works

1. **Watchdog** monitors `SOURCE_DIR` for new image files (recursively).
2. Each detected file is pushed to a queue and picked up by a background worker thread.
3. The worker waits until the device has finished writing the file (lock-check with retries).
4. It polls the active StudioVision Access form via COM to get the current patient (code, last name, first name).
5. It queries `PUBLIC.MDB` to resolve the patient's folder on the network drive using an existing `Photo externe` entry.
6. It moves the image into that folder, appending a timestamp suffix on name conflict.
7. It inserts a new row into `PUBLIC.MDB` so StudioVision registers the image.
8. It requeried the `SFDoc` subform (with a `Refresh()` fallback) and moves to the last record so the new image is immediately visible.
9. If no patient is found within the configured timeout, the file is moved to the orphan folder.

---

## Nidek device support (`box2.py` only)

Nidek devices save scans as a set of files inside a sub-folder (`SOURCE_DIR/<device>/<scan>/`). `box2.py` handles this layout:

- Waits 2 seconds after the first file event to let the full scan land.
- Deletes XML sidecar files automatically.
- Keeps only the **largest image** in the scan folder; all others (thumbnails) are deleted.
- Cleans up the scan folder and its parent once the main image has been processed.
- Tracks already-processed scan folders to drop any residual files that arrive late in the queue.

Files not inside a Nidek sub-folder are processed normally (same as the base version).

---

## Requirements

- **Windows only** — requires `win32com` (COM automation) and `pyodbc` (Access ODBC driver).
- Python 3.10+ (`studiovision_monitor.py`, `box2.py`) or Python 3.9+ (`windows7.py`).
- Microsoft Access ODBC driver installed on the machine.

```bash
pip install -r requirements.txt
```

| Package | Purpose |
|---|---|
| `watchdog` | File system monitoring |
| `pyodbc` | Access database connection via ODBC |
| `pywin32` | COM automation for interacting with Access |

---

## Configuration

Set the following paths at the top of whichever script you run:

| Variable | Description |
|---|---|
| `SOURCE_DIR` | Folder watched for new images (shared by the imaging device) |
| `ORPHAN_DIR` | Destination for files that could not be matched to a patient |
| `DEST_PHOTOS` | Root of the patient photo folders on the network drive |
| `PUBLIC_MDB` | Path to `PUBLIC.MDB` (StudioVision shared database) |
| `DOCUM_MDB` | Path to `DOCUM.MDB` (reserved, currently unused) |

Other tunable constants:

| Constant | Default | Description |
|---|---|---|
| `FILE_LOCK_RETRY_DELAY` | `3` s | Delay between retries when a file is still locked |
| `FILE_LOCK_MAX_ATTEMPTS` | `15` | Max retries before giving up on a locked file |
| `PATIENT_POLL_INTERVAL` | `3` s | How often to poll Access for an open patient |
| `PATIENT_WAIT_TIMEOUT` | `900` s | Time before orphaning a file if no patient is found (15 min) |
| `SFDOC_SUBFORM_NAME` | `"SFDoc"` | Name of the Access subform listing documents — update if renamed |

---

## Watched extensions

The following file extensions are monitored by default:

`.jpg`, `.jpeg`, `.jfif`, `.png`, `.bmp`, `.tif`, `.tiff`, `.dcm`

File type → database description mapping:

| Extension | Description inserted |
|---|---|
| `.tif`, `.tiff` | `OCT` |
| `.dcm` | `DICOM` |
| all others | `Image` |

To add or remove extensions, edit `WATCHED_EXTENSIONS` and update `EXAM_DESCRIPTION` accordingly.

---

## Running

```bash
python src/box2.py
# or
python src/studiovision_monitorV2.py
# or
python src/studiovision_monitor.py
```

Logs are written to both the console and `image_router.log` in the working directory.  
Stop with `Ctrl+C` — the script will finish processing any remaining queued files before exiting.

---

## Patient folder resolution

The script finds the patient folder by querying `PUBLIC.MDB` for an existing `Photo externe` entry for the same patient code. That field stores a relative path:

```
\<group_folder>\<patient_folder>\filename.jpg
```

The group and patient folder names are extracted and combined with `DEST_PHOTOS` to build the absolute path on disk.

---

## Orphan files

A file is moved to `ORPHAN_DIR` when:

- No patient is open in StudioVision within the configured timeout.
- The patient folder cannot be resolved from the database.

All orphan events are logged as warnings and must be handled manually.

---

## Technical notes

- `pythoncom.CoInitialize()` / `CoUninitialize()` are called on the worker thread — COM objects cannot be shared across threads.
- `DOCUM.MDB` is read-only for inserts; all writes go to `PUBLIC.MDB`.
- The `windows7.py` variant is functionally identical to `studiovision_monitor.py` but avoids `X | None` union syntax for compatibility with Python 3.9.

---

## studiovision_monitorV2.py — what changed vs v3.5

### Batched UI refresh (burst debounce)

In v3.5, `refresh_ui()` was called immediately after every successful DB insert, which caused Access to freeze when a device sent several images in rapid succession.

`studiovision_monitorV2.py` replaces the blocking `file_queue.get()` with a `get(timeout=1.5)`. After each successful insert a `needs_refresh` flag is raised instead of calling `refresh_ui()` immediately. When the queue stays empty for 1.5 s (i.e. the burst is over), the refresh fires exactly once and the flag resets. This collapses N consecutive refreshes into a single one.

### SFDoc-only requery

In v3.5, `refresh_ui()` called `Requery()` / `Refresh()` on the entire active form, which reset the parent form's current-record pointer and sent the doctor back to record #1.

`studiovision_monitorV2.py` introduces `_find_sfdoc()`, which recursively walks the control tree to locate the `SFDoc` subform and requeried **only that subform**. The parent form's recordset is never touched.