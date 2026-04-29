# Studiovision-Autosync

Automatic image routing script for [StudioVision](https://www.studiodentaire.com/) — a dental practice management software.
When a medical imaging device saves a photo, the script detects it, identifies the open patient in StudioVision, moves the file to the correct patient folder on the network drive, and inserts a record in the Access database so the image appears in the patient's file immediately.

## How it works

1. **Watchdog** monitors a source folder for new image files.
2. The file is placed in a queue and picked up by a background worker thread.
3. The worker waits until the device has finished writing the file (lock check).
4. It polls the open StudioVision Access form via COM to get the current patient (code, last name, first name).
5. It queries `PUBLIC.MDB` to resolve the patient's folder on the network drive (using a previous `Photo externe` entry).
6. It moves the image into that folder, renaming it with a timestamp suffix if a conflict exists.
7. It inserts a new row in `PUBLIC.MDB` so StudioVision displays the image.
8. It requeried the `SFDoc` subform (with a `Refresh()` fallback) and moves the cursor to the last record so the new image is immediately visible.
9. If the insert fails, the already-moved file is sent to the orphan folder to avoid a ghost image with no database record.
10. If no patient is found within 15 minutes, the file is moved to the orphan folder.

## Requirements

- **Windows only** — relies on `win32com` (COM automation) and `pyodbc` (Access ODBC driver).
- Python 3.10+
- Microsoft Access ODBC driver installed on the machine.

Install dependencies:

```bash
pip install -r requirements.txt
```

`requirements.txt` includes:

| Package | Purpose |
|---|---|
| `watchdog` | File system monitoring |
| `pyodbc` | Access database connection via ODBC |
| `pywin32` | COM automation for interacting with Access |

## Configuration

Before running, set the following paths at the top of `src/studiovision_monitor.py`:

| Variable | Description |
|---|---|
| `SOURCE_DIR` | Folder watched for new images (shared by the imaging device) |
| `ORPHAN_DIR` | Destination for files that could not be matched to a patient |
| `DEST_PHOTOS` | Root of the patient photo folders on the network drive |
| `PUBLIC_MDB` | Path to `PUBLIC.MDB` (StudioVision shared database) |
| `DOCUM_MDB` | Path to `DOCUM.MDB` (currently unused) |

Other tunable constants:

| Constant | Default | Description |
|---|---|---|
| `FILE_LOCK_RETRY_DELAY` | `3` s | Delay between retries when a file is still locked |
| `FILE_LOCK_MAX_ATTEMPTS` | `15` | Max retries before giving up on a locked file |
| `PATIENT_POLL_INTERVAL` | `3` s | How often to poll Access for an open patient |
| `PATIENT_WAIT_TIMEOUT` | `900` s | Time before orphaning a file if no patient is found (15 min) |
| `SFDOC_SUBFORM_NAME` | `"SFDoc"` | Name of the Access subform that lists documents — update if renamed |

## Watched extensions

The following file extensions are monitored by default:

`.jpg`, `.jpeg`, `.jfif`, `.png`, `.bmp`, `.tif`, `.tiff`, `.dcm`

To add or remove extensions, edit the `WATCHED_EXTENSIONS` set and update `EXAM_DESCRIPTION` accordingly.

## Running

```bash
python src/studiovision_monitor.py
```

The script logs to both the console and `image_router.log` in the working directory.
Stop it with `Ctrl+C` — it will finish processing any remaining queued files before exiting.

## Patient folder resolution

The script finds the patient folder by querying `PUBLIC.MDB` for a previous `Photo externe` entry belonging to the same patient code.
The `Photo externe` field stores a relative path in this format:

```
\<group_folder>\<patient_folder>\filename.jpg
```

The group folder and patient folder are extracted from this path and combined with `DEST_PHOTOS` to build the absolute path on disk.

## Orphan files

A file is moved to `ORPHAN_DIR` in two situations:

- No patient is open in StudioVision within the configured timeout.
- The database insert succeeds but the file was already moved — if the insert then fails, the moved file is orphaned to prevent a ghost image with no record.

All orphan events are logged as warnings and must be handled manually.

## Technical notes

- `pythoncom.CoInitialize()` is called on the worker thread because COM objects cannot be shared across threads.
- Inserts go exclusively into `PUBLIC.MDB`. There is no fallback to another database.
- Only the `SFDoc` subform is requeried after an insert, leaving all other open forms untouched to avoid destabilizing StudioVision.
- `MoveLast()` is called on the `SFDoc` recordset after the requery so the newly inserted image is already selected when the user clicks to view it.
