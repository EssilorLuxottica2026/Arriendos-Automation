from __future__ import annotations

import atexit
import os
from pathlib import Path
import shutil
import sys
import tempfile


APP_NAME = "ArriendosAutomation"
DICTIONARY_FILENAME = "Diccionario_Conceptos_Simple.xlsx"
FROZEN = bool(getattr(sys, "frozen", False))
RESOURCE_DIR = (
    Path(getattr(sys, "_MEIPASS")).resolve()
    if FROZEN and getattr(sys, "_MEIPASS", None)
    else Path(__file__).resolve().parents[2]
)
EXECUTABLE_DIR = Path(sys.executable).resolve().parent if FROZEN else RESOURCE_DIR
LEARNING_DICTIONARY_PATH = EXECUTABLE_DIR / DICTIONARY_FILENAME

_session_dir: Path | None = None


def _windows_user_data_root() -> Path:
    override = os.environ.get("ARRIENDOS_USER_DATA_DIR")
    if override:
        return Path(override).expanduser().resolve()
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "EssilorLuxottica" / APP_NAME
    return Path.home() / "AppData" / "Local" / "EssilorLuxottica" / APP_NAME


if FROZEN:
    _session_dir = Path(tempfile.mkdtemp(prefix=f"{APP_NAME}_"))
    DATA_DIR = _session_dir / "data"
    UPLOAD_DIR = DATA_DIR / "uploads"
    OUTPUT_DIR = DATA_DIR / "output"
    USER_DATA_DIR = _windows_user_data_root()
    INPUT_DIR = USER_DATA_DIR / "masters"
    RAW_DIR = USER_DATA_DIR / "backups" / "masters"
    DICTIONARY_BACKUP_DIR = USER_DATA_DIR / "backups" / "dictionary"
else:
    DATA_DIR = RESOURCE_DIR / "data"
    UPLOAD_DIR = DATA_DIR / "uploads"
    OUTPUT_DIR = DATA_DIR / "output"
    INPUT_DIR = DATA_DIR / "input"
    RAW_DIR = DATA_DIR / "raw"
    USER_DATA_DIR = DATA_DIR
    DICTIONARY_BACKUP_DIR = RAW_DIR / "dictionary"


def initialize_runtime() -> None:
    for path in (
        DATA_DIR,
        UPLOAD_DIR,
        OUTPUT_DIR,
        INPUT_DIR,
        RAW_DIR,
        DICTIONARY_BACKUP_DIR,
    ):
        path.mkdir(parents=True, exist_ok=True)

def cleanup_runtime() -> None:
    if _session_dir and _session_dir.exists():
        shutil.rmtree(_session_dir, ignore_errors=True)


def session_dir() -> Path | None:
    return _session_dir


initialize_runtime()
atexit.register(cleanup_runtime)
