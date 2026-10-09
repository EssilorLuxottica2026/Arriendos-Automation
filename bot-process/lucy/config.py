"""Configuración y credenciales del bot Lucy.

Valores sensibles o dependientes del ambiente (LUCY_LOGIN_URL, LUCY_USERNAME,
LUCY_PASSWORD, LUCY_COMPANY_CODE y opcionalmente LUCY_BUNDLE_DIR) se leen de variables de entorno o de los archivos ``.env``
(primero bot-process/.env y luego el .env de la raíz del repositorio).
"""
import os
from dataclasses import dataclass
from pathlib import Path

BOT_DIR = Path(__file__).resolve().parent.parent
ENV_FILES = (BOT_DIR / ".env", BOT_DIR.parent / ".env")
CACHE_DIR = BOT_DIR / "cache"
DEFAULT_BUNDLE_DIR = Path.home() / "Downloads"
BUNDLE_PATTERN = "bot_single_vendor_*.zip"
BUNDLE_JSON_NAME = "control_facturas_bot.json"
LOGIN_TIMEOUT_MS = 30_000
SESSION_CHECK_TIMEOUT_MS = 3_000
FILTER_TIMEOUT_MS = 15_000
MAX_STEP_ATTEMPTS = 2


@dataclass
class Credentials:
    username: str
    password: str


def load_env(paths: tuple[Path, ...] = ENV_FILES) -> None:
    for path in paths:
        if not path.is_file():
            continue
        for line in path.read_text(encoding="utf-8-sig").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip().strip("\"'"))


def get_env(name: str) -> str:
    load_env()
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Falta la variable {name} en el entorno o en el archivo .env.")
    return value


def get_login_url() -> str:
    return get_env("LUCY_LOGIN_URL")


def get_company_code() -> str:
    return get_env("LUCY_COMPANY_CODE")


def get_bundle_dir() -> Path:
    """Carpeta donde se buscan los ZIP del bot (LUCY_BUNDLE_DIR, por defecto Descargas)."""
    load_env()
    value = os.environ.get("LUCY_BUNDLE_DIR", "").strip()
    return Path(value).expanduser() if value else DEFAULT_BUNDLE_DIR


def get_credentials() -> Credentials:
    return Credentials(
        username=get_env("LUCY_USERNAME"),
        password=get_env("LUCY_PASSWORD"),
    )
