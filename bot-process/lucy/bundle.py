"""Bundle del bot: ZIP generado por la app con el JSON de control y los CSV.

El ZIP se busca en la carpeta de descargas y se extrae a ``bot-process/cache/<nombre del zip>/``.
La copia en caché es la copia de trabajo: ahí se guarda el resultado de cada factura, de modo
que si el bot se interrumpe, la siguiente ejecución continúa con las facturas pendientes.
"""
import json
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

from .config import BUNDLE_JSON_NAME, BUNDLE_PATTERN, CACHE_DIR

LUCY_RESULT_KEY = "validacion_lucy"


@dataclass
class InvoiceBundle:
    directory: Path
    json_path: Path
    data: dict


def find_latest_bundle(source_dir: Path) -> Path:
    bundles = sorted(source_dir.glob(BUNDLE_PATTERN), key=lambda path: path.stat().st_mtime)
    if not bundles:
        raise FileNotFoundError(f"No se encontró ningún {BUNDLE_PATTERN} en {source_dir}.")
    return bundles[-1]


def extract_bundle(zip_path: Path, cache_dir: Path = CACHE_DIR) -> Path:
    """Extrae el ZIP a la caché; si ya estaba extraído se reutiliza para conservar el avance."""
    target = cache_dir / zip_path.stem
    if (target / BUNDLE_JSON_NAME).is_file():
        return target
    target.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path) as bundle:
        for member in bundle.namelist():
            destination = (target / member).resolve()
            if not destination.is_relative_to(target.resolve()):
                raise ValueError(f"Ruta no permitida dentro del ZIP: {member}")
        bundle.extractall(target)
    if not (target / BUNDLE_JSON_NAME).is_file():
        raise FileNotFoundError(f"El ZIP {zip_path.name} no contiene {BUNDLE_JSON_NAME}.")
    return target


def read_bundle(directory: Path) -> InvoiceBundle:
    json_path = directory / BUNDLE_JSON_NAME
    data = json.loads(json_path.read_text(encoding="utf-8"))
    return InvoiceBundle(directory=directory, json_path=json_path, data=data)


def load_bundle(source_dir: Path, zip_path: Optional[Path] = None) -> InvoiceBundle:
    zip_path = zip_path or find_latest_bundle(source_dir)
    return read_bundle(extract_bundle(zip_path))


def save_bundle(bundle: InvoiceBundle) -> None:
    serialized = json.dumps(bundle.data, ensure_ascii=False, indent=2, allow_nan=False)
    temporary_path = bundle.json_path.with_suffix(".json.tmp")
    try:
        temporary_path.write_text(serialized, encoding="utf-8")
        temporary_path.replace(bundle.json_path)
    finally:
        temporary_path.unlink(missing_ok=True)


def get_invoices(bundle: InvoiceBundle) -> list[dict]:
    return bundle.data.get("facturas", [])


def get_folio(invoice: dict) -> str:
    return str(invoice["folio"]).strip()


def is_ready_for_bot(invoice: dict) -> bool:
    return bool(invoice.get("lista_para_bot"))


def has_lucy_result(invoice: dict) -> bool:
    return bool((invoice.get(LUCY_RESULT_KEY) or {}).get("estado"))


def get_pending_invoices(bundle: InvoiceBundle) -> list[dict]:
    return [
        invoice
        for invoice in get_invoices(bundle)
        if is_ready_for_bot(invoice) and not has_lucy_result(invoice)
    ]


def mark_invoice(bundle: InvoiceBundle, invoice: dict, status: str, matches: int) -> None:
    invoice[LUCY_RESULT_KEY] = {
        "estado": status,
        "coincidencias": matches,
        "fecha": datetime.now(timezone.utc).isoformat(),
    }
    save_bundle(bundle)
