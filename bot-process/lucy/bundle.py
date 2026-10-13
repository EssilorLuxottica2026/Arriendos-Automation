"""Bundle del bot: ZIP generado por la app con el JSON de control y los CSV.

El ZIP se busca en la carpeta de descargas y se extrae a la caché local del usuario, bajo
``%LOCALAPPDATA%/Arriendos/bot-process/cache/<nombre del zip>/``. Cada ejecución borra la caché
anterior y extrae el ZIP desde cero (puede haber varias corridas al día). Ahí se guarda el
resultado de cada factura y se conserva hasta la siguiente ejecución.
"""
import json
import shutil
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path, PureWindowsPath
from typing import Optional

from .config import BUNDLE_JSON_NAME, BUNDLE_PATTERN, CACHE_DIR

LUCY_RESULT_KEY = "validacion_lucy"
INTERACTIVE_EXPORT_RESULT_KEY = "interactive_export_lucy"


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


def clear_cache(cache_dir: Path = CACHE_DIR) -> None:
    """Borra las corridas anteriores: cada ejecución del bot empieza con una caché limpia."""
    if cache_dir.exists():
        shutil.rmtree(cache_dir)


def extract_bundle(zip_path: Path, cache_dir: Path = CACHE_DIR) -> Path:
    clear_cache(cache_dir)
    target = cache_dir / zip_path.stem
    target.mkdir(parents=True)
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


def allocated_costs_relative_path(part: dict) -> Optional[str]:
    """Ruta relativa del CSV Allocated Costs de una parte.

    En el JSON de control la ruta está en ``csv.allocated_costs``; ``allocated_costs`` a nivel de
    parte es la lista de filas del CSV, no una ruta.
    """
    path = (part.get("csv") or {}).get("allocated_costs")
    if isinstance(path, str) and path.strip():
        # El JSON se genera en Windows con "\"; se normaliza para cualquier sistema.
        return PureWindowsPath(path.strip()).as_posix()
    return None


def bundle_allocated_costs_csv(bundle: InvoiceBundle, part: dict) -> Optional[Path]:
    relative_path = allocated_costs_relative_path(part)
    if relative_path is None:
        return None
    source = (bundle.directory / relative_path).resolve()
    if not source.is_relative_to(bundle.directory.resolve()):
        raise ValueError(f"Ruta CSV fuera del paquete de control: {relative_path}.")
    return source


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


def mark_interactive_export(
    bundle: InvoiceBundle,
    invoice: dict,
    status: str,
    matches: int,
    allocated_cost_summary: Optional[dict] = None,
) -> None:
    result = {
        "estado": status,
        "coincidencias": matches,
        "fecha": datetime.now(timezone.utc).isoformat(),
    }
    if allocated_cost_summary:
        result["resumen_allocated_cost"] = allocated_cost_summary
    invoice[INTERACTIVE_EXPORT_RESULT_KEY] = result
    save_bundle(bundle)
