"""Persistent, temporary history of Lucy invoices found during validation."""
import hashlib
import json
import shutil
from datetime import datetime, timezone
from pathlib import Path

from .bundle import InvoiceBundle, get_folio
from .config import HISTORY_DIR

HISTORY_PATH = HISTORY_DIR / "validation_history.json"
HISTORY_SCHEMA_VERSION = 1


def _load_history(path: Path) -> dict:
    if not path.exists():
        return {"schema_version": HISTORY_SCHEMA_VERSION, "facturas": {}}
    history = json.loads(path.read_text(encoding="utf-8"))
    if history.get("schema_version") != HISTORY_SCHEMA_VERSION:
        raise ValueError(f"Versión de historial no soportada: {history.get('schema_version')}.")
    if not isinstance(history.get("facturas"), dict):
        raise ValueError(f"Formato de historial inválido: {path}.")
    return history


def _write_history(path: Path, history: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = path.with_suffix(".json.tmp")
    try:
        temporary_path.write_text(
            json.dumps(history, ensure_ascii=False, indent=2, allow_nan=False),
            encoding="utf-8",
        )
        temporary_path.replace(path)
    finally:
        temporary_path.unlink(missing_ok=True)


def _invoice_key(invoice: dict) -> str:
    identity = f"{get_folio(invoice)}\0{str(invoice.get('vendor') or '').strip()}"
    return hashlib.sha256(identity.encode("utf-8")).hexdigest()[:24]


def _copy_allocated_costs(
    bundle: InvoiceBundle,
    invoice: dict,
    history_dir: Path,
    invoice_key: str,
) -> list[dict]:
    archived_files = []
    bundle_root = bundle.directory.resolve()
    invoice_dir = history_dir / invoice_key

    for part_index, part in enumerate(invoice.get("partes", []), start=1):
        relative_path = part.get("allocated_costs") or (part.get("csv") or {}).get("allocated_costs")
        if not relative_path:
            continue
        source = (bundle.directory / relative_path).resolve()
        if not source.is_relative_to(bundle_root):
            raise ValueError(f"Ruta CSV fuera del paquete de control: {relative_path}.")
        if not source.is_file():
            raise FileNotFoundError(f"No existe el CSV Allocated Costs requerido: {source}.")
        target = invoice_dir / f"parte_{part_index}_{source.name}"
        invoice_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
        archived_files.append({
            "parte": part.get("parte"),
            "tipo": "allocated_costs",
            "ruta": str(target.relative_to(history_dir)),
        })
    return archived_files


def record_found_invoice(
    bundle: InvoiceBundle,
    invoice: dict,
    status: str,
    error: str | None = None,
    *,
    new_attempt: bool = False,
    history_path: Path = HISTORY_PATH,
) -> None:
    """Save a found invoice and its CSVs until a later process consumes it."""
    history_dir = history_path.parent
    key = _invoice_key(invoice)
    archived_files = _copy_allocated_costs(bundle, invoice, history_dir, key)
    history = _load_history(history_path)
    now = datetime.now(timezone.utc).isoformat()
    record = history["facturas"].get(key)

    if record is None:
        record = {
            "folio": get_folio(invoice),
            "vendor": str(invoice.get("vendor") or "").strip() or None,
            "primera_deteccion": now,
            "intentos": [],
            "pendiente_siguiente_paso": True,
        }
        history["facturas"][key] = record
    if new_attempt:
        record["intentos"].append({"fecha": now, "estado": status})
    record.update({
        "ultima_actualizacion": now,
        "estado_validacion": status,
        "error": error,
        "ejecucion_origen": bundle.data.get("ejecucion"),
        "json_origen": bundle.json_path.name,
        "informacion_factura": invoice,
        "archivos_csv": archived_files,
        "pendiente_siguiente_paso": True,
    })
    _write_history(history_path, history)


def load_pending_validation_history(history_path: Path = HISTORY_PATH) -> list[dict]:
    """Return found invoices still waiting for the not-yet-implemented next step."""
    history = _load_history(history_path)
    return [
        record
        for record in history["facturas"].values()
        if record.get("pendiente_siguiente_paso")
    ]
