"""Lista de espera persistente de facturas que deben volver a buscarse en Lucy.

Motivos de espera:
- ``validada``: se encontró y completó en Validation; espera reflejarse en Interactive Export.
- ``error_validacion``: se encontró en Validation pero no se pudo completar.
- ``no_reflejada``: no apareció ni en Validation ni en Interactive Export.
- ``error_interactive_export``: apareció en Interactive Export pero no se pudo completar.

Una factura sale de la lista al importar su CSV en Interactive Export o al detectarse duplicada.
"""
import hashlib
import json
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

from .bundle import InvoiceBundle, bundle_allocated_costs_csv, get_folio
from .config import HISTORY_DIR

HISTORY_PATH = HISTORY_DIR / "validation_history.json"
HISTORY_SCHEMA_VERSION = 1

REASON_VALIDATED = "validada"
REASON_VALIDATION_ERROR = "error_validacion"
REASON_NOT_REFLECTED = "no_reflejada"
REASON_EXPORT_ERROR = "error_interactive_export"
# Motivos que se vuelven a buscar en Validation; el resto solo se busca en Interactive Export.
VALIDATION_RETRY_REASONS = {REASON_VALIDATION_ERROR, REASON_NOT_REFLECTED}

CLOSED_COMPLETED = "completada_interactive_export"
CLOSED_DUPLICATED = "duplicada_revision_manual"

_LEGACY_DUPLICATED = "Invoice number duplicated"
_LEGACY_NOT_PROCESSED = "found_but_not_processed"


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


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def invoice_key(invoice: dict) -> str:
    identity = f"{get_folio(invoice)}\0{str(invoice.get('vendor') or '').strip()}"
    return hashlib.sha256(identity.encode("utf-8")).hexdigest()[:24]


def waiting_reason(record: dict) -> Optional[str]:
    """Motivo de espera; ``None`` si el registro no está pendiente."""
    if not record.get("pendiente_siguiente_paso"):
        return None
    if record.get("motivo_espera"):
        return record["motivo_espera"]
    # Registros creados antes de existir "motivo_espera".
    status = record.get("estado_validacion")
    if status == _LEGACY_DUPLICATED:
        return None
    if status == _LEGACY_NOT_PROCESSED:
        return REASON_VALIDATION_ERROR
    return REASON_VALIDATED


def _copy_allocated_costs(
    bundle: InvoiceBundle,
    invoice: dict,
    history_dir: Path,
    key: str,
) -> list[dict]:
    archived_files = []
    invoice_dir = history_dir / key

    for part_index, part in enumerate(invoice.get("partes", []), start=1):
        source = bundle_allocated_costs_csv(bundle, part)
        if source is None:
            continue
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


def upsert_waiting_invoice(
    bundle: Optional[InvoiceBundle],
    invoice: dict,
    reason: str,
    status: str,
    error: Optional[str] = None,
    *,
    new_attempt: bool = False,
    history_path: Path = HISTORY_PATH,
) -> None:
    """Agrega o actualiza una factura en la lista de espera.

    Con ``bundle`` se copian sus CSV Allocated Costs; sin él (factura que viene del propio
    historial) se conservan los CSV ya archivados.
    """
    history_dir = history_path.parent
    key = invoice_key(invoice)
    history = _load_history(history_path)
    now = _now()
    record = history["facturas"].get(key)

    if record is None:
        if bundle is None:
            raise ValueError(f"La factura {get_folio(invoice)} no está en el historial y no tiene paquete.")
        record = {
            "folio": get_folio(invoice),
            "vendor": str(invoice.get("vendor") or "").strip() or None,
            "primera_deteccion": now,
            "intentos": [],
        }
        history["facturas"][key] = record
    if new_attempt:
        record.setdefault("intentos", []).append({"fecha": now, "estado": status, "motivo": reason})
    record.update({
        "ultima_actualizacion": now,
        "motivo_espera": reason,
        "estado_validacion": status,
        "error": error,
        "informacion_factura": invoice,
        "pendiente_siguiente_paso": True,
        "estado_final": None,
    })
    if bundle is not None:
        record.update({
            "ejecucion_origen": bundle.data.get("ejecucion"),
            "json_origen": bundle.json_path.name,
            "archivos_csv": _copy_allocated_costs(bundle, invoice, history_dir, key),
        })
    _write_history(history_path, history)


def record_interactive_export_result(
    invoice: dict,
    status: str,
    matches: int,
    error: Optional[str] = None,
    *,
    history_path: Path = HISTORY_PATH,
) -> bool:
    """Registra un intento en Interactive Export; devuelve False si la factura no está en la lista."""
    history = _load_history(history_path)
    record = history["facturas"].get(invoice_key(invoice))
    if record is None or waiting_reason(record) is None:
        return False
    now = _now()
    record.setdefault("intentos_interactive_export", []).append({
        "fecha": now,
        "estado": status,
        "coincidencias": matches,
        "error": error,
    })
    record.update({"ultima_actualizacion": now, "estado_interactive_export": status})
    _write_history(history_path, history)
    return True


def close_waiting_invoice(
    invoice: dict,
    final_status: str,
    *,
    history_path: Path = HISTORY_PATH,
) -> bool:
    """Saca una factura de la lista de espera; devuelve False si no estaba pendiente."""
    history = _load_history(history_path)
    record = history["facturas"].get(invoice_key(invoice))
    if record is None or waiting_reason(record) is None:
        return False
    now = _now()
    record.update({
        "ultima_actualizacion": now,
        "pendiente_siguiente_paso": False,
        "estado_final": final_status,
        "fecha_cierre": now,
    })
    _write_history(history_path, history)
    return True


def archived_allocated_costs_csv(
    invoice: dict,
    part_number,
    *,
    history_path: Path = HISTORY_PATH,
) -> Optional[Path]:
    """Copia archivada del CSV Allocated Costs de una parte; ``None`` si no existe en el historial."""
    record = _load_history(history_path)["facturas"].get(invoice_key(invoice))
    for archived in (record or {}).get("archivos_csv") or []:
        if archived.get("tipo") == "allocated_costs" and archived.get("parte") == part_number:
            path = (history_path.parent / archived["ruta"]).resolve()
            if not path.is_relative_to(history_path.parent.resolve()):
                raise ValueError(f"Ruta CSV fuera del historial: {archived['ruta']}.")
            return path
    return None


def load_waiting_list(history_path: Path = HISTORY_PATH) -> list[dict]:
    """Registros pendientes, con ``motivo_espera`` resuelto (incluye registros antiguos)."""
    history = _load_history(history_path)
    waiting = []
    for record in history["facturas"].values():
        reason = waiting_reason(record)
        if reason and (record.get("informacion_factura") or {}).get("folio"):
            waiting.append({**record, "motivo_espera": reason})
    return waiting
