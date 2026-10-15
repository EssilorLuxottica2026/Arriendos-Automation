"""Proceso: completar cada copia (parte/vendor) de una factura multi-vendor en INTERACTIVE EXPORT.

Reutiliza sin modificarlo el flujo single vendor (``complete_export_document``): importa el CSV
Allocated Costs de la parte (A3) y revisa importes, cuentas, CeCo y resumen de conceptos (A4).
Ese flujo trabaja con ``partes[0]``, por eso se le entrega una vista de la factura con una sola parte.

Por cada parte: abrir una copia -> confirmar o cambiar el vendor -> A3 -> A4 -> guardar el resultado
en la parte del JSON. Al terminar todas, se confirma que estén contabilizadas antes de la compensación.

El envío a SAP (A5) y simular/guardar/confirmar el documento SAP aún no están implementados: con
``REPOST_DRY_RUN`` activo no se pulsa nada (solo se verifica Save Document) y con el modo real el
flujo se bloquea antes de hacer cualquier cambio en Lucy.
"""
from datetime import datetime, timezone
from typing import Optional

from playwright.sync_api import Page, expect

from ..bundle import InvoiceBundle, get_folio, save_bundle
from ..config import REPOST_DRY_RUN
from ..logging import log
from ..notifications import check_lucy_errors, failure_message
from ..plan import WorkItem
from .allocated_costs_review import check_before_import
from .export_document import DETAIL_LOAD_TIMEOUT_MS, allocated_costs_csv, complete_export_document, labeled_field
from .invoice_management import back_to_list
from .invoice_repost import RepostError, get_vendor_count, list_barcode_matches, open_document
from .invoice_validation import VENDOR_CODE_LENGTH

COPY_RESULT_KEY = "copia_lucy"
PART_REVIEWED = "Copy reviewed (CSV imported, not posted)"
PART_POSTED = "Copy posted in SAP"
PART_ERROR = "Copy error"
PART_DONE_STATES = {PART_REVIEWED, PART_POSTED}
STEP_CSV_REVIEWED = "csv_importado_y_revisado"
STEP_POSTED = "contabilizada_sap"

VENDOR_CODE_LABEL = "Vendor Code"


class PartsError(RepostError):
    """El flujo debe detenerse y reportarse para esta factura."""


def part_vendor_code(part: dict) -> str:
    return str(part["vendor"]).strip().zfill(VENDOR_CODE_LENGTH)


def single_part_view(invoice: dict, part: dict) -> dict:
    """Factura con una sola parte: el flujo single vendor y su revisión de importes leen solo esa parte."""
    return {
        **invoice,
        "partes": [part],
        "vendor": part.get("vendor"),
        "vendors": [part.get("vendor")],
        "importe": part.get("importe"),
    }


def part_item(bundle: InvoiceBundle, invoice: dict, part: dict) -> WorkItem:
    return WorkItem(single_part_view(invoice, part), bundle, None)


def part_state(part: dict) -> str:
    return (part.get(COPY_RESULT_KEY) or {}).get("estado") or ""


def is_part_done(part: dict) -> bool:
    return part_state(part) in PART_DONE_STATES


def pick_copy_index(found: int, total: int, done: int, dry_run: bool = REPOST_DRY_RUN) -> int:
    """Qué fila de las copias con el mismo barcode abrir para la siguiente parte.

    Si Lucy conserva las copias ya procesadas hay ``total`` filas y la siguiente es la ``done``-ésima;
    si las retira de la lista quedan ``total - done`` y la siguiente es la primera. En dry-run no se
    hizo ningún repost, así que se usa la copia disponible (la original) para probar el flujo.
    """
    if found == total:
        return done
    if found == total - done:
        return 0
    if dry_run and found > 0:
        return min(done, found - 1)
    raise PartsError(f"Se encontraron {found} copias y se esperaban {total - done} (o {total}) sin procesar.")


def preflight_parts(bundle: InvoiceBundle, invoice: dict) -> None:
    """Revisa partes y CSV contra el control antes de tocar Lucy (sobre todo antes del repost)."""
    folio = get_folio(invoice)
    parts = invoice.get("partes") or []
    if len(parts) != get_vendor_count(invoice):
        raise PartsError(f"La factura {folio} tiene {len(parts)} partes y {get_vendor_count(invoice)} vendors.")
    if not REPOST_DRY_RUN:
        raise PartsError(
            "El envío a SAP (A5) y simular/guardar/confirmar el documento SAP aún no están implementados; "
            "el repost real queda bloqueado para no dejar facturas a medias."
        )
    for part in parts:
        csv_path = allocated_costs_csv(part_item(bundle, invoice, part))
        check_before_import(f"{folio} parte {part.get('parte')}", part, csv_path)


def ensure_vendor(page: Page, folio: str, part: dict) -> None:
    """Confirma el vendor de la copia o lo cambia por el de la parte."""
    vendor_code = part_vendor_code(part)
    expect(page.locator("#mov-type-dropdown")).to_be_visible(timeout=DETAIL_LOAD_TIMEOUT_MS)
    field = labeled_field(page, VENDOR_CODE_LABEL).locator("input").first
    expect(field).to_be_visible()
    current = field.input_value().strip()
    if current == vendor_code:
        log(f"[{folio}] Vendor de la copia confirmado: {vendor_code}.")
        return
    log(f"[{folio}] La copia tiene el vendor '{current}'; cambiando a {vendor_code}.")
    field.click()
    field.fill(vendor_code)
    field.press("Tab")
    expect(field).to_have_value(vendor_code)
    check_lucy_errors(page, folio, "cambiar el vendor")
    log(f"[{folio}] Vendor cambiado a {vendor_code}.")


def post_part(page: Page, folio: str, part: dict) -> bool:
    """A5, simular, guardar y confirmar en SAP. Devuelve True si la parte quedó contabilizada."""
    # TODO: implementar cuando se disponga de la grabación de estos pasos.
    save_button = page.get_by_role("button", name="Save Document")
    expect(save_button).to_be_visible()
    log(f"[{folio}] DRY-RUN: Save Document visible; no se pulsa. Envío a SAP (A5) y confirmación no ejecutados.")
    return False


def mark_part(
    bundle: InvoiceBundle,
    part: dict,
    status: str,
    summary: Optional[dict] = None,
    error: Optional[str] = None,
) -> None:
    now = datetime.now(timezone.utc).isoformat()
    result = {"estado": status, "fecha": now}
    if summary:
        result["resumen_allocated_cost"] = summary
    if error:
        result["error"] = error
    part[COPY_RESULT_KEY] = result
    steps = part.setdefault("pasos_confirmados", [])
    if status in PART_DONE_STATES and STEP_CSV_REVIEWED not in steps:
        steps.append(STEP_CSV_REVIEWED)
    if status == PART_POSTED and STEP_POSTED not in steps:
        steps.append(STEP_POSTED)
    part["fecha_actualizacion"] = now
    save_bundle(bundle)


def process_part(page: Page, bundle: InvoiceBundle, invoice: dict, part: dict, barcode: str, done: int) -> None:
    folio = get_folio(invoice)
    label = f"[{folio} parte {part.get('parte')}]"
    total = len(invoice["partes"])
    matches = list_barcode_matches(page, barcode)
    index = pick_copy_index(len(matches), total, done)
    log(f"{label} Abriendo la copia {index + 1} de {len(matches)} con barcode {barcode}.")
    open_document(page, barcode, matches[index])
    try:
        ensure_vendor(page, folio, part)
        summary = complete_export_document(page, part_item(bundle, invoice, part))
        posted = post_part(page, folio, part)
    except Exception as exc:
        mark_part(bundle, part, PART_ERROR, error=failure_message(page, exc))
        raise
    mark_part(bundle, part, PART_POSTED if posted else PART_REVIEWED, summary)
    log(f"{label} Resultado guardado en el JSON.")
    back_to_list(page)


def confirm_all_posted(invoice: dict) -> None:
    """Todas las partes deben estar contabilizadas antes de comenzar la compensación."""
    folio = get_folio(invoice)
    pending = [str(part.get("parte")) for part in invoice["partes"] if part_state(part) != PART_POSTED]
    if not pending:
        log(f"[{folio}] Todas las partes están contabilizadas; se puede iniciar la compensación.")
    elif REPOST_DRY_RUN:
        log(f"[{folio}] DRY-RUN: partes sin contabilizar ({', '.join(pending)}); no se inicia la compensación.")
    else:
        raise PartsError(f"Partes sin contabilizar ({', '.join(pending)}); no se puede iniciar la compensación.")


def process_parts(page: Page, bundle: InvoiceBundle, invoice: dict, barcode: str) -> None:
    folio = get_folio(invoice)
    for part in invoice["partes"]:
        if is_part_done(part):
            log(f"[{folio} parte {part.get('parte')}] Ya procesada ({part_state(part)}); se omite.")
            continue
        done = sum(is_part_done(other) for other in invoice["partes"])
        process_part(page, bundle, invoice, part, barcode, done)
    confirm_all_posted(invoice)
