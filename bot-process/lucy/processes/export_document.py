"""Proceso: completar una factura abierta en INTERACTIVE EXPORT e importar su CSV Allocated Costs.

No pulsa Save Document: la factura queda abierta y sin guardar para revisión.
"""
import re
import time
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError, Locator, Page, expect

from ..bundle import bundle_allocated_costs_csv, get_folio
from ..history import archived_allocated_costs_csv
from ..logging import log
from ..notifications import (
    LucyNotificationError,
    arm_notification_observer,
    check_lucy_errors,
    consume_notifications,
    error_messages,
)
from ..plan import WorkItem
from .invoice_validation import (
    VENDOR_CODE_LENGTH,
    format_invoice_date,
    react_select_option,
    select_movement_type,
)

DOCUMENT_TYPE_SAP_LABEL = "Document Type Sap"
VAT_CODE_LABEL = "Vat Code"
TOTAL_VAT_LABEL = "Total VAT Amount"
PAYMENT_TERMS_LABEL = "Payment Terms"

DOCUMENT_TYPE_SAP = "KR - Registro de Factura no PO"
VAT_CODE_WITH_VAT = "VQ - IVA Compra Servicios 19% - Gasto"
VAT_CODE_WITHOUT_VAT = "VW - IVA Descontable Compras Bienes y Servicios 19%"
PAYMENT_TERMS = "_ - _"

# "Balance:0" exacto (también "Balance: 0,00"); evita aceptar "Balance:0.5" o "Balance:05".
ZERO_BALANCE = re.compile(r"Balance:\s*-?0(?:[.,]0+)?(?![\d.,]*\d)")
DETAIL_LOAD_TIMEOUT_MS = 60_000
CSV_IMPORT_TIMEOUT_MS = 60_000
EXPORT_PANEL_WAIT_MS = 5_000
EXPORT_PANEL_ATTEMPTS = 2


def allocated_costs_csv(item: WorkItem) -> Path:
    """CSV de la parte 1: del paquete de la corrida o, si viene de la lista de espera, la copia archivada."""
    parts = item.invoice.get("partes") or []
    if not parts:
        raise ValueError(f"La factura {item.folio} no tiene partes en el control.")
    if len(parts) > 1:
        log(f"[{item.folio}] AVISO: el control tiene {len(parts)} partes; se importa solo el CSV de la parte 1.")
    part = parts[0]
    if item.bundle is not None:
        path = bundle_allocated_costs_csv(item.bundle, part)
        source = "paquete de la corrida"
    else:
        path = archived_allocated_costs_csv(item.invoice, part.get("parte"))
        source = "copia archivada en la lista de espera"
    if path is None or not path.is_file():
        raise FileNotFoundError(f"No se encontró el CSV Allocated Costs de {item.folio} ({source}): {path}.")
    log(f"[{item.folio}] CSV Allocated Costs a importar ({source}): {path.name}.")
    return path


def has_vat(total_vat: str) -> bool:
    """True si el Total VAT Amount mostrado por Lucy es mayor que 0."""
    digits = re.sub(r"\D", "", total_vat)
    if not digits:
        raise ValueError(f"Total VAT Amount sin valor numérico: '{total_vat}'.")
    if total_vat.strip().startswith("-"):
        raise ValueError(f"Total VAT Amount negativo: '{total_vat}'.")
    return int(digits) > 0


def vat_code_for(total_vat: str) -> str:
    return VAT_CODE_WITH_VAT if has_vat(total_vat) else VAT_CODE_WITHOUT_VAT


def labeled_field(page: Page, label: str) -> Locator:
    """Campo de Lucy junto a su etiqueta visible.

    Los ids (``react-select-NNNN``) y la posición (``div:nth-child(15)``) cambian; se usa la
    etiqueta y el contenedor más cercano que tenga un input de Lucy (clase ``InputQ1Wrapper``).
    """
    label_text = page.get_by_role("main").get_by_text(
        re.compile(rf"^\s*{re.escape(label)}\s*\*?\s*$")
    ).first
    expect(label_text).to_be_visible()
    container = label_text.locator("xpath=ancestor::div[.//*[contains(@class, 'InputQ1Wrapper')]][1]")
    wrappers = container.locator("[class*='InputQ1Wrapper']")
    count = wrappers.count()
    if count != 1:
        raise RuntimeError(f"Se esperaba un campo para la etiqueta '{label}' y se encontraron {count}.")
    return wrappers.first


def selected_value(field: Locator) -> Locator:
    return field.locator("[class*='singleValue']")


def select_labeled_option(page: Page, folio: str, label: str, option_text: str) -> None:
    log(f"[{folio}] Seleccionando {label}: {option_text}.")
    field = labeled_field(page, label)
    combo = field.locator("input[id^='react-select-']")
    expect(combo).to_be_visible()
    combo.click()
    combo.fill(option_text)
    option = react_select_option(page, option_text)
    expect(option).to_be_visible()
    option.click()
    expect(selected_value(field)).to_have_text(option_text)
    check_lucy_errors(page, folio, f"seleccionar {label}")
    log(f"[{folio}] {label} confirmado: {option_text}.")


def read_total_vat(page: Page, folio: str) -> str:
    total_vat = labeled_field(page, TOTAL_VAT_LABEL).locator("input").first
    expect(total_vat).to_be_visible()
    value = total_vat.input_value()
    log(f"[{folio}] {TOTAL_VAT_LABEL} en Lucy: '{value}'.")
    return value


def verify_header(page: Page, invoice: dict) -> None:
    folio = get_folio(invoice)
    part = invoice["partes"][0]
    vendor_code = str(part["vendor"]).strip().zfill(VENDOR_CODE_LENGTH)
    document_type = part["tipo_documento"]
    invoice_date = format_invoice_date(part["fecha_factura"])

    expect(page.get_by_role("textbox", name=vendor_code)).to_have_value(vendor_code)
    log(f"[{folio}] Vendor confirmado: {vendor_code}.")
    expect(page.locator("#documents-document-type-dropdown")).to_contain_text(document_type)
    log(f"[{folio}] Tipo de documento confirmado: {document_type}.")
    date_input = page.locator(".react-datepicker__input-container > input").first
    expect(date_input).to_have_value(invoice_date)
    log(f"[{folio}] Fecha confirmada: {invoice_date}.")
    expect(page.locator("#q1-text-input-NUMDOC")).to_have_value(folio)
    log(f"[{folio}] Número de documento confirmado: {folio}.")


def verify_selections(page: Page, folio: str, vat_code: str) -> None:
    for label, expected in (
        (DOCUMENT_TYPE_SAP_LABEL, DOCUMENT_TYPE_SAP),
        (VAT_CODE_LABEL, vat_code),
        (PAYMENT_TERMS_LABEL, PAYMENT_TERMS),
    ):
        expect(selected_value(labeled_field(page, label))).to_have_text(expected)
        log(f"[{folio}] {label} sigue con {expected}.")


def open_csv_import(page: Page, folio: str) -> None:
    export_button = page.get_by_role("button", name="Export", exact=True)
    import_button = page.get_by_role("button", name="Import csv template")
    expect(export_button).to_be_visible()
    for attempt in range(1, EXPORT_PANEL_ATTEMPTS + 1):
        log(f"[{folio}] Pulsando Export (intento {attempt}/{EXPORT_PANEL_ATTEMPTS}).")
        export_button.click()
        try:
            import_button.wait_for(state="visible", timeout=EXPORT_PANEL_WAIT_MS)
            break
        except PlaywrightError:
            log(f"[{folio}] 'Import csv template' aún no aparece.")
    else:
        check_lucy_errors(page, folio, "pulsar Export")
        raise RuntimeError("No apareció el botón 'Import csv template' tras pulsar Export.")
    check_lucy_errors(page, folio, "pulsar Export")
    import_button.click()
    check_lucy_errors(page, folio, "abrir Import csv template")
    log(f"[{folio}] Import csv template abierto.")


def wait_for_zero_balance(page: Page, folio: str) -> None:
    """Espera Balance:0; si Lucy notifica un error antes, falla con ese mensaje."""
    scroller = page.locator("#scroller")
    deadline = time.monotonic() + CSV_IMPORT_TIMEOUT_MS / 1000
    while time.monotonic() < deadline:
        notifications = consume_notifications(page)
        for item in notifications:
            log(f"[{folio}] Lucy notificó ({item['type'] or 'sin tipo'}) al importar el CSV: {item['text']}")
        errors = error_messages(notifications)
        if errors:
            raise LucyNotificationError(f"Lucy reportó error al importar el CSV: {' | '.join(errors)}")
        try:
            if ZERO_BALANCE.search(scroller.inner_text(timeout=1_000)):
                return
        except PlaywrightError:
            pass
        page.wait_for_timeout(500)
    current = re.search(r"Balance:\s*[^\s]+", scroller.inner_text()) if scroller.is_visible() else None
    raise TimeoutError(
        f"No se llegó a Balance:0 tras importar el CSV (actual: {current.group(0) if current else 'no visible'})."
    )


def import_allocated_costs(page: Page, folio: str, csv_path: Path) -> None:
    open_csv_import(page, folio)
    chooser = page.get_by_role("button", name="Choose File")
    expect(chooser).to_be_visible()
    chooser.set_input_files(str(csv_path))
    log(f"[{folio}] CSV {csv_path.name} cargado; esperando Balance:0.")
    wait_for_zero_balance(page, folio)
    log(f"[{folio}] Balance:0 confirmado tras importar el CSV.")


def complete_export_document(page: Page, item: WorkItem) -> None:
    folio = item.folio
    # Se resuelve antes de tocar la página: sin CSV no se modifica nada en Lucy.
    csv_path = allocated_costs_csv(item)

    log(f"[{folio}] Esperando el detalle de la factura en INTERACTIVE EXPORT.")
    expect(page.locator("#mov-type-dropdown")).to_be_visible(timeout=DETAIL_LOAD_TIMEOUT_MS)
    arm_notification_observer(page)
    log(f"[{folio}] Observador de mensajes de Lucy activo.")
    select_movement_type(page)
    check_lucy_errors(page, folio, "seleccionar FI")
    log(f"[{folio}] Verificando datos de cabecera contra el control.")
    verify_header(page, item.invoice)

    select_labeled_option(page, folio, DOCUMENT_TYPE_SAP_LABEL, DOCUMENT_TYPE_SAP)
    vat_code = vat_code_for(read_total_vat(page, folio))
    log(f"[{folio}] Vat Code según Total VAT Amount: {vat_code}.")
    select_labeled_option(page, folio, VAT_CODE_LABEL, vat_code)
    select_labeled_option(page, folio, PAYMENT_TERMS_LABEL, PAYMENT_TERMS)

    log(f"[{folio}] Revisando nuevamente todos los campos.")
    verify_header(page, item.invoice)
    verify_selections(page, folio, vat_code)

    import_allocated_costs(page, folio, csv_path)
    check_lucy_errors(page, folio, "importar el CSV")
    log(f"[{folio}] Factura completada en INTERACTIVE EXPORT; Save Document NO se pulsó (sin guardar).")
