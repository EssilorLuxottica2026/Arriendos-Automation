"""Proceso: búsqueda de cada factura del JSON por Invoice Number."""
import re
import time
from datetime import datetime

from playwright.sync_api import Error as PlaywrightError, Locator, Page, expect

from ..bundle import get_folio, mark_invoice
from ..config import FILTER_TIMEOUT_MS
from ..history import (
    CLOSED_DUPLICATED,
    REASON_VALIDATED,
    REASON_VALIDATION_ERROR,
    close_waiting_invoice,
    upsert_waiting_invoice,
)
from ..logging import log
from ..notifications import arm_notification_observer, check_lucy_errors, failure_message
from ..plan import WorkItem
from .invoice_management import ensure_document_list, is_document_detail_open

INVOICE_VALIDATED = "Invoice review complete (not saved)"
INVOICE_DUPLICATED = "Invoice number duplicated"
INVOICE_ALREADY_VALIDATED = "No matching invoice found"

INVOICE_NUMBER_COLUMN = "Invoice Number"
BARCODE_COLUMN = "Barcode"
RESULTS_PATTERN = re.compile(r"(\d+) results")
ZERO_RESULTS_STABLE_S = 1.5
OPEN_ATTEMPTS = 3
OPEN_WAIT_MS = 10_000
DETAIL_LOAD_TIMEOUT_MS = 60_000
ROW_SELECT_WAIT_MS = 1_000
MOVEMENT_TYPE = "FI"
VENDOR_CODE_LENGTH = 10


def invoice_number_filter(page: Page) -> Locator:
    # El prefijo numérico del id es dinámico ("95-Invoice Number"); el sufijo es estable.
    return page.locator(f"[id$='-{INVOICE_NUMBER_COLUMN}']")


def filter_by_invoice_number(page: Page, folio: str) -> None:
    # Se reutiliza el mismo input: el folio nuevo reemplaza al anterior.
    field = invoice_number_filter(page)
    expect(field).to_be_visible()
    field.click()
    field.fill(folio)
    log(f"Filtro Invoice Number aplicado para folio {folio}.")


def get_results_count(page: Page) -> int:
    match = RESULTS_PATTERN.search(page.locator("#scroller").inner_text())
    return int(match.group(1)) if match else -1


def column_cells(page: Page, column: str) -> Locator:
    column_index = page.evaluate(
        """(name) => [...document.querySelectorAll('#scroller thead th')]
            .findIndex(th => th.innerText.trim().split('\\n').pop().trim() === name)""",
        column,
    )
    if column_index < 0:
        raise RuntimeError(f"No se encontró la columna {column} en la lista de documentos.")
    return page.locator(f"#scroller tbody tr td:nth-child({column_index + 1})")


def wait_for_results_count(page: Page, folio: str) -> int:
    """Espera a que "N results" corresponda al filtro aplicado y devuelve N."""
    # Justo después de escribir, el conteo del folio anterior sigue en pantalla; se espera a que
    # las filas correspondan al folio nuevo, o a que "0 results" se mantenga estable.
    deadline = time.monotonic() + FILTER_TIMEOUT_MS / 1000
    needle = folio.casefold()
    zero_since = None
    while time.monotonic() < deadline:
        total = get_results_count(page)
        if total == 0:
            zero_since = zero_since or time.monotonic()
            if time.monotonic() - zero_since >= ZERO_RESULTS_STABLE_S:
                log(f"El filtro de {folio} muestra 0 resultados.")
                return 0
        else:
            zero_since = None
            values = [text.strip() for text in column_cells(page, INVOICE_NUMBER_COLUMN).all_inner_texts()]
            if values and total >= len(values) and all(needle in v.casefold() for v in values):
                log(f"El filtro de {folio} muestra {total} resultado(s) antes de validar coincidencia exacta.")
                return total
        page.wait_for_timeout(300)
    raise TimeoutError(f"La lista no se actualizó con el filtro de Invoice Number '{folio}'.")


def pagination(page: Page) -> Locator:
    return page.locator(".table-pagination")


def page_input(page: Page) -> Locator:
    return pagination(page).locator("input[inputmode='numeric']")


def current_page_number(page: Page) -> int:
    value = page_input(page).input_value()
    return int(value) if value.isdigit() else 1


def total_pages(page: Page) -> int:
    match = re.search(r"of (\d+)", pagination(page).inner_text())
    return int(match.group(1)) if match else 1


def has_next_page(page: Page) -> bool:
    return current_page_number(page) < total_pages(page)


def next_page_button(page: Page) -> Locator:
    # Habilitados en orden: [anterior], siguiente; el de "siguiente" siempre es el último.
    return pagination(page).locator("#icon-chevron-enabled").last


def first_page_button(page: Page) -> Locator:
    return pagination(page).locator("#icon-double-chevron-enabled").first


def page_rows_signature(page: Page) -> list[str]:
    # Barcode identifica cada fila en Invoice Management; si la pantalla no lo tiene
    # (p. ej. Interactive Export), se usa el texto completo de las filas.
    try:
        return column_cells(page, BARCODE_COLUMN).all_inner_texts()
    except RuntimeError:
        return page.locator("#scroller tbody tr").all_inner_texts()


def wait_for_rows_change(page: Page, previous: list[str], description: str) -> None:
    deadline = time.monotonic() + FILTER_TIMEOUT_MS / 1000
    while page_rows_signature(page) == previous:
        if time.monotonic() > deadline:
            raise TimeoutError(f"La lista no cambió al pasar a {description}.")
        page.wait_for_timeout(200)


def go_to_next_page(page: Page) -> None:
    expected = str(current_page_number(page) + 1)
    previous = page_rows_signature(page)
    next_page_button(page).click()
    expect(page_input(page)).to_have_value(expected)
    wait_for_rows_change(page, previous, f"la página {expected}")


def go_to_page(page: Page, number: int) -> None:
    if current_page_number(page) > number:
        previous = page_rows_signature(page)
        first_page_button(page).click()
        expect(page_input(page)).to_have_value("1")
        wait_for_rows_change(page, previous, "la página 1")
    while current_page_number(page) < number:
        go_to_next_page(page)


def is_exact_folio(value: str, folio: str) -> bool:
    return value.strip().casefold() == folio.strip().casefold()


def find_exact_matches(page: Page, folio: str) -> list[tuple[int, int]]:
    """Recorre todas las páginas y devuelve (página, fila) de cada Invoice Number exacto."""
    # El filtro de Lucy es "contiene" (4092 también trae FE4092), por eso se compara cada fila.
    matches = []
    while True:
        page_number = current_page_number(page)
        values = column_cells(page, INVOICE_NUMBER_COLUMN).all_inner_texts()
        exact_on_page = [(page_number, row) for row, value in enumerate(values) if is_exact_folio(value, folio)]
        matches += exact_on_page
        log(
            f"[{folio}] Revisada página {page_number}: {len(values)} fila(s), "
            f"{len(exact_on_page)} coincidencia(s) exacta(s)."
        )
        if not has_next_page(page):
            return matches
        go_to_next_page(page)

def status_for_results(results: int) -> str:
    if results == 0:
        return INVOICE_ALREADY_VALIDATED
    if results > 1:
        return INVOICE_DUPLICATED
    return INVOICE_VALIDATED


def back_to_list_button(page: Page) -> Locator:
    return page.get_by_role("button", name="Back to List")


def wait_for_document_detail(page: Page, timeout: int) -> bool:
    try:
        page.wait_for_url("**/invoice-management/edit**", timeout=timeout)
        back_to_list_button(page).wait_for(timeout=DETAIL_LOAD_TIMEOUT_MS)
        return True
    except PlaywrightError:
        return False


def open_result(page: Page, folio: str, page_number: int, row: int) -> None:
    log(f"[{folio}] Abriendo coincidencia exacta en página {page_number}, fila {row + 1}.")
    go_to_page(page, page_number)
    folio_text = page.locator("#scroller tbody tr").nth(row).get_by_text(folio, exact=True).first
    for attempt in range(1, OPEN_ATTEMPTS + 1):
        log(f"[{folio}] Intento {attempt}/{OPEN_ATTEMPTS} de abrir el detalle.")
        expect(folio_text).to_be_visible()
        # El primer clic selecciona la fila y la re-renderiza; el doble clic se hace ya sobre
        # la fila estable para que Lucy abra el detalle.
        folio_text.click()
        page.wait_for_timeout(ROW_SELECT_WAIT_MS)
        folio_text.dblclick()
        if wait_for_document_detail(page, OPEN_WAIT_MS):
            log(f"[{folio}] Detalle abierto.")
            return
        if is_document_detail_open(page):
            break
    raise RuntimeError(f"No se pudo abrir el detalle de la factura {folio}.")


def react_select_option(page: Page, text: str) -> Locator:
    # El id de las opciones de react-select es dinámico ("react-select-171-option-0"); se elige por texto.
    return page.locator("[id^='react-select-'][id*='-option-']").filter(
        has_text=re.compile(rf"^{re.escape(text)}$")
    ).first


def select_movement_type(page: Page) -> None:
    log(f"Seleccionando tipo de movimiento {MOVEMENT_TYPE}.")
    page.locator("#mov-type-dropdown").locator("[class*='indicatorContainer']").click()
    react_select_option(page, MOVEMENT_TYPE).click()


def select_document_type(page: Page, document_type: str) -> None:
    log(f"Seleccionando tipo de documento {document_type}.")
    page.locator("#documents-document-type-dropdown").locator("[class*='indicatorContainer']").click()
    react_select_option(page, document_type).click()


def format_invoice_date(value: str) -> str:
    return datetime.fromisoformat(value).strftime("%d-%m-%Y")


def fill_document_data(page: Page, invoice: dict) -> None:
    """Comprueba el detalle contra el control; no guarda la factura."""
    part = invoice["partes"][0]
    vendor_code = str(part["vendor"]).strip().zfill(VENDOR_CODE_LENGTH)
    document_type = part["tipo_documento"]

    log(f"[{get_folio(invoice)}] Verificando vendor {vendor_code} y tipo {document_type}.")
    select_movement_type(page)
    expect(page.get_by_role("textbox", name=vendor_code)).to_have_value(vendor_code)
    log(f"[{get_folio(invoice)}] Vendor confirmado.")
    select_document_type(page, document_type)
    page.get_by_role("button", name="Continue").click()
    check_lucy_errors(page, get_folio(invoice), "pulsar Continue")
    log(f"[{get_folio(invoice)}] Continuó al detalle contable para comprobar los campos.")

    expect(page.locator("#documents-document-type-dropdown")).to_contain_text(document_type)
    invoice_date = page.locator(".react-datepicker__input-container > input").first
    expect(invoice_date).to_be_visible()
    expect(invoice_date).to_have_value(format_invoice_date(part["fecha_factura"]))
    log(f"[{get_folio(invoice)}] Fecha confirmada: {format_invoice_date(part['fecha_factura'])}.")
    document_number = page.locator("#q1-text-input-NUMDOC")
    expect(document_number).to_be_visible()
    expect(document_number).to_have_value(get_folio(invoice))
    log(f"[{get_folio(invoice)}] Número de documento confirmado.")

    select_movement_type(page)
    save_button = page.get_by_role("button", name="Save Document")
    expect(save_button).to_be_visible()
    log(
        f"[{get_folio(invoice)}] Botón Save Document visible; NO se pulsó. "
        "La factura queda abierta para inspección y sin guardar."
    )


def validate_invoice(page: Page, item: WorkItem) -> str:
    invoice = item.invoice
    folio = item.folio
    parts = invoice.get("partes", [])
    log(
        f"[{folio}] Origen: {item.origin}. El control contiene {len(parts)} parte(s); "
        "la comprobación del detalle de Lucy utiliza la parte 1."
    )
    log(f"[{folio}] Preparando listado de Validation y verificando el folio.")
    ensure_document_list(page)
    filter_by_invoice_number(page, folio)
    matches = find_exact_matches(page, folio) if wait_for_results_count(page, folio) else []
    status = status_for_results(len(matches))
    log(f"[{folio}] Coincidencias exactas encontradas: {len(matches)}.")
    if status == INVOICE_VALIDATED:
        # Se registra como error antes de abrir: si el proceso se corta, la factura queda en espera.
        upsert_waiting_invoice(item.bundle, invoice, REASON_VALIDATION_ERROR, "en_proceso", new_attempt=True)
        log(f"[{folio}] Encontrada en Validation; agregada a la lista de espera.")
        try:
            open_result(page, folio, *matches[0])
            arm_notification_observer(page)
            log(f"[{folio}] Observador de mensajes de Lucy activo.")
            fill_document_data(page, invoice)
            check_lucy_errors(page, folio, "revisar el detalle en Validation")
        except Exception as exc:
            error = failure_message(page, exc)
            upsert_waiting_invoice(item.bundle, invoice, REASON_VALIDATION_ERROR, "found_but_not_processed", error)
            log(f"[{folio}] No se pudo completar en Validation; queda en lista de espera: {error}.")
            raise
        upsert_waiting_invoice(item.bundle, invoice, REASON_VALIDATED, status)
        log(f"[{folio}] Validation completado; en lista de espera hasta reflejarse en Interactive Export.")
    elif status == INVOICE_DUPLICATED:
        log(f"[{folio}] Se detectaron duplicados; no se abrió ni modificó ningún registro.")
        if close_waiting_invoice(invoice, CLOSED_DUPLICATED):
            log(f"[{folio}] Retirada de la lista de espera para revisión manual.")
    else:
        log(f"[{folio}] No aparece en Validation; se buscará en Interactive Export.")
    if item.bundle is not None:
        mark_invoice(item.bundle, invoice, status, len(matches))
        log(f"[{folio}] Resultado escrito solo en el JSON extraído para esta ejecución: {status}.")
    return status
