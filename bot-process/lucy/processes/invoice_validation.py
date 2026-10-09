"""Proceso: búsqueda de cada factura del JSON por Invoice Number."""
import re
import time
from datetime import datetime

from playwright.sync_api import Error as PlaywrightError, Locator, Page, expect

from ..bundle import InvoiceBundle, get_folio, mark_invoice
from ..config import FILTER_TIMEOUT_MS
from .invoice_management import back_to_list, ensure_document_list, is_document_detail_open

INVOICE_VALIDATED = "Invoice validated"
INVOICE_DUPLICATED = "Invoice number duplicated"
INVOICE_ALREADY_VALIDATED = "Invoice already validated"

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
                return 0
        else:
            zero_since = None
            values = [text.strip() for text in column_cells(page, INVOICE_NUMBER_COLUMN).all_inner_texts()]
            if values and total >= len(values) and all(needle in v.casefold() for v in values):
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


def go_to_next_page(page: Page) -> None:
    expected = str(current_page_number(page) + 1)
    barcodes = column_cells(page, BARCODE_COLUMN)
    previous = barcodes.all_inner_texts()
    next_page_button(page).click()
    expect(page_input(page)).to_have_value(expected)
    deadline = time.monotonic() + FILTER_TIMEOUT_MS / 1000
    while barcodes.all_inner_texts() == previous:
        if time.monotonic() > deadline:
            raise TimeoutError(f"La lista no cambió al pasar a la página {expected}.")
        page.wait_for_timeout(200)


def go_to_page(page: Page, number: int) -> None:
    if current_page_number(page) > number:
        barcodes = column_cells(page, BARCODE_COLUMN)
        previous = barcodes.all_inner_texts()
        first_page_button(page).click()
        expect(page_input(page)).to_have_value("1")
        expect(barcodes.first).not_to_have_text(previous[0])
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
        matches += [(page_number, row) for row, value in enumerate(values) if is_exact_folio(value, folio)]
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
    go_to_page(page, page_number)
    folio_text = page.locator("#scroller tbody tr").nth(row).get_by_text(folio, exact=True).first
    for _ in range(OPEN_ATTEMPTS):
        expect(folio_text).to_be_visible()
        # El primer clic selecciona la fila y la re-renderiza; el doble clic se hace ya sobre
        # la fila estable para que Lucy abra el detalle.
        folio_text.click()
        page.wait_for_timeout(ROW_SELECT_WAIT_MS)
        folio_text.dblclick()
        if wait_for_document_detail(page, OPEN_WAIT_MS):
            return
        if is_document_detail_open(page):
            break
    raise RuntimeError(f"No se pudo abrir el detalle de la factura {folio}.")


def select_movement_type(page: Page) -> None:
    page.locator("#mov-type-dropdown").locator("[class*='indicatorContainer']").click()
    page.get_by_text(MOVEMENT_TYPE, exact=True).click()


def select_document_type(page: Page, document_type: str) -> None:
    dropdown = page.locator("#documents-document-type-dropdown")
    dropdown.locator("[class*='indicatorContainer']").click()
    # El id de las opciones de react-select es dinámico; se elige por texto.
    page.locator("[id^='react-select-'][id*='-option-']").filter(has_text=document_type).first.click()


def format_invoice_date(value: str) -> str:
    return datetime.fromisoformat(value).strftime("%d-%m-%Y")


def fill_document_data(page: Page, invoice: dict) -> None:
    """Pasos grabados en la validación del documento: valida datos contra el JSON y guarda."""
    part = invoice["partes"][0]
    vendor_code = str(part["vendor"]).strip().zfill(VENDOR_CODE_LENGTH)
    document_type = part["tipo_documento"]

    select_movement_type(page)
    expect(page.get_by_role("textbox", name=vendor_code)).to_have_value(vendor_code)
    select_document_type(page, document_type)
    page.get_by_role("button", name="Continue").click()

    expect(page.locator("#documents-document-type-dropdown")).to_contain_text(document_type)
    invoice_date = page.locator(".react-datepicker__input-container > input").first
    expect(invoice_date).to_be_visible()
    expect(invoice_date).to_have_value(format_invoice_date(part["fecha_factura"]))
    document_number = page.locator("#q1-text-input-NUMDOC")
    expect(document_number).to_be_visible()
    expect(document_number).to_have_value(get_folio(invoice))

    select_movement_type(page)
    page.get_by_role("button", name="Save Document").click()


def validate_invoice(page: Page, bundle: InvoiceBundle, invoice: dict) -> str:
    folio = get_folio(invoice)
    ensure_document_list(page)
    filter_by_invoice_number(page, folio)
    matches = find_exact_matches(page, folio) if wait_for_results_count(page, folio) else []
    status = status_for_results(len(matches))
    if status == INVOICE_VALIDATED:
        open_result(page, folio, *matches[0])
        fill_document_data(page, invoice)
        ensure_document_list(page)
    mark_invoice(bundle, invoice, status, len(matches))
    return status
