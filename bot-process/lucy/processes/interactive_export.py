"""Proceso: búsqueda de cada factura en INTERACTIVE EXPORT por Invoice Number."""
import re
import time
from dataclasses import dataclass
from typing import Optional

from playwright.sync_api import Error as PlaywrightError, Locator, Page, expect

from ..bundle import mark_interactive_export
from ..config import get_company_code
from ..history import (
    CLOSED_COMPLETED,
    CLOSED_DUPLICATED,
    REASON_EXPORT_ERROR,
    REASON_NOT_REFLECTED,
    close_waiting_invoice,
    record_interactive_export_result,
    upsert_waiting_invoice,
)
from ..logging import log
from ..plan import WorkItem
from .invoice_management import company_code_pattern, search_documents, wait_for_document_list
from .invoice_validation import (
    filter_by_invoice_number,
    find_exact_matches,
    go_to_page,
    invoice_number_filter,
    wait_for_results_count,
)

EXPORT_FOUND = "Interactive Export: invoice opened"
EXPORT_NOT_FOUND = "Interactive Export: invoice not found"
EXPORT_DUPLICATED = "Interactive Export: invoice number duplicated"
EXPORT_NOT_OPENED = "Interactive Export: found but not opened"

MENU_TEXT = "INTERACTIVE EXPORT"
# Posición del control "Select..." de compañía según la grabación de Lucy; solo se usa si
# Interactive Export no expone el contenedor estable "#company".
COMPANY_SELECT_FALLBACK_INDEX = 3
OPEN_ATTEMPTS = 3
OPEN_WAIT_MS = 10_000
ROW_SELECT_WAIT_MS = 1_000


@dataclass
class ExportListState:
    """URL del listado de Interactive Export ya filtrado por compañía en esta sesión."""
    list_url: Optional[str] = None


def status_for_export_matches(matches: int) -> str:
    if matches == 0:
        return EXPORT_NOT_FOUND
    if matches > 1:
        return EXPORT_DUPLICATED
    return EXPORT_FOUND


def open_interactive_export(page: Page) -> None:
    log("Abriendo INTERACTIVE EXPORT.")
    link = page.get_by_text(MENU_TEXT, exact=True).first
    expect(link).to_be_visible()
    link.click()


def company_option(page: Page, company_code: str) -> Locator:
    # Las opciones de react-select tienen id dinámico; se elige la que empieza con "8140-".
    return page.locator("[id^='react-select-'][id*='-option-']").filter(
        has_text=company_code_pattern(company_code)
    ).first


def selected_company(page: Page, company_code: str) -> Locator:
    # Valor ya elegido en react-select (clase generada "...-singleValue").
    return page.locator("[class*='singleValue']").filter(has_text=company_code_pattern(company_code)).first


def stable_company_input(page: Page) -> Locator:
    return page.locator("#company input[id^='react-select-']")


def wait_visible(locator: Locator, timeout_ms: int) -> bool:
    try:
        locator.wait_for(state="visible", timeout=timeout_ms)
        return True
    except PlaywrightError:
        return False


def show_export_search_form(page: Page) -> None:
    # Si Interactive Export conserva una lista previa, el formulario se abre con "New Search".
    if wait_visible(stable_company_input(page), 5_000):
        return
    new_search = page.get_by_role("button", name="New Search")
    if new_search.is_visible():
        log("INTERACTIVE EXPORT muestra una lista previa; solicitando New Search.")
        new_search.click()
        wait_visible(stable_company_input(page), 5_000)


def company_input(page: Page) -> Locator:
    stable = stable_company_input(page)
    if stable.is_visible():
        log("Campo de compañía localizado por su contenedor estable.")
        return stable
    log("Contenedor de compañía no disponible; usando el selector 'Select...' de la grabación.")
    placeholder = page.locator("div").filter(has_text=re.compile(r"^Select\.\.\.$")).nth(
        COMPANY_SELECT_FALLBACK_INDEX
    )
    expect(placeholder).to_be_visible()
    placeholder.click()
    return page.locator("input[id^='react-select-'][id$='-input']:focus")


def select_company(page: Page, company_code: str) -> None:
    log(f"Filtrando INTERACTIVE EXPORT por company code {company_code}.")
    show_export_search_form(page)
    field = company_input(page)
    expect(field).to_be_visible()
    field.click()
    field.fill(company_code)
    option = company_option(page, company_code)
    expect(option).to_be_visible()
    log(f"Opción de compañía detectada: {option.inner_text().strip()}.")
    option.click()
    expect(selected_company(page, company_code)).to_be_visible()
    log(f"Company code {company_code} confirmado en el filtro.")


def search_export_documents(page: Page, state: ExportListState) -> None:
    log("Preparando búsqueda de INTERACTIVE EXPORT por compañía.")
    open_interactive_export(page)
    select_company(page, get_company_code())
    search_documents(page)
    expect(invoice_number_filter(page)).to_be_visible()
    wait_for_document_list(page)
    state.list_url = page.url
    log("Listado de INTERACTIVE EXPORT cargado; ya se puede buscar por Invoice Number.")


def is_export_list_ready(page: Page, state: ExportListState) -> bool:
    # Invoice Management usa la misma tabla; la URL guardada distingue ambos listados.
    return (
        state.list_url is not None
        and page.url == state.list_url
        and page.locator("#scroller").is_visible()
        and invoice_number_filter(page).is_visible()
    )


def ensure_export_list(page: Page, state: ExportListState) -> None:
    if is_export_list_ready(page, state):
        log("Listado de INTERACTIVE EXPORT ya disponible; se reutiliza.")
        return
    search_export_documents(page, state)


def wait_for_export_detail(page: Page, list_url: str, timeout_ms: int) -> bool:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        if page.url != list_url or not invoice_number_filter(page).is_visible():
            return True
        page.wait_for_timeout(300)
    return False


def open_export_result(page: Page, state: ExportListState, folio: str, page_number: int, row: int) -> None:
    log(f"[{folio}] Entrando a la factura en INTERACTIVE EXPORT (página {page_number}, fila {row + 1}).")
    go_to_page(page, page_number)
    folio_text = page.locator("#scroller tbody tr").nth(row).get_by_text(folio, exact=True).first
    for attempt in range(1, OPEN_ATTEMPTS + 1):
        log(f"[{folio}] Intento {attempt}/{OPEN_ATTEMPTS} de entrar a la factura.")
        expect(folio_text).to_be_visible()
        folio_text.click()
        page.wait_for_timeout(ROW_SELECT_WAIT_MS)
        folio_text.dblclick()
        if wait_for_export_detail(page, state.list_url or page.url, OPEN_WAIT_MS):
            log(f"[{folio}] Factura abierta en INTERACTIVE EXPORT ({page.url}).")
            return
    raise RuntimeError(f"No se pudo entrar a la factura {folio} en INTERACTIVE EXPORT.")


def update_waiting_list(item: WorkItem, status: str, matches: int, error: Optional[str] = None) -> None:
    folio = item.folio
    in_list = record_interactive_export_result(item.invoice, status, matches, error)
    if status == EXPORT_FOUND:
        if close_waiting_invoice(item.invoice, CLOSED_COMPLETED):
            log(f"[{folio}] Retirada de la lista de espera: proceso completado en Interactive Export.")
    elif status == EXPORT_DUPLICATED:
        if close_waiting_invoice(item.invoice, CLOSED_DUPLICATED):
            log(f"[{folio}] Retirada de la lista de espera para revisión manual (duplicada).")
    elif in_list:
        log(f"[{folio}] Sigue en lista de espera (motivo: {item.reason or 'actualizado en esta corrida'}).")
    else:
        # No está en la lista: no apareció en Validation en esta corrida.
        reason = REASON_EXPORT_ERROR if status == EXPORT_NOT_OPENED else REASON_NOT_REFLECTED
        upsert_waiting_invoice(item.bundle, item.invoice, reason, status, error, new_attempt=True)
        record_interactive_export_result(item.invoice, status, matches, error)
        log(f"[{folio}] Agregada a la lista de espera (motivo: {reason}); se reintentará en futuras corridas.")


def save_export_result(item: WorkItem, status: str, matches: int, error: Optional[str] = None) -> None:
    update_waiting_list(item, status, matches, error)
    if item.bundle is not None:
        mark_interactive_export(item.bundle, item.invoice, status, matches)
        log(f"[{item.folio}] Resultado de INTERACTIVE EXPORT escrito en el JSON extraído.")


def search_in_interactive_export(page: Page, state: ExportListState, item: WorkItem) -> str:
    folio = item.folio
    log(f"[{folio}] Buscando en INTERACTIVE EXPORT (origen: {item.origin}).")
    ensure_export_list(page, state)
    filter_by_invoice_number(page, folio)
    matches = find_exact_matches(page, folio) if wait_for_results_count(page, folio) else []
    status = status_for_export_matches(len(matches))
    log(f"[{folio}] Coincidencias exactas en INTERACTIVE EXPORT: {len(matches)}.")
    if status == EXPORT_FOUND:
        try:
            open_export_result(page, state, folio, *matches[0])
        except Exception as exc:
            error = f"{type(exc).__name__}: {str(exc).splitlines()[0] if str(exc) else ''}"
            save_export_result(item, EXPORT_NOT_OPENED, len(matches), error)
            raise
    elif status == EXPORT_DUPLICATED:
        log(f"[{folio}] Folio duplicado en INTERACTIVE EXPORT; no se abrió ningún registro.")
    else:
        log(f"[{folio}] No encontrada en INTERACTIVE EXPORT; aún no se refleja en Lucy.")
    save_export_result(item, status, len(matches))
    return status
