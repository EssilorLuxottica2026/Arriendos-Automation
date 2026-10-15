"""Proceso: repost de facturas multi-vendor (una copia en Lucy por cada vendor de la factura).

Flujo: Interactive Export -> Litigation Archive -> Reverse (Repost x N vendors) ->
Interactive Export (verificar N facturas) -> completar cada copia con su CSV (ver invoice_parts) ->
validar que el botón Export (SAP) exista (nunca se pulsa).

Con ``REPOST_DRY_RUN`` activo no se pulsan Litigation Archive, Reverse - Repost ni Repost:
solo se valida que existan y se continúa con el siguiente paso.
"""
from typing import Callable, Optional

from playwright.sync_api import Error as PlaywrightError, Locator, Page, expect

from ..bundle import REPOST_COMPLETED, InvoiceBundle, get_invoices, get_folio, has_completed_repost, mark_repost
from ..config import (
    REPOST_DRY_RUN,
    REPOST_VERIFY_ATTEMPTS,
    REPOST_VERIFY_WAIT_MS,
    get_company_code,
)
from .invoice_management import (
    back_to_list,
    enter_company_code,
    search_documents,
    show_search_form,
    wait_for_company_selected,
    wait_for_document_list,
)
from .invoice_validation import (
    BARCODE_COLUMN,
    DETAIL_LOAD_TIMEOUT_MS,
    OPEN_ATTEMPTS,
    ROW_SELECT_WAIT_MS,
    back_to_list_button,
    column_filter,
    filter_by_column,
    find_exact_matches,
    get_results_count,
    go_to_page,
    wait_for_results_count,
)

REPOST_DRY_RUN_COMPLETED = "Repost dry-run completed"
REPOST_ERROR = "Repost error"

INTERACTIVE_EXPORT = "INTERACTIVE EXPORT"
REVERSE = "REVERSE"

Location = tuple[int, int]
Step = tuple[str, Callable[[Page], None]]


class RepostError(RuntimeError):
    """El flujo debe detenerse y reportarse para esta factura."""


def get_vendor_count(invoice: dict) -> int:
    return len(invoice.get("vendors") or []) or len(invoice.get("partes") or [])


def get_barcode(invoice: dict) -> str:
    barcodes = {str(part["barcode"]).strip() for part in invoice["partes"] if part.get("barcode")}
    if len(barcodes) != 1:
        raise RepostError(
            f"La factura {get_folio(invoice)} debe tener exactamente un barcode en sus partes (tiene {len(barcodes)})."
        )
    return barcodes.pop()


def select_section(page: Page, name: str) -> None:
    # Los ids "react-select-XXX" cambian en cada carga: se usa el primer react-select (selector de sección).
    page.locator("input[id^='react-select-'][id$='-input']").first.click()
    page.get_by_text(name, exact=True).last.click()
    expect(page.get_by_text(name, exact=True).first).to_be_visible()


def list_barcode_matches(page: Page, barcode: str) -> list[Location]:
    """Filtra la tabla por Barcode y devuelve (página, fila) de cada coincidencia exacta."""
    # Se limpia antes de escribir para que Lucy vuelva a filtrar aunque el texto no haya cambiado.
    field = column_filter(page, BARCODE_COLUMN)
    expect(field).to_be_visible()
    field.fill("")
    filter_by_column(page, BARCODE_COLUMN, barcode)
    if not wait_for_results_count(page, barcode, BARCODE_COLUMN):
        return []
    return find_exact_matches(page, barcode, BARCODE_COLUMN)


def open_document(page: Page, barcode: str, location: Location) -> None:
    page_number, row = location
    go_to_page(page, page_number)
    barcode_text = page.locator("#scroller tbody tr").nth(row).get_by_text(barcode, exact=True).first
    for _ in range(OPEN_ATTEMPTS):
        expect(barcode_text).to_be_visible()
        # El primer clic selecciona la fila y la re-renderiza; el doble clic va sobre la fila estable.
        barcode_text.click()
        page.wait_for_timeout(ROW_SELECT_WAIT_MS)
        barcode_text.dblclick()
        try:
            back_to_list_button(page).wait_for(timeout=DETAIL_LOAD_TIMEOUT_MS)
            return
        except PlaywrightError:
            continue
    raise RepostError(f"No se pudo abrir el documento con barcode {barcode}.")


def expect_button_exists(page: Page, name: str, exact: bool = False) -> Locator:
    button = page.get_by_role("button", name=name, exact=exact)
    expect(button).to_be_visible()
    return button


def find_single_invoice(page: Page, barcode: str) -> Location:
    select_section(page, INTERACTIVE_EXPORT)
    matches = list_barcode_matches(page, barcode)
    if len(matches) != 1:
        raise RepostError(
            f"{len(matches)} resultados para el barcode {barcode}: duplicación previa desconocida. No se puede procesar."
        )
    return matches[0]


def litigation_archive(page: Page, barcode: str, location: Location) -> None:
    open_document(page, barcode, location)
    button = expect_button_exists(page, "Litigation Archive")
    if REPOST_DRY_RUN:
        print(f"[Repost {barcode}] DRY-RUN: botón Litigation Archive existe; no se pulsa.")
        back_to_list(page)
        return
    button.click()
    page.get_by_role("button", name="Confirm").click()


def show_reverse_form(page: Page) -> None:
    barcode_input = page.locator("#Barcode")
    try:
        barcode_input.wait_for(state="visible", timeout=5_000)
    except PlaywrightError:
        page.get_by_role("button", name="New Search").click()
        expect(barcode_input).to_be_visible()


def search_in_reverse(page: Page, barcode: str) -> int:
    show_reverse_form(page)
    page.locator("#Barcode").fill(barcode)
    search_documents(page)
    wait_for_document_list(page)
    return get_results_count(page)


def repost_documents(page: Page, barcode: str, repost_count: int) -> None:
    select_section(page, REVERSE)
    for number in range(1, repost_count + 1):
        results = search_in_reverse(page, barcode)
        if REPOST_DRY_RUN and results == 0:
            # Sin Litigation Archive real, el documento aún no está disponible en REVERSE.
            print(f"[Repost {barcode}] DRY-RUN: sin resultados en REVERSE; se omiten los reposts.")
            return
        if results != 1:
            raise RepostError(f"{results} resultados en REVERSE para el barcode {barcode}; se esperaba 1.")
        open_document(page, barcode, (1, 0))
        button = expect_button_exists(page, "Reverse - Repost")
        if REPOST_DRY_RUN:
            print(f"[Repost {barcode}] DRY-RUN: botón Reverse - Repost existe; no se pulsa ({number}/{repost_count}).")
            back_to_list(page)
            continue
        button.click()
        repost = expect_button_exists(page, "Repost", exact=True)
        repost.click()
        page.get_by_role("button", name="Close", exact=True).click()
        back_to_list(page)


def search_company_in_interactive_export(page: Page) -> None:
    select_section(page, INTERACTIVE_EXPORT)
    show_search_form(page)
    company_code = get_company_code()
    enter_company_code(page, company_code)
    wait_for_company_selected(page, company_code)
    search_documents(page)
    wait_for_document_list(page)


def verify_reposts(page: Page, barcode: str, expected: int) -> None:
    found = -1
    for _ in range(REPOST_VERIFY_ATTEMPTS):
        page.get_by_role("button", name="Refresh").click()
        found = len(list_barcode_matches(page, barcode))
        if found > expected:
            raise RepostError(f"Algo salió mal con el repost de {barcode}: {found} resultados, se esperaban {expected}.")
        if found == expected:
            return
        page.wait_for_timeout(REPOST_VERIFY_WAIT_MS)
    raise RepostError(
        f"Error al hacer el repost de {barcode}: tras {REPOST_VERIFY_ATTEMPTS} intentos hay {found} resultados, "
        f"se esperaban {expected}."
    )


def verify_export_button(page: Page) -> None:
    # No se exporta a SAP: solo se valida que el botón exista.
    expect_button_exists(page, "Export", exact=True)


def run_repost(page: Page, bundle: InvoiceBundle, invoice: dict) -> None:
    # Import local: invoice_parts importa los helpers de este módulo.
    from .invoice_parts import preflight_parts, process_parts

    barcode = get_barcode(invoice)
    vendor_count = get_vendor_count(invoice)
    preflight_parts(bundle, invoice)
    location = find_single_invoice(page, barcode)
    litigation_archive(page, barcode, location)
    repost_documents(page, barcode, vendor_count)
    search_company_in_interactive_export(page)
    # Litigation Archive elimina la original: tras los reposts deben quedar tantas facturas como vendors.
    verify_reposts(page, barcode, 1 if REPOST_DRY_RUN else vendor_count)
    process_parts(page, bundle, invoice, barcode)
    verify_export_button(page)


def repost_invoice(page: Page, bundle: InvoiceBundle, invoice: dict) -> str:
    folio = get_folio(invoice)
    try:
        run_repost(page, bundle, invoice)
    except Exception as exc:  # noqa: BLE001 - se deja constancia en el JSON y se vuelve a lanzar
        detail = f"{type(exc).__name__}: {exc}"
        mark_repost(bundle, invoice, REPOST_ERROR, detail)
        print(f"[Repost {folio}] ERROR {detail}")
        raise
    status = REPOST_DRY_RUN_COMPLETED if REPOST_DRY_RUN else REPOST_COMPLETED
    mark_repost(bundle, invoice, status)
    return status


def is_multi_vendor(invoice: dict) -> bool:
    return get_vendor_count(invoice) > 1


def get_repost_invoices(bundle: InvoiceBundle, folios: Optional[list[str]] = None) -> list[dict]:
    """Facturas multi-vendor sin repost completado; con ``folios`` se toman esas facturas tal cual."""
    invoices = get_invoices(bundle)
    if folios:
        by_folio = {get_folio(invoice): invoice for invoice in invoices}
        missing = [folio for folio in folios if folio not in by_folio]
        if missing:
            raise RepostError(f"Folios no encontrados en el JSON: {', '.join(missing)}")
        return [by_folio[folio] for folio in folios]
    return [invoice for invoice in invoices if is_multi_vendor(invoice) and not has_completed_repost(invoice)]


def repost_step(bundle: InvoiceBundle, invoice: dict) -> Step:
    folio = get_folio(invoice)

    def step(page: Page) -> None:
        status = repost_invoice(page, bundle, invoice)
        print(f"[Repost {folio}] {status}")

    return f"Repost {folio}", step


def build_repost_steps(bundle: InvoiceBundle, folios: Optional[list[str]] = None) -> list[Step]:
    return [repost_step(bundle, invoice) for invoice in get_repost_invoices(bundle, folios)]
