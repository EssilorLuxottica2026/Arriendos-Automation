"""Proceso: Invoice Management (validación de documentos)."""
import re

from playwright.sync_api import Error as PlaywrightError, Page, expect

from ..config import get_company_code
from ..login import app_menu_link


def open_invoice_management(page: Page) -> None:
    link = app_menu_link(page)
    expect(link).to_be_visible()
    link.click()


def company_code_pattern(company_code: str) -> re.Pattern:
    return re.compile(rf"^{re.escape(company_code)}-")


def company_code_input(page: Page):
    # El id del input de react-select es dinámico; el contenedor "#company" es estable.
    return page.locator("#company input[id^='react-select-']")


def enter_company_code(page: Page, company_code: str) -> None:
    field = company_code_input(page)
    expect(field).to_be_visible()
    field.fill(company_code)
    field.press("Enter")


def wait_for_company_selected(page: Page, company_code: str) -> None:
    expect(page.locator("#company")).to_contain_text(company_code_pattern(company_code))


def search_documents(page: Page) -> None:
    button = page.get_by_role("button", name="Search Documents")
    expect(button).to_be_visible()
    button.click()


def wait_for_document_list(page: Page) -> None:
    expect(page.locator("#scroller")).to_contain_text(re.compile(r"\d+ results"), timeout=30_000)


def is_document_detail_open(page: Page) -> bool:
    return "/invoice-management/edit" in page.url


def back_to_list(page: Page) -> None:
    button = page.get_by_role("button", name="Back to List")
    expect(button).to_be_visible()
    button.click()
    wait_for_document_list(page)


def is_document_list_ready(page: Page) -> bool:
    # La lista solo existe después de la búsqueda por company code hecha en esta sesión.
    return (
        page.url.rstrip("/").endswith("/invoice-management")
        and page.locator("#scroller").is_visible()
        and page.locator("[id$='-Invoice Number']").is_visible()
    )


def ensure_document_list(page: Page) -> None:
    """Deja la pantalla en la lista de documentos del company code, venga de donde venga."""
    if is_document_detail_open(page):
        back_to_list(page)
    if not is_document_list_ready(page):
        search_company_documents(page)


def show_search_form(page: Page) -> None:
    # Si Invoice Management muestra una lista previa, el formulario se abre con "New Search".
    company = company_code_input(page)
    try:
        company.wait_for(state="visible", timeout=5_000)
    except PlaywrightError:
        page.get_by_role("button", name="New Search").click()
        expect(company).to_be_visible()


def search_company_documents(page: Page) -> None:
    open_invoice_management(page)
    show_search_form(page)
    company_code = get_company_code()
    enter_company_code(page, company_code)
    wait_for_company_selected(page, company_code)
    search_documents(page)
    wait_for_document_list(page)
