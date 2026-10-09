"""Procesos del bot Lucy.

``build_steps`` arma la lista de pasos a ejecutar; cada paso es ``(nombre, función(page))``
y se ejecuta con ``run_step`` (validación de pantalla/sesión y reintento).
"""
from typing import Callable

from playwright.sync_api import Page

from ..bundle import InvoiceBundle, get_folio, get_pending_invoices
from .invoice_validation import validate_invoice

Step = tuple[str, Callable[[Page], None]]


def invoice_step(bundle: InvoiceBundle, invoice: dict) -> Step:
    folio = get_folio(invoice)

    def step(page: Page) -> None:
        status = validate_invoice(page, bundle, invoice)
        print(f"[Factura {folio}] {status}")

    return f"Factura {folio}", step


def build_steps(bundle: InvoiceBundle) -> list[Step]:
    return [invoice_step(bundle, invoice) for invoice in get_pending_invoices(bundle)]
