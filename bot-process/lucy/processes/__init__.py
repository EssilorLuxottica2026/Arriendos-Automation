"""Procesos del bot Lucy.

Cada paso es ``(nombre, función(page))`` y se ejecuta con ``run_step`` (validación de
pantalla/sesión y reintento). Primero se ejecutan los pasos de Validation y después los de
Interactive Export.
"""
from typing import Callable

from playwright.sync_api import Page

from ..plan import WorkItem
from .interactive_export import ExportListState, search_in_interactive_export
from .invoice_validation import INVOICE_DUPLICATED, validate_invoice

Step = tuple[str, Callable[[Page], None]]


def build_validation_steps(items: list[WorkItem], duplicated_keys: set[str]) -> list[Step]:
    """Pasos de Validation; los folios duplicados se agregan a ``duplicated_keys``."""
    def validation_step(item: WorkItem) -> Step:
        def step(page: Page) -> None:
            status = validate_invoice(page, item)
            if status == INVOICE_DUPLICATED:
                duplicated_keys.add(item.key)

        return f"Validation {item.folio}", step

    return [validation_step(item) for item in items]


def build_export_steps(items: list[WorkItem]) -> list[Step]:
    state = ExportListState()

    def export_step(item: WorkItem) -> Step:
        def step(page: Page) -> None:
            search_in_interactive_export(page, state, item)

        return f"Interactive Export {item.folio}", step

    return [export_step(item) for item in items]
