"""Plan de trabajo: qué facturas se buscan en Validation y cuáles en Interactive Export."""
from dataclasses import dataclass
from typing import Optional

from .bundle import InvoiceBundle, get_folio
from .history import VALIDATION_RETRY_REASONS, invoice_key


@dataclass
class WorkItem:
    invoice: dict
    # ``None`` cuando la factura viene solo de la lista de espera (no de la corrida actual).
    bundle: Optional[InvoiceBundle]
    reason: Optional[str]

    @property
    def folio(self) -> str:
        return get_folio(self.invoice)

    @property
    def key(self) -> str:
        return invoice_key(self.invoice)

    @property
    def origin(self) -> str:
        if self.bundle is None:
            return f"lista de espera ({self.reason})"
        return "corrida actual" + (f" + lista de espera ({self.reason})" if self.reason else "")


def _merge(
    bundle: InvoiceBundle,
    run_invoices: list[dict],
    waiting: list[dict],
) -> list[WorkItem]:
    reasons = {invoice_key(record["informacion_factura"]): record["motivo_espera"] for record in waiting}
    items: dict[str, WorkItem] = {}
    for invoice in run_invoices:
        key = invoice_key(invoice)
        items.setdefault(key, WorkItem(invoice, bundle, reasons.get(key)))
    for record in waiting:
        invoice = record["informacion_factura"]
        items.setdefault(invoice_key(invoice), WorkItem(invoice, None, record["motivo_espera"]))
    return list(items.values())


def build_validation_items(
    bundle: InvoiceBundle,
    run_invoices: list[dict],
    waiting: list[dict],
) -> list[WorkItem]:
    """Facturas nuevas y las de la lista de espera que aún pueden aparecer en Validation."""
    return [
        item
        for item in _merge(bundle, run_invoices, waiting)
        if item.reason is None or item.reason in VALIDATION_RETRY_REASONS
    ]


def build_export_items(
    bundle: InvoiceBundle,
    run_invoices: list[dict],
    waiting: list[dict],
    excluded_keys: set[str] = frozenset(),
) -> list[WorkItem]:
    """Todas las facturas de la corrida y de la lista de espera, salvo las excluidas (duplicadas)."""
    return [item for item in _merge(bundle, run_invoices, waiting) if item.key not in excluded_keys]
