"""Bot Lucy: orquestador principal.

1. Toma el ZIP más reciente del bot (``bot_single_vendor_*.zip``) de la carpeta de descargas,
   borra la caché local de la corrida anterior y lo extrae a ``%LOCALAPPDATA%/Arriendos/bot-process/cache/``.
2. Ejecuta un paso por cada factura pendiente del JSON. Antes de cada paso se valida que la
   pantalla y la sesión sigan activas; si no, se reabre el navegador y/o se vuelve a iniciar sesión.

Configuración (LUCY_LOGIN_URL, LUCY_USERNAME, LUCY_PASSWORD, LUCY_COMPANY_CODE y opcionalmente
LUCY_BUNDLE_DIR) desde variables de entorno o desde bot-process/.env o el .env de la raíz.
"""
import sys
from threading import Event
from pathlib import Path
from typing import Callable, Optional

from playwright.sync_api import Playwright, sync_playwright

# Permite importar el paquete ``lucy`` aunque el script se cargue por ruta (p. ej. desde web.py).
sys.path.insert(0, str(Path(__file__).resolve().parent))

from lucy.browser import BotSession, close_browser  # noqa: E402
from lucy.bundle import InvoiceBundle, load_bundle  # noqa: E402
from lucy.config import get_bundle_dir, get_credentials, get_login_url  # noqa: E402
from lucy.history import load_pending_validation_history  # noqa: E402
from lucy.logging import log  # noqa: E402
from lucy.logging import set_log_callback  # noqa: E402
from lucy.processes import build_steps  # noqa: E402
from lucy.session import run_step  # noqa: E402


def run(
    playwright: Playwright,
    bundle: InvoiceBundle,
    headless: bool = False,
    keep_open: bool = False,
    hold_open: bool = False,
    finish_event: Optional[Event] = None,
    on_hold: Optional[Callable[[], None]] = None,
) -> None:
    steps = build_steps(bundle)
    log(f"Paquete cargado: {bundle.directory.name}")
    log(f"Facturas pendientes para Lucy: {len(steps)}")
    if not steps:
        log("No hay facturas pendientes; no se abrió Edge.")
        return
    for invoice in bundle.data.get("facturas", []):
        folio = str(invoice.get("folio", "")).strip()
        result = invoice.get("validacion_lucy") or {}
        if invoice.get("lista_para_bot") and not result.get("estado"):
            for part in invoice.get("partes", []):
                log(
                    f"[{folio}] Detectada: vendor={part.get('vendor')}, "
                    f"tienda={part.get('tienda')}, CECO={part.get('ceco')}, "
                    f"base={part.get('base')}, IVA={part.get('iva')}, "
                    f"importe={part.get('importe')}, soporte="
                    f"{Path(part['archivo_factura']).name if part.get('archivo_factura') else 'no disponible'}."
                )
        elif not invoice.get("lista_para_bot"):
            log(f"[{folio}] Omitida: no está marcada como lista para Lucy.")
        elif result.get("estado"):
            log(f"[{folio}] Omitida: ya tiene resultado previo ({result['estado']}).")

    session = BotSession(
        playwright=playwright,
        credentials=get_credentials(),
        login_url=get_login_url(),
        headless=headless,
    )
    log("Configuración cargada; se iniciará Microsoft Edge y el acceso a Lucy.")
    failures = []
    try:
        for name, step in steps:
            try:
                log(f"[{name}] Iniciando revisión.")
                run_step(session, name, step)
            except Exception as exc:  # noqa: BLE001 - una factura fallida no detiene las demás
                failures.append(f"{name}: {type(exc).__name__}: {str(exc).splitlines()[0] if str(exc) else ''}")
                log(f"[{name}] ERROR {failures[-1]}")
    finally:
        if keep_open:
            input("Revisa el resultado en el navegador y presiona Enter para cerrarlo...")
        elif hold_open and session.browser is not None and finish_event is not None:
            log("Revisión terminada. Edge queda abierto en la última pantalla; la factura NO se guardó.")
            if on_hold is not None:
                on_hold()
            finish_event.wait()
            log("Se solicitó finalizar la inspección; cerrando Edge.")
        close_browser(session)
    if failures:
        raise RuntimeError("Facturas con error (quedan pendientes): " + " | ".join(failures))


def main(
    headless: bool = False,
    bundle_path: Optional[str] = None,
    keep_open: bool = False,
    hold_open: bool = False,
    finish_event: Optional[Event] = None,
    on_hold: Optional[Callable[[], None]] = None,
) -> None:
    bundle_dir = get_bundle_dir()
    log(f"Buscando el paquete más reciente en {bundle_dir}.")
    bundle = load_bundle(bundle_dir, Path(bundle_path) if bundle_path else None)
    log(f"JSON de control cargado: {bundle.json_path.name}.")
    pending_history = load_pending_validation_history()
    log(f"Folios encontrados anteriormente y pendientes del siguiente paso: {len(pending_history)}.")
    try:
        with sync_playwright() as playwright:
            run(
                playwright,
                bundle,
                headless=headless,
                keep_open=keep_open,
                hold_open=hold_open,
                finish_event=finish_event,
                on_hold=on_hold,
            )
    finally:
        set_log_callback(None)


if __name__ == "__main__":
    main(bundle_path=sys.argv[1] if len(sys.argv) > 1 else None, keep_open=True)
