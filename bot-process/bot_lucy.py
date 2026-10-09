"""Bot Lucy: orquestador principal.

1. Toma el ZIP más reciente del bot (``bot_single_vendor_*.zip``) de la carpeta de descargas
   y lo extrae a ``bot-process/cache/`` (ahí se guarda el avance de cada factura).
2. Ejecuta un paso por cada factura pendiente del JSON. Antes de cada paso se valida que la
   pantalla y la sesión sigan activas; si no, se reabre el navegador y/o se vuelve a iniciar sesión.

Configuración (LUCY_LOGIN_URL, LUCY_USERNAME, LUCY_PASSWORD, LUCY_COMPANY_CODE y opcionalmente
LUCY_BUNDLE_DIR) desde variables de entorno o desde bot-process/.env o el .env de la raíz.
"""
import sys
from pathlib import Path
from typing import Optional

from playwright.sync_api import Playwright, sync_playwright

# Permite importar el paquete ``lucy`` aunque el script se cargue por ruta (p. ej. desde web.py).
sys.path.insert(0, str(Path(__file__).resolve().parent))

from lucy.browser import BotSession, close_browser  # noqa: E402
from lucy.bundle import InvoiceBundle, load_bundle  # noqa: E402
from lucy.config import get_bundle_dir, get_credentials, get_login_url  # noqa: E402
from lucy.processes import build_steps  # noqa: E402
from lucy.session import run_step  # noqa: E402


def run(playwright: Playwright, bundle: InvoiceBundle, headless: bool = False) -> None:
    steps = build_steps(bundle)
    print(f"Bundle: {bundle.directory} | facturas pendientes: {len(steps)}")
    if not steps:
        return
    session = BotSession(
        playwright=playwright,
        credentials=get_credentials(),
        login_url=get_login_url(),
        headless=headless,
    )
    failures = []
    try:
        for name, step in steps:
            try:
                run_step(session, name, step)
            except Exception as exc:  # noqa: BLE001 - una factura fallida no detiene las demás
                failures.append(f"{name}: {type(exc).__name__}: {str(exc).splitlines()[0] if str(exc) else ''}")
                print(f"[{name}] ERROR {failures[-1]}")
    finally:
        close_browser(session)
    if failures:
        raise RuntimeError("Facturas con error (quedan pendientes): " + " | ".join(failures))


def main(headless: bool = False, bundle_path: Optional[str] = None) -> None:
    bundle = load_bundle(get_bundle_dir(), Path(bundle_path) if bundle_path else None)
    with sync_playwright() as playwright:
        run(playwright, bundle, headless=headless)


if __name__ == "__main__":
    main(bundle_path=sys.argv[1] if len(sys.argv) > 1 else None)
