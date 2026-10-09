"""Validación de sesión/pantalla y ejecución segura de procesos."""
from typing import Callable

from playwright.sync_api import Page

from .browser import BotSession, is_page_open, reopen_page
from .config import MAX_STEP_ATTEMPTS
from .login import is_logged_in, login


def ensure_session(session: BotSession) -> Page:
    """Garantiza pantalla abierta y sesión iniciada; devuelve la página lista."""
    if not is_page_open(session):
        reopen_page(session)
    if not is_logged_in(session.page):
        login(session.page, session.credentials, session.login_url)
    return session.page


def is_session_alive(session: BotSession) -> bool:
    return is_page_open(session) and is_logged_in(session.page)


def run_step(session: BotSession, name: str, step: Callable[[Page], None]) -> None:
    """Ejecuta un proceso validando la sesión; si falla por pérdida de pantalla/sesión, reintenta."""
    for attempt in range(1, MAX_STEP_ATTEMPTS + 1):
        page = ensure_session(session)
        try:
            step(page)
            return
        except Exception as exc:  # noqa: BLE001 - expect/timeout también pueden deberse a la sesión
            if attempt == MAX_STEP_ATTEMPTS or is_session_alive(session):
                raise
            print(f"[{name}] Pantalla/sesión perdida ({type(exc).__name__}); reintentando...")
