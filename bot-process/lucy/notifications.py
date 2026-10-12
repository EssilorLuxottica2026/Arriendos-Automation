"""Mensajes propios de Lucy (notificaciones tipo toast en ``#notification-portal``).

Lucy muestra ``div[data-testid="notification"]`` con atributo ``type`` (``fail`` en errores) y el
texto en un ``span``. Son transitorias, por eso se instala un ``MutationObserver`` que guarda
cada notificación en ``window.__lucyBotNotifications`` apenas aparece.
"""
from playwright.sync_api import Error as PlaywrightError, Page

from .logging import log

ERROR_TYPES = {"fail", "error"}

_OBSERVER_SCRIPT = """
() => {
  const selector = '[data-testid="notification"]';
  window.__lucyBotNotifications = [];
  const read = (node) => {
    const element = node instanceof Element ? node : node?.parentElement;
    const wrapper = element?.closest?.(selector) || element?.querySelector?.(selector);
    if (!wrapper) return;
    const text = (wrapper.querySelector('span')?.textContent || wrapper.textContent || '').trim();
    if (!text) return;
    const type = wrapper.getAttribute('type') || '';
    const last = window.__lucyBotNotifications[window.__lucyBotNotifications.length - 1];
    if (last && last.text === text && last.type === type) return;
    window.__lucyBotNotifications.push({ text, type });
  };
  if (!window.__lucyBotObserverInstalled) {
    window.__lucyBotObserverInstalled = true;
    new MutationObserver((mutations) => {
      for (const mutation of mutations) {
        mutation.addedNodes.forEach(read);
        read(mutation.target);
      }
    }).observe(document.body, { childList: true, subtree: true, characterData: true });
  }
}
"""

_CONSUME_SCRIPT = """
() => {
  const items = window.__lucyBotNotifications || [];
  window.__lucyBotNotifications = [];
  return items;
}
"""


class LucyNotificationError(RuntimeError):
    """Lucy mostró una notificación de error durante el proceso."""


def arm_notification_observer(page: Page) -> None:
    """Empieza a registrar notificaciones nuevas (descarta las capturadas antes)."""
    try:
        page.evaluate(_OBSERVER_SCRIPT)
    except PlaywrightError as exc:
        log(f"AVISO: no se pudo activar el observador de notificaciones de Lucy: {exc}.")


def consume_notifications(page: Page) -> list[dict]:
    """Notificaciones capturadas desde la última lectura; las elimina del registro."""
    try:
        items = page.evaluate(_CONSUME_SCRIPT) or []
    except PlaywrightError:
        return []
    return [
        {"text": str(item.get("text") or "").strip(), "type": str(item.get("type") or "").strip()}
        for item in items
        if str(item.get("text") or "").strip()
    ]


def visible_error_notifications(page: Page) -> list[str]:
    """Errores de Lucy visibles ahora mismo (respaldo si el observador no estaba activo)."""
    try:
        notifications = page.locator("[data-testid='notification']")
        messages = []
        for index in range(notifications.count()):
            notification = notifications.nth(index)
            if (notification.get_attribute("type") or "") in ERROR_TYPES and notification.is_visible():
                text = notification.inner_text().strip()
                if text:
                    messages.append(" ".join(text.split()))
        return messages
    except PlaywrightError:
        return []


def error_messages(notifications: list[dict]) -> list[str]:
    return [item["text"] for item in notifications if item["type"] in ERROR_TYPES]


def check_lucy_errors(page: Page, folio: str, context: str) -> None:
    """Registra las notificaciones nuevas de Lucy y falla si alguna es de error."""
    notifications = consume_notifications(page)
    for item in notifications:
        log(f"[{folio}] Lucy notificó ({item['type'] or 'sin tipo'}) tras {context}: {item['text']}")
    errors = error_messages(notifications)
    if errors:
        raise LucyNotificationError(f"Lucy reportó error tras {context}: {' | '.join(errors)}")


def describe_lucy_errors(page: Page) -> str:
    """Texto de errores de Lucy pendientes o visibles, para adjuntarlo al error de la factura."""
    messages = error_messages(consume_notifications(page)) or visible_error_notifications(page)
    return " | ".join(dict.fromkeys(messages))


def failure_message(page: Page, exc: Exception) -> str:
    """Error de la factura; incluye el mensaje de Lucy si lo hubo y no está ya en la excepción."""
    error = f"{type(exc).__name__}: {str(exc).splitlines()[0] if str(exc) else ''}"
    lucy_errors = "" if isinstance(exc, LucyNotificationError) else describe_lucy_errors(page)
    if lucy_errors:
        log(f"Mensaje de Lucy asociado al error: {lucy_errors}")
        error += f" | Lucy: {lucy_errors}"
    return error
