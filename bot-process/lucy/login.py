"""Login en Lucy: una función por acción."""
from playwright.sync_api import Error as PlaywrightError, Locator, Page

from .config import LOGIN_TIMEOUT_MS, SESSION_CHECK_TIMEOUT_MS, Credentials


def app_menu_link(page: Page) -> Locator:
    # Enlace del menú lateral; existe en todas las pantallas solo con sesión iniciada.
    return page.locator("a[href$='/invoice-management']").first


def go_to_login(page: Page, login_url: str) -> None:
    page.goto(login_url)


def fill_username(page: Page, username: str) -> None:
    field = page.get_by_role("textbox", name="Username")
    field.click()
    field.fill(username)


def fill_password(page: Page, password: str) -> None:
    field = page.get_by_role("textbox", name="password")
    field.click()
    field.fill(password)


def submit_login(page: Page) -> None:
    page.get_by_role("button", name="LOGIN", exact=True).click()


def wait_for_login_success(page: Page) -> None:
    app_menu_link(page).wait_for(state="visible", timeout=LOGIN_TIMEOUT_MS)


def is_logged_in(page: Page) -> bool:
    try:
        if "/login" in page.url:
            return False
        app_menu_link(page).wait_for(state="visible", timeout=SESSION_CHECK_TIMEOUT_MS)
        return True
    except PlaywrightError:
        return False


def login(page: Page, credentials: Credentials, login_url: str) -> None:
    go_to_login(page, login_url)
    fill_username(page, credentials.username)
    fill_password(page, credentials.password)
    submit_login(page)
    wait_for_login_success(page)
