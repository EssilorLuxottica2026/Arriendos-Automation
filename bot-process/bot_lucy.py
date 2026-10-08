"""Bot Lucy: automatización grabada con Playwright (flujo parcial).

Credenciales (LUCY_USERNAME, LUCY_PASSWORD) desde variables de entorno o
desde el archivo bot-process/.env (ignorado por git).
"""
import os
import re
from pathlib import Path

from playwright.sync_api import Playwright, sync_playwright

LUCY_LOGIN_URL = "https://lucy-frontend-nf4ebt4lna-ew.a.run.app/login"
ENV_FILE = Path(__file__).with_name(".env")


def load_env(path: Path = ENV_FILE) -> None:
    if not path.is_file():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip("\"'"))


def run(playwright: Playwright, headless: bool = False) -> None:
    load_env()
    username = os.environ["LUCY_USERNAME"]
    password = os.environ["LUCY_PASSWORD"]

    browser = playwright.chromium.launch(headless=headless)
    context = browser.new_context()
    page = context.new_page()

    page.goto(LUCY_LOGIN_URL)
    page.get_by_role("textbox", name="Username").click()
    page.get_by_role("textbox", name="Username").fill(username)
    page.get_by_role("textbox", name="password").click()
    page.get_by_role("textbox", name="password").fill(password)
    page.get_by_role("button", name="LOGIN", exact=True).click()

    page.get_by_role("link", name="Invoice Management").click()
    page.locator(".css-1d2ztni-indicatorContainer").first.click()
    page.locator("#react-select-3-option-0").click()
    page.locator("div").filter(has_text=re.compile(r"^Select\.\.\.$")).nth(3).click()
    page.locator("#react-select-6-input").fill("8140")
    page.get_by_text("8140-GMO Colombia", exact=True).click()
    page.get_by_role("button", name="Search Documents").click()

    page.locator('[id="91-Barcode"]').click()
    page.locator('[id="91-Barcode"]').fill("8")

    # TODO: continuar el flujo grabado.

    context.close()
    browser.close()


def main(headless: bool = False) -> None:
    with sync_playwright() as playwright:
        run(playwright, headless=headless)


if __name__ == "__main__":
    main()
