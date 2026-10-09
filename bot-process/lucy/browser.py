"""Manejo del navegador y la pantalla (Playwright)."""
from dataclasses import dataclass
from typing import Optional

from playwright.sync_api import Browser, BrowserContext, Error as PlaywrightError, Page, Playwright

from .config import Credentials


@dataclass
class BotSession:
    playwright: Playwright
    credentials: Credentials
    login_url: str
    headless: bool = False
    browser: Optional[Browser] = None
    context: Optional[BrowserContext] = None
    page: Optional[Page] = None


def launch_browser(playwright: Playwright, headless: bool) -> Browser:
    return playwright.chromium.launch(headless=headless)


def open_context(browser: Browser) -> BrowserContext:
    return browser.new_context()


def open_page(context: BrowserContext) -> Page:
    return context.new_page()


def is_browser_open(session: BotSession) -> bool:
    return session.browser is not None and session.browser.is_connected()


def is_page_open(session: BotSession) -> bool:
    return (
        is_browser_open(session)
        and session.page is not None
        and not session.page.is_closed()
    )


def reopen_page(session: BotSession) -> None:
    if not is_browser_open(session):
        close_browser(session)
        session.browser = launch_browser(session.playwright, session.headless)
        session.context = open_context(session.browser)
    elif session.context is None:
        session.context = open_context(session.browser)
    session.page = open_page(session.context)


def close_browser(session: BotSession) -> None:
    for closable in (session.context, session.browser):
        if closable is None:
            continue
        try:
            closable.close()
        except PlaywrightError:
            pass
    session.browser = session.context = session.page = None
