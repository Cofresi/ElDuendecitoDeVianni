from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from pathlib import Path

from .config import AppConfig


@dataclass
class MercuryRunResult:
    success: bool
    message: str
    downloaded_file: str = ""


class MercuryAutomationError(RuntimeError):
    pass


def run_mercury_login_test(config: AppConfig, password: str) -> MercuryRunResult:
    if not config.mercury_url.strip():
        raise MercuryAutomationError("Configure primero la direccion de Mercury.")
    if not config.mercury_username.strip():
        raise MercuryAutomationError("Configure primero el usuario de Mercury.")
    if not password:
        raise MercuryAutomationError("Guarde primero la contrasena de Mercury.")

    try:
        from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise MercuryAutomationError(
            "Playwright no esta instalado todavia. Instale las dependencias y los navegadores de Playwright."
        ) from exc

    downloads = Path(config.downloads_folder)
    downloads.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=config.mercury_headless)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        try:
            page.goto(config.mercury_url, wait_until="domcontentloaded", timeout=60_000)
            _fill_login_form(page, config.mercury_username, password)
            logging.info("Prueba de Mercury completada en %s", config.mercury_url)
            return MercuryRunResult(True, "Mercury abrio correctamente y se intento iniciar sesion.")
        except PlaywrightTimeoutError as exc:
            raise MercuryAutomationError(f"Mercury tardo demasiado en responder: {exc}") from exc
        finally:
            context.close()
            browser.close()


def _fill_login_form(page, username: str, password: str) -> None:
    username_locator = _first_visible(
        page,
        [
            "input[type='email']",
            "input[name*='user' i]",
            "input[id*='user' i]",
            "input[name*='login' i]",
            "input[id*='login' i]",
            "input[type='text']",
        ],
    )
    password_locator = _first_visible(page, ["input[type='password']"])
    username_locator.fill(username)
    password_locator.fill(password)

    submit = _first_visible(
        page,
        [
            "button[type='submit']",
            "input[type='submit']",
            "button",
        ],
        required=False,
    )
    if submit:
        submit.click()
    else:
        password_locator.press("Enter")
    page.wait_for_load_state("domcontentloaded", timeout=15_000)


def _first_visible(page, selectors: list[str], required: bool = True):
    for selector in selectors:
        locator = page.locator(selector).first
        try:
            if locator.count() and locator.is_visible(timeout=1_000):
                return locator
        except Exception:
            continue
    if required:
        names = ", ".join(selectors)
        raise MercuryAutomationError(f"No se encontro un campo esperado en Mercury: {names}")
    return None
