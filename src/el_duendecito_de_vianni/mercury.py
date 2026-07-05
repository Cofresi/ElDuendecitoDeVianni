from __future__ import annotations

import logging
import os
from urllib.parse import urlparse
from dataclasses import dataclass
from pathlib import Path

from .config import AppConfig
from .importer import EXPORT_NAMES


@dataclass
class MercuryRunResult:
    success: bool
    message: str
    downloaded_file: str = ""


class MercuryAutomationError(RuntimeError):
    pass


def run_mercury_login_test(config: AppConfig, password: str) -> MercuryRunResult:
    return run_mercury_export(config, password, download_report=False)


def run_mercury_export(config: AppConfig, password: str, download_report: bool = True) -> MercuryRunResult:
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
        browser = playwright.chromium.launch(headless=config.mercury_headless, **_browser_launch_options())
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        try:
            page.goto(config.mercury_url, wait_until="domcontentloaded", timeout=60_000)
            _fill_login_form(page, config.mercury_username, password)
            if not download_report:
                logging.info("Prueba de Mercury completada en %s", config.mercury_url)
                return MercuryRunResult(True, "Mercury abrio correctamente y se intento iniciar sesion.")
            downloaded = _download_existing_report(page, config)
            logging.info("Reporte de Mercury descargado: %s", downloaded)
            return MercuryRunResult(
                True,
                f"Mercury descargo el reporte {config.mercury_report_name}.",
                downloaded_file=str(downloaded),
            )
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


def _download_existing_report(page, config: AppConfig) -> Path:
    _select_company(page, config.mercury_company)
    page.goto(_report_generator_url(config.mercury_url), wait_until="domcontentloaded", timeout=60_000)
    _click_text(page, "Abrir reporte existente")
    _click_text(page, config.mercury_report_name)
    _click_text(page, "Aceptar")
    page.wait_for_load_state("domcontentloaded", timeout=15_000)
    _click_text(page, "Ver reporte")

    downloads = Path(config.downloads_folder)
    for name in EXPORT_NAMES:
        (downloads / name).unlink(missing_ok=True)

    with page.expect_download(timeout=60_000) as download_info:
        _click_text(page, "Exportar")
    download = download_info.value
    suffix = Path(download.suggested_filename or "").suffix or ".xlsx"
    target = downloads / f"GridViewExport{suffix}"
    download.save_as(target)
    return target


def _select_company(page, company_name: str) -> None:
    if not company_name:
        return
    select = page.locator("select").first
    try:
        if select.count() and select.is_visible(timeout=5_000):
            select.select_option(label=company_name)
            page.wait_for_load_state("domcontentloaded", timeout=10_000)
    except Exception:
        logging.info("No se selecciono compania en pantalla inicial; se continua con la sesion actual.")


def _report_generator_url(login_url: str) -> str:
    parsed = urlparse(login_url)
    if not parsed.scheme or not parsed.netloc:
        raise MercuryAutomationError("La direccion de Mercury no parece ser una URL valida.")
    return f"{parsed.scheme}://{parsed.netloc}/Mercury.RRHH/GeneradorReportes.aspx"


def _click_text(page, text: str) -> None:
    page.get_by_text(text, exact=True).click(timeout=20_000)


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


def _browser_launch_options() -> dict[str, str]:
    executable_path = _find_playwright_chromium()
    if executable_path:
        return {"executable_path": str(executable_path)}
    return {}


def _find_playwright_chromium() -> Path | None:
    override = os.getenv("EL_DUENDECITO_CHROMIUM_EXE")
    if override and Path(override).exists():
        return Path(override)

    local_app_data = os.getenv("LOCALAPPDATA")
    if not local_app_data:
        return None

    browser_root = Path(local_app_data) / "ms-playwright"
    if not browser_root.exists():
        return None

    matches = sorted(browser_root.glob("chromium-*/chrome-win64/chrome.exe"), reverse=True)
    return matches[0] if matches else None
