from __future__ import annotations

import logging
import os
import re
from urllib.parse import urlparse
from dataclasses import dataclass, field
from pathlib import Path

from .config import AppConfig
from .importer import EXPORT_NAMES


@dataclass
class MercuryRunResult:
    success: bool
    message: str
    downloaded_file: str = ""
    downloaded_files: list[str] = field(default_factory=list)
    companies_without_download: list[str] = field(default_factory=list)


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
        try:
            if not download_report:
                context = browser.new_context(accept_downloads=True)
                page = context.new_page()
                try:
                    _open_authenticated_session(page, config, password)
                    logging.info("Prueba de Mercury completada en %s", config.mercury_url)
                    return MercuryRunResult(True, "Mercury abrio correctamente y se intento iniciar sesion.")
                finally:
                    context.close()
            downloaded_files: list[str] = []
            empty_companies: list[str] = []
            for company in _configured_companies(config):
                context = browser.new_context(accept_downloads=True)
                page = context.new_page()
                try:
                    _open_authenticated_session(page, config, password)
                    downloaded = _download_existing_report(page, config, company)
                    if downloaded:
                        downloaded_files.append(str(downloaded))
                        logging.info("Reporte de Mercury descargado para %s: %s", company, downloaded)
                    else:
                        empty_companies.append(company)
                        logging.info("Mercury no descargo datos para %s.", company)
                finally:
                    context.close()
            if not downloaded_files:
                return MercuryRunResult(
                    True,
                    "Mercury no encontro nuevas entradas para las companias configuradas.",
                    downloaded_files=[],
                    companies_without_download=empty_companies,
                )
            return MercuryRunResult(
                True,
                f"Mercury descargo {len(downloaded_files)} reporte(s).",
                downloaded_file=downloaded_files[0],
                downloaded_files=downloaded_files,
                companies_without_download=empty_companies,
            )
        except PlaywrightTimeoutError as exc:
            raise MercuryAutomationError(f"Mercury tardo demasiado en responder: {exc}") from exc
        finally:
            browser.close()


def _open_authenticated_session(page, config: AppConfig, password: str) -> None:
    page.goto(config.mercury_url, wait_until="domcontentloaded", timeout=60_000)
    _fill_login_form(page, config.mercury_username, password)


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
    _wait_for_post_login_page(page)


def _download_existing_report(page, config: AppConfig, company_name: str) -> Path | None:
    _select_company(page, company_name)
    _open_human_resources(page)
    _open_report_generator(page, config)
    _click_text(page, "Abrir reporte existente")
    _click_text(page, config.mercury_report_name)
    _click_selector_or_text(page, "#ctl00_MainContent_cmdAceptarReporte", "Aceptar")
    page.wait_for_load_state("domcontentloaded", timeout=15_000)
    _click_text(page, "Ver reporte")
    page.wait_for_load_state("domcontentloaded", timeout=15_000)
    if _wait_for_no_records(page, timeout=8_000):
        logging.info("Mercury no encontro registros para %s.", company_name)
        return None

    downloads = Path(config.downloads_folder)
    for name in EXPORT_NAMES:
        (downloads / name).unlink(missing_ok=True)

    try:
        with page.expect_download(timeout=15_000) as download_info:
            _click_text(page, "Exportar")
        download = download_info.value
    except Exception:
        if _page_has_no_records(page):
            return None
        raise
    suffix = Path(download.suggested_filename or "").suffix or ".xlsx"
    target = downloads / f"GridViewExport_{_company_file_label(company_name)}{suffix}"
    download.save_as(target)
    return target


def _configured_companies(config: AppConfig) -> list[str]:
    raw = config.mercury_companies or config.mercury_company
    companies = [part.strip() for part in re.split(r"[;\n,]+", raw) if part.strip()]
    if not companies and config.mercury_company.strip():
        companies.append(config.mercury_company.strip())
    return companies or ["Supermercado Ines"]


def _company_file_label(company_name: str) -> str:
    normalized = company_name.casefold()
    if "ines" in normalized:
        return "Ines"
    if "brother" in normalized:
        return "Brothers"
    value = re.sub(r"[^A-Za-z0-9]+", "_", company_name).strip("_")
    return value or "Mercury"


def _page_has_no_records(page) -> bool:
    for frame in page.frames:
        try:
            text = frame.locator("body").inner_text(timeout=1_000).casefold()
            if any(message in text for message in _NO_RECORDS_MESSAGES):
                return True
        except Exception:
            continue
    return False


_NO_RECORDS_MESSAGES = (
    "no existen registros",
    "no existen datos",
    "no data",
    "sin registros",
)

_COMPANY_LABELS = ("Compañía", "Compania", "Company")
_HUMAN_RESOURCES_LABELS = (
    "Recursos Humanos",
    "Administración Recursos Humanos",
    "Administracion Recursos Humanos",
    "Human Resources",
    "Human Resources Management",
)
_POST_HR_MARKERS = ("Principal", "Generador de Reportes", "Recursos Humanos")
_REPORT_GENERATOR_LABELS = ("Generador de reportes", "Generador de Reportes", "Report Generator")


def _wait_for_no_records(page, timeout: int = 5_000) -> bool:
    elapsed = 0
    step = 500
    while elapsed <= timeout:
        if _page_has_no_records(page):
            return True
        page.wait_for_timeout(step)
        elapsed += step
    return False


def _select_company(page, company_name: str) -> None:
    if not company_name:
        return
    _wait_for_post_login_page(page)
    select = page.locator("select").first
    if not select.count() or not select.is_visible(timeout=5_000):
        raise MercuryAutomationError("Mercury no mostro el selector de compania.")

    result = select.evaluate(
        """
        (select, companyName) => {
            const wanted = companyName.trim().toLocaleLowerCase();
            const options = Array.from(select.options).map((option) => ({
                value: option.value,
                text: option.textContent.trim(),
            }));
            const selectable = options.filter((option) => {
                const text = option.text.toLocaleLowerCase();
                return text && text !== "seleccione";
            });
            const match =
                selectable.find((option) => option.text.toLocaleLowerCase() === wanted) ||
                selectable.find((option) => option.text.toLocaleLowerCase().includes(wanted)) ||
                selectable.find((option) => wanted.includes(option.text.toLocaleLowerCase()));
            if (!match) {
                return {
                    ok: false,
                    selectedText: select.options[select.selectedIndex]?.textContent.trim() || "",
                    options: selectable.map((option) => option.text),
                };
            }
            select.value = match.value;
            select.dispatchEvent(new Event("input", { bubbles: true }));
            select.dispatchEvent(new Event("change", { bubbles: true }));
            return {
                ok: true,
                selectedText: match.text,
                options: selectable.map((option) => option.text),
            };
        }
        """,
        company_name,
    )
    if not result.get("ok"):
        options = ", ".join(result.get("options") or [])
        raise MercuryAutomationError(
            f"Mercury no encontro la compania '{company_name}' en el selector. Opciones disponibles: {options}"
        )

    page.wait_for_timeout(1_500)
    page.wait_for_load_state("domcontentloaded", timeout=10_000)
    logging.info("Compania Mercury seleccionada: %s", result.get("selectedText") or company_name)


def _wait_for_post_login_page(page) -> None:
    if _wait_for_any_text(page, _COMPANY_LABELS, timeout=20_000):
        return
    if _wait_for_any_text(page, _HUMAN_RESOURCES_LABELS, timeout=10_000):
        return
    raise MercuryAutomationError("Mercury no mostro la pagina inicial despues de iniciar sesion.")


def _open_human_resources(page) -> None:
    if _is_human_resources_page(page):
        return
    _click_tile_by_any_text(page, _HUMAN_RESOURCES_LABELS)
    if not _wait_for_human_resources_page(page, timeout=20_000):
        raise MercuryAutomationError("Mercury no abrio el modulo de Recursos Humanos.")
    _raise_if_not_authenticated(page)


def _raise_if_not_authenticated(page) -> None:
    try:
        if page.get_by_text("User is not authenticated", exact=False).is_visible(timeout=2_000):
            raise MercuryAutomationError(
                "Mercury no acepto la sesion para Recursos Humanos. "
                "Revise que la compania este seleccionada antes de abrir reportes."
            )
    except MercuryAutomationError:
        raise
    except Exception:
        return


def _open_report_generator(page, config: AppConfig) -> None:
    if _is_report_generator_page(page):
        return
    dashboard_url = page.url
    if _try_open_report_generator_from_menu(page):
        return

    page.goto(_report_generator_url(config.mercury_url), wait_until="domcontentloaded", timeout=60_000)
    try:
        _raise_if_not_authenticated(page)
    except MercuryAutomationError:
        page.goto(dashboard_url, wait_until="domcontentloaded", timeout=60_000)
        if _try_open_report_generator_from_menu(page):
            return
        raise
    if not _is_report_generator_page(page):
        raise MercuryAutomationError("Mercury no abrio el generador de reportes.")


def _is_report_generator_page(page) -> bool:
    if "generadorreportes" in page.url.casefold():
        return True
    try:
        body = page.locator("body").inner_text(timeout=1_000).casefold()
    except Exception:
        return False
    return "generador de reportes" in body


def _try_open_report_generator_from_menu(page) -> bool:
    if _click_any_visible_text(page, _REPORT_GENERATOR_LABELS, timeout=2_000):
        return _wait_for_report_generator_page(page, timeout=10_000)

    if _click_likely_reports_menu(page):
        if _click_any_visible_text(page, _REPORT_GENERATOR_LABELS, timeout=5_000):
            return _wait_for_report_generator_page(page, timeout=10_000)
    return False


def _wait_for_report_generator_page(page, timeout: int = 10_000) -> bool:
    elapsed = 0
    step = 500
    while elapsed <= timeout:
        if _is_report_generator_page(page):
            return True
        page.wait_for_timeout(step)
        elapsed += step
    return False


def _click_any_visible_text(page, texts: tuple[str, ...], timeout: int = 5_000) -> bool:
    for text in texts:
        try:
            locator = page.get_by_text(text, exact=False).first
            if locator.is_visible(timeout=timeout):
                locator.click(timeout=timeout)
                return True
        except Exception:
            continue
    return False


def _click_likely_reports_menu(page) -> bool:
    clicked = page.evaluate(
        """
        () => {
            const keywords = ["reporte", "report", "listado", "consulta"];
            const candidates = Array.from(document.querySelectorAll("a, button, li, div, span"));
            const visible = (element) => {
                const rect = element.getBoundingClientRect();
                const style = window.getComputedStyle(element);
                return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
            };
            for (const element of candidates) {
                if (!visible(element)) {
                    continue;
                }
                const text = [
                    element.innerText,
                    element.textContent,
                    element.getAttribute("title"),
                    element.getAttribute("aria-label"),
                    element.getAttribute("href"),
                    element.className,
                    element.id,
                ].join(" ").toLocaleLowerCase();
                if (keywords.some((keyword) => text.includes(keyword))) {
                    element.click();
                    return true;
                }
            }
            const icons = candidates
                .filter((element) => visible(element))
                .map((element) => ({ element, rect: element.getBoundingClientRect() }))
                .filter((item) => item.rect.left < 90 && item.rect.top > 120 && item.rect.width <= 90 && item.rect.height <= 90)
                .sort((a, b) => b.rect.top - a.rect.top);
            if (icons.length) {
                icons[0].element.click();
                return true;
            }
            return false;
        }
        """
    )
    if clicked:
        page.wait_for_timeout(1_000)
    return bool(clicked)


def _report_generator_url(login_url: str) -> str:
    parsed = urlparse(login_url)
    if not parsed.scheme or not parsed.netloc:
        raise MercuryAutomationError("La direccion de Mercury no parece ser una URL valida.")
    return f"{parsed.scheme}://{parsed.netloc}/Mercury.RRHH/GeneradorReportes.aspx"


def _click_text(page, text: str) -> None:
    page.get_by_text(text, exact=True).click(timeout=20_000)


def _wait_for_any_text(page, texts: tuple[str, ...], timeout: int) -> bool:
    step = 500
    elapsed = 0
    while elapsed <= timeout:
        for text in texts:
            try:
                if page.get_by_text(text, exact=False).first.is_visible(timeout=500):
                    return True
            except Exception:
                continue
        elapsed += step
    return False


def _click_tile_by_any_text(page, texts: tuple[str, ...]) -> None:
    if _is_human_resources_page(page):
        return
    last_error: Exception | None = None
    for text in texts:
        try:
            _click_tile_by_text(page, text)
            return
        except Exception as exc:
            if "Mercury.RRHH" in page.url:
                return
            last_error = exc
    options = ", ".join(texts)
    raise MercuryAutomationError(f"Mercury no abrio ninguna de estas opciones: {options}") from last_error


def _is_human_resources_page(page) -> bool:
    if "mercury.rrhh" in page.url.casefold():
        return True
    try:
        body = page.locator("body").inner_text(timeout=1_000).casefold()
    except Exception:
        return False
    return any(marker.casefold() in body for marker in _POST_HR_MARKERS)


def _wait_for_human_resources_page(page, timeout: int = 20_000) -> bool:
    elapsed = 0
    step = 500
    while elapsed <= timeout:
        if _is_human_resources_page(page):
            return True
        page.wait_for_timeout(step)
        elapsed += step
    return False


def _click_selector_or_text(page, selector: str, text: str) -> None:
    locator = page.locator(selector).first
    try:
        if locator.count():
            locator.click(timeout=20_000)
            return
    except Exception:
        logging.info("No se pudo usar selector Mercury %s; se intenta por texto.", selector)
    page.get_by_role("button", name=text).click(timeout=20_000)


def _click_clickable_parent(page, text: str) -> None:
    locator = page.get_by_text(text, exact=True).first
    locator.wait_for(timeout=20_000)
    handle = locator.element_handle()
    if not handle:
        raise MercuryAutomationError(f"No se encontro la opcion {text}.")
    clicked = handle.evaluate(
        """
        (node) => {
            let element = node;
            for (let i = 0; i < 10 && element; i += 1) {
                const style = window.getComputedStyle(element);
                const tag = element.tagName.toLowerCase();
                if (
                    tag === "a" ||
                    tag === "button" ||
                    element.onclick ||
                    element.getAttribute("role") === "button" ||
                    style.cursor === "pointer"
                ) {
                    element.click();
                    return true;
                }
                element = element.parentElement;
            }
            node.click();
            return false;
        }
        """
    )
    if not clicked:
        logging.info("Se hizo click directamente sobre el texto %s.", text)


def _click_tile_by_text(page, text: str) -> None:
    locator = page.get_by_text(text, exact=False).first
    locator.wait_for(timeout=20_000)

    click_points = locator.evaluate(
        """
        (node) => {
            const points = [];
            let element = node;
            for (let i = 0; i < 12 && element; i += 1) {
                const rect = element.getBoundingClientRect();
                if (rect.width >= 180 && rect.height >= 120) {
                    points.push({
                        x: rect.left + rect.width / 2,
                        y: rect.top + rect.height / 2,
                    });
                }
                element = element.parentElement;
            }
            const label = node.getBoundingClientRect();
            points.push({
                x: label.left + label.width / 2,
                y: Math.max(10, label.top - 120),
            });
            points.push({
                x: label.left + label.width / 2,
                y: label.top + label.height / 2,
            });
            return points;
        }
        """
    )

    for point in click_points:
        page.mouse.click(point["x"], point["y"])
        try:
            page.wait_for_function(
                """
                (markers) => {
                    const text = document.body.innerText || "";
                    return location.href.includes("Mercury.RRHH") || markers.some((marker) => text.includes(marker));
                }
                """,
                list(_POST_HR_MARKERS),
                timeout=4_000,
            )
            return
        except Exception:
            continue

    raise MercuryAutomationError(f"Mercury no abrio la opcion {text}.")


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
    executable_path = _find_browser_executable()
    if executable_path:
        return {"executable_path": str(executable_path)}
    return {}


def _find_browser_executable() -> Path | None:
    return _find_playwright_chromium() or _find_installed_browser()


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


def _find_installed_browser() -> Path | None:
    env_candidates = [
        ("PROGRAMFILES", "Microsoft/Edge/Application/msedge.exe"),
        ("PROGRAMFILES(X86)", "Microsoft/Edge/Application/msedge.exe"),
        ("LOCALAPPDATA", "Microsoft/Edge/Application/msedge.exe"),
        ("PROGRAMFILES", "Google/Chrome/Application/chrome.exe"),
        ("PROGRAMFILES(X86)", "Google/Chrome/Application/chrome.exe"),
        ("LOCALAPPDATA", "Google/Chrome/Application/chrome.exe"),
    ]
    for env_name, relative_path in env_candidates:
        base = os.getenv(env_name)
        if not base:
            continue
        candidate = Path(base) / Path(relative_path)
        if candidate.exists():
            return candidate
    return None
