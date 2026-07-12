from __future__ import annotations

import logging
import os
import re
import shutil
import tempfile
import time
from datetime import date
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import quote, urljoin, urlparse

from .config import AppConfig
from .importer import EXPORT_NAMES
from .spreadsheet import read_employees


@dataclass
class MercuryRunResult:
    success: bool
    message: str
    downloaded_file: str = ""
    downloaded_files: list[str] = field(default_factory=list)
    companies_without_download: list[str] = field(default_factory=list)
    photo_files: dict[str, dict[str, str]] = field(default_factory=dict)
    photo_temp_dir: str = ""


class MercuryAutomationError(RuntimeError):
    pass


def run_mercury_login_test(config: AppConfig, password: str) -> MercuryRunResult:
    return run_mercury_export(config, password, download_report=False)


def run_mercury_export(
    config: AppConfig, password: str, download_report: bool = True, report_date: date | None = None
) -> MercuryRunResult:
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
            photo_files: dict[str, dict[str, str]] = {}
            photo_temp_dir = ""
            for company in _configured_companies(config):
                context = browser.new_context(accept_downloads=True)
                page = context.new_page()
                try:
                    _open_authenticated_session(page, config, password)
                    downloaded = _download_existing_report(page, config, company, report_date=report_date)
                    if downloaded:
                        downloaded_files.append(str(downloaded))
                        if not photo_temp_dir:
                            photo_temp_dir = tempfile.mkdtemp(prefix="duendecito-fotos-", dir=downloads)
                        photo_files[str(downloaded)] = _download_employee_photos(
                            page, config, downloaded, company, Path(photo_temp_dir)
                        )
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
                photo_files=photo_files,
                photo_temp_dir=photo_temp_dir,
            )
        except PlaywrightTimeoutError as exc:
            raise MercuryAutomationError(f"Mercury tardo demasiado en responder: {exc}") from exc
        except Exception:
            if "photo_temp_dir" in locals() and photo_temp_dir:
                shutil.rmtree(photo_temp_dir, ignore_errors=True)
            raise
        finally:
            browser.close()


def _download_employee_photos(page, config: AppConfig, spreadsheet: Path, company_name: str, photo_dir: Path) -> dict[str, str]:
    photos: dict[str, str] = {}
    try:
        employees = read_employees(spreadsheet)
    except Exception as exc:
        logging.warning("No se pudieron leer los empleados para buscar fotos: %s", exc)
        return photos

    for employee in employees:
        number = str(employee.get("Numero", "")).strip()
        if not number:
            continue
        try:
            page.goto(_employee_page_url(config.mercury_url, number), wait_until="domcontentloaded", timeout=30_000)
            image = page.locator("img[id$='Image1']").first
            image.wait_for(state="attached", timeout=5_000)
            image_info = image.evaluate(
                "element => ({src: element.getAttribute('src') || '', width: element.naturalWidth || 0, height: element.naturalHeight || 0})"
            )
            if (
                not image_info["src"]
                or image_info["width"] < 40
                or image_info["height"] < 40
            ):
                logging.info("Mercury no tiene foto para el empleado %s (%s).", number, company_name)
                continue
            target = photo_dir / f"{_company_file_label(company_name)}_{number}.png"
            image.screenshot(path=str(target), timeout=10_000)
            photos[number] = str(target)
            logging.info("Foto descargada para el empleado %s (%s).", number, company_name)
        except Exception as exc:
            logging.warning("No se pudo obtener la foto del empleado %s (%s): %s", number, company_name, exc)
    return photos


def _employee_page_url(base_url: str, employee_number: str) -> str:
    parsed = urlparse(base_url)
    if not parsed.scheme or not parsed.netloc:
        raise MercuryAutomationError("La direccion de Mercury no parece ser una URL valida.")
    return f"{parsed.scheme}://{parsed.netloc}/Mercury.RRHH/Empleados.aspx?Codigo={quote(employee_number, safe='')}"


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


def _download_existing_report(
    page, config: AppConfig, company_name: str, report_date: date | None = None
) -> Path | None:
    _select_company(page, company_name)
    _open_human_resources(page)
    _open_report_generator(page, config)
    _open_saved_report(page, config.mercury_report_name)
    if report_date:
        _apply_report_date_filter(page, report_date)
    result_panel = page.locator("#ctl00_MainContent_tabGeneral_ctl05").first
    result_panel_before = result_panel.inner_html() if result_panel.count() else ""
    process_button = page.locator("#ctl00_MainContent_lnkProcesar").first
    if process_button.count():
        process_button.wait_for(state="visible", timeout=8_000)
        if not process_button.is_enabled():
            raise MercuryAutomationError("Mercury no habilito Ver reporte despues de aplicar los filtros.")
        click_started = time.monotonic()
        logging.info("Mercury: haciendo click en Ver reporte.")
        process_button.click(timeout=8_000, no_wait_after=True)
        logging.info("Mercury: click en Ver reporte termino en %.2fs.", time.monotonic() - click_started)
    else:
        _click_text(page, "Ver reporte")
    result_started = time.monotonic()
    result_is_empty = _wait_for_report_result(
        page,
        timeout=10_000,
        result_panel=result_panel,
        result_panel_before=result_panel_before,
    )
    logging.info("Mercury: resultado del reporte detectado en %.2fs (sin registros: %s).", time.monotonic() - result_started, result_is_empty)
    if result_is_empty:
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


def _open_saved_report(page, report_name: str) -> None:
    open_button = page.locator("#ctl00_MainContent_btnAbrirReporte").first
    if open_button.count():
        open_button.click(timeout=8_000)
    else:
        _click_text(page, "Abrir reporte existente")
    report_list = page.locator("#ctl00_MainContent_lstReportes").first
    if report_list.count():
        report_list.select_option(report_name)
    else:
        _click_text(page, report_name)
    _click_selector_or_text(page, "#ctl00_MainContent_cmdAceptarReporte", "Aceptar")
    page.wait_for_load_state("domcontentloaded", timeout=15_000)
    try:
        page.wait_for_function(
            """(name) => document.querySelector('#ctl00_MainContent_txtNombreReporte')?.value === name""",
            report_name,
            timeout=10_000,
        )
    except Exception:
        logging.info("Mercury no confirmo el nombre del reporte mediante el campo directo; se continua.")
    logging.info("Reporte Mercury abierto: %s", report_name)


def _select_saved_report(page, report_name: str, double_click: bool = False) -> bool:
    return bool(
        page.evaluate(
            """
            ({ reportName, doubleClick }) => {
                const normalize = (value) => (value || "")
                    .normalize("NFD")
                    .replace(/[\\u0300-\\u036f]/g, "")
                    .replace(/\\s+/g, " ")
                    .trim()
                    .toLocaleLowerCase();
                const wanted = normalize(reportName);
                const visible = (element) => {
                    const rect = element.getBoundingClientRect();
                    const style = window.getComputedStyle(element);
                    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
                };
                const dispatch = (element, type) => {
                    element.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
                };

                for (const select of Array.from(document.querySelectorAll("select")).filter(visible)) {
                    const option = Array.from(select.options).find((item) => normalize(item.textContent) === wanted);
                    if (option) {
                        select.value = option.value;
                        select.dispatchEvent(new Event("input", { bubbles: true }));
                        select.dispatchEvent(new Event("change", { bubbles: true }));
                        return true;
                    }
                }

                const candidates = Array.from(document.querySelectorAll("option, li, tr, td, span, div, a"))
                    .filter(visible)
                    .map((element) => ({ element, text: normalize(element.innerText || element.textContent) }))
                    .filter((item) => item.text === wanted)
                    .sort((a, b) => {
                        const ar = a.element.getBoundingClientRect();
                        const br = b.element.getBoundingClientRect();
                        return (ar.width * ar.height) - (br.width * br.height);
                    });
                if (!candidates.length) {
                    return false;
                }
                const element = candidates[0].element;
                const target = element.closest("option, li, tr, td, a, [onclick], [role='option'], [role='button']") || element;
                target.focus?.();
                dispatch(target, "mousedown");
                dispatch(target, "mouseup");
                target.click();
                if (doubleClick) {
                    dispatch(target, "dblclick");
                }
                return true;
            }
            """,
            {"reportName": report_name, "doubleClick": double_click},
        )
    )


def _wait_for_saved_report_dialog_closed(page, timeout: int = 8_000) -> bool:
    elapsed = 0
    step = 500
    while elapsed <= timeout:
        if not _is_saved_report_dialog_open(page):
            return True
        page.wait_for_timeout(step)
        elapsed += step
    return False


def _is_saved_report_dialog_open(page) -> bool:
    try:
        return bool(
            page.evaluate(
                """
                () => {
                    const visible = (element) => {
                        const rect = element.getBoundingClientRect();
                        const style = window.getComputedStyle(element);
                        return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
                    };
                    return Array.from(document.querySelectorAll("*"))
                        .filter(visible)
                        .some((element) => (element.innerText || element.textContent || "").trim() === "Reportes");
                }
                """
            )
        )
    except Exception:
        return False


def _format_mercury_date(value: date) -> str:
    return value.strftime("%d.%m.%Y")


def _apply_report_date_filter(page, report_date: date) -> None:
    filter_tab = page.locator("#__tab_ctl00_MainContent_tabGeneral_ctl03").first
    if filter_tab.count():
        filter_tab.click(timeout=5_000)
        try:
            page.wait_for_function(
                """() => {
                    const state = document.querySelector('#ctl00_MainContent_tabGeneral_ClientState');
                    return state && JSON.parse(state.value || '{}').ActiveTabIndex === 3;
                }""",
                timeout=8_000,
            )
        except Exception:
            logging.info("Mercury no confirmo el indice activo de Filtros; se valida el panel directamente.")
    elif not _click_any_visible_text(page, ("Filtros", "Filters"), timeout=3_000):
        raise MercuryAutomationError("Mercury no mostro la pestana Filtros para ajustar la Fecha Ingreso.")
    display_date = _format_mercury_date(report_date)
    try:
        rows = (
            ("1", "Desde", "11", "chkSel1"),
            ("2", "Hasta", "12", "chkSel2"),
        )
        for row_number, operation_name, input_number, checkbox_name in rows:
            field = page.locator(f"select[id$='cboCampo{row_number}']").first
            operation = page.locator(f"select[id$='cboOperador{row_number}']").first
            value = page.locator(f"input[id$='txtCriterio{input_number}']").first
            checkbox = page.locator(f"input[id$='{checkbox_name}']").first

            field.wait_for(state="visible", timeout=12_000)
            operation.wait_for(state="visible", timeout=12_000)
            value.wait_for(state="visible", timeout=12_000)
            if not checkbox.is_checked():
                checkbox.check()
                page.wait_for_timeout(1_000)

            if field.input_value() != "Fecha_Ingreso":
                field.select_option("Fecha_Ingreso")
            if operation.input_value() != operation_name:
                operation.select_option(operation_name)
                page.wait_for_timeout(1_000)
            value.click()
            value.fill(display_date)
            value.press("Tab")
            if value.input_value() != display_date:
                raise MercuryAutomationError(f"Mercury no mantuvo la fecha {display_date} en {value.get_attribute('id')}.")
        page.wait_for_timeout(300)
        logging.info("Filtro Fecha Ingreso Mercury aplicado: %s desde/hasta.", display_date)
        return
    except Exception as exc:
        raise MercuryAutomationError(f"Mercury no pudo ajustar Fecha Ingreso en Filtros: {exc}") from exc

    # Fallback retained for older Mercury layouts without the known control IDs.
    filter_frame = _wait_for_filter_controls(page, timeout=12_000) or page.main_frame
    result = filter_frame.evaluate(
        """
        ({ displayDate }) => {
            const visible = (element) => {
                const rect = element.getBoundingClientRect();
                const style = window.getComputedStyle(element);
                return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
            };
            const normalize = (value) => (value || "")
                .normalize("NFD")
                .replace(/[\\u0300-\\u036f]/g, "")
                .replace(/\\s+/g, " ")
                .trim()
                .toLocaleLowerCase();
            const selectedText = (select) => normalize(select.options[select.selectedIndex]?.textContent || select.value);
            const sameRow = (a, b) => Math.abs((a.top + a.height / 2) - (b.top + b.height / 2)) < 28;
            const controls = Array.from(document.querySelectorAll("select, input, textarea"))
                .filter(visible)
                .map((element) => ({ element, rect: element.getBoundingClientRect() }));
            const selectControls = controls
                .filter((item) => item.element.tagName.toLowerCase() === "select");
            const fieldSelects = selectControls
                .filter((item) => Array.from(item.element.options)
                    .some((option) => normalize(option.textContent) === "fecha ingreso"));
            const inputs = controls
                .filter((item) => ["input", "textarea"].includes(item.element.tagName.toLowerCase()))
                .filter((item) => !["button", "submit", "reset", "hidden", "password", "checkbox", "radio"].includes((item.element.type || "").toLocaleLowerCase()));
            const checkboxes = controls
                .filter((item) => item.element.tagName.toLowerCase() === "input")
                .filter((item) => ["checkbox", "radio"].includes((item.element.type || "").toLocaleLowerCase()));

            const fillRow = (operationName, rowIndex) => {
                const field = fieldSelects
                    .slice()
                    .sort((a, b) => a.rect.top - b.rect.top)[rowIndex];
                if (!field) {
                    return { operation: operationName, ok: false, reason: "Fecha Ingreso no encontrada" };
                }
                const fieldOption = Array.from(field.element.options)
                    .find((option) => normalize(option.textContent) === "fecha ingreso");
                if (fieldOption && selectedText(field.element) !== "fecha ingreso") {
                    field.element.value = fieldOption.value;
                    field.element.dispatchEvent(new Event("input", { bubbles: true }));
                    field.element.dispatchEvent(new Event("change", { bubbles: true }));
                }
                const operation = selectControls
                    .filter((item) => sameRow(item.rect, field.rect))
                    .filter((item) => item.rect.left > field.rect.right)
                    .sort((a, b) => a.rect.left - b.rect.left)[0];
                if (!operation) {
                    return { operation: operationName, ok: false, reason: "operacion no encontrada" };
                }

                const wantedOption = Array.from(operation.element.options)
                    .find((option) => normalize(option.textContent) === operationName);
                if (wantedOption && selectedText(operation.element) !== operationName) {
                    operation.element.value = wantedOption.value;
                    operation.element.dispatchEvent(new Event("input", { bubbles: true }));
                    operation.element.dispatchEvent(new Event("change", { bubbles: true }));
                }
                const input = inputs
                    .filter((item) => sameRow(item.rect, operation.rect))
                    .filter((item) => item.rect.left > operation.rect.right)
                    .sort((a, b) => a.rect.left - b.rect.left)[0];
                if (!input) {
                    return { operation: operationName, ok: false, reason: "campo fecha no encontrado" };
                }
                const checkbox = checkboxes
                    .filter((item) => sameRow(item.rect, operation.rect))
                    .filter((item) => item.rect.left < field.rect.left)
                    .sort((a, b) => b.rect.left - a.rect.left)[0];
                if (checkbox && !checkbox.element.checked) {
                    checkbox.element.click();
                    checkbox.element.dispatchEvent(new Event("change", { bubbles: true }));
                }
                const setter = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(input.element), "value")?.set;
                if (setter) {
                    setter.call(input.element, displayDate);
                } else {
                    input.element.value = displayDate;
                }
                input.element.dispatchEvent(new Event("input", { bubbles: true }));
                input.element.dispatchEvent(new Event("change", { bubbles: true }));
                input.element.dispatchEvent(new Event("blur", { bubbles: true }));
                return { operation: operationName, ok: true };
            };
            return [fillRow("desde", 0), fillRow("hasta", 1)];
        }
        """,
        {"displayDate": display_date},
    )
    failed = [item for item in result if not item.get("ok")]
    if failed:
        details = "; ".join(f"{item.get('operation')}: {item.get('reason')}" for item in failed)
        raise MercuryAutomationError(f"Mercury no pudo ajustar Fecha Ingreso en Filtros ({details}).")
    logging.info("Filtro Fecha Ingreso Mercury aplicado: %s desde/hasta.", display_date)


def _wait_for_filter_controls(page, timeout: int = 12_000):
    elapsed = 0
    step = 500
    while elapsed <= timeout:
        for frame in page.frames:
            try:
                ready = frame.evaluate(
                    """
            () => {
                const normalize = (value) => (value || "")
                    .normalize("NFD")
                    .replace(/[\\u0300-\\u036f]/g, "")
                    .replace(/\\s+/g, " ")
                    .trim()
                    .toLocaleLowerCase();
                const visible = (element) => {
                    const rect = element.getBoundingClientRect();
                    const style = window.getComputedStyle(element);
                    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
                };
                return Array.from(document.querySelectorAll("select"))
                    .filter(visible)
                    .some((select) => Array.from(select.options)
                        .some((option) => normalize(option.textContent) === "fecha ingreso"));
            }
                    """
                )
            except Exception:
                continue
            if ready:
                return frame
        page.wait_for_timeout(step)
        elapsed += step
    logging.info("Mercury no mostro los controles Fecha Ingreso despues de esperar %sms.", timeout)
    return None


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
_POST_HR_MARKERS = (
    "Principal",
    "Generador de Reportes",
    "Cantidad de empleados activos",
    "Presupuestos pendientes",
)
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


def _wait_for_report_result(
    page, timeout: int = 10_000, result_panel=None, result_panel_before: str = ""
) -> bool:
    """Return True for no records, or continue as soon as Exportar is ready."""
    elapsed = 0
    step = 250
    while elapsed <= timeout:
        if _page_has_no_records(page):
            return True
        if _page_has_report_data(result_panel):
            return False
        try:
            export_button = page.locator("#ctl00_MainContent_lnkExportar").first
            if (
                not result_panel
                and export_button.count()
                and export_button.is_visible(timeout=100)
                and export_button.is_enabled()
            ):
                return False
        except Exception:
            pass
        try:
            export_text = page.get_by_text("Exportar", exact=True).first
            if not result_panel and export_text.is_visible(timeout=100) and export_text.is_enabled():
                return False
        except Exception:
            pass
        page.wait_for_timeout(step)
        elapsed += step
    return False


def _page_has_report_data(result_panel) -> bool:
    if not result_panel or not result_panel.count():
        return False
    try:
        tables = result_panel.locator("table")
        for index in range(tables.count()):
            table = tables.nth(index)
            if table.locator("tbody tr td").count() > 0:
                return True
    except Exception:
        return False
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
    direct_button = page.locator("#ContentPlaceHolder1_cmdRecursosHumanos").first
    try:
        if direct_button.count() and direct_button.is_visible(timeout=2_000):
            direct_button.click(timeout=8_000)
            logging.info("Mercury: modulo Recursos Humanos abierto mediante selector directo.")
        else:
            _click_tile_by_any_text(page, _HUMAN_RESOURCES_LABELS)
    except Exception:
        logging.info("No se pudo usar el selector directo de Recursos Humanos; se usa la busqueda por texto.")
        _click_tile_by_any_text(page, _HUMAN_RESOURCES_LABELS)
    if not _wait_for_human_resources_page(page, timeout=15_000):
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
    if _try_open_report_generator_from_menu(page):
        return
    if _is_human_resources_page(page):
        _goto_mercury_page(page, _report_generator_url(page.url or config.mercury_url))
        _raise_if_not_authenticated(page)
        if _is_report_generator_page(page):
            return
    raise MercuryAutomationError("Mercury no abrio el generador de reportes desde el menu.")


def _is_report_generator_page(page) -> bool:
    if "generadorreportes" in page.url.casefold():
        return True
    try:
        body = page.locator("body").inner_text(timeout=1_000).casefold()
    except Exception:
        return False
    return "generador de reportes" in body


def _try_open_report_generator_from_menu(page) -> bool:
    direct_link = page.locator("#ctl00_rh_cap24").first
    try:
        if direct_link.count():
            href = direct_link.get_attribute("href")
            if href and not href.casefold().startswith("javascript:"):
                _goto_mercury_page(page, urljoin(page.url, href))
                logging.info("Mercury: Generador de reportes abierto mediante href directo.")
                if _wait_for_report_generator_page(page, timeout=8_000):
                    return True
            if not direct_link.is_visible(timeout=1_000):
                _open_left_navigation(page)
            direct_link.click(timeout=8_000, force=True)
            logging.info("Mercury: Generador de reportes abierto mediante selector directo.")
            if _wait_for_report_generator_page(page, timeout=12_000):
                return True
    except Exception:
        logging.info("No se pudo usar el selector directo de Generador de reportes; se usa la busqueda por menu.")

    if _click_any_visible_text(page, _REPORT_GENERATOR_LABELS, timeout=2_000):
        return _wait_for_report_generator_page(page, timeout=10_000)

    _open_left_navigation(page)
    if _click_any_visible_text(page, _REPORT_GENERATOR_LABELS, timeout=2_000):
        return _wait_for_report_generator_page(page, timeout=10_000)

    if _click_likely_reports_menu(page):
        if _click_any_visible_text(page, _REPORT_GENERATOR_LABELS, timeout=5_000):
            return _wait_for_report_generator_page(page, timeout=10_000)
    if _click_sidebar_until_report_generator(page):
        return True
    return False


def _open_left_navigation(page) -> None:
    clicked = page.evaluate(
        """
        () => {
            const candidates = Array.from(document.querySelectorAll("*"));
            const visible = (element) => {
                const rect = element.getBoundingClientRect();
                const style = window.getComputedStyle(element);
                return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
            };
            const menu = candidates
                .filter((element) => visible(element))
                .map((element) => ({ element, rect: element.getBoundingClientRect() }))
                .find((item) => item.rect.left >= 0 && item.rect.left < 100 && item.rect.top >= 40 && item.rect.top < 140);
            if (!menu) {
                return false;
            }
            const target = menu.element.closest("a, button, li, [onclick], [role='button']") || menu.element;
            target.click();
            return true;
        }
        """
    )
    if clicked:
        page.wait_for_timeout(1_000)


def _wait_for_report_generator_page(page, timeout: int = 10_000) -> bool:
    elapsed = 0
    step = 500
    while elapsed <= timeout:
        if _is_report_generator_page(page):
            return True
        page.wait_for_timeout(step)
        elapsed += step
    return False


def _goto_mercury_page(page, url: str) -> None:
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=60_000)
    except Exception as exc:
        if "net::ERR_ABORTED" not in str(exc):
            raise
        logging.info("Mercury cancelo una navegacion directa, se valida la pagina actual: %s", url)
        page.wait_for_timeout(2_000)


def _click_any_visible_text(page, texts: tuple[str, ...], timeout: int = 5_000) -> bool:
    elapsed = 0
    step = 500
    while elapsed <= timeout:
        clicked = page.evaluate(
            """
            (texts) => {
                const normalize = (value) => (value || "")
                    .normalize("NFD")
                    .replace(/[\\u0300-\\u036f]/g, "")
                    .replace(/\\s+/g, " ")
                    .trim()
                    .toLocaleLowerCase();
                const wanted = texts.map((text) => normalize(text));
                const visible = (element) => {
                    const rect = element.getBoundingClientRect();
                    const style = window.getComputedStyle(element);
                    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
                };
                const candidates = Array.from(document.querySelectorAll("*"))
                    .filter((element) => visible(element))
                    .map((element) => {
                        const rect = element.getBoundingClientRect();
                        const text = normalize([
                            element.innerText,
                            element.textContent,
                            element.getAttribute("title"),
                            element.getAttribute("aria-label"),
                            element.getAttribute("href"),
                        ].join(" "));
                        return { element, rect, text };
                    })
                    .filter((item) => wanted.some((text) => item.text.includes(text)))
                    .filter((item) => item.rect.width < 800 && item.rect.height < 500)
                    .sort((a, b) => (a.rect.width * a.rect.height) - (b.rect.width * b.rect.height));
                if (!candidates.length) {
                    return false;
                }
                const element = candidates[0].element;
                const target = element.closest("a, button, li, [onclick], [role='button']") || element;
                target.click();
                return true;
            }
            """,
            list(texts),
        )
        if clicked:
            return True
        page.wait_for_timeout(step)
        elapsed += step
    return False


def _click_likely_reports_menu(page) -> bool:
    clicked = page.evaluate(
        """
        () => {
            const keywords = ["reporte", "report", "listado", "consulta"];
            const candidates = Array.from(document.querySelectorAll("*"));
            const visible = (element) => {
                const rect = element.getBoundingClientRect();
                const style = window.getComputedStyle(element);
                return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
            };
            const matches = candidates
                .filter((element) => visible(element))
                .map((element) => {
                    const rect = element.getBoundingClientRect();
                    const text = [
                        element.innerText,
                        element.textContent,
                        element.getAttribute("title"),
                        element.getAttribute("aria-label"),
                        element.getAttribute("href"),
                        element.className,
                        element.id,
                    ].join(" ").toLocaleLowerCase();
                    return { element, rect, text };
                })
                .filter((item) => keywords.some((keyword) => item.text.includes(keyword)))
                .filter((item) => item.rect.width < 800 && item.rect.height < 500)
                .sort((a, b) => (a.rect.width * a.rect.height) - (b.rect.width * b.rect.height));
            if (matches.length) {
                const target = matches[0].element.closest("a, button, li, [onclick], [role='button']") || matches[0].element;
                target.click();
                return true;
            }
            const icons = candidates
                .filter((element) => visible(element))
                .map((element) => ({ element, rect: element.getBoundingClientRect() }))
                .filter((item) => item.rect.left >= 0 && item.rect.left < 110 && item.rect.top > 120 && item.rect.top < 650)
                .filter((item) => item.rect.width <= 110 && item.rect.height <= 110)
                .sort((a, b) => b.rect.top - a.rect.top);
            if (icons.length) {
                const target = icons[0].element.closest("a, button, li, [onclick], [role='button']") || icons[0].element;
                target.click();
                return true;
            }
            return false;
        }
        """
    )
    if clicked:
        page.wait_for_timeout(1_000)
    return bool(clicked)


def _click_sidebar_until_report_generator(page) -> bool:
    points = page.evaluate(
        """
        () => {
            const candidates = Array.from(document.querySelectorAll("*"));
            const visible = (element) => {
                const rect = element.getBoundingClientRect();
                const style = window.getComputedStyle(element);
                return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
            };
            const seen = new Set();
            return candidates
                .filter((element) => visible(element))
                .map((element) => {
                    const rect = element.getBoundingClientRect();
                    const clickable = element.closest("a, button, li, [onclick], [role='button']") || element;
                    const clickableRect = clickable.getBoundingClientRect();
                    return {
                        x: Math.round(clickableRect.left + clickableRect.width / 2),
                        y: Math.round(clickableRect.top + clickableRect.height / 2),
                        left: clickableRect.left,
                        top: clickableRect.top,
                        width: clickableRect.width,
                        height: clickableRect.height,
                    };
                })
                .filter((point) => point.left >= 0 && point.left < 110 && point.top > 120 && point.top < 650)
                .filter((point) => point.width <= 110 && point.height <= 110)
                .filter((point) => {
                    const key = `${point.x}:${point.y}`;
                    if (seen.has(key)) {
                        return false;
                    }
                    seen.add(key);
                    return true;
                })
                .sort((a, b) => b.y - a.y);
        }
        """
    )

    for point in points:
        page.mouse.click(point["x"], point["y"])
        page.wait_for_timeout(1_500)
        if _click_any_visible_text(page, _REPORT_GENERATOR_LABELS, timeout=1_500):
            return _wait_for_report_generator_page(page, timeout=10_000)
        if _is_report_generator_page(page):
            return True
    return False


def _report_generator_url(login_url: str) -> str:
    parsed = urlparse(login_url)
    if not parsed.scheme or not parsed.netloc:
        raise MercuryAutomationError("La direccion de Mercury no parece ser una URL valida.")
    return f"{parsed.scheme}://{parsed.netloc}/Mercury.RRHH/GeneradorReportes.aspx"


def _click_text(page, text: str) -> None:
    page.get_by_text(text, exact=True).click(timeout=20_000)


def _click_enabled_text(page, text: str, timeout: int = 20_000) -> bool:
    elapsed = 0
    step = 500
    while elapsed <= timeout:
        locator = page.get_by_text(text, exact=True).first
        try:
            if locator.count() and locator.is_visible(timeout=500) and locator.is_enabled(timeout=500):
                locator.click(timeout=2_000)
                return True
        except Exception:
            pass
        page.wait_for_timeout(step)
        elapsed += step
    return False


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
            if _wait_for_human_resources_page(page, timeout=8_000):
                return
            last_error = exc
    if _click_first_management_tile(page):
        return
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
    click_points = page.evaluate(
        """
        (wantedText) => {
            const normalize = (value) => (value || "")
                .normalize("NFD")
                .replace(/[\\u0300-\\u036f]/g, "")
                .replace(/\\s+/g, " ")
                .trim()
                .toLocaleLowerCase();
            const wanted = normalize(wantedText);
            const wantedTokens = wanted.split(" ").filter(Boolean);
            const matchesWanted = (text) => text.includes(wanted) || wantedTokens.every((token) => text.includes(token));
            const visible = (element) => {
                const rect = element.getBoundingClientRect();
                const style = window.getComputedStyle(element);
                return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
            };
            const matches = Array.from(document.querySelectorAll("*"))
                .filter((element) => visible(element))
                .map((element) => ({ element, rect: element.getBoundingClientRect(), text: normalize(element.innerText || element.textContent) }))
                .filter((item) => matchesWanted(item.text))
                .filter((item) => item.rect.width < 700 && item.rect.height < 400)
                .sort((a, b) => (a.rect.width * a.rect.height) - (b.rect.width * b.rect.height));
            if (!matches.length) {
                return [];
            }
            let node = matches[0].element;
            const points = [];
            const seen = new Set();
            const addPoint = (rect) => {
                const x = Math.round(rect.left + rect.width / 2);
                const y = Math.round(rect.top + rect.height / 2);
                const key = `${x}:${y}`;
                if (!seen.has(key)) {
                    seen.add(key);
                    points.push({ x, y });
                }
            };
            let element = node;
            for (let i = 0; i < 12 && element; i += 1) {
                const rect = element.getBoundingClientRect();
                const style = window.getComputedStyle(element);
                const tag = element.tagName.toLowerCase();
                if (
                    tag === "a" ||
                    tag === "button" ||
                    element.onclick ||
                    element.getAttribute("role") === "button" ||
                    style.cursor === "pointer"
                ) {
                    addPoint(rect);
                }
                if (rect.width >= 180 && rect.height >= 120) {
                    addPoint(rect);
                }
                element = element.parentElement;
            }
            const label = node.getBoundingClientRect();
            addPoint({
                left: label.left,
                top: Math.max(10, label.top - 120),
                width: label.width,
                height: 1,
            });
            addPoint(label);
            return points;
        }
        """,
        text,
    )
    if not click_points:
        raise MercuryAutomationError(f"No se encontro la opcion {text}.")

    for point in click_points:
        page.mouse.click(point["x"], point["y"])
        try:
            page.wait_for_function(
                """
                (markers) => {
                    const text = document.body.innerText || "";
                    return location.href.toLocaleLowerCase().includes("mercury.rrhh") || markers.some((marker) => text.includes(marker));
                }
                """,
                list(_POST_HR_MARKERS),
                timeout=8_000,
            )
            return
        except Exception:
            continue

    raise MercuryAutomationError(f"Mercury no abrio la opcion {text}.")


def _click_first_management_tile(page) -> bool:
    try:
        points = page.evaluate(
            """
            () => {
                const visible = (element) => {
                    const rect = element.getBoundingClientRect();
                    const style = window.getComputedStyle(element);
                    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
                };
                const candidates = Array.from(document.querySelectorAll("*"))
                    .filter((element) => visible(element))
                    .map((element) => {
                        const rect = element.getBoundingClientRect();
                        return {
                            x: Math.round(rect.left + rect.width / 2),
                            y: Math.round(rect.top + rect.height / 2),
                            left: rect.left,
                            top: rect.top,
                            width: rect.width,
                            height: rect.height,
                        };
                    })
                    .filter((item) => item.width >= 180 && item.width <= 520)
                    .filter((item) => item.height >= 180 && item.height <= 420)
                    .filter((item) => item.left >= 20 && item.left < 520)
                    .filter((item) => item.top >= 250 && item.top < 760)
                    .sort((a, b) => (a.top - b.top) || (a.left - b.left));
                return candidates.map((item) => ({ x: item.x, y: item.y }));
            }
            """
        )
    except Exception:
        return _wait_for_human_resources_page(page, timeout=8_000)

    for point in points:
        try:
            page.mouse.click(point["x"], point["y"])
        except Exception:
            if _wait_for_human_resources_page(page, timeout=8_000):
                return True
            continue
        try:
            page.wait_for_function(
                """
                (markers) => {
                    const text = document.body.innerText || "";
                    return location.href.toLocaleLowerCase().includes("mercury.rrhh") || markers.some((marker) => text.includes(marker));
                }
                """,
                list(_POST_HR_MARKERS),
                timeout=8_000,
            )
            return True
        except Exception:
            continue
    return False


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
