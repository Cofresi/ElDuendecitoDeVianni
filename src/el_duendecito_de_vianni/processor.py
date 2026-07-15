from __future__ import annotations

import logging
import shutil
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Mapping
from zoneinfo import ZoneInfo

from .config import AppConfig
from .importer import find_export, import_export
from .processed_store import ProcessedStore, file_sha256
from .spreadsheet import read_employees
from .templates import process_template_copy
from .utils import company_template_subfolder, employee_folder_name, safe_filename, sorted_templates, template_folder_for_employee
from .work_schedule import WorkScheduleLookup, add_work_schedule_sentence


DR_TZ = ZoneInfo("America/Santo_Domingo")


@dataclass
class RunReport:
    source_spreadsheet: str = ""
    imported_spreadsheet: str = ""
    output_folder: str = ""
    employees_processed: int = 0
    generated_documents: list[str] = field(default_factory=list)
    missing_placeholders: set[str] = field(default_factory=set)
    skipped_files: list[str] = field(default_factory=list)
    already_processed: bool = False
    message: str = ""

    @property
    def document_count(self) -> int:
        return len(self.generated_documents)


class DocumentProcessor:
    def __init__(self, config: AppConfig):
        self.config = config
        self.store = ProcessedStore(Path(config.imported_folder) / "processed_files.json")
        self.work_schedule_lookup = WorkScheduleLookup.from_file(config.work_schedule_lookup)

    def process_next_export(
        self, force: bool = False, delete_original: bool = True, run_date: date | None = None
    ) -> RunReport:
        source = find_export(self.config.downloads_folder)
        if not source:
            return RunReport(message="No se encontro un archivo de exportacion en Descargas.")
        return self.process_export_file(source, force=force, delete_original=delete_original, run_date=run_date)

    def process_export_file(
        self,
        source: str | Path,
        force: bool = False,
        delete_original: bool = True,
        run_date: date | None = None,
        photo_paths: Mapping[str, str] | None = None,
        workflow: str = "entradas",
    ) -> RunReport:
        source = Path(source)
        source_hash = file_sha256(source)
        if self.store.is_processed_hash(source_hash) and not force:
            logging.info("Archivo ya procesado: %s", source)
            return RunReport(source_spreadsheet=str(source), already_processed=True, message="Este archivo ya fue procesado.")
        imported = import_export(source, self.config.imported_folder, run_date=run_date, workflow=workflow)
        report = self.process_imported_file(
            imported.imported_path,
            run_date=run_date,
            photo_paths=photo_paths,
            workflow=workflow,
        )
        report.source_spreadsheet = str(source)
        report.imported_spreadsheet = str(imported.imported_path)
        self.store.add(source, imported.imported_path, imported.file_hash)
        if delete_original:
            source.unlink(missing_ok=True)
        return report

    def process_imported_file(
        self,
        spreadsheet_path: str | Path,
        run_date: date | None = None,
        photo_paths: Mapping[str, str] | None = None,
        workflow: str = "entradas",
    ) -> RunReport:
        spreadsheet = Path(spreadsheet_path)
        employees = read_employees(spreadsheet)
        workflow = "salidas" if workflow.casefold() == "salidas" else "entradas"
        date_text = (run_date or datetime.now(DR_TZ).date()).strftime("%d.%m.%Y")
        output_label = "nuevasSalidas" if workflow == "salidas" else "nuevasEntradas"
        final_output = Path(self.config.output_folder) / f"{output_label}_{date_text}" / _run_company_folder(employees)
        temp_output = final_output.with_name(final_output.name + "_tmp")
        if temp_output.exists():
            shutil.rmtree(temp_output)
        temp_output.mkdir(parents=True, exist_ok=True)

        report = RunReport(imported_spreadsheet=str(spreadsheet), output_folder=str(final_output))
        try:
            for employee in employees:
                if workflow == "entradas":
                    add_work_schedule_sentence(employee, self.work_schedule_lookup)
                employee_values = dict(employee)
                _add_employee_name_aliases(employee_values)
                employee_values["Foto"] = (photo_paths or {}).get(employee.get("Numero", ""), "")
                employee_dir = temp_output / employee_folder_name(employee)
                employee_dir.mkdir(parents=True, exist_ok=True)
                template_folder = template_folder_for_employee(self.config.template_folder, employee, workflow)
                templates = sorted_templates(template_folder)
                if workflow == "salidas":
                    templates = _salidas_templates_for_employee(templates, employee)
                logging.info(
                    "Usando plantillas de %s para %s",
                    template_folder,
                    employee.get("Nombre Empleado", "empleado sin nombre"),
                )
                for template in templates:
                    destination = employee_dir / safe_filename(template.name)
                    result = process_template_copy(
                        template,
                        destination,
                        employee_values,
                        photo_path=employee_values["Foto"] or None,
                    )
                    report.generated_documents.extend(str(p) for p in result.generated_files)
                    report.missing_placeholders.update(result.missing_placeholders)
                    report.skipped_files.extend(result.skipped_files)
            generated_relative = [Path(p).relative_to(temp_output) for p in report.generated_documents]
            if final_output.exists():
                stamped = datetime.now(DR_TZ).strftime("%H%M%S")
                final_output = final_output.with_name(f"{final_output.name}_{stamped}")
                report.output_folder = str(final_output)
            temp_output.replace(final_output)
            report.generated_documents = [str(final_output / relative) for relative in generated_relative]
        except Exception:
            shutil.rmtree(temp_output, ignore_errors=True)
            raise

        report.employees_processed = len(employees)
        logging.info(
            "Procesados %s empleados, documentos generados: %s",
            report.employees_processed,
            report.document_count,
        )
        if report.missing_placeholders:
            logging.warning("Campos faltantes: %s", ", ".join(sorted(report.missing_placeholders)))
        company_name = _display_company_name(employees)
        if workflow == "salidas":
            report.message = (
                f"El duendecito de Vianni termino de procesar {report.employees_processed} salidas de {company_name} "
                f"y genero {report.document_count} documentos."
            )
        else:
            report.message = (
                f"El duendecito de Vianni termino de procesar {report.employees_processed} nuevos empleados de {company_name} "
                f"y genero {report.document_count} documentos."
            )
        return report


def _add_employee_name_aliases(employee: dict[str, str]) -> None:
    name = employee.get("Nombre Empleado") or employee.get("Nombres Empleado", "")
    if not employee.get("Nombre Empleado"):
        employee["Nombre Empleado"] = name
    if not employee.get("Nombres Empleado"):
        employee["Nombres Empleado"] = name


def _normalized_text(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    return "".join(char for char in text if not unicodedata.combining(char)).strip().casefold()


def _salidas_templates_for_employee(templates: list[Path], employee: dict[str, str]) -> list[Path]:
    action = _normalized_text(employee.get("Tipo Acción") or employee.get("Tipo Accion"))
    if "desahucio" in action:
        selected_action = "desahucio"
    elif "renuncia" in action:
        selected_action = "renuncia"
    else:
        selected_action = ""
        logging.warning(
            "Tipo Acción no reconocido para %s: %s",
            employee.get("Nombre Empleado") or employee.get("Nombres Empleado", "empleado sin nombre"),
            employee.get("Tipo Acción") or employee.get("Tipo Accion", ""),
        )

    selected: list[Path] = []
    for template in templates:
        template_name = _normalized_text(template.stem)
        if "notificacion de desahucio" in template_name:
            if selected_action == "desahucio":
                selected.append(template)
        elif "notificacion de renuncia" in template_name:
            if selected_action == "renuncia":
                selected.append(template)
        else:
            selected.append(template)
    return selected


def _run_company_folder(employees: list[dict[str, str]]) -> str:
    if not employees:
        return "Brothers"
    return company_template_subfolder(employees[0])


def _display_company_name(employees: list[dict[str, str]]) -> str:
    return "Supermercado Ines" if _run_company_folder(employees) == "Ines" else "Brothers & Investments"
