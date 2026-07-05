from __future__ import annotations

from copy import copy
from pathlib import Path

from docx import Document
from openpyxl import Workbook, load_workbook

from el_duendecito_de_vianni.config import AppConfig
from el_duendecito_de_vianni.importer import import_export
from el_duendecito_de_vianni.processed_store import ProcessedStore, file_sha256
from el_duendecito_de_vianni.processor import DocumentProcessor
from el_duendecito_de_vianni.spreadsheet import read_employees
from el_duendecito_de_vianni.templates import process_docx, replace_placeholders
from el_duendecito_de_vianni.utils import company_template_subfolder, employee_folder_name, sorted_templates


def make_config(tmp_path: Path) -> AppConfig:
    return AppConfig(
        downloads_folder=str(tmp_path / "Downloads"),
        template_folder=str(tmp_path / "templates"),
        output_folder=str(tmp_path / "output"),
        imported_folder=str(tmp_path / "imported_files"),
        logs_folder=str(tmp_path / "logs"),
    )


def make_employee_sheet(path: Path, rows: list[dict[str, object]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    headers = list(rows[0].keys())
    sheet.append(headers)
    for row in rows:
        sheet.append([row.get(header) for header in headers])
    workbook.save(path)


def make_docx(path: Path, text: str) -> None:
    document = Document()
    document.add_paragraph(text)
    document.save(path)


def test_one_employee_and_one_docx_template(tmp_path: Path) -> None:
    config = make_config(tmp_path)
    Path(config.template_folder).mkdir(parents=True)
    Path(config.output_folder).mkdir(parents=True)
    make_employee_sheet(tmp_path / "employees.xlsx", [{"Numero": 295, "Nombre Empleado": "ALANNA"}])
    make_docx(Path(config.template_folder) / "01_Contrato.docx", "Hola {{Nombre Empleado}}")

    report = DocumentProcessor(config).process_imported_file(tmp_path / "employees.xlsx")

    assert report.employees_processed == 1
    assert report.document_count == 1
    generated = Path(report.generated_documents[0])
    assert generated.exists()
    assert "Hola ALANNA" in "\n".join(p.text for p in Document(generated).paragraphs)


def test_multiple_employees_and_multiple_templates(tmp_path: Path) -> None:
    config = make_config(tmp_path)
    Path(config.template_folder).mkdir(parents=True)
    make_employee_sheet(
        tmp_path / "employees.xlsx",
        [
            {"Numero": 1, "Nombre Empleado": "ANA"},
            {"Numero": 2, "Nombre Empleado": "LUIS"},
        ],
    )
    make_docx(Path(config.template_folder) / "01_Contrato.docx", "{{Nombre Empleado}}")
    make_docx(Path(config.template_folder) / "02_Autorizacion.docx", "{{Numero}}")

    report = DocumentProcessor(config).process_imported_file(tmp_path / "employees.xlsx")

    assert report.employees_processed == 2
    assert report.document_count == 4


def test_placeholders_with_accents_and_spaces(tmp_path: Path) -> None:
    config = make_config(tmp_path)
    Path(config.template_folder).mkdir(parents=True)
    make_employee_sheet(
        tmp_path / "employees.xlsx",
        [{"Numero": 1, "Nombre Empleado": "ANA", "Posición": "Analista"}],
    )
    make_docx(Path(config.template_folder) / "01_Contrato.docx", "Cargo: {{Posición}}")

    report = DocumentProcessor(config).process_imported_file(tmp_path / "employees.xlsx")

    text = "\n".join(p.text for p in Document(report.generated_documents[0]).paragraphs)
    assert "Cargo: Analista" in text


def test_blank_excel_cells_become_blank_text(tmp_path: Path) -> None:
    path = tmp_path / "employees.xlsx"
    make_employee_sheet(path, [{"Numero": 1, "Nombre Empleado": None, "Telefono1": None}])

    rows = read_employees(path)

    assert rows[0]["Nombre Empleado"] == ""
    assert rows[0]["Telefono1"] == ""


def test_missing_placeholder_is_left_and_reported(tmp_path: Path) -> None:
    config = make_config(tmp_path)
    Path(config.template_folder).mkdir(parents=True)
    make_employee_sheet(tmp_path / "employees.xlsx", [{"Numero": 1, "Nombre Empleado": "ANA"}])
    make_docx(Path(config.template_folder) / "01_Contrato.docx", "{{Campo Inexistente}}")

    report = DocumentProcessor(config).process_imported_file(tmp_path / "employees.xlsx")

    assert "Campo Inexistente" in report.missing_placeholders
    text = "\n".join(p.text for p in Document(report.generated_documents[0]).paragraphs)
    assert "{{Campo Inexistente}}" in text


def test_placeholder_split_across_word_runs(tmp_path: Path) -> None:
    path = tmp_path / "split.docx"
    document = Document()
    paragraph = document.add_paragraph()
    paragraph.add_run("Hola {{Nombre")
    paragraph.add_run(" Empleado}}")
    document.save(path)

    missing: set[str] = set()
    process_docx(path, {"Nombre Empleado": "ANA"}, missing)

    text = "\n".join(p.text for p in Document(path).paragraphs)
    assert text == "Hola ANA"
    assert not missing


def test_placeholder_format_filters() -> None:
    missing: set[str] = set()
    values = {
        "Salario Base": "18800",
        "Numero": "295.0",
        "Fecha Ingreso": "2026-07-04",
        "Nombre Empleado": "ANA",
        "Sexo": "Femenino",
    }

    rendered = replace_placeholders(
        "{{Salario Base|money}} {{Numero|int}} {{Fecha Ingreso|date}} {{Sexo|tratamiento}} {{Nombre Empleado}}",
        values,
        missing,
    )

    assert rendered == "18,800.00 295 04/07/2026 Sra. ANA"
    assert not missing


def test_tratamiento_filter_gender_values() -> None:
    assert replace_placeholders("{{Sexo|tratamiento}}", {"Sexo": "F"}, set()) == "Sra."
    assert replace_placeholders("{{Sexo|tratamiento}}", {"Sexo": "Femenino"}, set()) == "Sra."
    assert replace_placeholders("{{Sexo|tratamiento}}", {"Sexo": "M"}, set()) == "Sr."
    assert replace_placeholders("{{Sexo|tratamiento}}", {"Sexo": "Masculino"}, set()) == "Sr."
    assert replace_placeholders("{{Sexo|tratamiento}}", {"Sexo": "Sin definir"}, set()) == ""


def test_docx_template_money_filter(tmp_path: Path) -> None:
    config = make_config(tmp_path)
    Path(config.template_folder).mkdir(parents=True)
    make_employee_sheet(
        tmp_path / "employees.xlsx",
        [{"Numero": 1, "Nombre Empleado": "ANA", "Salario Base": 18800}],
    )
    make_docx(Path(config.template_folder) / "01_Contrato.docx", "Salario: {{Salario Base|money}}")

    report = DocumentProcessor(config).process_imported_file(tmp_path / "employees.xlsx")

    text = "\n".join(p.text for p in Document(report.generated_documents[0]).paragraphs)
    assert "Salario: 18,800.00" in text


def test_docx_template_tratamiento_filter(tmp_path: Path) -> None:
    config = make_config(tmp_path)
    Path(config.template_folder).mkdir(parents=True)
    make_employee_sheet(
        tmp_path / "employees.xlsx",
        [{"Numero": 1, "Nombre Empleado": "ANA", "Sexo": "Femenino"}],
    )
    make_docx(Path(config.template_folder) / "01_Carta.docx", "{{Sexo|tratamiento}} {{Nombre Empleado}}")

    report = DocumentProcessor(config).process_imported_file(tmp_path / "employees.xlsx")

    text = "\n".join(p.text for p in Document(report.generated_documents[0]).paragraphs)
    assert "Sra. ANA" in text


def test_xlsx_template_preserves_formula_and_formatting(tmp_path: Path) -> None:
    config = make_config(tmp_path)
    Path(config.template_folder).mkdir(parents=True)
    make_employee_sheet(tmp_path / "employees.xlsx", [{"Numero": 1, "Nombre Empleado": "ANA"}])
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "{{Nombre Empleado}}"
    sheet["B1"] = "=1+1"
    bold_font = copy(sheet["A1"].font)
    bold_font.bold = True
    sheet["A1"].font = bold_font
    template = Path(config.template_folder) / "01_Formulario.xlsx"
    workbook.save(template)

    report = DocumentProcessor(config).process_imported_file(tmp_path / "employees.xlsx")
    generated = load_workbook(report.generated_documents[0], data_only=False)

    assert generated.active["A1"].value == "ANA"
    assert generated.active["A1"].font.bold
    assert generated.active["B1"].value == "=1+1"


def test_xlsx_template_format_filters(tmp_path: Path) -> None:
    config = make_config(tmp_path)
    Path(config.template_folder).mkdir(parents=True)
    make_employee_sheet(
        tmp_path / "employees.xlsx",
        [{"Numero": 295, "Nombre Empleado": "ANA", "Salario Base": 18800}],
    )
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "{{Salario Base|money}}"
    sheet["A2"] = "{{Numero|int}}"
    template = Path(config.template_folder) / "01_Formulario.xlsx"
    workbook.save(template)

    report = DocumentProcessor(config).process_imported_file(tmp_path / "employees.xlsx")
    generated = load_workbook(report.generated_documents[0])

    assert generated.active["A1"].value == "18,800.00"
    assert generated.active["A2"].value == "295"


def test_duplicate_input_export_detection(tmp_path: Path) -> None:
    downloads = tmp_path / "Downloads"
    imported = tmp_path / "imported_files"
    downloads.mkdir()
    source = downloads / "GridViewExports.xlsx"
    make_employee_sheet(source, [{"Numero": 1, "Nombre Empleado": "ANA"}])
    result = import_export(source, imported)
    store = ProcessedStore(imported / "processed_files.json")
    store.add(source, result.imported_path, file_sha256(result.imported_path))

    assert store.is_processed_hash(file_sha256(source))


def test_invalid_characters_in_employee_names() -> None:
    row = {"Numero": "12", "Nombre Empleado": 'ANA:LUIS/TEST*"'}

    assert employee_folder_name(row) == "12 - ANA-LUIS-TEST--"


def test_print_order_sorting() -> None:
    folder = Path("unused")
    names = ["10_Final.docx", "02_Formulario.xlsx", "01_Contrato.docx"]
    assert sorted(names, key=lambda value: value.lower()) == ["01_Contrato.docx", "02_Formulario.xlsx", "10_Final.docx"]


def test_template_sorting_ignores_office_lock_files(tmp_path: Path) -> None:
    make_docx(tmp_path / "01_Contrato.docx", "Contrato")
    (tmp_path / "~$01_Contrato.docx").write_text("office lock", encoding="utf-8")
    (tmp_path / "~$02_Formulario.xlsx").write_text("office lock", encoding="utf-8")

    assert [path.name for path in sorted_templates(tmp_path)] == ["01_Contrato.docx"]


def test_company_template_routing_uses_ines_folder(tmp_path: Path) -> None:
    config = make_config(tmp_path)
    ines_folder = Path(config.template_folder) / "Ines"
    brothers_folder = Path(config.template_folder) / "Brothers"
    ines_folder.mkdir(parents=True)
    brothers_folder.mkdir(parents=True)
    make_employee_sheet(
        tmp_path / "employees.xlsx",
        [{"Numero": 1, "Nombre Empleado": "ANA", "Compania": "Supermercado Ines"}],
    )
    make_docx(ines_folder / "01_Ines.docx", "Ines {{Nombre Empleado}}")
    make_docx(brothers_folder / "01_Brothers.docx", "Brothers {{Nombre Empleado}}")

    report = DocumentProcessor(config).process_imported_file(tmp_path / "employees.xlsx")

    assert report.document_count == 1
    generated = Path(report.generated_documents[0])
    assert generated.name == "01_Ines.docx"
    assert "Ines ANA" in "\n".join(p.text for p in Document(generated).paragraphs)


def test_company_template_routing_uses_brothers_for_other_companies(tmp_path: Path) -> None:
    config = make_config(tmp_path)
    ines_folder = Path(config.template_folder) / "Ines"
    brothers_folder = Path(config.template_folder) / "Brothers"
    ines_folder.mkdir(parents=True)
    brothers_folder.mkdir(parents=True)
    make_employee_sheet(
        tmp_path / "employees.xlsx",
        [{"Numero": 2, "Nombre Empleado": "LUIS", "Compania": "Otra Compania"}],
    )
    make_docx(ines_folder / "01_Ines.docx", "Ines {{Nombre Empleado}}")
    make_docx(brothers_folder / "01_Brothers.docx", "Brothers {{Nombre Empleado}}")

    report = DocumentProcessor(config).process_imported_file(tmp_path / "employees.xlsx")

    assert report.document_count == 1
    generated = Path(report.generated_documents[0])
    assert generated.name == "01_Brothers.docx"
    assert "Brothers LUIS" in "\n".join(p.text for p in Document(generated).paragraphs)


def test_company_template_subfolder_mapping() -> None:
    assert company_template_subfolder({"Compania": "Supermercado Ines"}) == "Ines"
    assert company_template_subfolder({"Compania": "Supermercado Inés"}) == "Brothers"
    assert company_template_subfolder({"Compania": ""}) == "Brothers"
