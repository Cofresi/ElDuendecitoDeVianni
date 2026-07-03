from __future__ import annotations

import logging
import re
import shutil
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

PLACEHOLDER_RE = re.compile(r"\{\{([^{}]+)\}\}")


@dataclass
class TemplateResult:
    generated_files: list[Path] = field(default_factory=list)
    missing_placeholders: set[str] = field(default_factory=set)
    skipped_files: list[str] = field(default_factory=list)


def placeholders_in_text(text: str) -> set[str]:
    return set(PLACEHOLDER_RE.findall(text or ""))


def replace_placeholders(text: str, values: dict[str, str], missing: set[str]) -> str:
    def repl(match: re.Match[str]) -> str:
        key = match.group(1)
        if key not in values:
            missing.add(key)
            return match.group(0)
        return values.get(key, "")

    return PLACEHOLDER_RE.sub(repl, text or "")


def _replace_paragraph(paragraph, values: dict[str, str], missing: set[str]) -> None:
    full_text = "".join(run.text for run in paragraph.runs)
    if "{{" not in full_text:
        return
    replaced = replace_placeholders(full_text, values, missing)
    if replaced == full_text:
        return
    if paragraph.runs:
        paragraph.runs[0].text = replaced
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(replaced)


def _iter_paragraphs(container):
    for paragraph in getattr(container, "paragraphs", []):
        yield paragraph
    for table in getattr(container, "tables", []):
        for row in table.rows:
            for cell in row.cells:
                yield from _iter_paragraphs(cell)


def process_docx(path: Path, values: dict[str, str], missing: set[str]) -> None:
    document = Document(path)
    for paragraph in _iter_paragraphs(document):
        _replace_paragraph(paragraph, values, missing)
    for section in document.sections:
        for part in (section.header, section.footer):
            for paragraph in _iter_paragraphs(part):
                _replace_paragraph(paragraph, values, missing)
    document.save(path)


def process_xlsx(path: Path, values: dict[str, str], missing: set[str]) -> None:
    workbook = load_workbook(path)
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and "{{" in cell.value:
                    cell.value = replace_placeholders(cell.value, values, missing)
        for attr in ("oddHeader", "evenHeader", "firstHeader", "oddFooter", "evenFooter", "firstFooter"):
            header_footer = getattr(sheet, attr, None)
            if header_footer:
                for side in ("left", "center", "right"):
                    item = getattr(header_footer, side, None)
                    if item and item.text:
                        item.text = replace_placeholders(item.text, values, missing)
    workbook.save(path)


def process_template_copy(template: Path, destination: Path, values: dict[str, str]) -> TemplateResult:
    result = TemplateResult()
    missing: set[str] = set()
    suffix = template.suffix.lower()
    if suffix in {".doc", ".xls"}:
        result.skipped_files.append(template.name)
        logging.warning(
            "Plantilla heredada omitida: %s. Conviertala a DOCX/XLSX o instale soporte de Microsoft Office.",
            template,
        )
        return result
    shutil.copy2(template, destination)
    try:
        if suffix == ".docx":
            process_docx(destination, values, missing)
        elif suffix == ".xlsx":
            process_xlsx(destination, values, missing)
        else:
            result.skipped_files.append(template.name)
    except Exception:
        destination.unlink(missing_ok=True)
        raise
    result.generated_files.append(destination)
    result.missing_placeholders.update(missing)
    return result
