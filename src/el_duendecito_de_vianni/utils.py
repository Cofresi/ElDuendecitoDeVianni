from __future__ import annotations

import re
from pathlib import Path


INVALID_WINDOWS_CHARS = r'<>:"/\|?*'


def safe_filename(value: object, fallback: str = "SIN_NOMBRE") -> str:
    text = str(value or "").strip()
    for char in INVALID_WINDOWS_CHARS:
        text = text.replace(char, "-")
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text or fallback


def employee_folder_name(row: dict[str, str]) -> str:
    number = safe_filename(row.get("Numero", ""), "SIN_NUMERO")
    name = safe_filename(row.get("Nombre Empleado", ""), "SIN_NOMBRE")
    return safe_filename(f"{number} - {name}")


def sorted_templates(template_folder: str | Path) -> list[Path]:
    supported = {".doc", ".docx", ".xls", ".xlsx"}
    folder = Path(template_folder)
    if not folder.exists():
        return []
    return sorted(
        (p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in supported),
        key=lambda p: p.name.lower(),
    )
