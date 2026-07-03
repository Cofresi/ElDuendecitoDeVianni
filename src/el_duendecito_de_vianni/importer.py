from __future__ import annotations

import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from .processed_store import file_sha256


DR_TZ = ZoneInfo("America/Santo_Domingo")
EXPORT_NAMES = {
    "GridViewExports.xls",
    "GridViewExport.xls",
    "GridViewExports.xlsx",
    "GridViewExport.xlsx",
}


@dataclass
class ImportResult:
    source_path: Path
    imported_path: Path
    file_hash: str
    already_processed: bool = False


def find_export(downloads_folder: str | Path) -> Path | None:
    folder = Path(downloads_folder)
    for name in EXPORT_NAMES:
        candidate = folder / name
        if candidate.exists() and candidate.is_file():
            return candidate
    return None


def _target_path(imported_folder: Path, suffix: str) -> Path:
    date_text = datetime.now(DR_TZ).strftime("%d.%m.%Y")
    base = imported_folder / f"nuevasEntradas_{date_text}{suffix}"
    if not base.exists():
        return base
    stamped = datetime.now(DR_TZ).strftime("%H%M%S")
    return imported_folder / f"nuevasEntradas_{date_text}_{stamped}{suffix}"


def import_export(source: str | Path, imported_folder: str | Path) -> ImportResult:
    source_path = Path(source)
    imported_dir = Path(imported_folder)
    imported_dir.mkdir(parents=True, exist_ok=True)
    target = _target_path(imported_dir, source_path.suffix.lower())
    temp_target = target.with_suffix(target.suffix + ".tmp")
    shutil.copy2(source_path, temp_target)
    if file_sha256(source_path) != file_sha256(temp_target):
        temp_target.unlink(missing_ok=True)
        raise IOError("La copia del archivo exportado no pudo verificarse.")
    temp_target.replace(target)
    return ImportResult(source_path=source_path, imported_path=target, file_hash=file_sha256(target))
