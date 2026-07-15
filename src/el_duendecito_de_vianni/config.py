from __future__ import annotations

import json
import os
import sys
from dataclasses import asdict, dataclass
from pathlib import Path


APP_DIR_NAME = "ElDuendecitoDeVianni"


@dataclass
class AppConfig:
    downloads_folder: str
    template_folder: str
    output_folder: str
    imported_folder: str
    logs_folder: str
    work_schedule_lookup: str
    mercury_url: str = ""
    mercury_username: str = ""
    mercury_company: str = "Supermercado Ines"
    mercury_companies: str = "Supermercado Ines;Brothers"
    mercury_report_name: str = "EntradasDeHoy"
    mercury_headless: bool = False
    scan_interval_minutes: int = 15
    ask_before_delete_original: bool = True
    selected_printer: str = ""
    start_minimized_to_tray: bool = True
    start_with_windows: bool = False
    monitoring_enabled: bool = True

    @classmethod
    def default(cls, app_root: Path) -> "AppConfig":
        return cls(
            downloads_folder=str(Path.home() / "Downloads"),
            template_folder=str(app_root / "plantillas"),
            output_folder=str(app_root / "output"),
            imported_folder=str(app_root / "imported_files"),
            logs_folder=str(app_root / "logs"),
            work_schedule_lookup=str(app_root / "politica_horario.xlsx"),
        )

    def resolved(self, app_root: Path) -> "AppConfig":
        data = asdict(self)
        for key in (
            "downloads_folder",
            "template_folder",
            "output_folder",
            "imported_folder",
            "logs_folder",
            "work_schedule_lookup",
        ):
            value = os.path.expandvars(data[key])
            path = Path(value)
            if not path.is_absolute():
                path = app_root / path
            data[key] = str(path)
        data["scan_interval_minutes"] = max(1, int(data["scan_interval_minutes"]))
        return AppConfig(**data)


def get_app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(os.getenv("EL_DUENDECITO_APP_ROOT", Path.cwd())).resolve()


def load_config(app_root: Path | None = None) -> AppConfig:
    root = app_root or get_app_root()
    config_path = root / "config.json"
    if not config_path.exists():
        config = AppConfig.default(root)
        save_config(config, root)
        return config

    raw = json.loads(config_path.read_text(encoding="utf-8"))
    defaults = asdict(AppConfig.default(root))
    defaults.update(raw)
    return AppConfig(**defaults).resolved(root)


def save_config(config: AppConfig, app_root: Path | None = None) -> None:
    root = app_root or get_app_root()
    root.mkdir(parents=True, exist_ok=True)
    (root / "config.json").write_text(
        json.dumps(asdict(config), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def ensure_directories(config: AppConfig) -> None:
    for folder in (
        config.template_folder,
        config.output_folder,
        config.imported_folder,
        config.logs_folder,
    ):
        Path(folder).mkdir(parents=True, exist_ok=True)
    for workflow_folder in ("Entradas", "Salidas"):
        for company_folder in ("Brothers", "Ines"):
            (Path(config.template_folder) / workflow_folder / company_folder).mkdir(parents=True, exist_ok=True)
