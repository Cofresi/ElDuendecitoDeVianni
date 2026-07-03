from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo


DR_TZ = ZoneInfo("America/Santo_Domingo")


def configure_logging(logs_folder: str | Path) -> Path:
    folder = Path(logs_folder)
    folder.mkdir(parents=True, exist_ok=True)
    log_path = folder / f"duendecito_{datetime.now(DR_TZ):%Y%m%d}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(),
        ],
        force=True,
    )
    return log_path
