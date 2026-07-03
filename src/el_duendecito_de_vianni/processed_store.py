from __future__ import annotations

import hashlib
import json
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo


DR_TZ = ZoneInfo("America/Santo_Domingo")


class ProcessedStore:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.records = self._load()

    def _load(self) -> list[dict]:
        if not self.path.exists():
            return []
        return json.loads(self.path.read_text(encoding="utf-8"))

    def save(self) -> None:
        self.path.write_text(json.dumps(self.records, ensure_ascii=False, indent=2), encoding="utf-8")

    def is_processed_hash(self, file_hash: str) -> bool:
        return any(record.get("hash") == file_hash for record in self.records)

    def add(self, source: Path, imported: Path, file_hash: str) -> None:
        self.records.append(
            {
                "source": str(source),
                "imported": str(imported),
                "hash": file_hash,
                "processed_at": datetime.now(DR_TZ).isoformat(),
            }
        )
        self.save()


def file_sha256(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()
