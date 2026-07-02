from __future__ import annotations

import json
import os
import threading
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, TypeVar

from app.config import get_settings

T = TypeVar("T")


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


class JsonStore:
    """Persistência simples em arquivo JSON.

    Não usa banco de dados. Os dados ficam em `data/ora_nfse_storage.json`, que
    é ignorado pelo Git. Escritas são atômicas para reduzir risco de corrupção.
    """

    def __init__(self, path: str | Path | None = None) -> None:
        settings = get_settings()
        self.path = Path(path or settings.storage_file)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.RLock()
        if not self.path.exists():
            self.write(self.default_data())

    @staticmethod
    def default_data() -> dict[str, Any]:
        return {
            "schema_version": 9,
            "created_at": utc_now_iso(),
            "updated_at": utc_now_iso(),
            "seq": {"cliente": 0, "certificado": 0, "log": 0, "job": 0},
            "clientes": [],
            "certificados": [],
            "notas": [],
            "eventos": [],
            "logs": [],
            "sync_jobs": [],
        }

    def read(self) -> dict[str, Any]:
        with self._lock:
            if not self.path.exists():
                data = self.default_data()
                self.write(data)
                return data
            with self.path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
            return self._migrate(data)

    def write(self, data: dict[str, Any]) -> None:
        with self._lock:
            data = deepcopy(data)
            data["updated_at"] = utc_now_iso()
            self.path.parent.mkdir(parents=True, exist_ok=True)
            tmp = self.path.with_suffix(self.path.suffix + ".tmp")
            with tmp.open("w", encoding="utf-8") as fh:
                json.dump(data, fh, ensure_ascii=False, indent=2, sort_keys=True)
            os.replace(tmp, self.path)

    def transaction(self, callback: Callable[[dict[str, Any]], T]) -> T:
        with self._lock:
            data = self.read()
            result = callback(data)
            self.write(data)
            return result

    @staticmethod
    def next_id(data: dict[str, Any], name: str) -> int:
        seq = data.setdefault("seq", {})
        seq[name] = int(seq.get(name, 0)) + 1
        return int(seq[name])

    def _migrate(self, data: dict[str, Any]) -> dict[str, Any]:
        changed = False
        default = self.default_data()
        for key, value in default.items():
            if key not in data:
                data[key] = value
                changed = True
        try:
            if int(data.get("schema_version") or 0) < 9:
                data["schema_version"] = 9
                changed = True
        except (TypeError, ValueError):
            data["schema_version"] = 9
            changed = True
        data.setdefault("seq", {})
        for name in ["cliente", "certificado", "log", "job"]:
            data["seq"].setdefault(name, 0)
        # Garante listas mesmo se arquivo for editado manualmente.
        for key in ["clientes", "certificados", "notas", "eventos", "logs", "sync_jobs"]:
            if not isinstance(data.get(key), list):
                data[key] = []
                changed = True
        if changed:
            self.write(data)
        return data


store = JsonStore()
