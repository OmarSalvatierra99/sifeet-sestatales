from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re
import sqlite3


BACKUP_DIRNAME = "backups"
SAFE_LABEL_RE = re.compile(r"[^A-Za-z0-9._-]+")


def sanitize_backup_label(label: str) -> str:
    clean = SAFE_LABEL_RE.sub("_", (label or "").strip())
    return clean.strip("._-")


def resolve_backup_dir(db_path: str | Path, backup_dir: str | Path | None = None) -> Path:
    db_file = Path(db_path).expanduser().resolve()
    if backup_dir is None:
        destination = db_file.parent / BACKUP_DIRNAME
    else:
        destination = Path(backup_dir).expanduser()
        if not destination.is_absolute():
            destination = db_file.parent / destination
        destination = destination.resolve()
    destination.mkdir(parents=True, exist_ok=True)
    return destination


def create_db_backup(
    db_path: str | Path,
    *,
    backup_dir: str | Path | None = None,
    label: str = "",
) -> Path:
    db_file = Path(db_path).expanduser().resolve()
    if not db_file.exists():
        raise FileNotFoundError(f"No existe la base de datos: {db_file}")

    destination_dir = resolve_backup_dir(db_file, backup_dir=backup_dir)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_label = sanitize_backup_label(label)
    suffix = f"_{safe_label}" if safe_label else ""
    backup_path = destination_dir / f"{db_file.name}.backup_{timestamp}{suffix}"

    source_conn = sqlite3.connect(str(db_file))
    backup_conn = sqlite3.connect(str(backup_path))
    try:
        source_conn.backup(backup_conn)
    finally:
        backup_conn.close()
        source_conn.close()
    return backup_path
