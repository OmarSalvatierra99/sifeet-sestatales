from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re
import sqlite3
from typing import Iterable


BACKUP_DIRNAME = "backups"
DEFAULT_BACKUP_PATTERNS = ("*.sqlite", "*.db", "*.backup_*")
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


def find_backup_files(
    base_dir: str | Path,
    *,
    patterns: Iterable[str] = DEFAULT_BACKUP_PATTERNS,
    recursive: bool = False,
) -> list[Path]:
    directory = Path(base_dir).expanduser().resolve()
    if not directory.exists():
        return []

    files_by_path: dict[Path, Path] = {}
    for pattern in patterns:
        iterator = directory.rglob(pattern) if recursive else directory.glob(pattern)
        for path in iterator:
            if not path.is_file():
                continue
            files_by_path[path.resolve()] = path

    return sorted(
        files_by_path.values(),
        key=lambda item: (item.stat().st_mtime, item.name),
        reverse=True,
    )


def prune_backup_dir(
    base_dir: str | Path,
    *,
    keep: int,
    patterns: Iterable[str] = DEFAULT_BACKUP_PATTERNS,
    recursive: bool = False,
    dry_run: bool = True,
) -> dict[str, object]:
    keep = max(int(keep), 0)
    files = find_backup_files(base_dir, patterns=patterns, recursive=recursive)
    keep_files = files[:keep]
    delete_files = files[keep:]
    deleted_bytes = sum(path.stat().st_size for path in delete_files)

    if not dry_run:
        for path in delete_files:
            path.unlink()

    return {
        "base_dir": Path(base_dir).expanduser().resolve(),
        "total_found": len(files),
        "kept": keep_files,
        "deleted": delete_files,
        "deleted_bytes": deleted_bytes,
        "dry_run": dry_run,
    }


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
