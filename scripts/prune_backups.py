#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path
import sys


PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from backup_utils import prune_backup_dir


DEFAULT_KEEP_PROJECT_ROOT = 2
DEFAULT_KEEP_MANUAL = 8
DEFAULT_KEEP_DB = 30


def format_bytes(size: int) -> str:
    units = ("B", "KB", "MB", "GB", "TB")
    value = float(size)
    for unit in units:
        if value < 1024 or unit == units[-1]:
            if unit == "B":
                return f"{int(value)} {unit}"
            return f"{value:.1f} {unit}"
        value /= 1024
    return f"{size} B"


def project_relative(path: Path) -> str:
    try:
        return str(path.resolve().relative_to(PROJECT_ROOT))
    except ValueError:
        return str(path)


def build_targets(args: argparse.Namespace) -> list[dict[str, object]]:
    return [
        {
            "name": "raiz-del-proyecto",
            "path": PROJECT_ROOT,
            "keep": args.keep_project_root,
            "patterns": ("*.backup_*", "*.before_*.sqlite", "*.before_*.db"),
        },
        {
            "name": "respaldos-manuales",
            "path": PROJECT_ROOT / "backups",
            "keep": args.keep_manual,
            "patterns": ("*.sqlite", "*.db", "*.backup_*"),
        },
        {
            "name": "snapshots-db",
            "path": PROJECT_ROOT / "backups" / "db",
            "keep": args.keep_db,
            "patterns": ("*.sqlite",),
        },
    ]


def print_report(name: str, report: dict[str, object], *, show: int) -> None:
    deleted = report["deleted"]
    kept = report["kept"]
    print(
        f"[{name}] encontrados={report['total_found']} "
        f"conservar={len(kept)} eliminar={len(deleted)} "
        f"liberar={format_bytes(report['deleted_bytes'])}"
    )

    for path in deleted[:show]:
        print(f"  eliminar -> {project_relative(path)}")
    hidden = len(deleted) - min(len(deleted), show)
    if hidden > 0:
        print(f"  ... y {hidden} más")


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Limpia respaldos viejos y conserva solo los más recientes. "
            "Por defecto corre en modo simulación."
        )
    )
    parser.add_argument("--apply", action="store_true", help="Elimina los archivos indicados.")
    parser.add_argument(
        "--keep-project-root",
        type=int,
        default=DEFAULT_KEEP_PROJECT_ROOT,
        help=f"Cantidad de respaldos legacy a conservar en la raíz. Default: {DEFAULT_KEEP_PROJECT_ROOT}.",
    )
    parser.add_argument(
        "--keep-manual",
        type=int,
        default=DEFAULT_KEEP_MANUAL,
        help=f"Cantidad de respaldos a conservar en backups/. Default: {DEFAULT_KEEP_MANUAL}.",
    )
    parser.add_argument(
        "--keep-db",
        type=int,
        default=DEFAULT_KEEP_DB,
        help=f"Cantidad de snapshots a conservar en backups/db/. Default: {DEFAULT_KEEP_DB}.",
    )
    parser.add_argument(
        "--show",
        type=int,
        default=10,
        help="Cuántos archivos listados mostrar por grupo. Default: 10.",
    )
    args = parser.parse_args()

    total_deleted_files = 0
    total_deleted_bytes = 0
    mode = "APLICAR" if args.apply else "SIMULACION"
    print(f"Modo: {mode}")

    for target in build_targets(args):
        report = prune_backup_dir(
            target["path"],
            keep=target["keep"],
            patterns=target["patterns"],
            dry_run=not args.apply,
        )
        print_report(target["name"], report, show=max(args.show, 0))
        total_deleted_files += len(report["deleted"])
        total_deleted_bytes += report["deleted_bytes"]

    if args.apply:
        summary = (
            f"Se eliminaron {total_deleted_files} archivos y se liberaron "
            f"{format_bytes(total_deleted_bytes)}."
        )
    else:
        summary = (
            f"Se eliminarian {total_deleted_files} archivos y se liberarian "
            f"{format_bytes(total_deleted_bytes)}."
        )
    print(summary)
    if not args.apply:
        print("Vuelve a ejecutar con --apply para aplicar la limpieza.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
