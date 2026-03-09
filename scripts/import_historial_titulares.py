#!/usr/bin/env python3
"""
Bulk import into historial_titulares from CSV.

CSV columns (required unless noted):
- ejercicio (int) [optional if --ejercicio is passed]
- ente_id (str) [optional if ente is provided; recommended]
- ente (str) [optional if ente_id is provided]
- tipo_auditoria (str) [optional; default: Financiera]
- tipo_registro (str): titular | director_administrativo
- nombre (str)
- cargo (str)
- fecha_inicio (YYYY-MM-DD)
- fecha_fin (YYYY-MM-DD)
"""

from __future__ import annotations

import argparse
import csv
import sqlite3
import sys
from datetime import date
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from backup_utils import create_db_backup


ALLOWED_TIPO_REGISTRO = {"titular", "director_administrativo"}


def _clean(s: object) -> str:
    if s is None:
        return ""
    return str(s).strip()


def _parse_date(s: str) -> str:
    # Keep canonical ISO dates in DB.
    if not s:
        raise ValueError("fecha vacia")
    return date.fromisoformat(s).isoformat()


def resolve_ente_nombre(conn: sqlite3.Connection, ejercicio: int, ente_id: str, ente: str) -> str:
    cur = conn.cursor()
    if ente_id:
        row = cur.execute(
            """
            SELECT ente_nombre
            FROM entes_detalle
            WHERE ejercicio = ? AND ente_id = ?
            """,
            (str(ejercicio), ente_id),
        ).fetchone()
        if row is None:
            raise ValueError(f"ente_id no existe en entes_detalle para ejercicio={ejercicio}: {ente_id!r}")
        return str(row[0]).strip()

    if not ente:
        raise ValueError("falta ente_id o ente")

    # Best-effort resolve to the exact entes_detalle name, if present.
    row = cur.execute(
        """
        SELECT ente_nombre
        FROM entes_detalle
        WHERE ejercicio = ? AND LOWER(TRIM(ente_nombre)) = LOWER(TRIM(?))
        """,
        (str(ejercicio), ente),
    ).fetchone()
    if row is not None:
        return str(row[0]).strip()

    # Fall back to raw 'ente' string (keeps import flexible, but may break UI joins).
    return ente


def row_exists(conn: sqlite3.Connection, ejercicio: int, ente: str, tipo_auditoria: str, nombre: str, cargo: str, fi: str, ff: str, tipo_registro: str) -> bool:
    cur = conn.cursor()
    r = cur.execute(
        """
        SELECT 1
        FROM historial_titulares
        WHERE ejercicio = ?
          AND ente = ?
          AND tipo_auditoria = ?
          AND nombre = ?
          AND cargo = ?
          AND fecha_inicio = ?
          AND fecha_fin = ?
          AND tipo_registro = ?
        LIMIT 1
        """,
        (ejercicio, ente, tipo_auditoria, nombre, cargo, fi, ff, tipo_registro),
    ).fetchone()
    return r is not None


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--db", default="sifeet.db", help="SQLite DB path (default: sifeet.db)")
    p.add_argument("--csv", required=True, help="Input CSV path")
    p.add_argument("--ejercicio", type=int, help="Default ejercicio for rows without 'ejercicio'")
    p.add_argument("--dry-run", action="store_true", help="Validate and report, but do not write")
    p.add_argument(
        "--replace",
        action="store_true",
        help="Delete existing historial_titulares for the targeted ejercicio(s) before insert",
    )
    args = p.parse_args()

    conn = sqlite3.connect(args.db)
    conn.row_factory = sqlite3.Row

    # Figure targeted ejercicios for --replace safety.
    targeted_ejercicios: set[int] = set()
    parsed_rows: list[dict[str, str]] = []
    skipped_placeholders = 0

    with open(args.csv, newline="", encoding="utf-8-sig") as f:
        r = csv.DictReader(f)
        if r.fieldnames is None:
            raise SystemExit("CSV sin encabezados")
        for i, raw in enumerate(r, start=2):  # header is row 1
            ejercicio_s = _clean(raw.get("ejercicio"))
            ejercicio = int(ejercicio_s) if ejercicio_s else args.ejercicio
            if not ejercicio:
                raise SystemExit(f"Fila {i}: falta 'ejercicio' y no se paso --ejercicio")

            ente_id = _clean(raw.get("ente_id"))
            ente = _clean(raw.get("ente"))
            tipo_auditoria = _clean(raw.get("tipo_auditoria")) or "Financiera"
            tipo_registro = _clean(raw.get("tipo_registro"))
            nombre = _clean(raw.get("nombre"))
            cargo = _clean(raw.get("cargo"))
            fi = _parse_date(_clean(raw.get("fecha_inicio")))
            ff = _parse_date(_clean(raw.get("fecha_fin")))

            if tipo_registro not in ALLOWED_TIPO_REGISTRO:
                raise SystemExit(
                    f"Fila {i}: tipo_registro invalido {tipo_registro!r}. "
                    f"Permitidos: {sorted(ALLOWED_TIPO_REGISTRO)}"
                )
            # Allow template/placeholder rows to be present but not imported yet.
            if not nombre and not cargo:
                skipped_placeholders += 1
                continue
            if not nombre or not cargo:
                raise SystemExit(f"Fila {i}: 'nombre' y 'cargo' son obligatorios (o deja ambos vacios)")

            ente_nombre = resolve_ente_nombre(conn, ejercicio, ente_id, ente)

            parsed_rows.append(
                {
                    "ejercicio": str(ejercicio),
                    "ente": ente_nombre,
                    "tipo_auditoria": tipo_auditoria,
                    "nombre": nombre,
                    "cargo": cargo,
                    "fecha_inicio": fi,
                    "fecha_fin": ff,
                    "tipo_registro": tipo_registro,
                }
            )
            targeted_ejercicios.add(ejercicio)

    if not parsed_rows:
        print(f"Nada que importar (0 filas). Filas placeholder omitidas={skipped_placeholders}.")
        return 0

    if not args.dry_run:
        backup_path = create_db_backup(args.db, label="before_historial_csv_import")
        print(f"Respaldo creado: {backup_path}")

    if args.replace and not args.dry_run:
        cur = conn.cursor()
        for e in sorted(targeted_ejercicios):
            cur.execute("DELETE FROM historial_titulares WHERE ejercicio = ?", (e,))
        conn.commit()

    inserted = 0
    skipped = 0
    cur = conn.cursor()
    for row in parsed_rows:
        ejercicio = int(row["ejercicio"])
        if row_exists(
            conn,
            ejercicio,
            row["ente"],
            row["tipo_auditoria"],
            row["nombre"],
            row["cargo"],
            row["fecha_inicio"],
            row["fecha_fin"],
            row["tipo_registro"],
        ):
            skipped += 1
            continue

        if not args.dry_run:
            cur.execute(
                """
                INSERT INTO historial_titulares (
                    ejercicio, ente, tipo_auditoria, nombre, cargo, fecha_inicio, fecha_fin, tipo_registro
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    ejercicio,
                    row["ente"],
                    row["tipo_auditoria"],
                    row["nombre"],
                    row["cargo"],
                    row["fecha_inicio"],
                    row["fecha_fin"],
                    row["tipo_registro"],
                ),
            )
        inserted += 1

    if not args.dry_run:
        conn.commit()

    mode = "DRY-RUN" if args.dry_run else "OK"
    print(
        f"{mode}: filas={len(parsed_rows)} insertadas={inserted} omitidas_por_duplicado={skipped} "
        f"omitidas_placeholder={skipped_placeholders} "
        f"ejercicios={sorted(targeted_ejercicios)}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
