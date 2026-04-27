#!/usr/bin/env python3
"""
Generate a CSV template to bulk-load historial_titulares.

This project stores ente name (not ente_id) in historial_titulares, but the UI
looks up the name from entes_detalle by (ejercicio, ente_id). To avoid mismatches
we generate the template from entes_detalle.
"""

from __future__ import annotations

import argparse
import csv
import sqlite3
from pathlib import Path


def ente_numero_sort_sql(column: str) -> str:
    clean = f"TRIM(COALESCE({column}, ''))"
    return (
        "CASE "
        f"WHEN {clean} = '' THEN 0 "
        f"WHEN INSTR({clean}, '.') > 0 THEN "
        f"CAST(SUBSTR({clean}, 1, INSTR({clean}, '.') - 1) AS REAL) * 1000 "
        f"+ CAST(SUBSTR({clean}, INSTR({clean}, '.') + 1) AS REAL) "
        f"ELSE CAST({clean} AS REAL) * 1000 "
        "END"
    )


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--db", default="sifeet.db", help="SQLite DB path (default: sifeet.db)")
    p.add_argument("--ejercicio", type=int, required=True, help="Ejercicio to generate (e.g. 2023)")
    p.add_argument(
        "--out",
        default="bases/historial_titulares_template.csv",
        help="Output CSV path",
    )
    p.add_argument(
        "--tipo-auditoria",
        default="Financiera",
        help="Default tipo_auditoria to place in template",
    )
    args = p.parse_args()

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(args.db)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    rows = cur.execute(
        f"""
        SELECT ente_id, ente_nombre
        FROM entes_detalle
        WHERE ejercicio = ?
        ORDER BY {ente_numero_sort_sql('ente_numero')} ASC, ente_numero ASC
        """,
        (str(args.ejercicio),),
    ).fetchall()

    if not rows:
        raise SystemExit(f"No hay entes_detalle para ejercicio={args.ejercicio} en {args.db}")

    # Create two default rows per ente (titular + director_administrativo).
    fieldnames = [
        "ejercicio",
        "ente_id",
        "ente",
        "tipo_auditoria",
        "tipo_registro",
        "nombre",
        "cargo",
        "fecha_inicio",
        "fecha_fin",
    ]

    with out_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for r in rows:
            for tipo_registro in ("titular", "director_administrativo"):
                w.writerow(
                    {
                        "ejercicio": args.ejercicio,
                        "ente_id": (r["ente_id"] or "").strip(),
                        "ente": (r["ente_nombre"] or "").strip(),
                        "tipo_auditoria": args.tipo_auditoria,
                        "tipo_registro": tipo_registro,
                        "nombre": "",
                        "cargo": "",
                        "fecha_inicio": f"{args.ejercicio}-01-01",
                        "fecha_fin": f"{args.ejercicio}-12-31",
                    }
                )

    print(f"OK: plantilla generada -> {out_path} ({len(rows)} entes, {len(rows)*2} filas)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
