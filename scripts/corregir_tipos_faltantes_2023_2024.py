#!/usr/bin/env python3
"""
Rellena tipos de auditoria faltantes en historial_titulares para 2023 y 2024.

Regla:
- Si un ente tiene observaciones en un tipo de auditoria y en historial_titulares
  no existe ese tipo para el mismo ejercicio/ente, duplica filas base de historial
  (preferentemente del tipo "Financiera") hacia el tipo faltante.
- No elimina ni sobrescribe filas existentes.
"""

from __future__ import annotations

import argparse
import sqlite3


TARGET_YEARS = ("2023", "2024")
PREFERRED_SOURCE_TYPE = "Financiera"


def get_missing_types(conn: sqlite3.Connection, ejercicio: str):
    query = """
    WITH obs AS (
      SELECT DISTINCT
        TRIM(RTRIM(o.ente_id, '.')) AS ente_id,
        TRIM(COALESCE(o.tipo_auditoria, '')) AS tipo_obs
      FROM observaciones o
      WHERE TRIM(COALESCE(o.ejercicio, '')) = ?
        AND TRIM(COALESCE(o.periodo_cedula, '')) != ''
        AND TRIM(COALESCE(o.tipo_auditoria, '')) != ''
    ),
    uid AS (
      SELECT DISTINCT
        TRIM(RTRIM(ente_id, '.')) AS ente_id,
        TRIM(COALESCE(ente_uid, '')) AS ente_uid,
        TRIM(COALESCE(ente_nombre, '')) AS ente_nombre
      FROM entes_detalle
      WHERE TRIM(COALESCE(ejercicio, '')) = ?
        AND TRIM(COALESCE(ente_uid, '')) != ''
    ),
    hist AS (
      SELECT DISTINCT
        TRIM(COALESCE(ente_uid, '')) AS ente_uid,
        TRIM(COALESCE(tipo_auditoria, '')) AS tipo_hist
      FROM historial_titulares
      WHERE TRIM(COALESCE(ejercicio, '')) = ?
    )
    SELECT
      ? AS ejercicio,
      u.ente_uid,
      u.ente_nombre,
      o.tipo_obs AS missing_tipo
    FROM obs o
    JOIN uid u ON u.ente_id = o.ente_id
    LEFT JOIN hist h
      ON h.ente_uid = u.ente_uid
     AND h.tipo_hist = o.tipo_obs
    WHERE h.tipo_hist IS NULL
      AND EXISTS (
        SELECT 1
        FROM historial_titulares hx
        WHERE TRIM(COALESCE(hx.ejercicio, '')) = ?
          AND TRIM(COALESCE(hx.ente_uid, '')) = u.ente_uid
      )
    ORDER BY u.ente_nombre, o.tipo_obs
    """
    return conn.execute(
        query,
        (ejercicio, ejercicio, ejercicio, ejercicio, ejercicio),
    ).fetchall()


def get_source_rows(conn: sqlite3.Connection, ejercicio: str, ente_uid: str):
    preferred = conn.execute(
        """
        SELECT id, ente, nombre, cargo, fecha_inicio, fecha_fin, tipo_registro, tipo_auditoria, ente_uid
        FROM historial_titulares
        WHERE TRIM(COALESCE(ejercicio, '')) = ?
          AND TRIM(COALESCE(ente_uid, '')) = ?
          AND TRIM(COALESCE(tipo_auditoria, '')) = ?
        ORDER BY fecha_inicio, id
        """,
        (ejercicio, ente_uid, PREFERRED_SOURCE_TYPE),
    ).fetchall()
    if preferred:
        return preferred

    return conn.execute(
        """
        SELECT id, ente, nombre, cargo, fecha_inicio, fecha_fin, tipo_registro, tipo_auditoria, ente_uid
        FROM historial_titulares
        WHERE TRIM(COALESCE(ejercicio, '')) = ?
          AND TRIM(COALESCE(ente_uid, '')) = ?
        ORDER BY fecha_inicio, id
        """,
        (ejercicio, ente_uid),
    ).fetchall()


def target_exists(
    conn: sqlite3.Connection,
    ejercicio: str,
    ente_uid: str,
    missing_tipo: str,
    base_row: sqlite3.Row,
) -> bool:
    row = conn.execute(
        """
        SELECT 1
        FROM historial_titulares
        WHERE TRIM(COALESCE(ejercicio, '')) = ?
          AND TRIM(COALESCE(ente_uid, '')) = ?
          AND TRIM(COALESCE(tipo_auditoria, '')) = ?
          AND TRIM(COALESCE(tipo_registro, '')) = TRIM(COALESCE(?, ''))
          AND TRIM(COALESCE(nombre, '')) = TRIM(COALESCE(?, ''))
          AND COALESCE(fecha_inicio, '') = COALESCE(?, '')
          AND COALESCE(fecha_fin, '') = COALESCE(?, '')
        LIMIT 1
        """,
        (
            ejercicio,
            ente_uid,
            missing_tipo,
            base_row["tipo_registro"],
            base_row["nombre"],
            base_row["fecha_inicio"],
            base_row["fecha_fin"],
        ),
    ).fetchone()
    return bool(row)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--db", default="sifeet.db")
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()

    conn = sqlite3.connect(args.db)
    conn.row_factory = sqlite3.Row

    inserts: list[tuple] = []
    total_missing_types = 0

    for ejercicio in TARGET_YEARS:
        missing = get_missing_types(conn, ejercicio)
        total_missing_types += len(missing)
        print(f"\nEjercicio {ejercicio}: tipos faltantes detectados = {len(missing)}")
        for item in missing:
            source_rows = get_source_rows(conn, ejercicio, item["ente_uid"])
            print(
                f"- ente_uid={item['ente_uid']} | ente={item['ente_nombre']} | "
                f"tipo_faltante={item['missing_tipo']} | filas_base={len(source_rows)}"
            )
            for base in source_rows:
                if target_exists(conn, ejercicio, item["ente_uid"], item["missing_tipo"], base):
                    continue
                inserts.append(
                    (
                        ejercicio,
                        base["ente"],
                        base["nombre"],
                        base["cargo"],
                        base["fecha_inicio"],
                        base["fecha_fin"],
                        base["tipo_registro"],
                        item["missing_tipo"],
                        base["ente_uid"],
                    )
                )

    print(f"\nTipos faltantes (ejercicio+ente+tipo): {total_missing_types}")
    print(f"Filas nuevas a insertar: {len(inserts)}")
    print(f"Modo: {'APPLY' if args.apply else 'DRY-RUN'}")

    if not args.apply or not inserts:
        return 0

    conn.execute("BEGIN")
    conn.executemany(
        """
        INSERT INTO historial_titulares (
            ejercicio,
            ente,
            nombre,
            cargo,
            fecha_inicio,
            fecha_fin,
            tipo_registro,
            tipo_auditoria,
            ente_uid
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        inserts,
    )
    conn.commit()
    print(f"Inserciones aplicadas: {len(inserts)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
