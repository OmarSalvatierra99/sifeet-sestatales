#!/usr/bin/env python3
"""
Corrige tipo_auditoria en historial_titulares con dry-run por defecto.

Uso recomendado:
1) Preview:
   python scripts/corregir_tipo_historial.py --ejercicio 2023 --ente-id 5 --from-type Administrativo --to-type Financiera
2) Aplicar:
   python scripts/corregir_tipo_historial.py --ejercicio 2023 --ente-id 5 --from-type Administrativo --to-type Financiera --apply
"""

from __future__ import annotations

import argparse
import sqlite3


def normalize_ente_id(value: str) -> str:
    return (value or "").strip().rstrip(".").strip()


def get_scope(conn: sqlite3.Connection, ejercicio: str, ente_id: str):
    if not ente_id:
        return "", [], []
    row = conn.execute(
        """
        SELECT
            TRIM(COALESCE(ente_uid, '')) AS ente_uid,
            TRIM(COALESCE(ente_nombre, '')) AS ente_nombre
        FROM entes_detalle
        WHERE TRIM(COALESCE(ejercicio, '')) = ?
          AND TRIM(RTRIM(COALESCE(ente_id, ''), '.')) = ?
        LIMIT 1
        """,
        (ejercicio, ente_id),
    ).fetchone()
    if not row:
        return "", [], []

    ente_uid = row["ente_uid"]
    aliases = []
    if row["ente_nombre"]:
        aliases.append(row["ente_nombre"])
    if ente_uid:
        rows = conn.execute(
            """
            SELECT DISTINCT TRIM(COALESCE(ente_nombre, '')) AS ente_nombre
            FROM entes_detalle
            WHERE TRIM(COALESCE(ente_uid, '')) = ?
              AND TRIM(COALESCE(ente_nombre, '')) != ''
            ORDER BY ejercicio
            """,
            (ente_uid,),
        ).fetchall()
        for item in rows:
            name = item["ente_nombre"]
            if name and name not in aliases:
                aliases.append(name)
    return ente_uid, aliases, [row["ente_nombre"]]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--db", default="sifeet.db")
    parser.add_argument("--ejercicio", required=True)
    parser.add_argument("--ente-id", default="")
    parser.add_argument("--from-type", action="append", required=True)
    parser.add_argument("--to-type", required=True)
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()

    conn = sqlite3.connect(args.db)
    conn.row_factory = sqlite3.Row

    ejercicio = (args.ejercicio or "").strip()
    ente_id = normalize_ente_id(args.ente_id)
    from_types = sorted(set((item or "").strip() for item in args.from_type if (item or "").strip()))
    to_type = (args.to_type or "").strip()
    if not from_types:
        raise SystemExit("Debes indicar al menos un --from-type.")
    if not to_type:
        raise SystemExit("Debes indicar --to-type.")

    ente_uid, aliases, _ = get_scope(conn, ejercicio, ente_id)

    where = ["TRIM(COALESCE(h.ejercicio, '')) = ?"]
    params: list[str] = [ejercicio]

    from_placeholders = ", ".join(["?"] * len(from_types))
    where.append(f"TRIM(COALESCE(h.tipo_auditoria, '')) IN ({from_placeholders})")
    params.extend(from_types)

    if ente_id:
        if ente_uid and aliases:
            alias_ph = ", ".join(["?"] * len(aliases))
            where.append(
                f"(TRIM(COALESCE(h.ente_uid, '')) = ? OR TRIM(COALESCE(h.ente, '')) IN ({alias_ph}))"
            )
            params.extend([ente_uid, *aliases])
        elif ente_uid:
            where.append("TRIM(COALESCE(h.ente_uid, '')) = ?")
            params.append(ente_uid)
        elif aliases:
            alias_ph = ", ".join(["?"] * len(aliases))
            where.append(f"TRIM(COALESCE(h.ente, '')) IN ({alias_ph})")
            params.extend(aliases)
        else:
            print("No se encontro alcance para el ente_id indicado.")
            return 1

    where_sql = " AND ".join(where)
    rows = conn.execute(
        f"""
        SELECT
            h.id,
            h.ejercicio,
            TRIM(COALESCE(h.ente_uid, '')) AS ente_uid,
            TRIM(COALESCE(h.ente, '')) AS ente,
            TRIM(COALESCE(h.tipo_auditoria, '')) AS tipo_auditoria,
            TRIM(COALESCE(h.tipo_registro, '')) AS tipo_registro,
            TRIM(COALESCE(h.nombre, '')) AS nombre,
            h.fecha_inicio,
            h.fecha_fin
        FROM historial_titulares h
        WHERE {where_sql}
        ORDER BY h.ejercicio, h.ente, h.fecha_inicio, h.id
        """,
        params,
    ).fetchall()

    print("=== Correccion tipo_auditoria en historial_titulares ===")
    print(f"db: {args.db}")
    print(f"ejercicio: {ejercicio}")
    print(f"ente_id: {ente_id or '(todos)'}")
    print(f"from_type: {', '.join(from_types)}")
    print(f"to_type: {to_type}")
    print(f"filas_objetivo: {len(rows)}")
    print(f"modo: {'APPLY' if args.apply else 'DRY-RUN'}")

    if not rows:
        return 0

    print("")
    print("Muestra (max 20):")
    for item in rows[:20]:
        print(
            f"- id={item['id']} | ente={item['ente']} | "
            f"tipo={item['tipo_auditoria']} -> {to_type} | "
            f"registro={item['tipo_registro']} | "
            f"periodo={item['fecha_inicio']}..{item['fecha_fin']} | "
            f"nombre={item['nombre']}"
        )

    if not args.apply:
        return 0

    ids = [str(item["id"]) for item in rows]
    id_placeholders = ", ".join(["?"] * len(ids))
    conn.execute("BEGIN")
    conn.execute(
        f"""
        UPDATE historial_titulares
        SET tipo_auditoria = ?
        WHERE id IN ({id_placeholders})
        """,
        [to_type, *ids],
    )
    conn.commit()
    print("")
    print(f"Actualizacion aplicada. Filas actualizadas: {len(ids)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
