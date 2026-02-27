#!/usr/bin/env python3
"""
Diagnostica por que no aparecen titulares/administrativos en
`/observaciones-responsables` para un ejercicio.

Replica la logica principal del endpoint:
- cruza observaciones por ejercicio/ente/tipo/periodo_cedula
- busca historial_titulares por ente (uid o alias), tipo_auditoria y traslape de fechas
"""

from __future__ import annotations

import argparse
import re
import sqlite3
import unicodedata
from datetime import datetime

MONTHS = {
    "enero": "01",
    "febrero": "02",
    "marzo": "03",
    "abril": "04",
    "mayo": "05",
    "junio": "06",
    "julio": "07",
    "agosto": "08",
    "septiembre": "09",
    "setiembre": "09",
    "octubre": "10",
    "noviembre": "11",
    "diciembre": "12",
}


def normalize_text_key(value: str) -> str:
    raw = (value or "").strip().lower()
    if not raw:
        return ""
    raw = unicodedata.normalize("NFKD", raw)
    raw = "".join(ch for ch in raw if not unicodedata.combining(ch))
    raw = re.sub(r"\s+", " ", raw).strip()
    return raw


def normalize_tipo_auditoria(value: str) -> str:
    clean = (value or "").strip()
    key = normalize_text_key(clean)
    if key in {"auditoria", "auditoria financiera", "financiera", "financiero"}:
        return "Financiera"
    if key in {"obra publica", "obra"}:
        return "Obra Pública"
    if key == "cuenta publica":
        return "Cuenta Pública"
    return clean


def parse_periodo_cedula(ejercicio: str, periodo_cedula: str):
    if not ejercicio or not periodo_cedula:
        return None, None
    match = re.match(
        r"^\s*(\d{1,2})\s+de\s+([a-zA-ZáéíóúÁÉÍÓÚñÑ]+)\s+al\s+(\d{1,2})\s+de\s+([a-zA-ZáéíóúÁÉÍÓÚñÑ]+)\s*$",
        periodo_cedula,
    )
    if not match:
        return None, None

    start_day = int(match.group(1))
    start_month = MONTHS.get(normalize_text_key(match.group(2)))
    end_day = int(match.group(3))
    end_month = MONTHS.get(normalize_text_key(match.group(4)))
    if not start_month or not end_month:
        return None, None

    try:
        year = int(str(ejercicio).strip())
        start = datetime(year, int(start_month), start_day).strftime("%Y-%m-%d")
        end = datetime(year, int(end_month), end_day).strftime("%Y-%m-%d")
    except ValueError:
        return None, None
    return start, end


def normalize_ente_id(value: str) -> str:
    return (value or "").strip().rstrip(".").strip()


def get_aliases(conn: sqlite3.Connection, ejercicio: str, ente_id: str, fallback: list[str]) -> list[str]:
    aliases: list[str] = []
    for name in fallback:
        clean = (name or "").strip()
        if clean and clean not in aliases:
            aliases.append(clean)

    row = conn.execute(
        """
        SELECT TRIM(COALESCE(ente_uid, '')) AS ente_uid, TRIM(COALESCE(ente_nombre, '')) AS ente_nombre
        FROM entes_detalle
        WHERE TRIM(COALESCE(ejercicio, '')) = ?
          AND TRIM(RTRIM(COALESCE(ente_id, ''), '.')) = ?
        LIMIT 1
        """,
        (ejercicio, ente_id),
    ).fetchone()
    if not row:
        return aliases
    if row["ente_nombre"] and row["ente_nombre"] not in aliases:
        aliases.append(row["ente_nombre"])

    ente_uid = row["ente_uid"]
    if not ente_uid:
        return aliases

    uid_rows = conn.execute(
        """
        SELECT DISTINCT TRIM(COALESCE(ente_nombre, '')) AS ente_nombre
        FROM entes_detalle
        WHERE TRIM(COALESCE(ente_uid, '')) = ?
          AND TRIM(COALESCE(ente_nombre, '')) != ''
        ORDER BY ejercicio
        """,
        (ente_uid,),
    ).fetchall()
    for item in uid_rows:
        name = item["ente_nombre"]
        if name and name not in aliases:
            aliases.append(name)
    return aliases


def get_uid(conn: sqlite3.Connection, ejercicio: str, ente_id: str) -> str:
    row = conn.execute(
        """
        SELECT TRIM(COALESCE(ente_uid, '')) AS ente_uid
        FROM entes_detalle
        WHERE TRIM(COALESCE(ejercicio, '')) = ?
          AND TRIM(RTRIM(COALESCE(ente_id, ''), '.')) = ?
        LIMIT 1
        """,
        (ejercicio, ente_id),
    ).fetchone()
    return (row["ente_uid"] if row else "") or ""


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--db", default="sifeet.db")
    parser.add_argument("--ejercicio", required=True)
    parser.add_argument("--ente-id", default="")
    parser.add_argument("--max-examples", type=int, default=25)
    args = parser.parse_args()

    conn = sqlite3.connect(args.db)
    conn.row_factory = sqlite3.Row

    params = [args.ejercicio]
    ente_sql = ""
    ente_norm = normalize_ente_id(args.ente_id)
    if ente_norm:
        ente_sql = "AND TRIM(RTRIM(COALESCE(o.ente_id, ''), '.')) = ?"
        params.append(ente_norm)

    scopes = conn.execute(
        f"""
        SELECT DISTINCT
            TRIM(COALESCE(o.ejercicio, '')) AS ejercicio,
            TRIM(COALESCE(o.ente_id, '')) AS ente_id,
            TRIM(COALESCE(o.ente_nombre, '')) AS ente_nombre,
            TRIM(COALESCE(ed.ente_nombre, '')) AS ente_detalle_nombre,
            TRIM(COALESCE(ed.ente_uid, '')) AS ente_uid,
            TRIM(COALESCE(o.tipo_auditoria, '')) AS tipo_auditoria,
            TRIM(COALESCE(o.periodo_cedula, '')) AS periodo_cedula
        FROM observaciones o
        LEFT JOIN entes_detalle ed
          ON TRIM(RTRIM(COALESCE(o.ente_id, ''), '.')) = TRIM(RTRIM(COALESCE(ed.ente_id, ''), '.'))
         AND TRIM(COALESCE(o.ejercicio, '')) = TRIM(COALESCE(ed.ejercicio, ''))
        WHERE TRIM(COALESCE(o.ejercicio, '')) = ?
          AND TRIM(COALESCE(o.periodo_cedula, '')) != ''
          {ente_sql}
        ORDER BY o.ente_nombre, o.tipo_auditoria, o.periodo_cedula
        """,
        params,
    ).fetchall()

    total = 0
    parse_fail = 0
    ok = 0
    missing_same_type = 0
    missing_any_type = 0
    mismatch_tipo = 0
    examples = []

    for row in scopes:
        total += 1
        ini, fin = parse_periodo_cedula(row["ejercicio"], row["periodo_cedula"])
        if not ini or not fin:
            parse_fail += 1
            continue

        ente_id = normalize_ente_id(row["ente_id"])
        aliases = get_aliases(
            conn,
            row["ejercicio"],
            ente_id,
            [row["ente_nombre"], row["ente_detalle_nombre"]],
        )
        ente_uid = get_uid(conn, row["ejercicio"], ente_id) or row["ente_uid"]
        if not ente_uid and not aliases:
            missing_any_type += 1
            if len(examples) < args.max_examples:
                examples.append(
                    {
                        "ente_id": row["ente_id"],
                        "ente_nombre": row["ente_nombre"],
                        "tipo_obs": row["tipo_auditoria"],
                        "periodo_cedula": row["periodo_cedula"],
                        "motivo": "sin_uid_ni_alias",
                        "tipos_historial_en_rango": "",
                    }
                )
            continue

        scope_clause = ""
        scope_params: list[str] = []
        if ente_uid and aliases:
            placeholders = ", ".join(["?"] * len(aliases))
            scope_clause = (
                f"AND (TRIM(COALESCE(h.ente_uid, '')) = ? OR TRIM(COALESCE(h.ente, '')) IN ({placeholders}))"
            )
            scope_params.extend([ente_uid, *aliases])
        elif ente_uid:
            scope_clause = "AND TRIM(COALESCE(h.ente_uid, '')) = ?"
            scope_params.append(ente_uid)
        else:
            placeholders = ", ".join(["?"] * len(aliases))
            scope_clause = f"AND TRIM(COALESCE(h.ente, '')) IN ({placeholders})"
            scope_params.extend(aliases)

        overlap_rows = conn.execute(
            f"""
            SELECT
                TRIM(COALESCE(h.tipo_auditoria, '')) AS tipo_auditoria,
                COUNT(*) AS total
            FROM historial_titulares h
            WHERE TRIM(COALESCE(h.ejercicio, '')) = ?
              {scope_clause}
              AND h.tipo_registro IN ('titular', 'director_administrativo')
              AND date(h.fecha_inicio) <= date(?)
              AND date(h.fecha_fin) >= date(?)
            GROUP BY TRIM(COALESCE(h.tipo_auditoria, ''))
            """,
            [row["ejercicio"], *scope_params, fin, ini],
        ).fetchall()

        same_type = 0
        tipos_historial = []
        for item in overlap_rows:
            tipo = item["tipo_auditoria"]
            tipo_norm = normalize_tipo_auditoria(tipo)
            if tipo:
                tipos_historial.append(tipo)
            if tipo_norm == normalize_tipo_auditoria(row["tipo_auditoria"]):
                same_type += int(item["total"] or 0)

        any_type = sum(int(item["total"] or 0) for item in overlap_rows)
        if same_type > 0:
            ok += 1
            continue

        missing_same_type += 1
        if any_type == 0:
            missing_any_type += 1
            motivo = "sin_historial_en_rango"
        else:
            mismatch_tipo += 1
            motivo = "mismatch_tipo_auditoria"

        if len(examples) < args.max_examples:
            examples.append(
                {
                    "ente_id": row["ente_id"],
                    "ente_nombre": row["ente_nombre"],
                    "tipo_obs": row["tipo_auditoria"],
                    "periodo_cedula": row["periodo_cedula"],
                    "motivo": motivo,
                    "tipos_historial_en_rango": ", ".join(sorted(set(tipos_historial))),
                }
            )

    print("=== Diagnostico Responsables por Periodo ===")
    print(f"db: {args.db}")
    print(f"ejercicio: {args.ejercicio}")
    print(f"filtro_ente_id: {ente_norm or '(todos)'}")
    print("")
    print(f"scopes_observaciones: {total}")
    print(f"periodos_no_parseables: {parse_fail}")
    print(f"scopes_con_match_completo: {ok}")
    print(f"scopes_sin_match_mismo_tipo: {missing_same_type}")
    print(f"scopes_sin_historial_en_rango: {missing_any_type}")
    print(f"scopes_con_historial_pero_tipo_distinto: {mismatch_tipo}")

    if examples:
        print("")
        print("=== Ejemplos ===")
        for item in examples:
            print(
                "- "
                f"ente_id={item['ente_id']} | ente={item['ente_nombre']} | "
                f"tipo_obs={item['tipo_obs']} | cedula={item['periodo_cedula']} | "
                f"motivo={item['motivo']} | tipos_historial={item['tipos_historial_en_rango'] or '-'}"
            )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
