#!/usr/bin/env python3
"""
Import historial_titulares from JSON payloads like:
{
  "ejercicio": 2023,
  "nombre_ente": "...",
  "periodos_informe": [
    {
      "periodo": "01 de Enero al 31 de Diciembre",
      "titular": "...",
      "administrativo": [
        {
          "periodo": "01 de Enero al 13 de Abril",
          "director_administrativo": "..."
        }
      ]
    }
  ]
}
"""

from __future__ import annotations

import argparse
import difflib
import json
import re
import sqlite3
import sys
import unicodedata
from datetime import date
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from backup_utils import create_db_backup


PERIODO_RE = re.compile(
    r"^\s*(\d{1,2})\s+de\s+([A-Za-zÁÉÍÓÚáéíóúÜüÑñ]+)(?:\s+de\s+(\d{4}))?\s+al\s+(\d{1,2})\s+de\s+([A-Za-zÁÉÍÓÚáéíóúÜüÑñ]+)(?:\s+de\s+(\d{4}))?\s*$",
    re.IGNORECASE,
)
PERIODO_SHORT_RE = re.compile(
    r"^\s*(\d{1,2})\s+al\s+(\d{1,2})\s+de\s+([A-Za-zÁÉÍÓÚáéíóúÜüÑñ]+)(?:\s+de\s+(\d{4}))?\s*$",
    re.IGNORECASE,
)

MONTHS = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "setiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}


def _clean(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_text(value: str) -> str:
    value = _clean(value)
    if not value:
        return ""
    normalized = unicodedata.normalize("NFKD", value)
    normalized = "".join(char for char in normalized if not unicodedata.combining(char))
    normalized = re.sub(r"[^A-Za-z0-9]+", " ", normalized)
    return normalized.strip().lower()


def parse_periodo(periodo: str, ejercicio: int) -> tuple[str, str]:
    raw = _clean(periodo)
    match = PERIODO_RE.match(raw)
    if match:
        d1, m1, y1, d2, m2, y2 = match.groups()
        m1_num = MONTHS.get(normalize_text(m1))
        m2_num = MONTHS.get(normalize_text(m2))
        if not m1_num or not m2_num:
            raise ValueError(f"mes invalido en periodo: {raw!r}")

        if y1 and y2:
            yi = int(y1)
            yf = int(y2)
        elif y1 and not y2:
            yi = int(y1)
            yf = int(y1)
        elif y2 and not y1:
            yi = int(y2)
            yf = int(y2)
        else:
            yi = ejercicio
            yf = ejercicio

        fi = date(yi, m1_num, int(d1)).isoformat()
        ff = date(yf, m2_num, int(d2)).isoformat()
        return fi, ff

    short_match = PERIODO_SHORT_RE.match(raw)
    if short_match:
        d1, d2, m, y = short_match.groups()
        m_num = MONTHS.get(normalize_text(m))
        if not m_num:
            raise ValueError(f"mes invalido en periodo: {raw!r}")
        year = int(y) if y else ejercicio
        fi = date(year, m_num, int(d1)).isoformat()
        ff = date(year, m_num, int(d2)).isoformat()
        return fi, ff

    raise ValueError(f"periodo invalido: {raw!r}")


def resolve_ente_nombre(conn: sqlite3.Connection, ejercicio: int, nombre_ente: str) -> str:
    nombre_ente = _clean(nombre_ente)
    if not nombre_ente:
        raise ValueError("nombre_ente vacio")

    cur = conn.cursor()
    row = cur.execute(
        """
        SELECT ente_nombre
        FROM entes_detalle
        WHERE ejercicio = ? AND LOWER(TRIM(ente_nombre)) = LOWER(TRIM(?))
        LIMIT 1
        """,
        (str(ejercicio), nombre_ente),
    ).fetchone()
    if row is not None:
        return _clean(row[0])

    rows = cur.execute(
        """
        SELECT ente_nombre
        FROM entes_detalle
        WHERE ejercicio = ?
        """,
        (str(ejercicio),),
    ).fetchall()
    target = normalize_text(nombre_ente)
    normalized_rows = [(_clean(r[0]), normalize_text(_clean(r[0]))) for r in rows]

    matches = [raw_name for raw_name, norm_name in normalized_rows if norm_name == target]
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise ValueError(f"nombre_ente ambiguo para ejercicio {ejercicio}: {nombre_ente!r}")

    contains_matches = [
        raw_name
        for raw_name, norm_name in normalized_rows
        if target and (target in norm_name or norm_name in target)
    ]
    # Allow partial matching only when unique to avoid wrong inserts.
    if len(contains_matches) == 1:
        return contains_matches[0]
    if len(contains_matches) > 1:
        raise ValueError(
            f"nombre_ente coincide con multiples entes para ejercicio {ejercicio}: {nombre_ente!r}"
        )

    # Fuzzy fallback for small wording variations; only accept strong unique match.
    candidate_norms = [norm_name for _, norm_name in normalized_rows if norm_name]
    close = difflib.get_close_matches(target, candidate_norms, n=2, cutoff=0.86)
    if len(close) == 1:
        selected_norm = close[0]
        selected = [raw_name for raw_name, norm_name in normalized_rows if norm_name == selected_norm]
        if len(selected) == 1:
            return selected[0]

    raise ValueError(f"nombre_ente no existe en entes_detalle para ejercicio {ejercicio}: {nombre_ente!r}")


def row_exists(
    conn: sqlite3.Connection,
    ejercicio: int,
    ente: str,
    tipo_auditoria: str,
    nombre: str,
    cargo: str,
    fi: str,
    ff: str,
    tipo_registro: str,
) -> bool:
    row = conn.execute(
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
    return row is not None


def iter_payloads(raw: object) -> list[dict]:
    if isinstance(raw, dict):
        if (
            "periodos_informe" in raw
            or "periodos" in raw
            or ("periodo_informe" in raw and "titular" in raw)
            or (isinstance(raw.get("periodos_informe"), str) and "titular" in raw)
        ):
            return [raw]
        raise ValueError("JSON objeto no contiene 'periodos_informe', 'periodos' ni formato plano")
    if isinstance(raw, list):
        payloads = []
        for i, item in enumerate(raw, start=1):
            if not isinstance(item, dict):
                raise ValueError(f"item {i}: formato invalido (se esperaba objeto)")
            if "periodos_informe" not in item and "periodos" not in item:
                raise ValueError(
                    f"item {i}: formato invalido (se esperaba objeto con 'periodos_informe' o 'periodos')"
                )
            payloads.append(item)
        return payloads
    raise ValueError("JSON invalido: se esperaba objeto o lista")


def parse_rows(payload: dict, default_ejercicio: int | None, default_tipo_auditoria: str, conn: sqlite3.Connection) -> list[dict[str, str]]:
    ejercicio_raw = payload.get("ejercicio", default_ejercicio)
    if not ejercicio_raw:
        raise ValueError("falta 'ejercicio'")
    ejercicio = int(ejercicio_raw)

    nombre_ente = _clean(payload.get("nombre_ente"))
    ente = resolve_ente_nombre(conn, ejercicio, nombre_ente)

    periodos = payload.get("periodos_informe")
    periodos_key = "periodos_informe"
    if periodos is None and "periodos" in payload:
        periodos = payload.get("periodos")
        periodos_key = "periodos"
    if isinstance(periodos, str) and _clean(payload.get("titular")):
        periodos = [
            {
                "periodo": _clean(periodos),
                "titular": _clean(payload.get("titular")),
                "administrativo": payload.get("administrativo", []),
            }
        ]
        periodos_key = "periodo_plano"
    if periodos is None and _clean(payload.get("periodo_informe")) and _clean(payload.get("titular")):
        periodos = [
            {
                "periodo": _clean(payload.get("periodo_informe")),
                "titular": _clean(payload.get("titular")),
                "administrativo": payload.get("administrativo", []),
            }
        ]
        periodos_key = "periodo_plano"
    if not isinstance(periodos, list) or not periodos:
        raise ValueError("falta 'periodos_informe/periodos' o esta vacio")

    tipo_auditoria = (
        _clean(payload.get("tipo_auditoria"))
        or _clean(payload.get("area"))
        or _clean(payload.get("tipo"))
        or default_tipo_auditoria
    )

    rows: list[dict[str, str]] = []

    def parse_admin_block(admin_raw: object, source_label: str, idx_i: int, idx_j_start: int = 1) -> list[dict[str, str]]:
        if admin_raw is None:
            administrativos = []
        elif isinstance(admin_raw, dict):
            administrativos = [admin_raw]
        elif isinstance(admin_raw, list):
            administrativos = admin_raw
        else:
            raise ValueError(f"{source_label} debe ser objeto o lista")

        out: list[dict[str, str]] = []
        for j, admin in enumerate(administrativos, start=idx_j_start):
            if not isinstance(admin, dict):
                raise ValueError(f"{source_label}[{j}] debe ser objeto")

            director_saneamiento = _clean(admin.get("director_saneamiento"))
            director_agua_potable = _clean(admin.get("director_agua_potable"))
            if director_saneamiento or director_agua_potable:
                if director_saneamiento:
                    out.append(
                        {
                            "ejercicio": str(ejercicio),
                            "ente": ente,
                            "tipo_auditoria": tipo_auditoria,
                            "nombre": director_saneamiento,
                            "cargo": "Director de Saneamiento",
                            "fecha_inicio": f"{ejercicio}-01-01",
                            "fecha_fin": f"{ejercicio}-06-30",
                            "tipo_registro": "director_administrativo",
                        }
                    )
                if director_agua_potable:
                    out.append(
                        {
                            "ejercicio": str(ejercicio),
                            "ente": ente,
                            "tipo_auditoria": tipo_auditoria,
                            "nombre": director_agua_potable,
                            "cargo": "Director de Agua Potable",
                            "fecha_inicio": f"{ejercicio}-07-01",
                            "fecha_fin": f"{ejercicio}-12-31",
                            "tipo_registro": "director_administrativo",
                        }
                    )
                continue

            admin_nombre = (
                _clean(admin.get("director_administrativo"))
                or _clean(admin.get("director_obras"))
                or _clean(admin.get("jefe_departamento_ejecucion_supervision"))
                or _clean(admin.get("jefe_infraestructura"))
            )
            admin_periodo = _clean(admin.get("periodo"))
            # Allow placeholder blocks with missing name to pass through without insertion.
            if not admin_nombre and admin_periodo:
                continue
            if not admin_nombre or not admin_periodo:
                raise ValueError(
                    f"{source_label}[{j}] requiere 'periodo' y "
                    f"'director_administrativo/director_obras/jefe_departamento_ejecucion_supervision/jefe_infraestructura'"
                )
            admin_fi, admin_ff = parse_periodo(admin_periodo, ejercicio)
            admin_cargo = _clean(admin.get("cargo"))
            if not admin_cargo:
                admin_cargo = "Director Administrativo"
                if _clean(admin.get("director_obras")):
                    admin_cargo = "Director de Obras"
                elif _clean(admin.get("jefe_departamento_ejecucion_supervision")):
                    admin_cargo = "Director de Vivienda"
                elif _clean(admin.get("jefe_infraestructura")):
                    admin_cargo = "Jefe de Infraestructura"
            out.append(
                {
                    "ejercicio": str(ejercicio),
                    "ente": ente,
                    "tipo_auditoria": tipo_auditoria,
                    "nombre": admin_nombre,
                    "cargo": admin_cargo,
                    "fecha_inicio": admin_fi,
                    "fecha_fin": admin_ff,
                    "tipo_registro": "director_administrativo",
                }
            )
        return out

    for i, periodo_item in enumerate(periodos, start=1):
        if not isinstance(periodo_item, dict):
            raise ValueError(f"{periodos_key}[{i}] debe ser objeto")

        periodo_titular = _clean(periodo_item.get("periodo")) or _clean(periodo_item.get("periodo_informe"))
        titular = _clean(periodo_item.get("titular"))
        if not periodo_titular or not titular:
            raise ValueError(f"{periodos_key}[{i}] requiere 'periodo/periodo_informe' y 'titular'")
        fi, ff = parse_periodo(periodo_titular, ejercicio)
        rows.append(
            {
                "ejercicio": str(ejercicio),
                "ente": ente,
                "tipo_auditoria": tipo_auditoria,
                "nombre": titular,
                "cargo": "Titular",
                "fecha_inicio": fi,
                "fecha_fin": ff,
                "tipo_registro": "titular",
            }
        )

        administrativos_raw = periodo_item.get("administrativo")
        admin_source = "administrativo"
        if administrativos_raw is None:
            administrativos_raw = periodo_item.get("obra_publica")
            admin_source = "obra_publica"
        if administrativos_raw is not None:
            rows.extend(parse_admin_block(administrativos_raw, f"{periodos_key}[{i}].{admin_source}", i))

    # Some payloads provide administrativo/obra_publica once at root level.
    if _clean(payload.get("administrativo")) or isinstance(payload.get("administrativo"), (dict, list)):
        rows.extend(parse_admin_block(payload.get("administrativo"), "administrativo", 0))
    elif _clean(payload.get("obra_publica")) or isinstance(payload.get("obra_publica"), (dict, list)):
        rows.extend(parse_admin_block(payload.get("obra_publica"), "obra_publica", 0))
    return rows


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--db", default="sifeet.db", help="SQLite DB path (default: sifeet.db)")
    parser.add_argument("--json", nargs="+", required=True, help="JSON file path(s)")
    parser.add_argument("--ejercicio", type=int, help="Fallback ejercicio if missing in JSON")
    parser.add_argument("--tipo-auditoria", default="Financiera", help="Default tipo_auditoria")
    parser.add_argument("--dry-run", action="store_true", help="Validate and report; do not write")
    args = parser.parse_args()

    conn = sqlite3.connect(args.db)
    conn.row_factory = sqlite3.Row

    parsed_rows: list[dict[str, str]] = []
    for json_path in args.json:
        path = Path(json_path)
        with path.open("r", encoding="utf-8") as f:
            raw = json.load(f)
        payloads = iter_payloads(raw)
        for payload in payloads:
            parsed_rows.extend(parse_rows(payload, args.ejercicio, args.tipo_auditoria, conn))

    if parsed_rows and not args.dry_run:
        backup_path = create_db_backup(args.db, label="before_historial_json_import")
        print(f"Respaldo creado: {backup_path}")

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
    print(f"{mode}: filas={len(parsed_rows)} insertadas={inserted} omitidas_por_duplicado={skipped}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
