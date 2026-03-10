from datetime import datetime
import json
import re
import sys

from flask import jsonify, render_template, request


def split_periodo_tokens(raw_value: str) -> list[str]:
    raw = (raw_value or "").strip()
    if not raw:
        return []
    tokens = re.split(r"(?:\r?\n|;|\|)+", raw)
    return [" ".join((token or "").strip().split()) for token in tokens if (token or "").strip()]


def normalize_periodo_key(ejercicio: str, periodo: str, *, label: str, strict: bool = True) -> str:
    clean = " ".join((periodo or "").strip().split())
    if not clean:
        return ""
    fecha_inicio, fecha_fin = parse_periodo_cedula(ejercicio, clean)
    if not fecha_inicio or not fecha_fin:
        if strict:
            raise ValueError(
                f"Titulares: {label} debe usar formato '01 de enero al 31 de diciembre'."
            )
        return clean.lower()
    return f"{fecha_inicio}|{fecha_fin}"


def parse_cedula_periodos(
    ejercicio: str,
    raw_value: str,
    *,
    strict: bool = True,
) -> tuple[list[str], list[str]]:
    tokens = split_periodo_tokens(raw_value)
    periodos: list[str] = []
    keys: list[str] = []
    seen = set()
    for periodo in tokens:
        key = normalize_periodo_key(
            ejercicio,
            periodo,
            label="cada partición de cédula",
            strict=strict,
        )
        if key in seen:
            continue
        seen.add(key)
        periodos.append(periodo)
        keys.append(key)
    return periodos, keys


def register_gabo_routes(app, deps):
    globals().update(deps)
    TITULAR_EJERCICIO_FIJO = "2025"
    OBSERVACION_ESTADOS_VALIDOS = {"Emitido", "Pendiente", "Solventado"}
    TITULAR_MONTHS_ES = {
        "01": "enero",
        "02": "febrero",
        "03": "marzo",
        "04": "abril",
        "05": "mayo",
        "06": "junio",
        "07": "julio",
        "08": "agosto",
        "09": "septiembre",
        "10": "octubre",
        "11": "noviembre",
        "12": "diciembre",
    }
    def _tipo_auditoria_options(tipo_auditoria: str) -> list[str]:
        clean = " ".join((tipo_auditoria or "").split())
        if clean == "Financiera y Obra Pública":
            return ["Financiera", "Obra Pública"]
        if clean in {"Financiera", "Obra Pública"}:
            return [clean]
        return []

    def _normalize_observacion_estado(value: str) -> str:
        raw = " ".join((value or "").split())
        key = raw.lower()
        if key in {"e", "emitido"}:
            return "Emitido"
        if key in {"p", "pendiente"}:
            return "Pendiente"
        if key in {"s", "solventado"}:
            return "Solventado"
        return raw

    def _build_titular_form_data(source, *, default_ejercicio: str) -> dict[str, str]:
        ejercicio = " ".join(
            (
                source.get("titular_ejercicio")
                or source.get("ejercicio")
                or default_ejercicio
                or ""
            ).split()
        )
        form_data = {
            "titular_ejercicio": ejercicio,
            "titular_ente_id": normalize_ente_id(
                source.get("titular_ente_id") or source.get("ente_id") or ""
            ),
            "titular_tipo_auditoria": normalize_tipo_auditoria(
                source.get("titular_tipo_auditoria")
                or source.get("tipo_auditoria")
                or "Financiera"
            ) or "Financiera",
            "titular_periodo_informe": " ".join(
                (source.get("titular_periodo_informe") or "").split()
            ),
            "titular_nombre": " ".join((source.get("titular_nombre") or "").split()),
            "titular_periodo_administrativo": " ".join(
                (source.get("titular_periodo_administrativo") or "").split()
            ),
            "titular_administrativo": " ".join(
                (source.get("titular_administrativo") or "").split()
            ),
            "titular_fecha_inicio": " ".join(
                (source.get("titular_fecha_inicio") or "").split()
            ),
            "titular_fecha_fin": " ".join(
                (source.get("titular_fecha_fin") or "").split()
            ),
            "titular_admin_fecha_inicio": " ".join(
                (source.get("titular_admin_fecha_inicio") or "").split()
            ),
            "titular_admin_fecha_fin": " ".join(
                (source.get("titular_admin_fecha_fin") or "").split()
            ),
            "titular_admin_mismo_periodo": (
                "1" if (source.get("titular_admin_mismo_periodo") or "").strip() else "0"
            ),
            "titular_cedula_resultados": " ".join(
                (source.get("titular_cedula_resultados") or "").split()
            ),
        }
        if form_data["titular_periodo_informe"] and (
            not form_data["titular_fecha_inicio"] or not form_data["titular_fecha_fin"]
        ):
            fecha_inicio, fecha_fin = parse_periodo_cedula(
                ejercicio,
                form_data["titular_periodo_informe"],
            )
            if fecha_inicio and fecha_fin:
                form_data["titular_fecha_inicio"] = fecha_inicio
                form_data["titular_fecha_fin"] = fecha_fin
        if form_data["titular_periodo_administrativo"] and (
            not form_data["titular_admin_fecha_inicio"] or not form_data["titular_admin_fecha_fin"]
        ):
            fecha_inicio, fecha_fin = parse_periodo_cedula(
                ejercicio,
                form_data["titular_periodo_administrativo"],
            )
            if fecha_inicio and fecha_fin:
                form_data["titular_admin_fecha_inicio"] = fecha_inicio
                form_data["titular_admin_fecha_fin"] = fecha_fin
        if (
            form_data["titular_admin_mismo_periodo"] != "1"
            and form_data["titular_fecha_inicio"]
            and form_data["titular_fecha_fin"]
            and form_data["titular_admin_fecha_inicio"] == form_data["titular_fecha_inicio"]
            and form_data["titular_admin_fecha_fin"] == form_data["titular_fecha_fin"]
        ):
            form_data["titular_admin_mismo_periodo"] = "1"
        return form_data

    def _format_periodo_label_from_dates(fecha_inicio_iso: str, fecha_fin_iso: str) -> str:
        fecha_inicio = parse_historial_date(fecha_inicio_iso)
        fecha_fin = parse_historial_date(fecha_fin_iso)
        if not fecha_inicio or not fecha_fin:
            return ""
        mes_inicio = TITULAR_MONTHS_ES.get(fecha_inicio.strftime("%m"), "")
        mes_fin = TITULAR_MONTHS_ES.get(fecha_fin.strftime("%m"), "")
        if not mes_inicio or not mes_fin:
            return ""
        return (
            f"{fecha_inicio.day:02d} de {mes_inicio} al "
            f"{fecha_fin.day:02d} de {mes_fin}"
        )

    def _resolve_period_range(
        ejercicio: str,
        *,
        fecha_inicio_raw: str,
        fecha_fin_raw: str,
        periodo_raw: str,
        label: str,
    ) -> tuple[str, str, str]:
        fecha_inicio_clean = " ".join((fecha_inicio_raw or "").split())
        fecha_fin_clean = " ".join((fecha_fin_raw or "").split())
        periodo_clean = " ".join((periodo_raw or "").split())
        if fecha_inicio_clean or fecha_fin_clean:
            fecha_inicio = parse_historial_date(fecha_inicio_clean)
            fecha_fin = parse_historial_date(fecha_fin_clean)
            if not fecha_inicio or not fecha_fin:
                raise ValueError(f"Titulares: {label} requiere fecha inicial y final válidas.")
            if fecha_inicio > fecha_fin:
                raise ValueError(f"Titulares: {label} no puede tener fecha inicial mayor a la final.")
            fecha_inicio_iso = fecha_inicio.isoformat()
            fecha_fin_iso = fecha_fin.isoformat()
            return (
                fecha_inicio_iso,
                fecha_fin_iso,
                _format_periodo_label_from_dates(fecha_inicio_iso, fecha_fin_iso),
            )
        if not periodo_clean:
            raise ValueError(f"Titulares: {label} requerido.")
        fecha_inicio_iso, fecha_fin_iso = parse_periodo_cedula(
            ejercicio,
            periodo_clean,
        )
        if not fecha_inicio_iso or not fecha_fin_iso:
            raise ValueError(
                f"Titulares: {label} debe usar formato '01 de enero al 31 de diciembre'."
            )
        return fecha_inicio_iso, fecha_fin_iso, periodo_clean

    def _load_titular_entes(db, ejercicio: str) -> list[dict]:
        if not ejercicio:
            return []
        rows = db.execute(
            """
            SELECT
                TRIM(COALESCE(ente_id, '')) AS ente_id,
                TRIM(COALESCE(ente_numero, '')) AS ente_numero,
                TRIM(COALESCE(ente_nombre, '')) AS ente_nombre
            FROM entes_detalle
            WHERE TRIM(COALESCE(ejercicio, '')) = ?
              AND TRIM(COALESCE(ente_id, '')) != ''
            ORDER BY CAST(COALESCE(NULLIF(ente_numero, ''), '0') AS REAL), ente_numero, ente_nombre
            """,
            (ejercicio,),
        ).fetchall()
        return [dict(row) for row in rows]

    def _get_ente_row_by_ejercicio_id(db, ejercicio: str, ente_id_norm: str):
        if not ejercicio or not ente_id_norm:
            return None
        return db.execute(
            f"""
            SELECT
                TRIM(COALESCE(ente_id, '')) AS ente_id,
                TRIM(COALESCE(ente_numero, '')) AS ente_numero,
                TRIM(COALESCE(ente_nombre, '')) AS ente_nombre,
                TRIM(COALESCE(ente_uid, '')) AS ente_uid
            FROM entes_detalle
            WHERE TRIM(COALESCE(ejercicio, '')) = ?
              AND {normalize_ente_id_sql('ente_id')} = ?
            LIMIT 1
            """,
            (ejercicio, ente_id_norm),
        ).fetchone()

    def _list_historial_titulares_rows(
        db,
        *,
        ejercicio: str,
        ente_id_norm: str = "",
        tipo_auditoria: str = "",
    ) -> list[dict]:
        if not ejercicio:
            return []

        where_clauses = ["CAST(h.ejercicio AS TEXT) = ?"]
        params: list[str] = [ejercicio]

        if tipo_auditoria:
            where_clauses.append("TRIM(COALESCE(h.tipo_auditoria, '')) = ?")
            params.append(tipo_auditoria)

        if ente_id_norm:
            ente_row = _get_ente_row_by_ejercicio_id(db, ejercicio, ente_id_norm)
            if not ente_row:
                return []
            ente_uid = (ente_row["ente_uid"] or "").strip()
            ente_nombre = (ente_row["ente_nombre"] or "").strip()
            if ente_uid:
                where_clauses.append(
                    "(TRIM(COALESCE(h.ente_uid, '')) = ? OR TRIM(COALESCE(h.ente, '')) = ?)"
                )
                params.extend([ente_uid, ente_nombre])
            else:
                where_clauses.append("TRIM(COALESCE(h.ente, '')) = ?")
                params.append(ente_nombre)

        where_sql = " AND ".join(where_clauses)
        rows = db.execute(
            f"""
            SELECT
                h.id,
                CAST(h.ejercicio AS TEXT) AS ejercicio,
                TRIM(COALESCE(ed.ente_id, '')) AS ente_id,
                TRIM(COALESCE(ed.ente_numero, '')) AS ente_numero,
                TRIM(COALESCE(ed.ente_nombre, h.ente, '')) AS ente_nombre,
                TRIM(COALESCE(h.tipo_auditoria, '')) AS tipo_auditoria,
                TRIM(COALESCE(h.nombre, '')) AS nombre,
                TRIM(COALESCE(h.cargo, '')) AS cargo,
                TRIM(COALESCE(h.fecha_inicio, '')) AS fecha_inicio,
                TRIM(COALESCE(h.fecha_fin, '')) AS fecha_fin,
                TRIM(COALESCE(h.tipo_registro, '')) AS tipo_registro
            FROM historial_titulares AS h
            LEFT JOIN entes_detalle AS ed
              ON ed.id = (
                SELECT ed2.id
                FROM entes_detalle AS ed2
                WHERE TRIM(COALESCE(ed2.ejercicio, '')) = CAST(h.ejercicio AS TEXT)
                  AND (
                    (
                      TRIM(COALESCE(h.ente_uid, '')) != ''
                      AND TRIM(COALESCE(ed2.ente_uid, '')) = TRIM(COALESCE(h.ente_uid, ''))
                    )
                    OR TRIM(COALESCE(ed2.ente_nombre, '')) = TRIM(COALESCE(h.ente, ''))
                  )
                ORDER BY
                  CASE
                    WHEN TRIM(COALESCE(h.ente_uid, '')) != ''
                     AND TRIM(COALESCE(ed2.ente_uid, '')) = TRIM(COALESCE(h.ente_uid, ''))
                    THEN 0
                    ELSE 1
                  END,
                  CAST(COALESCE(NULLIF(ed2.ente_numero, ''), '0') AS REAL),
                  ed2.id
                LIMIT 1
              )
            WHERE {where_sql}
            ORDER BY
              CAST(COALESCE(NULLIF(ed.ente_numero, ''), '0') AS REAL) ASC,
              ed.ente_numero ASC,
              ente_nombre ASC,
              h.tipo_auditoria ASC,
              h.tipo_registro ASC,
              h.fecha_inicio DESC,
              h.id DESC
            """,
            params,
        ).fetchall()
        return [dict(row) for row in rows]

    def _upsert_historial_titular(
        db,
        *,
        ejercicio: str,
        ente_uid: str,
        ente_nombre: str,
        tipo_auditoria: str,
        nombre: str,
        cargo: str,
        fecha_inicio: str,
        fecha_fin: str,
        tipo_registro: str,
    ) -> str:
        if ente_uid:
            existing = db.execute(
                """
                SELECT
                    id,
                    TRIM(COALESCE(ente_uid, '')) AS ente_uid,
                    TRIM(COALESCE(ente, '')) AS ente,
                    TRIM(COALESCE(tipo_auditoria, '')) AS tipo_auditoria,
                    TRIM(COALESCE(nombre, '')) AS nombre,
                    TRIM(COALESCE(cargo, '')) AS cargo
                FROM historial_titulares
                WHERE CAST(ejercicio AS TEXT) = ?
                  AND tipo_auditoria = ?
                  AND tipo_registro = ?
                  AND fecha_inicio = ?
                  AND fecha_fin = ?
                  AND (
                    TRIM(COALESCE(ente_uid, '')) = ?
                    OR TRIM(COALESCE(ente, '')) = ?
                  )
                ORDER BY id DESC
                LIMIT 1
                """,
                (
                    ejercicio,
                    tipo_auditoria,
                    tipo_registro,
                    fecha_inicio,
                    fecha_fin,
                    ente_uid,
                    ente_nombre,
                ),
            ).fetchone()
        else:
            existing = db.execute(
                """
                SELECT
                    id,
                    TRIM(COALESCE(ente_uid, '')) AS ente_uid,
                    TRIM(COALESCE(ente, '')) AS ente,
                    TRIM(COALESCE(tipo_auditoria, '')) AS tipo_auditoria,
                    TRIM(COALESCE(nombre, '')) AS nombre,
                    TRIM(COALESCE(cargo, '')) AS cargo
                FROM historial_titulares
                WHERE CAST(ejercicio AS TEXT) = ?
                  AND tipo_auditoria = ?
                  AND tipo_registro = ?
                  AND fecha_inicio = ?
                  AND fecha_fin = ?
                  AND TRIM(COALESCE(ente, '')) = ?
                ORDER BY id DESC
                LIMIT 1
                """,
                (
                    ejercicio,
                    tipo_auditoria,
                    tipo_registro,
                    fecha_inicio,
                    fecha_fin,
                    ente_nombre,
                ),
            ).fetchone()

        if existing:
            changed = (
                (existing["ente_uid"] or "").strip() != (ente_uid or "").strip()
                or (existing["ente"] or "").strip() != ente_nombre
                or (existing["tipo_auditoria"] or "").strip() != tipo_auditoria
                or (existing["nombre"] or "").strip() != nombre
                or (existing["cargo"] or "").strip() != cargo
            )
            db.execute(
                """
                UPDATE historial_titulares
                SET ente_uid = ?,
                    ente = ?,
                    tipo_auditoria = ?,
                    nombre = ?,
                    cargo = ?
                WHERE id = ?
                """,
                (
                    ente_uid or None,
                    ente_nombre,
                    tipo_auditoria,
                    nombre,
                    cargo,
                    existing["id"],
                ),
            )
            return "updated" if changed else "unchanged"

        db.execute(
            """
            INSERT INTO historial_titulares (
                ejercicio, ente_uid, ente, tipo_auditoria, nombre, cargo, fecha_inicio, fecha_fin, tipo_registro
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                int(ejercicio),
                ente_uid or None,
                ente_nombre,
                tipo_auditoria,
                nombre,
                cargo,
                fecha_inicio,
                fecha_fin,
                tipo_registro,
            ),
        )
        return "inserted"

    def _save_titulares_capture(db, user, form_data: dict[str, str]) -> dict[str, object]:
        titular_ejercicio = " ".join((form_data.get("titular_ejercicio") or "").split())
        titular_ente_id = normalize_ente_id(form_data.get("titular_ente_id", ""))
        titular_tipo_auditoria = normalize_tipo_auditoria(
            form_data.get("titular_tipo_auditoria", "")
        )
        titular_nombre = " ".join((form_data.get("titular_nombre") or "").split())
        titular_administrativo = " ".join(
            (form_data.get("titular_administrativo") or "").split()
        )
        titular_admin_mismo_periodo = (
            "1" if (form_data.get("titular_admin_mismo_periodo") or "").strip() else "0"
        )
        titular_cedula_raw = " ".join(
            (form_data.get("titular_cedula_resultados") or "").split()
        )

        if not titular_ejercicio:
            raise ValueError("Titulares: ejercicio requerido.")
        ejercicio_exists = db.execute(
            """
            SELECT 1
            FROM entes_detalle
            WHERE TRIM(COALESCE(ejercicio, '')) = ?
            LIMIT 1
            """,
            (titular_ejercicio,),
        ).fetchone()
        if not ejercicio_exists:
            raise ValueError("Titulares: el ejercicio seleccionado no está disponible.")
        if not titular_ente_id:
            raise ValueError("Titulares: selecciona un ente.")
        if titular_tipo_auditoria not in {"Financiera", "Obra Pública"}:
            raise ValueError("Titulares: tipo de auditoría inválido.")
        if not titular_nombre:
            raise ValueError("Titulares: nombre del titular requerido.")
        if not titular_administrativo:
            raise ValueError("Titulares: nombre del administrativo requerido.")

        titular_inicio, titular_fin, titular_periodo_informe = _resolve_period_range(
            titular_ejercicio,
            fecha_inicio_raw=form_data.get("titular_fecha_inicio", ""),
            fecha_fin_raw=form_data.get("titular_fecha_fin", ""),
            periodo_raw=form_data.get("titular_periodo_informe", ""),
            label="periodo del titular",
        )
        admin_fecha_inicio_raw = form_data.get("titular_admin_fecha_inicio", "")
        admin_fecha_fin_raw = form_data.get("titular_admin_fecha_fin", "")
        admin_periodo_raw = form_data.get("titular_periodo_administrativo", "")
        if titular_admin_mismo_periodo == "1":
            admin_fecha_inicio_raw = titular_inicio
            admin_fecha_fin_raw = titular_fin
            admin_periodo_raw = titular_periodo_informe
        admin_inicio, admin_fin, titular_periodo_administrativo = _resolve_period_range(
            titular_ejercicio,
            fecha_inicio_raw=admin_fecha_inicio_raw,
            fecha_fin_raw=admin_fecha_fin_raw,
            periodo_raw=admin_periodo_raw,
            label="periodo del administrativo",
        )

        titular_periodo_informe_key = f"{titular_inicio}|{titular_fin}"
        if titular_cedula_raw:
            titular_cedula_periodos, titular_cedula_keys = parse_cedula_periodos(
                titular_ejercicio,
                titular_cedula_raw,
            )
        else:
            titular_cedula_periodos = [titular_periodo_informe]
            titular_cedula_keys = [titular_periodo_informe_key]
        titular_cedula_resultados = " | ".join(titular_cedula_periodos)

        ente_row = db.execute(
            f"""
            SELECT
                TRIM(COALESCE(ente_nombre, '')) AS ente_nombre,
                TRIM(COALESCE(ente_uid, '')) AS ente_uid
            FROM entes_detalle
            WHERE TRIM(COALESCE(ejercicio, '')) = ?
              AND {normalize_ente_id_sql('ente_id')} = ?
            LIMIT 1
            """,
            (titular_ejercicio, titular_ente_id),
        ).fetchone()
        if not ente_row:
            raise ValueError("Titulares: el ente seleccionado no existe para ese ejercicio.")

        ente_nombre = (ente_row["ente_nombre"] or "").strip()
        ente_uid = (ente_row["ente_uid"] or "").strip()

        historial_states = [
            _upsert_historial_titular(
                db,
                ejercicio=titular_ejercicio,
                ente_uid=ente_uid,
                ente_nombre=ente_nombre,
                tipo_auditoria=titular_tipo_auditoria,
                nombre=titular_nombre,
                cargo="Titular",
                fecha_inicio=titular_inicio,
                fecha_fin=titular_fin,
                tipo_registro="titular",
            ),
            _upsert_historial_titular(
                db,
                ejercicio=titular_ejercicio,
                ente_uid=ente_uid,
                ente_nombre=ente_nombre,
                tipo_auditoria=titular_tipo_auditoria,
                nombre=titular_administrativo,
                cargo="Director Administrativo",
                fecha_inicio=admin_inicio,
                fecha_fin=admin_fin,
                tipo_registro="director_administrativo",
            ),
        ]

        duplicate_capture = db.execute(
            """
            SELECT id, created_at
            FROM cargas_titulares
            WHERE ejercicio = ?
              AND ente_id = ?
              AND tipo_auditoria = ?
              AND periodo_informe = ?
              AND titular = ?
              AND periodo_administrativo = ?
              AND administrativo = ?
              AND cedula_resultados = ?
            ORDER BY id DESC
            LIMIT 1
            """,
            (
                titular_ejercicio,
                titular_ente_id,
                titular_tipo_auditoria,
                titular_periodo_informe,
                titular_nombre,
                titular_periodo_administrativo,
                titular_administrativo,
                titular_cedula_resultados,
            ),
        ).fetchone()

        if duplicate_capture and all(state == "unchanged" for state in historial_states):
            return {
                "ok": False,
                "level": "info",
                "message": (
                    f"Titulares: ya existe una captura idéntica "
                    f"(ID {duplicate_capture['id']}, fecha {duplicate_capture['created_at']})."
                ),
            }

        if not duplicate_capture:
            db.execute(
                """
                INSERT INTO cargas_titulares (
                    ejercicio,
                    ente_id,
                    ente_nombre,
                    tipo_auditoria,
                    periodo_informe,
                    titular,
                    periodo_administrativo,
                    administrativo,
                    cedula_resultados,
                    created_by,
                    created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    titular_ejercicio,
                    titular_ente_id,
                    ente_nombre,
                    titular_tipo_auditoria,
                    titular_periodo_informe,
                    titular_nombre,
                    titular_periodo_administrativo,
                    titular_administrativo,
                    titular_cedula_resultados,
                    user["username"],
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                ),
            )

        db.commit()

        if duplicate_capture:
            return {
                "ok": True,
                "level": "success",
                "message": "Titulares: historial actualizado; la captura idéntica ya existía en bitácora.",
            }

        return {
            "ok": True,
            "level": "success",
            "message": "Titulares: registro guardado correctamente en historial y bitácora.",
        }

    def fuentes_por_ente(
        db,
        ejercicio: str,
        ente_id_norm: str,
        tipo_auditoria: str = "",
    ) -> list[dict]:
        if not ejercicio or not ente_id_norm:
            return []
        tipo_options = _tipo_auditoria_options(tipo_auditoria)
        tipo_filter_sql = ""
        tipo_filter_params: list[str] = []
        if tipo_options:
            placeholders = ", ".join(["?"] * len(tipo_options))
            tipo_filter_sql = f" AND TRIM(COALESCE(o.tipo_auditoria, '')) IN ({placeholders})"
            tipo_filter_params = tipo_options.copy()

        ente_base = db.execute(
            f"""
            SELECT TRIM(COALESCE(ente_uid, '')) AS ente_uid
            FROM entes_detalle
            WHERE TRIM(COALESCE(ejercicio, '')) = ?
              AND {normalize_ente_id_sql('ente_id')} = ?
            LIMIT 1
            """,
            (ejercicio, ente_id_norm),
        ).fetchone()
        ente_uid = (ente_base["ente_uid"] or "").strip() if ente_base else ""

        catalogo_rows = db.execute(
            """
            SELECT id, TRIM(COALESCE(nombre, '')) AS nombre
            FROM fuentes_financiamiento
            WHERE TRIM(COALESCE(nombre, '')) != ''
            ORDER BY nombre ASC
            """
        ).fetchall()
        catalogo_por_nombre = {row["nombre"].lower(): dict(row) for row in catalogo_rows}

        if ente_uid:
            registros_rows = db.execute(
                f"""
                SELECT DISTINCT ff.id, ff.nombre
                FROM registros AS r
                JOIN fuentes_financiamiento AS ff
                  ON r.fuente_id = ff.id
                JOIN entes_detalle AS ed
                  ON TRIM(COALESCE(r.ejercicio, '')) = TRIM(COALESCE(ed.ejercicio, ''))
                 AND {normalize_ente_id_sql('r.ente_id')} = {normalize_ente_id_sql('ed.ente_id')}
                WHERE TRIM(COALESCE(ed.ente_uid, '')) = ?
                ORDER BY ff.nombre ASC
                """,
                (ente_uid,),
            ).fetchall()
            observaciones_rows = db.execute(
                f"""
                SELECT DISTINCT TRIM(COALESCE(o.fuente_financiamiento, '')) AS nombre
                FROM observaciones AS o
                JOIN entes_detalle AS ed
                  ON TRIM(COALESCE(o.ejercicio, '')) = TRIM(COALESCE(ed.ejercicio, ''))
                 AND {normalize_ente_id_sql('o.ente_id')} = {normalize_ente_id_sql('ed.ente_id')}
                WHERE TRIM(COALESCE(ed.ente_uid, '')) = ?
                  AND TRIM(COALESCE(o.fuente_financiamiento, '')) != ''
                  {tipo_filter_sql}
                ORDER BY nombre ASC
                """,
                [ente_uid, *tipo_filter_params],
            ).fetchall()
        else:
            registros_rows = db.execute(
                f"""
                SELECT DISTINCT ff.id, ff.nombre
                FROM registros AS r
                JOIN fuentes_financiamiento AS ff
                  ON r.fuente_id = ff.id
                WHERE TRIM(COALESCE(r.ejercicio, '')) = ?
                  AND {normalize_ente_id_sql('r.ente_id')} = ?
                ORDER BY ff.nombre ASC
                """,
                (ejercicio, ente_id_norm),
            ).fetchall()
            observaciones_rows = db.execute(
                f"""
                SELECT DISTINCT TRIM(COALESCE(fuente_financiamiento, '')) AS nombre
                FROM observaciones
                WHERE TRIM(COALESCE(ejercicio, '')) = ?
                  AND {normalize_ente_id_sql('ente_id')} = ?
                  AND TRIM(COALESCE(fuente_financiamiento, '')) != ''
                  {tipo_filter_sql.replace("o.", "")}
                ORDER BY nombre ASC
                """,
                [ejercicio, ente_id_norm, *tipo_filter_params],
            ).fetchall()

        resultado = []
        seen_keys = set()
        for row in registros_rows:
            nombre = (row["nombre"] or "").strip()
            if not nombre:
                continue
            key = nombre.lower()
            if key in seen_keys:
                continue
            seen_keys.add(key)
            resultado.append({"id": str(row["id"]), "nombre": nombre})

        for row in observaciones_rows:
            nombre = (row["nombre"] or "").strip()
            if not nombre:
                continue
            key = nombre.lower()
            if key in seen_keys:
                continue
            seen_keys.add(key)
            catalogo_row = catalogo_por_nombre.get(key)
            if catalogo_row:
                resultado.append({"id": str(catalogo_row["id"]), "nombre": catalogo_row["nombre"]})
            else:
                resultado.append({"id": f"__obs__:{nombre}", "nombre": nombre})

        if not resultado:
            for row in catalogo_rows:
                resultado.append({"id": str(row["id"]), "nombre": row["nombre"]})

        if not resultado:
            observaciones_global_rows = db.execute(
                """
                SELECT DISTINCT TRIM(COALESCE(fuente_financiamiento, '')) AS nombre
                FROM observaciones
                WHERE TRIM(COALESCE(fuente_financiamiento, '')) != ''
                  {tipo_filter_sql}
                ORDER BY nombre ASC
                """.format(tipo_filter_sql=tipo_filter_sql.replace("o.", ""))
                if tipo_filter_sql
                else """
                SELECT DISTINCT TRIM(COALESCE(fuente_financiamiento, '')) AS nombre
                FROM observaciones
                WHERE TRIM(COALESCE(fuente_financiamiento, '')) != ''
                ORDER BY nombre ASC
                """,
                tipo_filter_params,
            ).fetchall()
            for row in observaciones_global_rows:
                nombre = (row["nombre"] or "").strip()
                if not nombre:
                    continue
                key = nombre.lower()
                if key in seen_keys:
                    continue
                seen_keys.add(key)
                catalogo_row = catalogo_por_nombre.get(key)
                if catalogo_row:
                    resultado.append({"id": str(catalogo_row["id"]), "nombre": catalogo_row["nombre"]})
                else:
                    resultado.append({"id": f"__obs__:{nombre}", "nombre": nombre})

        return sorted(resultado, key=lambda item: item["nombre"].lower())

    def parse_pdp_amounts(raw_value: str) -> list[float]:
        tokens = re.split(r"(?:\r?\n|;|\|)+", (raw_value or "").strip())
        parsed: list[float] = []
        for token in tokens:
            clean = (token or "").strip()
            if not clean:
                continue
            clean = clean.replace("$", "").replace(",", "").strip()
            if clean in {"-", "—"}:
                parsed.append(0.0)
                continue
            try:
                amount = float(clean)
            except ValueError as exc:
                raise ValueError(f"Monto PDP inválido en detalle: '{token}'.") from exc
            if amount < 0:
                raise ValueError("Los montos PDP no pueden ser negativos.")
            parsed.append(amount)
        return parsed

    def materialize_observaciones_from_manual(
        db,
        *,
        ejercicio: str,
        ente_id: str,
        ente_numero: str,
        ente_nombre: str,
        tipo_auditoria: str,
        fuente_nombre: str,
        ramo_33: str,
        estado: str,
        periodo_cedula: str,
        periodo_titular: str,
        oficio: str,
        fecha_notificacion: str,
        cantidad_sa: int,
        cantidad_pdp: int,
        cantidad_pras: int,
        cantidad_pefcf: int,
        cantidad_r: int,
        monto_pdp_solventado: float,
        monto_pdp_pendiente: float,
        pdp_amounts: list[float],
        pdp_details: list[dict] | None = None,
        solventacion_totales_by_anexo: dict[str, dict] | None = None,
        replace_scope: bool = False,
    ) -> None:
        if replace_scope:
            db.execute(
                """
                DELETE FROM observaciones
                WHERE TRIM(COALESCE(ejercicio, '')) = TRIM(COALESCE(?, ''))
                  AND TRIM(COALESCE(ente_id, '')) = TRIM(COALESCE(?, ''))
                  AND TRIM(COALESCE(tipo_auditoria, '')) = TRIM(COALESCE(?, ''))
                  AND LOWER(TRIM(COALESCE(oficio, ''))) = LOWER(TRIM(COALESCE(?, '')))
                  AND LOWER(TRIM(COALESCE(fuente_financiamiento, ''))) = LOWER(TRIM(COALESCE(?, '')))
                  AND LOWER(TRIM(COALESCE(periodo_cedula, ''))) = LOWER(TRIM(COALESCE(?, '')))
                """,
                (ejercicio, ente_id, tipo_auditoria, oficio, fuente_nombre, periodo_cedula),
            )

        counts = {
            "SA": cantidad_sa,
            "PDP": cantidad_pdp,
            "PRAS": cantidad_pras,
            "PEFCF": cantidad_pefcf,
            "R": cantidad_r,
        }
        pdp_index = 0
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        for tipo_anexo, total in counts.items():
            for numero_observacion in range(1, total + 1):
                monto_emitido = None
                monto_solventado = None
                monto_pendiente = None
                pdp_concepto = None
                pdp_subconcepto = None
                fuente_row_nombre = fuente_nombre
                totales_tipo = (
                    (solventacion_totales_by_anexo or {}).get(tipo_anexo)
                    if isinstance(solventacion_totales_by_anexo, dict)
                    else None
                )
                if tipo_anexo == "PDP":
                    detalle = pdp_details[pdp_index] if pdp_details and pdp_index < len(pdp_details) else {}
                    monto_emitido = detalle.get("monto")
                    if monto_emitido is None:
                        monto_emitido = pdp_amounts[pdp_index] if pdp_index < len(pdp_amounts) else 0.0
                    pdp_concepto = (detalle.get("concepto") or "").strip() or None
                    pdp_subconcepto = (detalle.get("subconcepto") or "").strip() or None
                    fuente_detalle = " ".join(str(detalle.get("fuente") or "").split())
                    if fuente_detalle:
                        fuente_row_nombre = fuente_detalle
                    if totales_tipo and numero_observacion == 1:
                        monto_solventado = float(totales_tipo.get("solventado", 0.0) or 0.0)
                        monto_pendiente = float(totales_tipo.get("pendiente", 0.0) or 0.0)
                    elif numero_observacion == 1:
                        monto_solventado = monto_pdp_solventado
                        monto_pendiente = monto_pdp_pendiente
                    else:
                        monto_solventado = 0.0
                        monto_pendiente = 0.0
                    pdp_index += 1
                elif totales_tipo:
                    monto_emitido = float(totales_tipo.get("emitido", 0.0) or 0.0) if numero_observacion == 1 else 0.0
                    monto_solventado = float(totales_tipo.get("solventado", 0.0) or 0.0) if numero_observacion == 1 else 0.0
                    monto_pendiente = float(totales_tipo.get("pendiente", 0.0) or 0.0) if numero_observacion == 1 else 0.0
                db.execute(
                    """
                    INSERT INTO observaciones (
                        ejercicio,
                        auditoria,
                        ente_id,
                        ente_numero,
                        ente_numero_sort,
                        ente_nombre,
                        tipo_auditoria,
                        fuente_financiamiento,
                        ramo_33,
                        periodo,
                        periodo_cedula,
                        periodo_titular,
                        oficio,
                        fecha_notificacion,
                        tipo_anexo,
                        numero_observacion,
                        estatus,
                        estado,
                        monto,
                        monto_pdp_emitido,
                        monto_pdp_solventado,
                        monto_pdp_pendiente,
                        pdp_concepto_irregularidad,
                        pdp_subconcepto_irregularidad,
                        created_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        ejercicio,
                        tipo_auditoria,
                        ente_id,
                        ente_numero,
                        parse_ente_numero_sort(ente_numero),
                        ente_nombre,
                        tipo_auditoria,
                        fuente_row_nombre,
                        ramo_33,
                        periodo_cedula,
                        periodo_cedula,
                        periodo_titular or periodo_cedula,
                        oficio,
                        fecha_notificacion,
                        tipo_anexo,
                        numero_observacion,
                        estado,
                        estado,
                        monto_emitido,
                        monto_emitido,
                        monto_solventado,
                        monto_pendiente,
                        pdp_concepto if tipo_anexo == "PDP" else None,
                        pdp_subconcepto if tipo_anexo == "PDP" else None,
                        now,
                    ),
                )

    def parse_manual_pdp_details(raw_value: str, cantidad_pdp: int) -> list[dict]:
        if cantidad_pdp <= 0:
            return []
        if not (raw_value or "").strip():
            raise ValueError(
                "Debes capturar concepto y monto en todas las observaciones PDP."
            )
        try:
            payload = json.loads(raw_value)
        except json.JSONDecodeError as exc:
            raise ValueError("El detalle PDP tiene un formato inválido.") from exc
        if not isinstance(payload, list):
            raise ValueError("El detalle PDP tiene un formato inválido.")
        if len(payload) != cantidad_pdp:
            raise ValueError(
                "Debes capturar concepto y monto para cada observación PDP."
            )

        details: list[dict] = []
        for item in payload[:cantidad_pdp]:
            data = item if isinstance(item, dict) else {}
            concepto = " ".join(str(data.get("concepto") or "").split())
            subconcepto = " ".join(str(data.get("subconcepto") or "").split())
            fuente = " ".join(str(data.get("fuente") or "").split())
            monto_field = data.get("monto")
            monto_raw = "" if monto_field is None else str(monto_field).strip()
            monto_val = None
            if monto_raw:
                monto_clean = monto_raw.replace("$", "").replace(",", "").strip()
                try:
                    monto_val = float(monto_clean)
                except ValueError as exc:
                    raise ValueError("Monto PDP inválido en el detalle de observaciones.") from exc
                if monto_val < 0:
                    raise ValueError("Los montos PDP no pueden ser negativos.")
            if not concepto:
                raise ValueError("Debes capturar concepto en todas las observaciones PDP.")
            if monto_val is None:
                raise ValueError("Debes capturar monto en todas las observaciones PDP.")
            details.append(
                {
                    "concepto": concepto,
                    "subconcepto": subconcepto,
                    "monto": monto_val,
                    "fuente": fuente,
                }
            )

        return details

    def parse_manual_fuentes_detalle(raw_value: str) -> list[dict]:
        raw = (raw_value or "").strip()
        if not raw:
            return []
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError:
            return []
        if not isinstance(payload, list):
            return []
        rows: list[dict] = []
        for item in payload:
            if not isinstance(item, dict):
                continue
            fuente_nombre = " ".join(str(item.get("fuente_nombre") or "").split())
            tipo_auditoria = " ".join(str(item.get("tipo_auditoria") or "").split())
            if not fuente_nombre:
                continue
            if tipo_auditoria not in {"Financiera", "Obra Pública"}:
                tipo_auditoria = "Financiera"
            cantidad_sa = parse_non_negative_int(str(item.get("cantidad_sa", "0")), "Cantidad SA")
            cantidad_pdp = parse_non_negative_int(str(item.get("cantidad_pdp", "0")), "Cantidad PDP")
            cantidad_pras = parse_non_negative_int(str(item.get("cantidad_pras", "0")), "Cantidad PRAS")
            cantidad_pefcf = parse_non_negative_int(str(item.get("cantidad_pefcf", "0")), "Cantidad PEFCF")
            cantidad_r = parse_non_negative_int(str(item.get("cantidad_r", "0")), "Cantidad R")
            if (cantidad_sa + cantidad_pdp + cantidad_pras + cantidad_pefcf + cantidad_r) <= 0:
                continue
            rows.append(
                {
                    "fuente_nombre": fuente_nombre,
                    "tipo_auditoria": tipo_auditoria,
                    "cantidad_sa": cantidad_sa,
                    "cantidad_pdp": cantidad_pdp,
                    "cantidad_pras": cantidad_pras,
                    "cantidad_pefcf": cantidad_pefcf,
                    "cantidad_r": cantidad_r,
                }
            )
        return rows

    def count_existing_observaciones_scope(
        db,
        *,
        ejercicio: str,
        ente_id: str,
        tipo_auditoria: str,
        fuente_nombre: str,
        periodo: str,
        oficio: str,
    ) -> tuple[int, dict[str, int]]:
        rows = db.execute(
            f"""
            SELECT
                TRIM(COALESCE(tipo_anexo, '')) AS tipo_anexo,
                COUNT(*) AS total
            FROM observaciones
            WHERE TRIM(COALESCE(ejercicio, '')) = ?
              AND {normalize_ente_id_sql('ente_id')} = ?
              AND TRIM(COALESCE(tipo_auditoria, '')) = ?
              AND LOWER(TRIM(COALESCE(fuente_financiamiento, ''))) = LOWER(TRIM(COALESCE(?, '')))
              AND LOWER(TRIM(COALESCE(periodo_cedula, ''))) = LOWER(TRIM(COALESCE(?, '')))
              AND LOWER(TRIM(COALESCE(oficio, ''))) = LOWER(TRIM(COALESCE(?, '')))
            GROUP BY TRIM(COALESCE(tipo_anexo, ''))
            """,
            (ejercicio, ente_id, tipo_auditoria, fuente_nombre, periodo, oficio),
        ).fetchall()
        counts: dict[str, int] = {}
        total = 0
        for row in rows:
            tipo = (row["tipo_anexo"] or "").strip().upper()
            qty = int(row["total"] or 0)
            if not tipo or qty <= 0:
                continue
            counts[tipo] = qty
            total += qty
        return total, counts

    def count_existing_observaciones_by_clave(
        db,
        *,
        ejercicio: str,
        ente_id: str,
        periodo: str,
        oficio: str,
    ) -> tuple[int, dict[str, int]]:
        rows = db.execute(
            f"""
            SELECT
                TRIM(COALESCE(tipo_anexo, '')) AS tipo_anexo,
                COUNT(*) AS total
            FROM observaciones
            WHERE TRIM(COALESCE(ejercicio, '')) = ?
              AND {normalize_ente_id_sql('ente_id')} = ?
              AND LOWER(TRIM(COALESCE(periodo_cedula, ''))) = LOWER(TRIM(COALESCE(?, '')))
              AND LOWER(TRIM(COALESCE(oficio, ''))) = LOWER(TRIM(COALESCE(?, '')))
            GROUP BY TRIM(COALESCE(tipo_anexo, ''))
            """,
            (ejercicio, ente_id, periodo, oficio),
        ).fetchall()
        counts: dict[str, int] = {}
        total = 0
        for row in rows:
            tipo = (row["tipo_anexo"] or "").strip().upper()
            qty = int(row["total"] or 0)
            if not tipo or qty <= 0:
                continue
            counts[tipo] = qty
            total += qty
        return total, counts

    def build_observaciones_admin_scope(raw_scope) -> tuple[list[str], list[str], dict, int]:
        ejercicio = " ".join((raw_scope.get("ejercicio") or "").split())
        ente_id = normalize_ente_id(raw_scope.get("ente_id", ""))
        tipo_auditoria = " ".join((raw_scope.get("tipo_auditoria") or "").split())
        fuente = " ".join((raw_scope.get("fuente") or "").split())
        periodo = " ".join((raw_scope.get("periodo") or "").split())
        oficio = " ".join((raw_scope.get("oficio") or "").split())
        estado_raw = " ".join((raw_scope.get("estado") or "").split())
        tipo_anexo = " ".join((raw_scope.get("tipo_anexo") or "").split()).upper()

        if tipo_anexo == "PEFCT":
            tipo_anexo = "PEFCF"

        where_clauses: list[str] = []
        params: list[str] = []
        extra_filters = 0

        if ejercicio:
            where_clauses.append("TRIM(COALESCE(ejercicio, '')) = ?")
            params.append(ejercicio)

        if ente_id:
            where_clauses.append(f"{normalize_ente_id_sql('ente_id')} = ?")
            params.append(ente_id)
            extra_filters += 1

        if tipo_auditoria:
            tipo_options = _tipo_auditoria_options(tipo_auditoria)
            if not tipo_options:
                tipo_options = [tipo_auditoria]
            placeholders = ", ".join(["?"] * len(tipo_options))
            where_clauses.append(
                f"TRIM(COALESCE(tipo_auditoria, '')) IN ({placeholders})"
            )
            params.extend(tipo_options)
            extra_filters += 1

        if fuente:
            where_clauses.append(
                "LOWER(TRIM(COALESCE(fuente_financiamiento, ''))) = LOWER(TRIM(COALESCE(?, '')))"
            )
            params.append(fuente)
            extra_filters += 1

        if periodo:
            where_clauses.append(
                "LOWER(TRIM(COALESCE(periodo_cedula, ''))) = LOWER(TRIM(COALESCE(?, '')))"
            )
            params.append(periodo)
            extra_filters += 1

        if oficio:
            where_clauses.append(
                "LOWER(TRIM(COALESCE(oficio, ''))) = LOWER(TRIM(COALESCE(?, '')))"
            )
            params.append(oficio)
            extra_filters += 1

        if estado_raw:
            estado = _normalize_observacion_estado(estado_raw)
            if estado not in OBSERVACION_ESTADOS_VALIDOS:
                raise ValueError("Estado inválido para filtrar observaciones.")
            if estado == "Emitido":
                where_clauses.append(
                    "LOWER(TRIM(COALESCE(estado, ''))) IN ('emitido', 'e')"
                )
            elif estado == "Pendiente":
                where_clauses.append(
                    "LOWER(TRIM(COALESCE(estado, ''))) IN ('pendiente', 'p')"
                )
            else:
                where_clauses.append(
                    "LOWER(TRIM(COALESCE(estado, ''))) IN ('solventado', 's')"
                )
            extra_filters += 1

        if tipo_anexo:
            if tipo_anexo not in {"SA", "PDP", "PRAS", "PEFCF", "R"}:
                raise ValueError("Tipo de anexo inválido para filtrar observaciones.")
            where_clauses.append("UPPER(TRIM(COALESCE(tipo_anexo, ''))) = ?")
            params.append(tipo_anexo)
            extra_filters += 1

        return (
            where_clauses,
            params,
            {
                "ejercicio": ejercicio,
                "ente_id": ente_id,
                "tipo_auditoria": tipo_auditoria,
                "fuente": fuente,
                "periodo": periodo,
                "oficio": oficio,
                "estado": estado_raw,
                "tipo_anexo": tipo_anexo,
            },
            extra_filters,
        )

    @app.get("/carga/entes")
    @gabo_required
    def carga_entes_por_ejercicio():
        ejercicio = (request.args.get("ejercicio") or "").strip()
        if not ejercicio:
            return jsonify([])

        db = get_db()
        rows = db.execute(
            """
            SELECT
                TRIM(COALESCE(ente_id, '')) AS ente_id,
                TRIM(COALESCE(ente_numero, '')) AS ente_numero,
                TRIM(COALESCE(ente_nombre, '')) AS ente_nombre
            FROM entes_detalle
            WHERE TRIM(COALESCE(ejercicio, '')) = ?
              AND TRIM(COALESCE(ente_id, '')) != ''
            ORDER BY CAST(COALESCE(NULLIF(ente_numero, ''), '0') AS REAL), ente_numero, ente_nombre
            """,
            (ejercicio,),
        ).fetchall()
        return jsonify([dict(row) for row in rows])

    @app.get("/carga/fuentes-ente")
    @gabo_required
    def carga_fuentes_por_ente():
        ejercicio = (request.args.get("ejercicio") or "").strip()
        ente_id = normalize_ente_id(request.args.get("ente_id", ""))
        tipo_auditoria = (request.args.get("tipo_auditoria") or "").strip()
        if not ejercicio or not ente_id:
            return jsonify([])

        db = get_db()
        rows = fuentes_por_ente(db, ejercicio, ente_id, tipo_auditoria=tipo_auditoria)
        return jsonify([dict(row) for row in rows])

    @app.get("/carga/pdp-catalogo")
    @gabo_required
    def carga_pdp_catalogo():
        db = get_db()
        rows = db.execute(
            """
            SELECT DISTINCT
                TRIM(COALESCE(pdp_concepto_irregularidad, '')) AS concepto,
                TRIM(COALESCE(pdp_subconcepto_irregularidad, '')) AS subconcepto
            FROM observaciones
            WHERE TRIM(COALESCE(ejercicio, '')) IN ('2023', '2024')
              AND TRIM(COALESCE(tipo_anexo, '')) = 'PDP'
              AND TRIM(COALESCE(pdp_concepto_irregularidad, '')) != ''
            ORDER BY concepto ASC, subconcepto ASC
            """
        ).fetchall()

        conceptos: list[str] = []
        subconceptos_by_concept: dict[str, list[str]] = {}
        seen_conceptos = set()
        for row in rows:
            concepto = (row["concepto"] or "").strip()
            subconcepto = (row["subconcepto"] or "").strip()
            if not concepto:
                continue
            if concepto not in seen_conceptos:
                seen_conceptos.add(concepto)
                conceptos.append(concepto)
            if subconcepto:
                subconceptos_by_concept.setdefault(concepto, [])
                if subconcepto not in subconceptos_by_concept[concepto]:
                    subconceptos_by_concept[concepto].append(subconcepto)

        return jsonify(
            {
                "conceptos": conceptos,
                "subconceptos_by_concepto": subconceptos_by_concept,
            }
        )

    @app.get("/carga/observaciones-cargadas")
    @gabo_required
    def carga_observaciones_cargadas():
        ejercicio = (request.args.get("ejercicio") or "").strip()
        ente_id = normalize_ente_id(request.args.get("ente_id", ""))
        tipo_auditoria = (request.args.get("tipo_auditoria") or "").strip()
        fuente = " ".join((request.args.get("fuente") or "").split())
        periodo = " ".join((request.args.get("periodo") or "").split())
        oficio = " ".join((request.args.get("oficio") or "").split())
        if not ejercicio or not ente_id:
            return jsonify([])

        db = get_db()
        params: list[str] = [ejercicio, ente_id]
        tipo_clause = ""
        if tipo_auditoria:
            tipo_options = _tipo_auditoria_options(tipo_auditoria)
            if not tipo_options:
                tipo_options = [tipo_auditoria]
            tipo_placeholders = ", ".join(["?"] * len(tipo_options))
            tipo_clause = f" AND TRIM(COALESCE(tipo_auditoria, '')) IN ({tipo_placeholders})"
            params.extend(tipo_options)
        where_extra: list[str] = []
        if fuente:
            where_extra.append("LOWER(TRIM(COALESCE(fuente_financiamiento, ''))) = LOWER(TRIM(COALESCE(?, '')))")
            params.append(fuente)
        if periodo:
            where_extra.append("LOWER(TRIM(COALESCE(periodo_cedula, ''))) = LOWER(TRIM(COALESCE(?, '')))")
            params.append(periodo)
        if oficio:
            where_extra.append("LOWER(TRIM(COALESCE(oficio, ''))) = LOWER(TRIM(COALESCE(?, '')))")
            params.append(oficio)

        where_sql = ""
        if where_extra:
            where_sql = " AND " + " AND ".join(where_extra)

        rows = db.execute(
            f"""
            SELECT
                id,
                TRIM(COALESCE(tipo_anexo, '')) AS tipo_anexo,
                numero_observacion,
                TRIM(COALESCE(estado, '')) AS estado,
                COALESCE(reclasificada, 0) AS reclasificada,
                TRIM(COALESCE(tipo_auditoria, '')) AS tipo_auditoria,
                TRIM(COALESCE(fuente_financiamiento, '')) AS fuente_financiamiento,
                monto_pdp_emitido,
                monto_pdp_solventado,
                monto_pdp_pendiente,
                TRIM(COALESCE(pdp_concepto_irregularidad, '')) AS pdp_concepto_irregularidad,
                TRIM(COALESCE(pdp_subconcepto_irregularidad, '')) AS pdp_subconcepto_irregularidad
            FROM observaciones
            WHERE TRIM(COALESCE(ejercicio, '')) = ?
              AND {normalize_ente_id_sql('ente_id')} = ?
              {tipo_clause}
              {where_sql}
            ORDER BY
              CASE TRIM(COALESCE(tipo_anexo, ''))
                WHEN 'SA' THEN 1
                WHEN 'PDP' THEN 2
                WHEN 'PRAS' THEN 3
                WHEN 'PEFCF' THEN 4
                WHEN 'R' THEN 5
                ELSE 9
              END,
              COALESCE(numero_observacion, 0)
            LIMIT 1200
            """,
            params,
        ).fetchall()
        payload = []
        for row in rows:
            item = dict(row)
            item["estado"] = _normalize_observacion_estado(item.get("estado", ""))
            item["reclasificada"] = 1 if int(item.get("reclasificada") or 0) else 0
            payload.append(item)
        return jsonify(payload)

    @app.get("/carga/observaciones-claves")
    @gabo_required
    def carga_observaciones_claves():
        ejercicio = (request.args.get("ejercicio") or "").strip()
        ente_id = normalize_ente_id(request.args.get("ente_id", ""))
        tipo_auditoria = (request.args.get("tipo_auditoria") or "").strip()
        fuente = " ".join((request.args.get("fuente") or "").split())
        if not ejercicio or not ente_id:
            return jsonify([])

        db = get_db()
        params: list[str] = [ejercicio, ente_id]
        tipo_clause = ""
        if tipo_auditoria:
            tipo_options = _tipo_auditoria_options(tipo_auditoria)
            if not tipo_options:
                tipo_options = [tipo_auditoria]
            tipo_placeholders = ", ".join(["?"] * len(tipo_options))
            tipo_clause = f" AND TRIM(COALESCE(tipo_auditoria, '')) IN ({tipo_placeholders})"
            params.extend(tipo_options)

        where_extra: list[str] = []
        if fuente:
            where_extra.append(
                "LOWER(TRIM(COALESCE(fuente_financiamiento, ''))) = LOWER(TRIM(COALESCE(?, '')))"
            )
            params.append(fuente)

        where_sql = ""
        if where_extra:
            where_sql = " AND " + " AND ".join(where_extra)

        rows = db.execute(
            f"""
            SELECT DISTINCT
                TRIM(COALESCE(periodo_cedula, '')) AS periodo,
                TRIM(COALESCE(oficio, '')) AS oficio
            FROM observaciones
            WHERE TRIM(COALESCE(ejercicio, '')) = ?
              AND {normalize_ente_id_sql('ente_id')} = ?
              {tipo_clause}
              {where_sql}
              AND TRIM(COALESCE(periodo_cedula, '')) != ''
              AND TRIM(COALESCE(oficio, '')) != ''
            ORDER BY
              LOWER(TRIM(COALESCE(periodo_cedula, ''))) ASC,
              LOWER(TRIM(COALESCE(oficio, ''))) ASC
            LIMIT 1200
            """,
            params,
        ).fetchall()

        payload = []
        for row in rows:
            periodo = (row["periodo"] or "").strip()
            oficio = (row["oficio"] or "").strip()
            if not periodo or not oficio:
                continue
            payload.append({"periodo": periodo, "oficio": oficio})
        return jsonify(payload)

    @app.post("/carga/observaciones-cargadas/<int:observacion_id>/actualizar")
    @gabo_required
    def carga_observacion_actualizar(observacion_id: int):
        payload = request.get_json(silent=True) or {}
        accion = " ".join(str(payload.get("accion", "") or "").lower().split())
        estado = _normalize_observacion_estado(payload.get("estado", ""))
        monto_emitido_raw = payload.get("monto_pdp_emitido", "")
        monto_solventado_raw = payload.get("monto_pdp_solventado", "")

        db = get_db()
        current = db.execute(
            """
            SELECT id, TRIM(COALESCE(tipo_anexo, '')) AS tipo_anexo
            FROM observaciones
            WHERE id = ?
            LIMIT 1
            """,
            (observacion_id,),
        ).fetchone()
        if not current:
            return jsonify({"ok": False, "error": "Observación no encontrada."}), 404

        tipo_anexo = (current["tipo_anexo"] or "").strip().upper()
        if accion in {"reclasificar_pdp_pras", "reclasificar_pras", "pdp_a_pras"}:
            if tipo_anexo != "PDP":
                return jsonify(
                    {"ok": False, "error": "Solo se puede reclasificar observaciones PDP."},
                    400,
                )
            db.execute(
                """
                UPDATE observaciones
                SET tipo_anexo = 'PRAS',
                    estado = 'Pendiente',
                    reclasificada = 1,
                    pdp_concepto_irregularidad = '',
                    pdp_subconcepto_irregularidad = '',
                    monto_pdp_emitido = 0,
                    monto_pdp_solventado = 0,
                    monto_pdp_pendiente = 0,
                    monto = 0
                WHERE id = ?
                """,
                (observacion_id,),
            )
            db.commit()
            return jsonify(
                {
                    "ok": True,
                    "id": observacion_id,
                    "estado": "Pendiente",
                    "tipo_anexo": "PRAS",
                    "reclasificada": 1,
                    "accion": "reclasificar_pdp_pras",
                }
            )

        if estado not in OBSERVACION_ESTADOS_VALIDOS:
            return jsonify({"ok": False, "error": "Estado inválido."}), 400

        if tipo_anexo == "PDP":
            try:
                monto_emitido = parse_non_negative_float(str(monto_emitido_raw), "Monto PDP emitido")
                monto_solventado = parse_non_negative_float(str(monto_solventado_raw), "Monto PDP solventado")
            except ValueError:
                return jsonify({"ok": False, "error": "Montos PDP inválidos."}), 400
            if monto_solventado > monto_emitido:
                return jsonify(
                    {"ok": False, "error": "Monto solventado no puede ser mayor a emitido."},
                    400,
                )
            monto_pendiente = monto_emitido - monto_solventado
            db.execute(
                """
                UPDATE observaciones
                SET estado = ?,
                    monto_pdp_emitido = ?,
                    monto_pdp_solventado = ?,
                    monto_pdp_pendiente = ?,
                    monto = ?
                WHERE id = ?
                """,
                (
                    estado,
                    monto_emitido,
                    monto_solventado,
                    monto_pendiente,
                    monto_emitido,
                    observacion_id,
                ),
            )
        else:
            db.execute(
                """
                UPDATE observaciones
                SET estado = ?
                WHERE id = ?
                """,
                (estado, observacion_id),
            )
        db.commit()
        return jsonify(
            {
                "ok": True,
                "id": observacion_id,
                "estado": estado,
                "tipo_anexo": tipo_anexo,
            }
        )

    @app.get("/carga/observaciones-admin")
    @gabo_required
    def carga_observaciones_admin():
        user = get_current_user()
        db = get_db()
        return_vista = (request.args.get("return_vista") or "").strip().lower()
        if return_vista not in {"manual", "titulares"}:
            return_vista = "manual"
        ejercicios_rows = db.execute(
            """
            SELECT DISTINCT TRIM(COALESCE(ejercicio, '')) AS ejercicio
            FROM entes_detalle
            WHERE TRIM(COALESCE(ejercicio, '')) != ''
            ORDER BY CAST(TRIM(COALESCE(ejercicio, '')) AS INTEGER) DESC, ejercicio DESC
            """
        ).fetchall()
        ejercicios = [row["ejercicio"] for row in ejercicios_rows if (row["ejercicio"] or "").strip()]
        if not ejercicios:
            ejercicios = [TITULAR_EJERCICIO_FIJO]
        return render_template(
            "carga_observaciones_admin.html",
            user=user,
            ejercicios=ejercicios,
            ejercicio_default=ejercicios[0],
            return_vista=return_vista,
        )

    @app.route("/carga/titulares", methods=["GET", "POST"])
    @gabo_required
    def carga_titulares():
        user = get_current_user()
        db = get_db()
        ejercicios_rows = db.execute(
            """
            SELECT DISTINCT TRIM(COALESCE(ejercicio, '')) AS ejercicio
            FROM entes_detalle
            WHERE TRIM(COALESCE(ejercicio, '')) != ''
            ORDER BY CAST(TRIM(COALESCE(ejercicio, '')) AS INTEGER) DESC, ejercicio DESC
            """
        ).fetchall()
        ejercicios = [row["ejercicio"] for row in ejercicios_rows if (row["ejercicio"] or "").strip()]
        if not ejercicios:
            ejercicios = [TITULAR_EJERCICIO_FIJO]

        form_source = request.form if request.method == "POST" else request.args
        requested_default = (
            (request.form.get("titular_ejercicio") if request.method == "POST" else request.args.get("ejercicio"))
            or TITULAR_EJERCICIO_FIJO
        )
        default_ejercicio = " ".join((requested_default or "").split())
        if default_ejercicio not in ejercicios:
            default_ejercicio = ejercicios[0]

        form_data = _build_titular_form_data(
            form_source,
            default_ejercicio=default_ejercicio,
        )
        if form_data["titular_ejercicio"] not in ejercicios:
            form_data["titular_ejercicio"] = default_ejercicio

        titular_result = None
        if request.method == "POST":
            try:
                titular_result = _save_titulares_capture(db, user, form_data)
                if titular_result.get("ok"):
                    form_data["titular_nombre"] = ""
                    form_data["titular_administrativo"] = ""
            except ValueError as exc:
                titular_result = {
                    "ok": False,
                    "level": "error",
                    "message": str(exc),
                }

        entes = _load_titular_entes(db, form_data["titular_ejercicio"])

        return render_template(
            "carga_titulares.html",
            user=user,
            ejercicios=ejercicios,
            entes=entes,
            form_data=form_data,
            titular_result=titular_result,
        )

    @app.get("/carga/titulares/historial")
    @gabo_required
    def carga_titulares_historial():
        ejercicio = " ".join((request.args.get("ejercicio") or "").split())
        ente_id = normalize_ente_id(request.args.get("ente_id", ""))
        tipo_auditoria = normalize_tipo_auditoria(request.args.get("tipo_auditoria", ""))
        if not ejercicio:
            return jsonify({"ok": False, "error": "Selecciona un ejercicio para consultar."}), 400

        db = get_db()
        rows = _list_historial_titulares_rows(
            db,
            ejercicio=ejercicio,
            ente_id_norm=ente_id,
            tipo_auditoria=tipo_auditoria,
        )
        return jsonify({"ok": True, "rows": rows})

    @app.post("/carga/titulares/historial/<int:historial_id>/actualizar")
    @gabo_required
    def carga_titulares_historial_actualizar(historial_id: int):
        payload = request.get_json(silent=True)
        if not isinstance(payload, dict):
            payload = request.form

        ejercicio = " ".join((payload.get("ejercicio") or "").split())
        ente_id = normalize_ente_id(payload.get("ente_id", ""))
        tipo_auditoria = normalize_tipo_auditoria(payload.get("tipo_auditoria", ""))
        tipo_registro = " ".join((payload.get("tipo_registro") or "").split()).lower()
        nombre = " ".join((payload.get("nombre") or "").split())
        cargo = " ".join((payload.get("cargo") or "").split())
        fecha_inicio_raw = " ".join((payload.get("fecha_inicio") or "").split())
        fecha_fin_raw = " ".join((payload.get("fecha_fin") or "").split())

        if not ejercicio:
            return jsonify({"ok": False, "error": "Selecciona un ejercicio."}), 400
        if not ente_id:
            return jsonify({"ok": False, "error": "Selecciona un ente."}), 400
        if tipo_auditoria not in {"Financiera", "Obra Pública"}:
            return jsonify({"ok": False, "error": "Tipo de auditoría inválido."}), 400
        if tipo_registro not in {"titular", "director_administrativo"}:
            return jsonify({"ok": False, "error": "Tipo de registro inválido."}), 400
        if not nombre:
            return jsonify({"ok": False, "error": "Captura el nombre del responsable."}), 400

        if not cargo:
            cargo = "Director Administrativo" if tipo_registro == "director_administrativo" else "Titular"

        fecha_inicio = parse_historial_date(fecha_inicio_raw)
        fecha_fin = parse_historial_date(fecha_fin_raw)
        if not fecha_inicio or not fecha_fin:
            return jsonify({"ok": False, "error": "Captura fechas válidas en formato ISO."}), 400
        if fecha_inicio > fecha_fin:
            return jsonify({"ok": False, "error": "La fecha inicial no puede ser mayor que la final."}), 400

        db = get_db()
        existing_row = db.execute(
            "SELECT id FROM historial_titulares WHERE id = ? LIMIT 1",
            (historial_id,),
        ).fetchone()
        if not existing_row:
            return jsonify({"ok": False, "error": "El registro solicitado no existe."}), 404

        ente_row = _get_ente_row_by_ejercicio_id(db, ejercicio, ente_id)
        if not ente_row:
            return jsonify({"ok": False, "error": "El ente seleccionado no existe en ese ejercicio."}), 400

        fecha_inicio_iso = fecha_inicio.isoformat()
        fecha_fin_iso = fecha_fin.isoformat()
        ente_nombre = (ente_row["ente_nombre"] or "").strip()
        ente_uid = (ente_row["ente_uid"] or "").strip()

        if ente_uid:
            duplicate_row = db.execute(
                """
                SELECT id
                FROM historial_titulares
                WHERE id != ?
                  AND CAST(ejercicio AS TEXT) = ?
                  AND tipo_auditoria = ?
                  AND TRIM(COALESCE(nombre, '')) = ?
                  AND TRIM(COALESCE(cargo, '')) = ?
                  AND fecha_inicio = ?
                  AND fecha_fin = ?
                  AND tipo_registro = ?
                  AND (
                    TRIM(COALESCE(ente_uid, '')) = ?
                    OR TRIM(COALESCE(ente, '')) = ?
                  )
                LIMIT 1
                """,
                (
                    historial_id,
                    ejercicio,
                    tipo_auditoria,
                    nombre,
                    cargo,
                    fecha_inicio_iso,
                    fecha_fin_iso,
                    tipo_registro,
                    ente_uid,
                    ente_nombre,
                ),
            ).fetchone()
        else:
            duplicate_row = db.execute(
                """
                SELECT id
                FROM historial_titulares
                WHERE id != ?
                  AND CAST(ejercicio AS TEXT) = ?
                  AND TRIM(COALESCE(ente, '')) = ?
                  AND tipo_auditoria = ?
                  AND TRIM(COALESCE(nombre, '')) = ?
                  AND TRIM(COALESCE(cargo, '')) = ?
                  AND fecha_inicio = ?
                  AND fecha_fin = ?
                  AND tipo_registro = ?
                LIMIT 1
                """,
                (
                    historial_id,
                    ejercicio,
                    ente_nombre,
                    tipo_auditoria,
                    nombre,
                    cargo,
                    fecha_inicio_iso,
                    fecha_fin_iso,
                    tipo_registro,
                ),
            ).fetchone()
        if duplicate_row:
            return jsonify({"ok": False, "error": "Ya existe otro registro idéntico en historial."}), 400

        db.execute(
            """
            UPDATE historial_titulares
            SET ejercicio = ?,
                ente_uid = ?,
                ente = ?,
                tipo_auditoria = ?,
                nombre = ?,
                cargo = ?,
                fecha_inicio = ?,
                fecha_fin = ?,
                tipo_registro = ?
            WHERE id = ?
            """,
            (
                int(ejercicio),
                ente_uid or None,
                ente_nombre,
                tipo_auditoria,
                nombre,
                cargo,
                fecha_inicio_iso,
                fecha_fin_iso,
                tipo_registro,
                historial_id,
            ),
        )
        db.commit()

        updated_rows = _list_historial_titulares_rows(
            db,
            ejercicio=ejercicio,
            ente_id_norm=ente_id,
            tipo_auditoria=tipo_auditoria,
        )
        updated_row = next(
            (item for item in updated_rows if int(item["id"]) == historial_id),
            None,
        )
        if updated_row is None:
            updated_row = {
                "id": historial_id,
                "ejercicio": ejercicio,
                "ente_id": (ente_row["ente_id"] or "").strip(),
                "ente_numero": (ente_row["ente_numero"] or "").strip(),
                "ente_nombre": ente_nombre,
                "tipo_auditoria": tipo_auditoria,
                "nombre": nombre,
                "cargo": cargo,
                "fecha_inicio": fecha_inicio_iso,
                "fecha_fin": fecha_fin_iso,
                "tipo_registro": tipo_registro,
            }

        return jsonify(
            {
                "ok": True,
                "message": "Historial de titulares actualizado correctamente.",
                "row": updated_row,
            }
        )

    @app.get("/carga/observaciones-admin/datos")
    @gabo_required
    def carga_observaciones_admin_datos():
        try:
            where_clauses, params, scope, _ = build_observaciones_admin_scope(request.args)
        except ValueError as exc:
            return jsonify({"ok": False, "error": str(exc)}), 400
        if not scope["ejercicio"]:
            return jsonify({"ok": False, "error": "Selecciona un ejercicio para consultar."}), 400

        where_sql = " AND ".join(where_clauses) if where_clauses else "1=1"
        db = get_db()
        rows = db.execute(
            f"""
            SELECT
                id,
                TRIM(COALESCE(ejercicio, '')) AS ejercicio,
                TRIM(COALESCE(ente_id, '')) AS ente_id,
                TRIM(COALESCE(ente_numero, '')) AS ente_numero,
                TRIM(COALESCE(ente_nombre, '')) AS ente_nombre,
                TRIM(COALESCE(tipo_auditoria, '')) AS tipo_auditoria,
                TRIM(COALESCE(fuente_financiamiento, '')) AS fuente_financiamiento,
                TRIM(COALESCE(periodo_cedula, '')) AS periodo_cedula,
                TRIM(COALESCE(oficio, '')) AS oficio,
                TRIM(COALESCE(fecha_notificacion, '')) AS fecha_notificacion,
                TRIM(COALESCE(created_at, '')) AS created_at,
                TRIM(COALESCE(tipo_anexo, '')) AS tipo_anexo,
                COALESCE(numero_observacion, 0) AS numero_observacion,
                TRIM(COALESCE(estado, '')) AS estado,
                COALESCE(reclasificada, 0) AS reclasificada,
                COALESCE(monto_pdp_emitido, 0) AS monto_pdp_emitido,
                COALESCE(monto_pdp_solventado, 0) AS monto_pdp_solventado,
                COALESCE(monto_pdp_pendiente, 0) AS monto_pdp_pendiente,
                TRIM(COALESCE(pdp_concepto_irregularidad, '')) AS pdp_concepto_irregularidad,
                TRIM(COALESCE(pdp_subconcepto_irregularidad, '')) AS pdp_subconcepto_irregularidad
            FROM observaciones
            WHERE {where_sql}
            ORDER BY
                id DESC
            LIMIT 2500
            """,
            params,
        ).fetchall()
        payload = []
        for row in rows:
            item = dict(row)
            item["estado"] = _normalize_observacion_estado(item.get("estado", ""))
            item["reclasificada"] = 1 if int(item.get("reclasificada") or 0) else 0
            payload.append(item)
        return jsonify({"ok": True, "rows": payload})

    @app.post("/carga/observaciones-admin/<int:observacion_id>/borrar")
    @gabo_required
    def carga_observaciones_admin_borrar(observacion_id: int):
        db = get_db()
        found = db.execute(
            "SELECT id FROM observaciones WHERE id = ? LIMIT 1",
            (observacion_id,),
        ).fetchone()
        if not found:
            return jsonify({"ok": False, "error": "Observación no encontrada."}), 404
        db.execute("DELETE FROM observaciones WHERE id = ?", (observacion_id,))
        db.commit()
        return jsonify({"ok": True, "deleted": 1, "id": observacion_id})

    @app.post("/carga/observaciones-admin/borrar")
    @gabo_required
    def carga_observaciones_admin_borrar_multiples():
        payload = request.get_json(silent=True) or {}
        raw_ids = payload.get("ids")
        if not isinstance(raw_ids, list) or not raw_ids:
            return jsonify({"ok": False, "error": "No se recibieron observaciones para borrar."}), 400

        ids: list[int] = []
        for raw_id in raw_ids:
            try:
                item_id = int(raw_id)
            except (TypeError, ValueError):
                continue
            if item_id > 0:
                ids.append(item_id)
        ids = sorted(set(ids))
        if not ids:
            return jsonify({"ok": False, "error": "Lista de observaciones inválida."}), 400

        placeholders = ", ".join(["?"] * len(ids))
        db = get_db()
        count_row = db.execute(
            f"SELECT COUNT(*) AS total FROM observaciones WHERE id IN ({placeholders})",
            ids,
        ).fetchone()
        total = int((count_row["total"] if count_row else 0) or 0)
        if total <= 0:
            return jsonify({"ok": False, "error": "No se encontraron observaciones para borrar."}), 404

        db.execute(f"DELETE FROM observaciones WHERE id IN ({placeholders})", ids)
        db.commit()
        return jsonify({"ok": True, "deleted": total})

    @app.post("/carga/observaciones-admin/borrar-todo")
    @gabo_required
    def carga_observaciones_admin_borrar_todo():
        payload = request.get_json(silent=True) or {}
        try:
            where_clauses, params, scope, extra_filters = build_observaciones_admin_scope(payload)
        except ValueError as exc:
            return jsonify({"ok": False, "error": str(exc)}), 400

        if not scope["ejercicio"]:
            return jsonify({"ok": False, "error": "Debes seleccionar ejercicio para borrar."}), 400
        if extra_filters <= 0:
            return jsonify(
                {
                    "ok": False,
                    "error": (
                        "Agrega al menos un filtro adicional (ente, tipo, fuente, periodo, oficio, estado o anexo)."
                    ),
                }
            ), 400

        where_sql = " AND ".join(where_clauses)
        db = get_db()
        count_row = db.execute(
            f"SELECT COUNT(*) AS total FROM observaciones WHERE {where_sql}",
            params,
        ).fetchone()
        total = int((count_row["total"] if count_row else 0) or 0)
        if total <= 0:
            return jsonify({"ok": True, "deleted": 0})

        db.execute(
            f"DELETE FROM observaciones WHERE {where_sql}",
            params,
        )
        db.commit()
        return jsonify({"ok": True, "deleted": total})

    @app.route("/carga", methods=["GET", "POST"])
    @gabo_required
    def carga():
        user = get_current_user()
        db = get_db()
        requested_loader_view = (request.args.get("vista") or "").strip().lower()
        initial_loader_mode = "titulares" if requested_loader_view == "titulares" else "manual"
        manual_ejercicios_rows = db.execute(
            """
            SELECT DISTINCT TRIM(COALESCE(ejercicio, '')) AS ejercicio
            FROM entes_detalle
            WHERE TRIM(COALESCE(ejercicio, '')) != ''
            ORDER BY CAST(TRIM(COALESCE(ejercicio, '')) AS INTEGER) DESC, ejercicio DESC
            """
        ).fetchall()
        manual_ejercicios = [row["ejercicio"] for row in manual_ejercicios_rows]
        if not manual_ejercicios:
            manual_ejercicios = [TITULAR_EJERCICIO_FIJO]
        manual_ejercicio_default = manual_ejercicios[0] if manual_ejercicios else TITULAR_EJERCICIO_FIJO
        fuentes_rows = db.execute(
            """
            SELECT id, nombre
            FROM fuentes_financiamiento
            ORDER BY nombre ASC
            """
        ).fetchall()
        fuentes = [dict(row) for row in fuentes_rows]
    
        titular_ejercicios = manual_ejercicios.copy()

        script_result = None
        manual_result = None
        titular_result = None
        form_data = {
            "template_ejercicio": "",
            "template_out": "bases/historial_titulares_template.csv",
            "template_tipo_auditoria": "Financiera",
            "csv_path": "",
            "csv_ejercicio": "",
            "csv_replace": "0",
            "json_path": "",
            "json_ejercicio": "",
            "json_tipo_auditoria": "Financiera",
            "manual_id": "",
            "manual_ente_id": "",
            "manual_tipo_auditoria": "",
            "manual_tipo_responsable": "Titular",
            "manual_titular_nombre": "",
            "manual_administrativo_nombre": "",
            "manual_numero_oficio": "",
            "manual_oficio_base": "",
            "manual_asunto": "Notificación de Cédula de Resultados",
            "manual_ejercicio": manual_ejercicio_default,
            "manual_fuente_id": "",
            "manual_fuente_nueva": "",
            "manual_fuentes_detalle_json": "",
            "manual_periodo": "",
            "manual_periodo_titular": "",
            "manual_ramo_33": "No",
            "manual_estado": "E",
            "manual_fecha_notificacion": "",
            "manual_cantidad_sa": "0",
            "manual_cantidad_pdp": "0",
            "manual_cantidad_pras": "0",
            "manual_cantidad_pefcf": "0",
            "manual_cantidad_r": "0",
            "manual_monto_pdp_emitido": "0",
            "manual_monto_pdp_solventado": "0",
            "manual_monto_pdp_pendiente": "0",
            "manual_montos_pdp": "",
            "manual_pdp_detalle_json": "",
            "titular_ejercicio": manual_ejercicio_default,
            "titular_ente_id": "",
            "titular_tipo_auditoria": "Financiera",
            "titular_periodo_informe": "",
            "titular_nombre": "",
            "titular_periodo_administrativo": "",
            "titular_administrativo": "",
            "titular_cedula_resultados": "",
        }
    
        action = ""
        if request.method == "POST":
            action = (request.form.get("action") or "").strip()
            form_data.update(
                {
                    "template_ejercicio": (request.form.get("template_ejercicio") or "").strip(),
                    "template_out": (request.form.get("template_out") or "").strip() or "bases/historial_titulares_template.csv",
                    "template_tipo_auditoria": (request.form.get("template_tipo_auditoria") or "").strip() or "Financiera",
                    "csv_path": (request.form.get("csv_path") or "").strip(),
                    "csv_ejercicio": (request.form.get("csv_ejercicio") or "").strip(),
                    "csv_replace": "1" if request.form.get("csv_replace") else "0",
                    "json_path": (request.form.get("json_path") or "").strip(),
                    "json_ejercicio": (request.form.get("json_ejercicio") or "").strip(),
                    "json_tipo_auditoria": (request.form.get("json_tipo_auditoria") or "").strip() or "Financiera",
                    "manual_id": (request.form.get("manual_id") or "").strip(),
                    "manual_ente_id": (request.form.get("manual_ente_id") or "").strip(),
                    "manual_tipo_auditoria": (request.form.get("manual_tipo_auditoria") or "").strip(),
                    "manual_tipo_responsable": (request.form.get("manual_tipo_responsable") or "").strip() or "Titular",
                    "manual_titular_nombre": (request.form.get("manual_titular_nombre") or "").strip(),
                    "manual_administrativo_nombre": (request.form.get("manual_administrativo_nombre") or "").strip(),
                    "manual_numero_oficio": (request.form.get("manual_numero_oficio") or "").strip(),
                    "manual_oficio_base": " ".join((request.form.get("manual_oficio_base") or "").split()),
                    "manual_asunto": (request.form.get("manual_asunto") or "").strip() or "Notificación de Cédula de Resultados",
                    "manual_ejercicio": (request.form.get("manual_ejercicio") or "").strip() or manual_ejercicio_default,
                    "manual_fuente_id": (request.form.get("manual_fuente_id") or "").strip(),
                    "manual_fuente_nueva": (request.form.get("manual_fuente_nueva") or "").strip(),
                    "manual_fuentes_detalle_json": (request.form.get("manual_fuentes_detalle_json") or "").strip(),
                    "manual_periodo": (request.form.get("manual_periodo") or "").strip(),
                    "manual_periodo_titular": (request.form.get("manual_periodo_titular") or "").strip(),
                    "manual_ramo_33": (request.form.get("manual_ramo_33") or "").strip() or "No",
                    "manual_estado": (request.form.get("manual_estado") or "").strip() or "E",
                    "manual_fecha_notificacion": (request.form.get("manual_fecha_notificacion") or "").strip(),
                    "manual_cantidad_sa": (request.form.get("manual_cantidad_sa") or "").strip() or "0",
                    "manual_cantidad_pdp": (request.form.get("manual_cantidad_pdp") or "").strip() or "0",
                    "manual_cantidad_pras": (request.form.get("manual_cantidad_pras") or "").strip() or "0",
                    "manual_cantidad_pefcf": (request.form.get("manual_cantidad_pefcf") or "").strip() or "0",
                    "manual_cantidad_r": (request.form.get("manual_cantidad_r") or "").strip() or "0",
                    "manual_monto_pdp_emitido": (request.form.get("manual_monto_pdp_emitido") or "").strip() or "0",
                    "manual_monto_pdp_solventado": (request.form.get("manual_monto_pdp_solventado") or "").strip() or "0",
                    "manual_monto_pdp_pendiente": (request.form.get("manual_monto_pdp_pendiente") or "").strip() or "0",
                    "manual_montos_pdp": (request.form.get("manual_montos_pdp") or "").strip(),
                    "manual_pdp_detalle_json": (request.form.get("manual_pdp_detalle_json") or "").strip(),
                }
            )
            form_data.update(
                _build_titular_form_data(
                    request.form,
                    default_ejercicio=form_data["manual_ejercicio"],
                )
            )
    
            try:
                if action == "titular_save":
                    titular_result = _save_titulares_capture(db, user, form_data)
                elif action in {"manual_check", "manual_save"}:
                    manual_id_raw = form_data["manual_id"]
                    manual_ente_id = form_data["manual_ente_id"]
                    tipo_auditoria = form_data["manual_tipo_auditoria"]
                    tipo_responsable = "Titular"
                    titular_nombre = ""
                    administrativo_nombre = ""
                    numero_oficio = form_data["manual_numero_oficio"]
                    oficio_base = " ".join(form_data["manual_oficio_base"].split())
                    asunto = form_data["manual_asunto"]
                    ejercicio = form_data["manual_ejercicio"]
                    fuente_id_raw = form_data["manual_fuente_id"]
                    fuente_nueva = " ".join(form_data["manual_fuente_nueva"].split()) if fuente_id_raw == "__new__" else ""
                    periodo = " ".join(form_data["manual_periodo"].split())
                    periodo_titular = form_data["manual_periodo_titular"]
                    ramo_33 = "No"
                    estado = "E"
                    fecha_notificacion = form_data["manual_fecha_notificacion"]
                    raw_montos_pdp = form_data["manual_montos_pdp"]
                    raw_pdp_detalle_json = form_data["manual_pdp_detalle_json"]
                    fuentes_detalle_rows = parse_manual_fuentes_detalle(form_data["manual_fuentes_detalle_json"])
                    manual_edit_id = None
                    if manual_id_raw:
                        try:
                            manual_edit_id = int(manual_id_raw)
                        except ValueError as exc:
                            raise ValueError("ID de edición manual inválido.") from exc

                    if not manual_ente_id:
                        raise ValueError("Debes seleccionar un ente.")
                    if not numero_oficio:
                        raise ValueError("Debes capturar el número de oficio.")
                    if asunto not in ASUNTOS_MANUALES:
                        raise ValueError("Debes seleccionar un asunto válido.")
                    if not ejercicio:
                        raise ValueError("Debes capturar el ejercicio.")
                    if ejercicio not in manual_ejercicios:
                        raise ValueError("El ejercicio seleccionado no está disponible.")
                    if not fecha_notificacion:
                        raise ValueError("Debes capturar la fecha de notificación.")
                    try:
                        int(ejercicio)
                    except ValueError as exc:
                        raise ValueError("Ejercicio inválido.") from exc
                    if not periodo:
                        raise ValueError("Debes capturar el periodo.")

                    if not tipo_auditoria:
                        raise ValueError("Debes seleccionar el tipo de auditoría.")
                    if tipo_auditoria not in {"Financiera", "Obra Pública"}:
                        raise ValueError("Tipo de auditoría inválido.")
                    fuente_id = None
                    fuente_nombre = ""
                    usa_fuentes_detalle = (
                        asunto == "Notificación de Cédula de Resultados"
                        and len(fuentes_detalle_rows) > 0
                    )
                    if usa_fuentes_detalle and manual_edit_id:
                        raise ValueError("La edición con múltiples fuentes no está soportada.")
                    if not usa_fuentes_detalle and not fuente_id_raw:
                        raise ValueError("Debes seleccionar una fuente.")
                    if not usa_fuentes_detalle and fuente_id_raw == "__new__" and not fuente_nueva:
                        raise ValueError("Debes escribir la nueva fuente.")
                    if usa_fuentes_detalle:
                        fuente_nombre = fuentes_detalle_rows[0]["fuente_nombre"]
                        fuente_row = db.execute(
                            """
                            SELECT id
                            FROM fuentes_financiamiento
                            WHERE LOWER(TRIM(nombre)) = LOWER(TRIM(?))
                            LIMIT 1
                            """,
                            (fuente_nombre,),
                        ).fetchone()
                        if fuente_row:
                            fuente_id = int(fuente_row["id"])
                        elif action == "manual_save":
                            cursor_fuente = db.execute(
                                """
                                INSERT INTO fuentes_financiamiento (nombre, created_at)
                                VALUES (?, ?)
                                """,
                                (
                                    fuente_nombre,
                                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                                ),
                            )
                            fuente_id = int(cursor_fuente.lastrowid)
                        else:
                            fuente_id = -1
                    elif fuente_id_raw == "__new__":
                        fuente_nombre = fuente_nueva
                        fuente_row = db.execute(
                            """
                            SELECT id
                            FROM fuentes_financiamiento
                            WHERE LOWER(TRIM(nombre)) = LOWER(TRIM(?))
                            LIMIT 1
                            """,
                            (fuente_nueva,),
                        ).fetchone()
                        if fuente_row:
                            fuente_id = int(fuente_row["id"])
                        elif action == "manual_save":
                            cursor_fuente = db.execute(
                                """
                                INSERT INTO fuentes_financiamiento (nombre, created_at)
                                VALUES (?, ?)
                                """,
                                (
                                    fuente_nueva,
                                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                                ),
                            )
                            fuente_id = int(cursor_fuente.lastrowid)
                        else:
                            # En verificación no insertamos todavía una fuente nueva.
                            fuente_id = -1
                    else:
                        if fuente_id_raw.startswith("__obs__:"):
                            fuente_obs = " ".join(fuente_id_raw.replace("__obs__:", "", 1).split())
                            if not fuente_obs:
                                raise ValueError("Debes seleccionar una fuente válida.")
                            fuente_nombre = fuente_obs
                            fuente_row = db.execute(
                                """
                                SELECT id
                                FROM fuentes_financiamiento
                                WHERE LOWER(TRIM(nombre)) = LOWER(TRIM(?))
                                LIMIT 1
                                """,
                                (fuente_obs,),
                            ).fetchone()
                            if fuente_row:
                                fuente_id = int(fuente_row["id"])
                            elif action == "manual_save":
                                cursor_fuente = db.execute(
                                    """
                                    INSERT INTO fuentes_financiamiento (nombre, created_at)
                                    VALUES (?, ?)
                                    """,
                                    (
                                        fuente_obs,
                                        datetime.now().strftime("%Y-%m-%d %H:%M"),
                                    ),
                                )
                                fuente_id = int(cursor_fuente.lastrowid)
                            else:
                                fuente_id = -1
                        else:
                            try:
                                fuente_id = int(fuente_id_raw)
                            except ValueError as exc:
                                raise ValueError("Debes seleccionar una fuente válida.") from exc
                            fuente_exists = db.execute(
                                "SELECT 1 FROM fuentes_financiamiento WHERE id = ? LIMIT 1",
                                (fuente_id,),
                            ).fetchone()
                            if not fuente_exists:
                                raise ValueError("La fuente seleccionada no existe.")
                    ente_row = db.execute(
                        """
                        SELECT
                            TRIM(COALESCE(ente_id, '')) AS ente_id,
                            TRIM(COALESCE(ente_numero, '')) AS ente_numero,
                            TRIM(COALESCE(ente_nombre, '')) AS ente_nombre
                        FROM entes_detalle
                        WHERE TRIM(COALESCE(ejercicio, '')) = ?
                          AND TRIM(COALESCE(ente_id, '')) = ?
                        LIMIT 1
                        """,
                        (ejercicio, manual_ente_id),
                    ).fetchone()
                    if not ente_row:
                        raise ValueError("El ente seleccionado no existe para ejercicio 2025.")
                    if fuente_id is not None and fuente_id >= 0:
                        fuente_nombre_row = db.execute(
                            """
                            SELECT TRIM(COALESCE(nombre, '')) AS nombre
                            FROM fuentes_financiamiento
                            WHERE id = ?
                            LIMIT 1
                            """,
                            (fuente_id,),
                        ).fetchone()
                        if not fuente_nombre_row or not (fuente_nombre_row["nombre"] or "").strip():
                            raise ValueError("No se pudo resolver el nombre de la fuente seleccionada.")
                        fuente_nombre = (fuente_nombre_row["nombre"] or "").strip()
                    elif not fuente_nombre:
                        raise ValueError("No se pudo resolver el nombre de la fuente seleccionada.")

                    # Regla automática para carga Gabo:
                    # Ramo XXXIII fijo en "No" y E/R en "R" para fuentes de remanentes.
                    if re.match(r"^(remanentes|rea)\b", fuente_nombre.strip(), flags=re.IGNORECASE):
                        estado = "R"
                    else:
                        estado = "E"
                    if asunto == "Notificación de Cédula de Resultados":
                        pdp_details_by_fuente: list[list[dict]] = []
                        if usa_fuentes_detalle:
                            first_row = fuentes_detalle_rows[0]
                            tipo_auditoria = str(first_row["tipo_auditoria"])
                            form_data["manual_tipo_auditoria"] = tipo_auditoria
                            cantidad_sa = int(first_row["cantidad_sa"])
                            cantidad_pdp = int(first_row["cantidad_pdp"])
                            cantidad_pras = int(first_row["cantidad_pras"])
                            cantidad_pefcf = int(first_row["cantidad_pefcf"])
                            cantidad_r = int(first_row["cantidad_r"])
                        else:
                            cantidad_sa = parse_non_negative_int(form_data["manual_cantidad_sa"], "Cantidad SA")
                            cantidad_pdp = parse_non_negative_int(form_data["manual_cantidad_pdp"], "Cantidad PDP")
                            cantidad_pras = parse_non_negative_int(form_data["manual_cantidad_pras"], "Cantidad PRAS")
                            cantidad_pefcf = parse_non_negative_int(form_data["manual_cantidad_pefcf"], "Cantidad PEFCF")
                            cantidad_r = parse_non_negative_int(form_data["manual_cantidad_r"], "Cantidad R")
                        monto_pdp_emitido = parse_non_negative_float(form_data["manual_monto_pdp_emitido"], "Monto PDP emitido")
                        monto_pdp_solventado = parse_non_negative_float(form_data["manual_monto_pdp_solventado"], "Monto PDP solventado")
                        monto_pdp_pendiente = parse_non_negative_float(form_data["manual_monto_pdp_pendiente"], "Monto PDP pendiente")
                        if usa_fuentes_detalle:
                            total_pdp_fuentes = sum(
                                int(row["cantidad_pdp"]) for row in fuentes_detalle_rows
                            )
                            pdp_details = parse_manual_pdp_details(
                                raw_pdp_detalle_json, total_pdp_fuentes
                            )
                            pdp_amounts = [
                                float(item.get("monto") or 0.0) for item in pdp_details
                            ]
                            pdp_offset = 0
                            for fuente_row in fuentes_detalle_rows:
                                cantidad_row = int(fuente_row["cantidad_pdp"])
                                next_offset = pdp_offset + cantidad_row
                                pdp_details_by_fuente.append(pdp_details[pdp_offset:next_offset])
                                pdp_offset = next_offset
                        else:
                            pdp_amounts = parse_pdp_amounts(raw_montos_pdp)
                            pdp_details = parse_manual_pdp_details(
                                raw_pdp_detalle_json, cantidad_pdp
                            )
                        solventacion_totales_by_anexo = {}
                        if not usa_fuentes_detalle and cantidad_pdp == 0 and pdp_amounts:
                            raise ValueError("Capturaste montos PDP pero la cantidad PDP es 0.")
                        if not usa_fuentes_detalle and pdp_amounts and len(pdp_amounts) != cantidad_pdp:
                            raise ValueError(
                                "La cantidad de montos PDP no coincide con 'Cantidad PDP'. "
                                "Captura un monto por línea."
                            )
                        if not usa_fuentes_detalle and not pdp_amounts and cantidad_pdp > 0:
                            # Sin detalle por observación: concentrar el total en la primera PDP.
                            pdp_amounts = [monto_pdp_emitido] + [0.0] * (cantidad_pdp - 1)
                        if not usa_fuentes_detalle and cantidad_pdp > 0 and pdp_details:
                            has_detail_amount = any((item.get("monto") is not None) for item in pdp_details)
                            if has_detail_amount:
                                pdp_amounts = [(item.get("monto") or 0.0) for item in pdp_details]
                    tipos_auditoria = [tipo_auditoria]
                    if manual_edit_id and len(tipos_auditoria) > 1:
                        raise ValueError(
                            "Para editar una captura existente, selecciona solo 'Financiero'."
                        )
                    if manual_edit_id:
                        owned_row = db.execute(
                            """
                            SELECT id
                            FROM cargas_manuales
                            WHERE id = ? AND created_by = ?
                            LIMIT 1
                            """,
                            (manual_edit_id, user["username"]),
                        ).fetchone()
                        if not owned_row:
                            raise ValueError("La captura manual a editar no existe o no te pertenece.")
                    existing_rows = []
                    for tipo_item in tipos_auditoria:
                        if manual_edit_id:
                            found = db.execute(
                                """
                                SELECT id, created_at, created_by, tipo_auditoria
                                FROM cargas_manuales
                                WHERE id != ?
                                  AND LOWER(TRIM(numero_oficio)) = LOWER(TRIM(?))
                                  AND TRIM(COALESCE(ente_id, '')) = TRIM(COALESCE(?, ''))
                                  AND asunto = ?
                                  AND ejercicio = ?
                                  AND fuente_id = ?
                                  AND LOWER(TRIM(periodo)) = LOWER(TRIM(?))
                                  AND tipo_auditoria = ?
                                ORDER BY id DESC
                                LIMIT 1
                                """,
                                (
                                    manual_edit_id,
                                    numero_oficio,
                                    manual_ente_id,
                                    asunto,
                                    ejercicio,
                                    fuente_id,
                                    periodo,
                                    tipo_item,
                                ),
                            ).fetchone()
                        else:
                            found = db.execute(
                                """
                                SELECT id, created_at, created_by, tipo_auditoria
                                FROM cargas_manuales
                                WHERE LOWER(TRIM(numero_oficio)) = LOWER(TRIM(?))
                                  AND TRIM(COALESCE(ente_id, '')) = TRIM(COALESCE(?, ''))
                                  AND asunto = ?
                                  AND ejercicio = ?
                                  AND fuente_id = ?
                                  AND LOWER(TRIM(periodo)) = LOWER(TRIM(?))
                                  AND tipo_auditoria = ?
                                ORDER BY id DESC
                                LIMIT 1
                                """,
                                (
                                    numero_oficio,
                                    manual_ente_id,
                                    asunto,
                                    ejercicio,
                                    fuente_id,
                                    periodo,
                                    tipo_item,
                                ),
                            ).fetchone()
                        if found:
                            existing_rows.append(found)

                    if action == "manual_check":
                        if existing_rows:
                            resumen = ", ".join(
                                f"{row['tipo_auditoria']} (ID {row['id']}, {row['created_at']})"
                                for row in existing_rows
                            )
                            manual_result = {
                                "ok": False,
                                "level": "info",
                                "message": f"Ya existe registro para: {resumen}.",
                            }
                        else:
                            manual_result = {
                                "ok": True,
                                "level": "success",
                                "message": "No se encontró registro previo con esta clave. Puedes guardar.",
                            }
                    else:
                        if manual_edit_id:
                            if existing_rows:
                                resumen = ", ".join(
                                    f"{row['tipo_auditoria']} (ID {row['id']}, {row['created_at']})"
                                    for row in existing_rows
                                )
                                manual_result = {
                                    "ok": False,
                                    "level": "info",
                                    "message": f"No se puede actualizar por duplicado: {resumen}.",
                                }
                            else:
                                db.execute(
                                    """
                                    UPDATE cargas_manuales
                                    SET ente_id = ?,
                                        ente_nombre = ?,
                                        tipo_auditoria = ?,
                                        tipo_responsable = ?,
                                        titular_nombre = ?,
                                        administrativo_nombre = ?,
                                        numero_oficio = ?,
                                        asunto = ?,
                                        ejercicio = ?,
                                        fuente_id = ?,
                                        periodo = ?,
                                        cantidad_sa = ?,
                                        cantidad_pdp = ?,
                                        cantidad_pras = ?,
                                        cantidad_pefcf = ?,
                                        cantidad_r = ?,
                                        monto_pdp_emitido = ?,
                                        monto_pdp_solventado = ?,
                                        monto_pdp_pendiente = ?,
                                        created_at = ?
                                    WHERE id = ? AND created_by = ?
                                    """,
                                    (
                                        manual_ente_id,
                                        (ente_row["ente_nombre"] or "").strip(),
                                        tipos_auditoria[0],
                                        tipo_responsable,
                                        titular_nombre or None,
                                        administrativo_nombre or None,
                                        numero_oficio,
                                        asunto,
                                        ejercicio,
                                        fuente_id,
                                        periodo,
                                        cantidad_sa,
                                        cantidad_pdp,
                                        cantidad_pras,
                                        cantidad_pefcf,
                                        cantidad_r,
                                        monto_pdp_emitido,
                                        monto_pdp_solventado,
                                        monto_pdp_pendiente,
                                        datetime.now().strftime("%Y-%m-%d %H:%M"),
                                        manual_edit_id,
                                        user["username"],
                                    ),
                                )
                                materialize_observaciones_from_manual(
                                    db,
                                    ejercicio=ejercicio,
                                    ente_id=manual_ente_id,
                                    ente_numero=(ente_row["ente_numero"] or "").strip(),
                                    ente_nombre=(ente_row["ente_nombre"] or "").strip(),
                                    tipo_auditoria=tipos_auditoria[0],
                                    fuente_nombre=fuente_nombre,
                                    ramo_33=ramo_33,
                                    estado=estado,
                                    periodo_cedula=periodo,
                                    periodo_titular=periodo_titular,
                                    oficio=numero_oficio,
                                    fecha_notificacion=fecha_notificacion,
                                    cantidad_sa=cantidad_sa,
                                    cantidad_pdp=cantidad_pdp,
                                    cantidad_pras=cantidad_pras,
                                    cantidad_pefcf=cantidad_pefcf,
                                    cantidad_r=cantidad_r,
                                    monto_pdp_solventado=monto_pdp_solventado,
                                    monto_pdp_pendiente=monto_pdp_pendiente,
                                    pdp_amounts=pdp_amounts,
                                    pdp_details=pdp_details,
                                    solventacion_totales_by_anexo=solventacion_totales_by_anexo,
                                    replace_scope=True,
                                )
                                db.commit()
                                form_data["manual_id"] = str(manual_edit_id)
                                manual_result = {
                                    "ok": True,
                                    "level": "success",
                                    "message": "Captura manual actualizada. Puedes seguir editando.",
                                }
                        else:
                            tipos_existentes = {row["tipo_auditoria"] for row in existing_rows}
                            tipos_por_insertar = [tipo for tipo in tipos_auditoria if tipo not in tipos_existentes]
                            if not tipos_por_insertar:
                                resumen = ", ".join(
                                    f"{row['tipo_auditoria']} (ID {row['id']}, {row['created_at']})"
                                    for row in existing_rows
                                )
                                manual_result = {
                                    "ok": False,
                                    "level": "info",
                                    "message": f"Registro duplicado detectado para: {resumen}.",
                                }
                            else:
                                inserted_ids = []
                                for tipo_item in tipos_por_insertar:
                                    pdp_details_tipo_item = pdp_details
                                    pdp_amounts_tipo_item = pdp_amounts
                                    if usa_fuentes_detalle:
                                        pdp_details_tipo_item = (
                                            pdp_details_by_fuente[0] if pdp_details_by_fuente else []
                                        )
                                        pdp_amounts_tipo_item = [
                                            float(item.get("monto") or 0.0)
                                            for item in pdp_details_tipo_item
                                        ]
                                    cursor = db.execute(
                                        """
                                        INSERT INTO cargas_manuales (
                                            ente_id,
                                            ente_nombre,
                                            tipo_auditoria,
                                            tipo_responsable,
                                            titular_nombre,
                                            administrativo_nombre,
                                            numero_oficio,
                                            asunto,
                                            ejercicio,
                                            fuente_id,
                                            periodo,
                                            cantidad_sa,
                                            cantidad_pdp,
                                            cantidad_pras,
                                            cantidad_pefcf,
                                            cantidad_r,
                                            monto_pdp_emitido,
                                            monto_pdp_solventado,
                                            monto_pdp_pendiente,
                                            created_by,
                                            created_at
                                        )
                                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                        """,
                                        (
                                            manual_ente_id,
                                            (ente_row["ente_nombre"] or "").strip(),
                                            tipo_item,
                                            tipo_responsable,
                                            titular_nombre or None,
                                            administrativo_nombre or None,
                                            numero_oficio,
                                            asunto,
                                            ejercicio,
                                            fuente_id,
                                            periodo,
                                            cantidad_sa,
                                            cantidad_pdp,
                                            cantidad_pras,
                                            cantidad_pefcf,
                                            cantidad_r,
                                            monto_pdp_emitido,
                                            monto_pdp_solventado,
                                            monto_pdp_pendiente,
                                            user["username"],
                                            datetime.now().strftime("%Y-%m-%d %H:%M"),
                                        ),
                                    )
                                    inserted_ids.append(int(cursor.lastrowid))
                                    materialize_observaciones_from_manual(
                                        db,
                                        ejercicio=ejercicio,
                                        ente_id=manual_ente_id,
                                        ente_numero=(ente_row["ente_numero"] or "").strip(),
                                        ente_nombre=(ente_row["ente_nombre"] or "").strip(),
                                        tipo_auditoria=tipo_item,
                                        fuente_nombre=fuente_nombre,
                                        ramo_33=ramo_33,
                                        estado=estado,
                                        periodo_cedula=periodo,
                                        periodo_titular=periodo_titular,
                                        oficio=numero_oficio,
                                        fecha_notificacion=fecha_notificacion,
                                        cantidad_sa=cantidad_sa,
                                        cantidad_pdp=cantidad_pdp,
                                        cantidad_pras=cantidad_pras,
                                        cantidad_pefcf=cantidad_pefcf,
                                        cantidad_r=cantidad_r,
                                        monto_pdp_solventado=monto_pdp_solventado,
                                        monto_pdp_pendiente=monto_pdp_pendiente,
                                        pdp_amounts=pdp_amounts_tipo_item,
                                        pdp_details=pdp_details_tipo_item,
                                        solventacion_totales_by_anexo=solventacion_totales_by_anexo,
                                        replace_scope=False,
                                    )
                                db.commit()
                                if existing_rows:
                                    existentes = ", ".join(sorted(tipos_existentes))
                                    insertados = ", ".join(tipos_por_insertar)
                                    manual_result = {
                                        "ok": True,
                                        "level": "success",
                                        "message": (
                                            f"Registro manual guardado para: {insertados}. "
                                            f"Ya existían: {existentes}."
                                        ),
                                    }
                                    form_data["manual_id"] = ""
                                else:
                                    form_data["manual_id"] = str(inserted_ids[0]) if len(inserted_ids) == 1 else ""
                                    manual_result = {
                                        "ok": True,
                                        "level": "success",
                                        "message": (
                                            "Registro manual guardado correctamente. "
                                            "Puedes editar los campos y volver a guardar."
                                            if len(inserted_ids) == 1
                                            else "Registros manuales guardados correctamente."
                                        ),
                                    }
                                if (
                                    action == "manual_save"
                                    and asunto == "Notificación de Cédula de Resultados"
                                    and usa_fuentes_detalle
                                    and len(fuentes_detalle_rows) > 1
                                    and manual_result
                                    and manual_result.get("ok")
                                ):
                                    extra_inserted = 0
                                    extra_skipped = 0
                                    for extra_idx, extra_row in enumerate(
                                        fuentes_detalle_rows[1:], start=1
                                    ):
                                        extra_fuente_nombre = " ".join(str(extra_row["fuente_nombre"]).split())
                                        if not extra_fuente_nombre:
                                            continue
                                        extra_fuente_db = db.execute(
                                            """
                                            SELECT id
                                            FROM fuentes_financiamiento
                                            WHERE LOWER(TRIM(nombre)) = LOWER(TRIM(?))
                                            LIMIT 1
                                            """,
                                            (extra_fuente_nombre,),
                                        ).fetchone()
                                        if extra_fuente_db:
                                            extra_fuente_id = int(extra_fuente_db["id"])
                                        else:
                                            cursor_fuente_extra = db.execute(
                                                """
                                                INSERT INTO fuentes_financiamiento (nombre, created_at)
                                                VALUES (?, ?)
                                                """,
                                                (
                                                    extra_fuente_nombre,
                                                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                                                ),
                                            )
                                            extra_fuente_id = int(cursor_fuente_extra.lastrowid)
                                        extra_tipo = extra_row["tipo_auditoria"]
                                        extra_tipos = [extra_tipo]
                                        for extra_tipo_item in extra_tipos:
                                            duplicate_extra = db.execute(
                                                """
                                                SELECT id
                                                FROM cargas_manuales
                                                WHERE LOWER(TRIM(numero_oficio)) = LOWER(TRIM(?))
                                                  AND TRIM(COALESCE(ente_id, '')) = TRIM(COALESCE(?, ''))
                                                  AND asunto = ?
                                                  AND ejercicio = ?
                                                  AND fuente_id = ?
                                                  AND LOWER(TRIM(periodo)) = LOWER(TRIM(?))
                                                  AND tipo_auditoria = ?
                                                LIMIT 1
                                                """,
                                                (
                                                    numero_oficio,
                                                    manual_ente_id,
                                                    asunto,
                                                    ejercicio,
                                                    extra_fuente_id,
                                                    periodo,
                                                    extra_tipo_item,
                                                ),
                                            ).fetchone()
                                            if duplicate_extra:
                                                extra_skipped += 1
                                                continue
                                            cantidad_sa_extra = int(extra_row["cantidad_sa"])
                                            cantidad_pdp_extra = int(extra_row["cantidad_pdp"])
                                            cantidad_pras_extra = int(extra_row["cantidad_pras"])
                                            cantidad_pefcf_extra = int(extra_row["cantidad_pefcf"])
                                            cantidad_r_extra = int(extra_row["cantidad_r"])
                                            pdp_details_extra = (
                                                pdp_details_by_fuente[extra_idx]
                                                if extra_idx < len(pdp_details_by_fuente)
                                                else []
                                            )
                                            pdp_amounts_extra = [
                                                float(item.get("monto") or 0.0)
                                                for item in pdp_details_extra
                                            ]
                                            db.execute(
                                                """
                                                INSERT INTO cargas_manuales (
                                                    ente_id,
                                                    ente_nombre,
                                                    tipo_auditoria,
                                                    tipo_responsable,
                                                    titular_nombre,
                                                    administrativo_nombre,
                                                    numero_oficio,
                                                    asunto,
                                                    ejercicio,
                                                    fuente_id,
                                                    periodo,
                                                    cantidad_sa,
                                                    cantidad_pdp,
                                                    cantidad_pras,
                                                    cantidad_pefcf,
                                                    cantidad_r,
                                                    monto_pdp_emitido,
                                                    monto_pdp_solventado,
                                                    monto_pdp_pendiente,
                                                    created_by,
                                                    created_at
                                                )
                                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                                """,
                                                (
                                                    manual_ente_id,
                                                    (ente_row["ente_nombre"] or "").strip(),
                                                    extra_tipo_item,
                                                    tipo_responsable,
                                                    titular_nombre or None,
                                                    administrativo_nombre or None,
                                                    numero_oficio,
                                                    asunto,
                                                    ejercicio,
                                                    extra_fuente_id,
                                                    periodo,
                                                    cantidad_sa_extra,
                                                    cantidad_pdp_extra,
                                                    cantidad_pras_extra,
                                                    cantidad_pefcf_extra,
                                                    cantidad_r_extra,
                                                    0.0,
                                                    0.0,
                                                    0.0,
                                                    user["username"],
                                                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                                                ),
                                            )
                                            materialize_observaciones_from_manual(
                                                db,
                                                ejercicio=ejercicio,
                                                ente_id=manual_ente_id,
                                                ente_numero=(ente_row["ente_numero"] or "").strip(),
                                                ente_nombre=(ente_row["ente_nombre"] or "").strip(),
                                                tipo_auditoria=extra_tipo_item,
                                                fuente_nombre=extra_fuente_nombre,
                                                ramo_33=ramo_33,
                                                estado=estado,
                                                periodo_cedula=periodo,
                                                periodo_titular=periodo_titular,
                                                oficio=numero_oficio,
                                                fecha_notificacion=fecha_notificacion,
                                                cantidad_sa=cantidad_sa_extra,
                                                cantidad_pdp=cantidad_pdp_extra,
                                                cantidad_pras=cantidad_pras_extra,
                                                cantidad_pefcf=cantidad_pefcf_extra,
                                                cantidad_r=cantidad_r_extra,
                                                monto_pdp_solventado=0.0,
                                                monto_pdp_pendiente=0.0,
                                                pdp_amounts=pdp_amounts_extra,
                                                pdp_details=pdp_details_extra,
                                                solventacion_totales_by_anexo={},
                                                replace_scope=False,
                                            )
                                            extra_inserted += 1
                                    if extra_inserted > 0:
                                        db.commit()
                                    if manual_result and manual_result.get("ok"):
                                        manual_result["message"] = (
                                            f"{manual_result['message']} "
                                            f"Fuentes adicionales: {extra_inserted} agregadas"
                                            + (f", {extra_skipped} omitidas por duplicado." if extra_skipped else ".")
                                        )
                else:
                    command = [sys.executable]
                    if action == "template_generate":
                        if not form_data["template_ejercicio"]:
                            raise ValueError("Debes indicar el ejercicio para generar plantilla.")
                        out_path = resolve_project_path(form_data["template_out"], must_exist=False)
                        out_path.parent.mkdir(parents=True, exist_ok=True)
                        command.extend(
                            [
                                "scripts/make_historial_titulares_template.py",
                                "--db",
                                DB_PATH,
                                "--ejercicio",
                                str(int(form_data["template_ejercicio"])),
                                "--tipo-auditoria",
                                form_data["template_tipo_auditoria"],
                                "--out",
                                str(out_path),
                            ]
                        )
                    elif action in {"csv_validate", "csv_import"}:
                        csv_path = resolve_project_path(form_data["csv_path"], must_exist=True)
                        command.extend(
                            [
                                "scripts/import_historial_titulares.py",
                                "--db",
                                DB_PATH,
                                "--csv",
                                str(csv_path),
                            ]
                        )
                        if form_data["csv_ejercicio"]:
                            command.extend(["--ejercicio", str(int(form_data["csv_ejercicio"]))])
                        if form_data["csv_replace"] == "1":
                            command.append("--replace")
                        if action == "csv_validate":
                            command.append("--dry-run")
                    elif action in {"json_validate", "json_import"}:
                        json_path = resolve_project_path(form_data["json_path"], must_exist=True)
                        command.extend(
                            [
                                "scripts/import_historial_titulares_json.py",
                                "--db",
                                DB_PATH,
                                "--json",
                                str(json_path),
                                "--tipo-auditoria",
                                form_data["json_tipo_auditoria"],
                            ]
                        )
                        if form_data["json_ejercicio"]:
                            command.extend(["--ejercicio", str(int(form_data["json_ejercicio"]))])
                        if action == "json_validate":
                            command.append("--dry-run")
                    else:
                        raise ValueError("Acción de carga no soportada.")
                    script_result = run_loader_command(command)
            except ValueError as exc:
                if action == "manual_save":
                    manual_result = {
                        "ok": False,
                        "level": "error",
                        "message": str(exc),
                    }
                elif action == "manual_check":
                    manual_result = {
                        "ok": False,
                        "level": "error",
                        "message": str(exc),
                    }
                elif action == "titular_save":
                    titular_result = {
                        "ok": False,
                        "level": "error",
                        "message": str(exc),
                    }
                else:
                    script_result = {
                        "ok": False,
                        "returncode": 1,
                        "command": "",
                        "stdout": "",
                        "stderr": str(exc),
                    }

        if action == "titular_save" or titular_result is not None:
            initial_loader_mode = "titulares"
        elif action in {"manual_save", "manual_check"} or manual_result is not None:
            initial_loader_mode = "manual"
    
        titular_entes_rows = db.execute(
            """
            SELECT
                TRIM(COALESCE(ente_id, '')) AS ente_id,
                TRIM(COALESCE(ente_numero, '')) AS ente_numero,
                TRIM(COALESCE(ente_nombre, '')) AS ente_nombre
            FROM entes_detalle
            WHERE TRIM(COALESCE(ejercicio, '')) = ?
              AND TRIM(COALESCE(ente_id, '')) != ''
            ORDER BY CAST(COALESCE(NULLIF(ente_numero, ''), '0') AS REAL), ente_numero, ente_nombre
            """,
            (form_data["titular_ejercicio"],),
        ).fetchall()

        fuentes_rows = db.execute(
            """
            SELECT id, nombre
            FROM fuentes_financiamiento
            ORDER BY nombre ASC
            """
        ).fetchall()
        fuentes = [dict(row) for row in fuentes_rows]
        manual_ente_id_norm = normalize_ente_id(form_data["manual_ente_id"])
        manual_entes_rows = db.execute(
            """
            SELECT
                TRIM(COALESCE(ente_id, '')) AS ente_id,
                TRIM(COALESCE(ente_numero, '')) AS ente_numero,
                TRIM(COALESCE(ente_nombre, '')) AS ente_nombre
            FROM entes_detalle
            WHERE TRIM(COALESCE(ejercicio, '')) = ?
              AND TRIM(COALESCE(ente_id, '')) != ''
            ORDER BY CAST(COALESCE(NULLIF(ente_numero, ''), '0') AS REAL), ente_numero, ente_nombre
            """,
            (form_data["manual_ejercicio"],),
        ).fetchall()
        if manual_ente_id_norm:
            manual_fuentes = fuentes_por_ente(
                db,
                form_data["manual_ejercicio"],
                manual_ente_id_norm,
                tipo_auditoria=form_data["manual_tipo_auditoria"],
            )
        else:
            manual_fuentes = []
    
        return render_template(
            "carga.html",
            user=user,
            result=script_result,
            manual_result=manual_result,
            titular_result=titular_result,
            initial_loader_mode=initial_loader_mode,
            form_data=form_data,
            fuentes=fuentes,
            manual_fuentes=manual_fuentes,
            manual_ejercicios=manual_ejercicios,
            titular_ejercicios=titular_ejercicios,
            titular_entes=[dict(row) for row in titular_entes_rows],
            manual_entes=[dict(row) for row in manual_entes_rows],
            asuntos=[
                "Notificación de Cédula de Resultados",
            ],
            tipos_responsable=["Titular", "Administrativo", "Ambos"],
        )
    
