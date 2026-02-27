from datetime import datetime
from io import BytesIO
import time

from flask import jsonify, redirect, render_template, request, send_file, url_for
from openpyxl import Workbook
from openpyxl.styles import Font


def register_luis_routes(app, deps):
    globals().update(deps)
    dashboard_cache: dict[str, tuple[float, dict]] = {}
    dashboard_cache_ttl_seconds = 45
    @app.route("/entes", methods=["GET", "POST"])
    @luis_required
    def entes():
        if request.method == "POST":
            if get_current_user()["role"] != "editor":
                return redirect(url_for("index", notice="no_permission"))
            ejercicio = request.form.get("ente_ejercicio", "").strip()
            ente_id = request.form.get("ente_id", "").strip()
            ente_numero = request.form.get("ente_numero", "").strip()
            ente_nombre = request.form.get("ente_nombre", "").strip()
            responsable = request.form.get("ente_responsable", "").strip()
            clasificacion = request.form.get("ente_clasificacion", "").strip()
            ramo33 = request.form.get("ente_ramo33", "").strip()
    
            if not all([ejercicio, ente_id, ente_numero, ente_nombre]):
                return redirect(url_for("index", notice="ente_error"))
    
            db = get_db()
            ente_uid = resolve_ente_uid(db, ente_nombre)
            db.execute(
                """
                INSERT INTO entes_detalle (
                    ente_uid, ente_id, ejercicio, ente_numero, ente_nombre,
                    responsable, clasificacion, ramo33, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(ente_id, ejercicio) DO UPDATE SET
                    ente_uid = COALESCE(entes_detalle.ente_uid, excluded.ente_uid),
                    ente_numero = excluded.ente_numero,
                    ente_nombre = excluded.ente_nombre,
                    responsable = excluded.responsable,
                    clasificacion = excluded.clasificacion,
                    ramo33 = excluded.ramo33
                """,
                (
                    ente_uid,
                    ente_id,
                    ejercicio,
                    ente_numero,
                    ente_nombre,
                    responsable or "",
                    clasificacion or "",
                    ramo33 or "",
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                ),
            )
            db.commit()
            return redirect(url_for("index", notice="ente_saved"))
    
        ejercicio = request.args.get("ejercicio", "").strip()
        if not ejercicio:
            return jsonify([])
    
        db = get_db()
        rows = db.execute(
            """
            SELECT ente_id, ente_numero, ente_nombre, responsable, clasificacion, ramo33
            FROM entes_detalle
            WHERE ejercicio = ?
            ORDER BY CAST(ente_numero AS REAL) ASC, ente_numero ASC
            """,
            (ejercicio,),
        ).fetchall()
        return jsonify([dict(row) for row in rows])
    
    
    @app.route("/historial", methods=["GET", "POST"])
    @luis_required
    def historial():
        if request.method == "POST":
            if get_current_user()["role"] != "editor":
                return redirect(url_for("index", notice="no_permission"))
            ejercicio = request.form.get("historial_ejercicio", "").strip()
            ente_id = normalize_ente_id(request.form.get("historial_ente_id", ""))
            nombre = request.form.get("historial_nombre", "").strip()
            cargo = request.form.get("historial_cargo", "").strip()
            fecha_inicio = request.form.get("historial_fecha_inicio", "").strip()
            fecha_fin = request.form.get("historial_fecha_fin", "").strip()
            tipo_registro = request.form.get("historial_tipo_registro", "").strip()
            tipo_auditoria = request.form.get("historial_tipo_auditoria", "").strip() or "Financiera"
    
            if not all([ejercicio, ente_id, nombre, cargo, fecha_inicio, fecha_fin, tipo_registro]):
                return redirect(url_for("index", notice="historial_error"))
    
            db = get_db()
            ente_row = db.execute(
                f"""
                SELECT ente_nombre, ente_uid
                FROM entes_detalle
                WHERE ejercicio = ? AND {normalize_ente_id_sql('ente_id')} = ?
                """,
                (ejercicio, ente_id),
            ).fetchone()
            if ente_row is None:
                return redirect(url_for("index", notice="historial_error"))
            ente_nombre = ente_row["ente_nombre"]
            ente_uid = (ente_row["ente_uid"] or "").strip()
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
            db.commit()
            return redirect(url_for("index", notice="historial_saved"))
    
        ejercicio = request.args.get("ejercicio", "").strip()
        ente_id = normalize_ente_id(request.args.get("ente_id", ""))
        if not ejercicio or not ente_id:
            return jsonify([])
    
        db = get_db()
        ente_info = db.execute(
            f"""
            SELECT ente_uid, ente_nombre
            FROM entes_detalle
            WHERE ejercicio = ? AND {normalize_ente_id_sql('ente_id')} = ?
            LIMIT 1
            """,
            (ejercicio, ente_id),
        ).fetchone()
        if not ente_info:
            return jsonify([])
    
        ente_uid = (ente_info["ente_uid"] or "").strip()
        ente_aliases = get_ente_aliases_by_uid(
            db,
            ejercicio,
            ente_id,
            fallback_names=[ente_info["ente_nombre"]],
        )
    
        filter_clause = ""
        filter_params = []
        if ente_uid and ente_aliases:
            placeholders = ", ".join(["?"] * len(ente_aliases))
            filter_clause = (
                f"AND (TRIM(COALESCE(ente_uid, '')) = ? OR TRIM(COALESCE(ente, '')) IN ({placeholders}))"
            )
            filter_params.extend([ente_uid, *ente_aliases])
        elif ente_uid:
            filter_clause = "AND TRIM(COALESCE(ente_uid, '')) = ?"
            filter_params.append(ente_uid)
        elif ente_aliases:
            placeholders = ", ".join(["?"] * len(ente_aliases))
            filter_clause = f"AND TRIM(COALESCE(ente, '')) IN ({placeholders})"
            filter_params.extend(ente_aliases)
    
        rows = db.execute(
            f"""
            SELECT id, nombre, cargo, fecha_inicio, fecha_fin, tipo_registro, tipo_auditoria
            FROM historial_titulares
            WHERE ejercicio = ?
            {filter_clause}
            ORDER BY id DESC
            """,
            [ejercicio, *filter_params],
        ).fetchall()
        return jsonify([dict(row) for row in rows])
    
    
    @app.post("/oficios")
    @luis_required
    def oficios():
        user = get_current_user()
        if not user or user["username"] != "omar":
            return redirect(url_for("index", notice="no_permission"))
    
        ejercicio = request.form.get("oficio_ejercicio", "").strip()
        ente_id = request.form.get("oficio_ente_id", "").strip()
        oficio_numero = request.form.get("oficio_numero", "").strip()
        tipo_auditoria = request.form.get("tipo_auditoria", "").strip()
        fecha_notificacion = request.form.get("fecha_notificacion", "").strip()
        fuente_id = request.form.get("oficio_fuente_id", "").strip()
    
        if not all(
            [
                ejercicio,
                ente_id,
                oficio_numero,
                tipo_auditoria,
                fecha_notificacion,
                fuente_id,
            ]
        ):
            return redirect(url_for("index", notice="oficio_error"))
    
        db = get_db()
        ente_row = db.execute(
            """
            SELECT ente_id
            FROM entes_detalle
            WHERE ente_id = ? AND ejercicio = ?
            """,
            (ente_id, ejercicio),
        ).fetchone()
    
        if ente_row is None:
            return redirect(url_for("index", notice="oficio_error"))
    
        fuente_row = db.execute(
            """
            SELECT 1
            FROM registros
            WHERE ejercicio = ? AND ente_id = ? AND fuente_id = ?
            LIMIT 1
            """,
            (ejercicio, ente_id, fuente_id),
        ).fetchone()
        if fuente_row is None:
            return redirect(url_for("index", notice="oficio_error"))
    
        db.execute(
            """
            INSERT INTO oficios (
                ente_id, ejercicio, oficio, tipo_auditoria, fecha_notificacion, observaciones,
                fuente_id, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                ente_id,
                ejercicio,
                oficio_numero,
                tipo_auditoria,
                fecha_notificacion,
                "",
                int(fuente_id),
                datetime.now().strftime("%Y-%m-%d %H:%M"),
            ),
        )
        db.commit()
        return redirect(url_for("index", notice="oficio_saved"))
    
    
    @app.get("/fuentes-ente")
    @luis_required
    def fuentes_ente():
        ejercicio = request.args.get("ejercicio", "").strip()
        ente_id = request.args.get("ente_id", "").strip()
        if not ejercicio or not ente_id:
            return jsonify([])
    
        db = get_db()
        rows = db.execute(
            """
            SELECT DISTINCT fuentes_financiamiento.id, fuentes_financiamiento.nombre
            FROM registros
            JOIN fuentes_financiamiento
                ON registros.fuente_id = fuentes_financiamiento.id
            WHERE registros.ejercicio = ? AND registros.ente_id = ?
            ORDER BY fuentes_financiamiento.nombre ASC
            """,
            (ejercicio, ente_id),
        ).fetchall()
        return jsonify([dict(row) for row in rows])
    
    
    @app.get("/ejercicios-disponibles")
    @luis_required
    def ejercicios_disponibles():
        db = get_db()
        observaciones_rows = db.execute(
            """
            SELECT ejercicio, COUNT(*) AS total
            FROM observaciones
            GROUP BY ejercicio
            """
        ).fetchall()
        entes_rows = db.execute(
            """
            SELECT ejercicio, COUNT(*) AS total
            FROM entes_detalle
            GROUP BY ejercicio
            """
        ).fetchall()
    
        resumen = {}
        for row in observaciones_rows:
            ejercicio = row["ejercicio"]
            resumen.setdefault(
                ejercicio,
                {"ejercicio": ejercicio, "total_observaciones": 0, "total_entes": 0},
            )
            resumen[ejercicio]["total_observaciones"] = row["total"]
        for row in entes_rows:
            ejercicio = row["ejercicio"]
            resumen.setdefault(
                ejercicio,
                {"ejercicio": ejercicio, "total_observaciones": 0, "total_entes": 0},
            )
            resumen[ejercicio]["total_entes"] = row["total"]
    
        ordered = sorted(resumen.values(), key=lambda item: item["ejercicio"], reverse=True)
        return jsonify(ordered)


    @app.get("/observaciones-dashboard")
    @luis_required
    def observaciones_dashboard():
        ejercicio = request.args.get("ejercicio", "").strip()
        if not ejercicio:
            return jsonify({
                "rows": [],
                "total_rows": 0,
                "page": 1,
                "page_size": 40,
                "total_pages": 1,
                "filtros": {},
                "summary": {
                    "tipos": [],
                    "totals": {"emitidas": 0, "solventadas": 0, "pendientes": 0},
                    "pdp_montos": {"emitido": 0, "solventado": 0, "pendiente": 0},
                    "top_pendientes": [],
                },
            })

        ente_id = normalize_ente_id(request.args.get("ente_id", ""))
        tipo_anexo = request.args.get("tipo_anexo", "").strip()
        tipo_auditoria = normalize_tipo_auditoria(request.args.get("tipo_auditoria", ""))
        estado = request.args.get("estado", "").strip()
        fuente = request.args.get("fuente_financiamiento", "").strip()
        ramo_33 = request.args.get("ramo_33", "").strip()
        concepto_irregularidad = request.args.get("concepto_irregularidad", "").strip()
        periodo_cedula = request.args.get("periodo_cedula", "").strip()

        try:
            page = max(1, int(request.args.get("page", "1")))
        except ValueError:
            page = 1
        try:
            page_size = int(request.args.get("page_size", "40"))
        except ValueError:
            page_size = 40
        page_size = max(10, min(200, page_size))

        cache_key = request.query_string.decode("utf-8")
        now = time.time()
        cached = dashboard_cache.get(cache_key)
        if cached and now - cached[0] < dashboard_cache_ttl_seconds:
            return jsonify(cached[1])

        db = get_db()
        selected = {
            "tipo_auditoria": tipo_auditoria,
            "tipo_anexo": tipo_anexo,
            "estado": estado,
            "fuente_financiamiento": fuente,
            "ramo_33": ramo_33,
            "concepto_irregularidad": concepto_irregularidad,
            "periodo_cedula": periodo_cedula,
        }

        def build_scope(*, exclude_key: str = "", include_ente: bool = True):
            clauses = ["ejercicio = ?"]
            params = [ejercicio]
            if include_ente and ente_id:
                clauses.append(f"{normalize_ente_id_sql('ente_id')} = ?")
                params.append(ente_id)
            for key, value in selected.items():
                if key == exclude_key or not value:
                    continue
                if key == "tipo_auditoria":
                    clauses.append("tipo_auditoria = ?")
                    params.append(value)
                elif key == "tipo_anexo":
                    clauses.append("tipo_anexo = ?")
                    params.append(value)
                elif key == "estado":
                    clauses.append("estado = ?")
                    params.append(value)
                elif key == "fuente_financiamiento":
                    clauses.append("fuente_financiamiento = ?")
                    params.append(value)
                elif key == "ramo_33":
                    clauses.append("ramo_33 = ?")
                    params.append(value)
                elif key == "periodo_cedula":
                    clauses.append("periodo_cedula = ?")
                    params.append(value)
                elif key == "concepto_irregularidad":
                    clauses.append(
                        "(pdp_concepto_irregularidad = ? OR pdp_subconcepto_irregularidad = ?)"
                    )
                    params.extend([value, value])
            return " AND ".join(clauses), params

        scope_sql, scope_params = build_scope()
        total_rows = db.execute(
            f"SELECT COUNT(*) FROM observaciones WHERE {scope_sql}",
            scope_params,
        ).fetchone()[0]
        total_pages = max(1, (total_rows + page_size - 1) // page_size)
        page = min(page, total_pages)
        offset = (page - 1) * page_size

        rows = db.execute(
            f"""
            SELECT
                id,
                ejercicio,
                ente_id,
                ente_numero,
                ente_nombre,
                tipo_auditoria,
                fuente_financiamiento,
                ramo_33,
                periodo_cedula,
                periodo_titular,
                oficio,
                fecha_notificacion,
                tipo_anexo,
                numero_observacion,
                estado,
                monto_pdp_emitido,
                monto_pdp_solventado,
                monto_pdp_pendiente,
                pdp_no_irregularidad,
                pdp_concepto_irregularidad,
                pdp_subconcepto_irregularidad
            FROM observaciones
            WHERE {scope_sql}
            ORDER BY
                COALESCE(ente_numero_sort, 0) ASC,
                ente_numero ASC,
                ente_id ASC,
                tipo_anexo ASC,
                numero_observacion ASC
            LIMIT ? OFFSET ?
            """,
            [*scope_params, page_size, offset],
        ).fetchall()

        summary_rows = db.execute(
            f"""
            SELECT
                tipo_anexo,
                COUNT(*) AS emitidas,
                SUM(CASE WHEN LOWER(TRIM(COALESCE(estado, ''))) = 'solventado' THEN 1 ELSE 0 END) AS solventadas,
                SUM(CASE WHEN LOWER(TRIM(COALESCE(estado, ''))) = 'pendiente' THEN 1 ELSE 0 END) AS pendientes
            FROM observaciones
            WHERE {scope_sql}
              AND TRIM(COALESCE(tipo_anexo, '')) != ''
            GROUP BY tipo_anexo
            ORDER BY tipo_anexo
            """,
            scope_params,
        ).fetchall()
        pdp_montos_row = db.execute(
            f"""
            SELECT
                SUM(COALESCE(monto_pdp_emitido, 0)) AS emitido,
                SUM(COALESCE(monto_pdp_solventado, 0)) AS solventado,
                SUM(COALESCE(monto_pdp_pendiente, 0)) AS pendiente
            FROM observaciones
            WHERE {scope_sql}
              AND tipo_anexo = 'PDP'
            """,
            scope_params,
        ).fetchone()
        top_pendientes_rows = db.execute(
            f"""
            SELECT
                TRIM(COALESCE(ente_nombre, 'Sin ente')) AS ente_nombre,
                COUNT(*) AS pendientes
            FROM observaciones
            WHERE {scope_sql}
              AND LOWER(TRIM(COALESCE(estado, ''))) = 'pendiente'
            GROUP BY ente_id, ente_nombre
            ORDER BY pendientes DESC, ente_nombre ASC
            LIMIT 5
            """,
            scope_params,
        ).fetchall()

        def query_distinct(column: str, exclude_key: str):
            where_sql, where_params = build_scope(exclude_key=exclude_key)
            return db.execute(
                f"""
                SELECT DISTINCT {column} AS value
                FROM observaciones
                WHERE {where_sql}
                  AND {column} IS NOT NULL AND TRIM({column}) != ''
                ORDER BY value
                """,
                where_params,
            ).fetchall()

        entes_where, entes_params = build_scope(include_ente=False)
        entes = db.execute(
            f"""
            SELECT DISTINCT
                TRIM(COALESCE(ente_id, '')) AS ente_id,
                TRIM(COALESCE(ente_numero, '')) AS ente_numero,
                TRIM(COALESCE(ente_nombre, '')) AS ente_nombre
            FROM observaciones
            WHERE {entes_where}
              AND TRIM(COALESCE(ente_id, '')) != ''
            ORDER BY COALESCE(ente_numero_sort, 0), ente_numero, ente_nombre
            """,
            entes_params,
        ).fetchall()

        concepto_where, concepto_params = build_scope(exclude_key="concepto_irregularidad")
        conceptos = db.execute(
            f"""
            SELECT DISTINCT concepto
            FROM (
                SELECT pdp_concepto_irregularidad AS concepto
                FROM observaciones
                WHERE {concepto_where}
                  AND pdp_concepto_irregularidad IS NOT NULL AND TRIM(pdp_concepto_irregularidad) != ''
                UNION
                SELECT pdp_subconcepto_irregularidad AS concepto
                FROM observaciones
                WHERE {concepto_where}
                  AND pdp_subconcepto_irregularidad IS NOT NULL AND TRIM(pdp_subconcepto_irregularidad) != ''
            )
            ORDER BY concepto
            """,
            concepto_params + concepto_params,
        ).fetchall()

        filtros = {
            "tipo_anexo": [row[0] for row in query_distinct("tipo_anexo", "tipo_anexo")],
            "tipo_auditoria": [row[0] for row in query_distinct("tipo_auditoria", "tipo_auditoria")],
            "estado": [row[0] for row in query_distinct("estado", "estado")],
            "fuente_financiamiento": [row[0] for row in query_distinct("fuente_financiamiento", "fuente_financiamiento")],
            "ramo_33": [row[0] for row in query_distinct("ramo_33", "ramo_33")],
            "cedulas": [row[0] for row in query_distinct("periodo_cedula", "periodo_cedula")],
            "conceptos_irregularidad": [row[0] for row in conceptos],
            "entes": [dict(row) for row in entes],
        }

        tipos_summary = [dict(row) for row in summary_rows]
        totals = {"emitidas": 0, "solventadas": 0, "pendientes": 0}
        for row in tipos_summary:
            totals["emitidas"] += int(row.get("emitidas") or 0)
            totals["solventadas"] += int(row.get("solventadas") or 0)
            totals["pendientes"] += int(row.get("pendientes") or 0)

        payload = {
            "rows": [dict(row) for row in rows],
            "total_rows": total_rows,
            "page": page,
            "page_size": page_size,
            "total_pages": total_pages,
            "filtros": filtros,
            "summary": {
                "tipos": tipos_summary,
                "totals": totals,
                "pdp_montos": {
                    "emitido": float((pdp_montos_row["emitido"] or 0) if pdp_montos_row else 0),
                    "solventado": float((pdp_montos_row["solventado"] or 0) if pdp_montos_row else 0),
                    "pendiente": float((pdp_montos_row["pendiente"] or 0) if pdp_montos_row else 0),
                },
                "top_pendientes": [
                    {
                        "ente_nombre": row["ente_nombre"],
                        "pendientes": int(row["pendientes"] or 0),
                    }
                    for row in top_pendientes_rows
                ],
            },
        }
        dashboard_cache[cache_key] = (now, payload)
        if len(dashboard_cache) > 250:
            # Prevent unbounded growth in long-running processes.
            oldest_key = min(dashboard_cache.items(), key=lambda item: item[1][0])[0]
            dashboard_cache.pop(oldest_key, None)
        return jsonify(payload)
    
    
    @app.get("/observaciones")
    @luis_required
    def observaciones_api():
        ejercicio = request.args.get("ejercicio", "").strip()
        ente_id = normalize_ente_id(request.args.get("ente_id", ""))
        tipo_anexo = request.args.get("tipo_anexo", "").strip()
        tipo_auditoria = normalize_tipo_auditoria(request.args.get("tipo_auditoria", ""))
        estado = request.args.get("estado", "").strip()
        fuente = request.args.get("fuente_financiamiento", "").strip()
        ramo_33 = request.args.get("ramo_33", "").strip()
        concepto_irregularidad = request.args.get("concepto_irregularidad", "").strip()
        periodo_informe = request.args.get("periodo_informe", "").strip()
        titular = request.args.get("titular", "").strip()
        periodo_admin = request.args.get("periodo_admin", "").strip()
        administrativo = request.args.get("administrativo", "").strip()
        periodo_cedula = request.args.get("periodo_cedula", "").strip()
        search = request.args.get("search", "").strip()
        if not ejercicio:
            return jsonify([])
    
        db = get_db()
        params = [ejercicio]
        filter_clauses = []
        if ente_id:
            filter_clauses.append(f"{normalize_ente_id_sql('observaciones.ente_id')} = ?")
            params.append(ente_id)
        if tipo_anexo:
            filter_clauses.append("observaciones.tipo_anexo = ?")
            params.append(tipo_anexo)
        if tipo_auditoria:
            filter_clauses.append("observaciones.tipo_auditoria = ?")
            params.append(tipo_auditoria)
        if estado:
            filter_clauses.append("observaciones.estado = ?")
            params.append(estado)
        if fuente:
            filter_clauses.append("observaciones.fuente_financiamiento = ?")
            params.append(fuente)
        if ramo_33:
            filter_clauses.append("observaciones.ramo_33 = ?")
            params.append(ramo_33)
        if concepto_irregularidad:
            filter_clauses.append(
                "(observaciones.pdp_concepto_irregularidad = ? OR observaciones.pdp_subconcepto_irregularidad = ?)"
            )
            params.extend([concepto_irregularidad, concepto_irregularidad])
        if periodo_cedula:
            filter_clauses.append("observaciones.periodo_cedula = ?")
            params.append(periodo_cedula)
        if search:
            filter_clauses.append(
                """
                (
                    observaciones.ente_nombre LIKE ?
                    OR observaciones.oficio LIKE ?
                    OR observaciones.fuente_financiamiento LIKE ?
                    OR observaciones.pdp_concepto_irregularidad LIKE ?
                    OR observaciones.pdp_subconcepto_irregularidad LIKE ?
                    OR CAST(observaciones.numero_observacion AS TEXT) LIKE ?
                )
                """
            )
            search_term = f"%{search}%"
            params.extend([search_term] * 6)
    
        filter_sql = ""
        if filter_clauses:
            filter_sql = " AND " + " AND ".join(filter_clauses)
        needs_historial_join = bool(periodo_informe or titular or periodo_admin or administrativo)
    
        if needs_historial_join:
            if periodo_informe:
                filter_clauses.append(f"{periodo_sql('resp')} = ?")
                params.append(periodo_informe)
            if titular:
                filter_clauses.append("resp.nombre = ?")
                params.append(titular)
            if periodo_admin:
                filter_clauses.append(f"{periodo_sql('admin')} = ?")
                params.append(periodo_admin)
            if administrativo:
                filter_clauses.append("admin.nombre = ?")
                params.append(administrativo)
    
            join_filter_sql = " AND " + " AND ".join(filter_clauses) if filter_clauses else ""
            rows = db.execute(
                f"""
                SELECT
                    observaciones.id,
                    observaciones.ejercicio,
                    observaciones.ente_id,
                    observaciones.ente_numero,
                    observaciones.ente_nombre,
                    observaciones.tipo_auditoria,
                    observaciones.fuente_financiamiento,
                    observaciones.ramo_33,
                    observaciones.periodo_cedula,
                    observaciones.periodo_titular,
                    observaciones.oficio,
                    observaciones.fecha_notificacion,
                    observaciones.tipo_anexo,
                    observaciones.numero_observacion,
                    observaciones.estado,
                    observaciones.monto_pdp_emitido,
                    observaciones.monto_pdp_solventado,
                    observaciones.monto_pdp_pendiente,
                    observaciones.pdp_no_irregularidad,
                    observaciones.pdp_concepto_irregularidad,
                    observaciones.pdp_subconcepto_irregularidad
                FROM observaciones
                LEFT JOIN entes_detalle
                    ON {normalize_ente_id_sql("observaciones.ente_id")} = {normalize_ente_id_sql("entes_detalle.ente_id")}
                    AND observaciones.ejercicio = entes_detalle.ejercicio
                LEFT JOIN historial_titulares AS resp
                    ON resp.id = (
                        SELECT id
                        FROM historial_titulares
                        WHERE ejercicio = observaciones.ejercicio
                          AND tipo_auditoria = observaciones.tipo_auditoria
                          AND (
                              (
                                  TRIM(COALESCE(entes_detalle.ente_uid, '')) != ''
                                  AND TRIM(COALESCE(historial_titulares.ente_uid, '')) = TRIM(entes_detalle.ente_uid)
                              )
                              OR ente = COALESCE(observaciones.ente_nombre, entes_detalle.ente_nombre)
                          )
                          AND tipo_registro = 'titular'
                        ORDER BY id DESC
                        LIMIT 1
                    )
                LEFT JOIN historial_titulares AS admin
                    ON admin.id = (
                        SELECT id
                        FROM historial_titulares
                        WHERE ejercicio = observaciones.ejercicio
                          AND tipo_auditoria = observaciones.tipo_auditoria
                          AND (
                              (
                                  TRIM(COALESCE(entes_detalle.ente_uid, '')) != ''
                                  AND TRIM(COALESCE(historial_titulares.ente_uid, '')) = TRIM(entes_detalle.ente_uid)
                              )
                              OR ente = COALESCE(observaciones.ente_nombre, entes_detalle.ente_nombre)
                          )
                          AND tipo_registro = 'director_administrativo'
                        ORDER BY id DESC
                        LIMIT 1
                    )
                WHERE observaciones.ejercicio = ?
                {join_filter_sql}
                ORDER BY
                    COALESCE(observaciones.ente_numero_sort, 0) ASC,
                    observaciones.ente_numero ASC,
                    observaciones.ente_id ASC,
                    observaciones.tipo_anexo ASC,
                    observaciones.numero_observacion ASC
                """,
                params,
            ).fetchall()
        else:
            rows = db.execute(
                f"""
                SELECT
                    observaciones.id,
                    observaciones.ejercicio,
                    observaciones.ente_id,
                    observaciones.ente_numero,
                    observaciones.ente_nombre,
                    observaciones.tipo_auditoria,
                    observaciones.fuente_financiamiento,
                    observaciones.ramo_33,
                    observaciones.periodo_cedula,
                    observaciones.periodo_titular,
                    observaciones.oficio,
                    observaciones.fecha_notificacion,
                    observaciones.tipo_anexo,
                    observaciones.numero_observacion,
                    observaciones.estado,
                    observaciones.monto_pdp_emitido,
                    observaciones.monto_pdp_solventado,
                    observaciones.monto_pdp_pendiente,
                    observaciones.pdp_no_irregularidad,
                    observaciones.pdp_concepto_irregularidad,
                    observaciones.pdp_subconcepto_irregularidad
                FROM observaciones
                WHERE observaciones.ejercicio = ?
                {filter_sql}
                ORDER BY
                    COALESCE(observaciones.ente_numero_sort, 0) ASC,
                    observaciones.ente_numero ASC,
                    observaciones.ente_id ASC,
                    observaciones.tipo_anexo ASC,
                    observaciones.numero_observacion ASC
                """,
                params,
            ).fetchall()
    
        return jsonify([dict(row) for row in rows])
    
    
    @app.get("/observaciones-filtros")
    @luis_required
    def observaciones_filtros():
        ejercicio = request.args.get("ejercicio", "").strip()
        ente_id = normalize_ente_id(request.args.get("ente_id", ""))
        tipo_auditoria = normalize_tipo_auditoria(request.args.get("tipo_auditoria", ""))
        tipo_anexo = request.args.get("tipo_anexo", "").strip()
        estado = request.args.get("estado", "").strip()
        fuente = request.args.get("fuente_financiamiento", "").strip()
        ramo_33 = request.args.get("ramo_33", "").strip()
        concepto_irregularidad = request.args.get("concepto_irregularidad", "").strip()
        periodo_cedula = request.args.get("periodo_cedula", "").strip()
        titular_seleccionado = request.args.get("titular", "").strip()
        administrativo_seleccionado = request.args.get("administrativo", "").strip()
        include_historial = request.args.get("include_historial", "").strip() == "1"
        if not ejercicio:
            return jsonify({})
    
        db = get_db()
        filtros = {}
        selected = {
            "tipo_auditoria": tipo_auditoria,
            "tipo_anexo": tipo_anexo,
            "estado": estado,
            "fuente_financiamiento": fuente,
            "ramo_33": ramo_33,
            "concepto_irregularidad": concepto_irregularidad,
            "periodo_cedula": periodo_cedula,
        }
    
        def build_observaciones_scope(exclude_key: str = "", include_ente: bool = True):
            clauses = ["ejercicio = ?"]
            params = [ejercicio]
            if include_ente and ente_id:
                clauses.append(f"{normalize_ente_id_sql('ente_id')} = ?")
                params.append(ente_id)
    
            for key, value in selected.items():
                if key == exclude_key or not value:
                    continue
                if key == "tipo_auditoria":
                    clauses.append("tipo_auditoria = ?")
                    params.append(value)
                elif key == "tipo_anexo":
                    clauses.append("tipo_anexo = ?")
                    params.append(value)
                elif key == "estado":
                    clauses.append("estado = ?")
                    params.append(value)
                elif key == "fuente_financiamiento":
                    clauses.append("fuente_financiamiento = ?")
                    params.append(value)
                elif key == "ramo_33":
                    clauses.append("ramo_33 = ?")
                    params.append(value)
                elif key == "periodo_cedula":
                    clauses.append("periodo_cedula = ?")
                    params.append(value)
                elif key == "concepto_irregularidad":
                    clauses.append("(pdp_concepto_irregularidad = ? OR pdp_subconcepto_irregularidad = ?)")
                    params.extend([value, value])
            return " AND ".join(clauses), params
    
        def query_distinct(column: str, exclude_key: str):
            where_sql, where_params = build_observaciones_scope(exclude_key)
            return db.execute(
                f"""
                SELECT DISTINCT {column} AS value
                FROM observaciones
                WHERE {where_sql}
                  AND {column} IS NOT NULL AND TRIM({column}) != ''
                ORDER BY value
                """,
                where_params,
            ).fetchall()
    
        auditorias = query_distinct("tipo_auditoria", "tipo_auditoria")
        tipos = query_distinct("tipo_anexo", "tipo_anexo")
        estados = query_distinct("estado", "estado")
        fuentes = query_distinct("fuente_financiamiento", "fuente_financiamiento")
        ramos = query_distinct("ramo_33", "ramo_33")
        cedulas = query_distinct("periodo_cedula", "periodo_cedula")
        concepto_where, concepto_params = build_observaciones_scope("concepto_irregularidad")
        conceptos = db.execute(
            f"""
            SELECT DISTINCT concepto
            FROM (
                SELECT pdp_concepto_irregularidad AS concepto
                FROM observaciones
                WHERE {concepto_where}
                  AND pdp_concepto_irregularidad IS NOT NULL AND TRIM(pdp_concepto_irregularidad) != ''
                UNION
                SELECT pdp_subconcepto_irregularidad AS concepto
                FROM observaciones
                WHERE {concepto_where}
                  AND pdp_subconcepto_irregularidad IS NOT NULL AND TRIM(pdp_subconcepto_irregularidad) != ''
            )
            ORDER BY concepto
            """,
            concepto_params + concepto_params,
        ).fetchall()
        entes_where, entes_params = build_observaciones_scope(include_ente=False)
        entes = db.execute(
            f"""
            SELECT DISTINCT
                TRIM(COALESCE(ente_id, '')) AS ente_id,
                TRIM(COALESCE(ente_numero, '')) AS ente_numero,
                TRIM(COALESCE(ente_nombre, '')) AS ente_nombre
            FROM observaciones
            WHERE {entes_where}
              AND TRIM(COALESCE(ente_id, '')) != ''
            ORDER BY COALESCE(ente_numero_sort, 0), ente_numero, ente_nombre
            """,
            entes_params,
        ).fetchall()
        periodos_informe = []
        titulares = []
        periodos_admin = []
        administrativos = []
        if include_historial:
            ente_aliases = get_ente_aliases_by_uid(db, ejercicio, ente_id)
            ente_uid = get_ente_uid_by_ejercicio_id(db, ejercicio, ente_id)
    
            titular_params = [ejercicio]
            titular_clause = ""
            if ente_uid and ente_aliases:
                placeholders = ", ".join(["?"] * len(ente_aliases))
                titular_clause = (
                    f"AND (TRIM(COALESCE(ente_uid, '')) = ? OR TRIM(COALESCE(ente, '')) IN ({placeholders}))"
                )
                titular_params.extend([ente_uid, *ente_aliases])
            elif ente_uid:
                titular_clause = "AND TRIM(COALESCE(ente_uid, '')) = ?"
                titular_params.append(ente_uid)
            elif ente_aliases:
                placeholders = ", ".join(["?"] * len(ente_aliases))
                titular_clause = f"AND TRIM(COALESCE(ente, '')) IN ({placeholders})"
                titular_params.extend(ente_aliases)
    
            periodos_params = titular_params.copy()
            periodo_titular_clause = ""
            if titular_seleccionado:
                periodo_titular_clause = "AND nombre = ?"
                periodos_params.append(titular_seleccionado)
    
            titular_tipo_clause = ""
            titular_tipo_params = []
            if tipo_auditoria:
                titular_tipo_clause = "AND tipo_auditoria = ?"
                titular_tipo_params.append(tipo_auditoria)
    
            periodos_informe = db.execute(
                f"""
                SELECT DISTINCT {periodo_sql("historial_titulares")} AS periodo
                FROM historial_titulares
                WHERE ejercicio = ? {titular_clause} {periodo_titular_clause}
                  {titular_tipo_clause}
                  AND tipo_registro = 'titular'
                  AND fecha_inicio IS NOT NULL AND fecha_fin IS NOT NULL
                ORDER BY fecha_inicio
                """,
                periodos_params + titular_tipo_params,
            ).fetchall()
            titulares = db.execute(
                f"""
                SELECT DISTINCT nombre
                FROM historial_titulares
                WHERE ejercicio = ? {titular_clause}
                  {titular_tipo_clause}
                  AND tipo_registro = 'titular'
                  AND nombre IS NOT NULL AND nombre != ''
                ORDER BY nombre
                """,
                titular_params + titular_tipo_params,
            ).fetchall()
    
            admin_params = [ejercicio]
            admin_clause = ""
            if ente_uid and ente_aliases:
                placeholders = ", ".join(["?"] * len(ente_aliases))
                admin_clause = (
                    f"AND (TRIM(COALESCE(ente_uid, '')) = ? OR TRIM(COALESCE(ente, '')) IN ({placeholders}))"
                )
                admin_params.extend([ente_uid, *ente_aliases])
            elif ente_uid:
                admin_clause = "AND TRIM(COALESCE(ente_uid, '')) = ?"
                admin_params.append(ente_uid)
            elif ente_aliases:
                placeholders = ", ".join(["?"] * len(ente_aliases))
                admin_clause = f"AND TRIM(COALESCE(ente, '')) IN ({placeholders})"
                admin_params.extend(ente_aliases)
    
            admin_periodos_params = admin_params.copy()
            periodo_admin_clause = ""
            if administrativo_seleccionado:
                periodo_admin_clause = "AND nombre = ?"
                admin_periodos_params.append(administrativo_seleccionado)
    
            admin_tipo_clause = ""
            admin_tipo_params = []
            if tipo_auditoria:
                admin_tipo_clause = "AND tipo_auditoria = ?"
                admin_tipo_params.append(tipo_auditoria)
    
            periodos_admin = db.execute(
                f"""
                SELECT DISTINCT {periodo_sql("historial_titulares")} AS periodo
                FROM historial_titulares
                WHERE ejercicio = ? {admin_clause} {periodo_admin_clause}
                  {admin_tipo_clause}
                  AND tipo_registro = 'director_administrativo'
                  AND fecha_inicio IS NOT NULL AND fecha_fin IS NOT NULL
                ORDER BY fecha_inicio
                """,
                admin_periodos_params + admin_tipo_params,
            ).fetchall()
            administrativos = db.execute(
                f"""
                SELECT DISTINCT nombre
                FROM historial_titulares
                WHERE ejercicio = ? {admin_clause}
                  {admin_tipo_clause}
                  AND tipo_registro = 'director_administrativo'
                  AND nombre IS NOT NULL AND nombre != ''
                ORDER BY nombre
                """,
                admin_params + admin_tipo_params,
            ).fetchall()
        filtros["tipo_anexo"] = [row[0] for row in tipos]
        filtros["tipo_auditoria"] = [row[0] for row in auditorias]
        filtros["entes"] = [dict(row) for row in entes]
        filtros["estado"] = [row[0] for row in estados]
        filtros["fuente_financiamiento"] = [row[0] for row in fuentes]
        filtros["ramo_33"] = [row[0] for row in ramos]
        filtros["conceptos_irregularidad"] = [row[0] for row in conceptos]
        filtros["periodo_informe"] = [row[0] for row in periodos_informe]
        filtros["titulares"] = [row[0] for row in titulares]
        filtros["periodo_admin"] = [row[0] for row in periodos_admin]
        filtros["administrativos"] = [row[0] for row in administrativos]
        filtros["cedulas"] = [row[0] for row in cedulas]
    
        return jsonify(filtros)
    
    
    @app.get("/observaciones-responsables")
    @luis_required
    def observaciones_responsables():
        ejercicio = request.args.get("ejercicio", "").strip()
        ente_id = normalize_ente_id(request.args.get("ente_id", ""))
        tipo_auditoria = normalize_tipo_auditoria(request.args.get("tipo_auditoria", ""))
        estado = request.args.get("estado", "").strip()
        fuente = request.args.get("fuente_financiamiento", "").strip()
        ramo_33 = request.args.get("ramo_33", "").strip()
        periodo_cedula = request.args.get("periodo_cedula", "").strip()
        if not ejercicio or not periodo_cedula:
            return jsonify([])
    
        db = get_db()
        filter_clauses = ["o.ejercicio = ?"]
        params = [ejercicio]
        if ente_id:
            filter_clauses.append(f"{normalize_ente_id_sql('o.ente_id')} = ?")
            params.append(ente_id)
        if tipo_auditoria:
            filter_clauses.append("o.tipo_auditoria = ?")
            params.append(tipo_auditoria)
        if estado:
            filter_clauses.append("o.estado = ?")
            params.append(estado)
        if fuente:
            filter_clauses.append("o.fuente_financiamiento = ?")
            params.append(fuente)
        if ramo_33:
            filter_clauses.append("o.ramo_33 = ?")
            params.append(ramo_33)
        if periodo_cedula:
            filter_clauses.append("o.periodo_cedula = ?")
            params.append(periodo_cedula)
    
        where_sql = " AND ".join(filter_clauses)
        observaciones_rows = db.execute(
            f"""
            SELECT DISTINCT
                o.ejercicio,
                o.ente_id,
                o.ente_nombre,
                ed.ente_nombre AS ente_detalle_nombre,
                ed.ente_uid AS ente_uid,
                o.tipo_auditoria,
                o.periodo_cedula
            FROM observaciones AS o
            LEFT JOIN entes_detalle AS ed
                ON {normalize_ente_id_sql("o.ente_id")} = {normalize_ente_id_sql("ed.ente_id")}
                AND o.ejercicio = ed.ejercicio
            WHERE {where_sql}
            ORDER BY o.ente_nombre ASC, o.tipo_auditoria ASC
            """,
            params,
        ).fetchall()
    
        resultado = []
        for row in observaciones_rows:
            cedula_inicio, cedula_fin = parse_periodo_cedula(row["ejercicio"], row["periodo_cedula"])
            if not cedula_inicio or not cedula_fin:
                continue
            cedula_inicio_date = parse_historial_date(cedula_inicio)
            cedula_fin_date = parse_historial_date(cedula_fin)
            if not cedula_inicio_date or not cedula_fin_date:
                continue
    
            ente_id_norm = normalize_ente_id(row["ente_id"])
            nombres_ente = get_ente_aliases_by_uid(
                db,
                row["ejercicio"],
                ente_id_norm,
                fallback_names=[row["ente_nombre"], row["ente_detalle_nombre"]],
            )
            ente_uid = (
                get_ente_uid_by_ejercicio_id(db, row["ejercicio"], ente_id_norm)
                or (row["ente_uid"] or "").strip()
            )
            scope_clause = ""
            scope_params = []
            if ente_uid and nombres_ente:
                placeholders = ", ".join(["?"] * len(nombres_ente))
                scope_clause = (
                    f"AND (TRIM(COALESCE(h.ente_uid, '')) = ? OR TRIM(COALESCE(h.ente, '')) IN ({placeholders}))"
                )
                scope_params.extend([ente_uid, *nombres_ente])
            elif ente_uid:
                scope_clause = "AND TRIM(COALESCE(h.ente_uid, '')) = ?"
                scope_params.append(ente_uid)
            elif nombres_ente:
                placeholders = ", ".join(["?"] * len(nombres_ente))
                scope_clause = f"AND TRIM(COALESCE(h.ente, '')) IN ({placeholders})"
                scope_params.extend(nombres_ente)
            else:
                continue
    
            historial_rows = db.execute(
                f"""
                SELECT
                    h.tipo_registro,
                    h.tipo_auditoria,
                    h.nombre,
                    h.fecha_inicio,
                    h.fecha_fin,
                    {periodo_sql("h")} AS periodo
                FROM historial_titulares AS h
                WHERE h.ejercicio = ?
                  {scope_clause}
                  AND h.tipo_registro IN ('titular', 'director_administrativo')
                  AND h.nombre IS NOT NULL AND h.nombre != ''
                ORDER BY h.tipo_registro ASC, h.fecha_inicio ASC, h.nombre ASC
                """,
                [row["ejercicio"], *scope_params],
            ).fetchall()
    
            titulares = []
            administrativos = []
            titulares_seen = set()
            administrativos_seen = set()
            for item in historial_rows:
                if normalize_tipo_auditoria(item["tipo_auditoria"] or "") != normalize_tipo_auditoria(row["tipo_auditoria"] or ""):
                    continue
                inicio = parse_historial_date(item["fecha_inicio"])
                fin = parse_historial_date(item["fecha_fin"])
                if not inicio or not fin:
                    continue
                # Inclusive overlap between [inicio, fin] and cedula range
                if inicio > cedula_fin_date or fin < cedula_inicio_date:
                    continue
                payload = {
                    "nombre": item["nombre"],
                    "periodo": item["periodo"],
                }
                key = (payload["nombre"], payload["periodo"])
                if item["tipo_registro"] == "titular":
                    if key in titulares_seen:
                        continue
                    titulares_seen.add(key)
                    titulares.append(payload)
                elif item["tipo_registro"] == "director_administrativo":
                    if key in administrativos_seen:
                        continue
                    administrativos_seen.add(key)
                    administrativos.append(payload)
    
            resultado.append(
                {
                    "ejercicio": row["ejercicio"],
                    "ente_id": row["ente_id"],
                    "ente_nombre": row["ente_nombre"] or row["ente_detalle_nombre"] or "—",
                    "tipo_auditoria": row["tipo_auditoria"],
                    "periodo_cedula": row["periodo_cedula"],
                    "titulares": titulares,
                    "administrativos": administrativos,
                }
            )
    
        return jsonify(resultado)
    
    
    @app.get("/observaciones-exportar")
    @luis_required
    def observaciones_exportar():
        ejercicio = request.args.get("ejercicio", "").strip()
        ente_id = normalize_ente_id(request.args.get("ente_id", ""))
        tipo_anexo = request.args.get("tipo_anexo", "").strip()
        tipo_auditoria = normalize_tipo_auditoria(request.args.get("tipo_auditoria", ""))
        estado = request.args.get("estado", "").strip()
        fuente = request.args.get("fuente_financiamiento", "").strip()
        ramo_33 = request.args.get("ramo_33", "").strip()
        concepto_irregularidad = request.args.get("concepto_irregularidad", "").strip()
        periodo_cedula = request.args.get("periodo_cedula", "").strip()
        if not ejercicio:
            return jsonify({"error": "ejercicio requerido"}), 400
    
        db = get_db()
        params = [ejercicio]
        filter_clauses = []
        if ente_id:
            filter_clauses.append(f"{normalize_ente_id_sql('observaciones.ente_id')} = ?")
            params.append(ente_id)
        if tipo_anexo:
            filter_clauses.append("observaciones.tipo_anexo = ?")
            params.append(tipo_anexo)
        if tipo_auditoria:
            filter_clauses.append("observaciones.tipo_auditoria = ?")
            params.append(tipo_auditoria)
        if estado:
            filter_clauses.append("observaciones.estado = ?")
            params.append(estado)
        if fuente:
            filter_clauses.append("observaciones.fuente_financiamiento = ?")
            params.append(fuente)
        if ramo_33:
            filter_clauses.append("observaciones.ramo_33 = ?")
            params.append(ramo_33)
        if concepto_irregularidad:
            filter_clauses.append(
                "(observaciones.pdp_concepto_irregularidad = ? OR observaciones.pdp_subconcepto_irregularidad = ?)"
            )
            params.extend([concepto_irregularidad, concepto_irregularidad])
        if periodo_cedula:
            filter_clauses.append("observaciones.periodo_cedula = ?")
            params.append(periodo_cedula)
    
        filter_sql = ""
        if filter_clauses:
            filter_sql = " AND " + " AND ".join(filter_clauses)
    
        rows = db.execute(
            f"""
            SELECT
                observaciones.ente_nombre,
                observaciones.tipo_anexo,
                observaciones.numero_observacion,
                observaciones.estado,
                observaciones.fecha_notificacion,
                observaciones.fuente_financiamiento,
                observaciones.pdp_concepto_irregularidad,
                observaciones.monto_pdp_emitido,
                observaciones.monto_pdp_solventado,
                observaciones.monto_pdp_pendiente,
                observaciones.tipo_auditoria
            FROM observaciones
            LEFT JOIN entes_detalle
                ON {normalize_ente_id_sql("observaciones.ente_id")} = {normalize_ente_id_sql("entes_detalle.ente_id")}
                AND observaciones.ejercicio = entes_detalle.ejercicio
            WHERE observaciones.ejercicio = ?
            {filter_sql}
            ORDER BY
                CAST(entes_detalle.ente_numero AS REAL) ASC,
                entes_detalle.ente_numero ASC,
                observaciones.ente_id ASC,
                observaciones.tipo_anexo ASC,
                observaciones.numero_observacion ASC
            """,
            params,
        ).fetchall()
    
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Observaciones"
    
        headers = [
            "Ente",
            "Tipo auditoria",
            "Anexo",
            "No. Obs",
            "Estado",
            "Fecha",
            "Fuente",
            "Concepto de Irregularidad",
            "Monto emitido",
            "Monto solventado",
            "Monto pendiente",
        ]
        sheet.append(headers)
        for cell in sheet[1]:
            cell.font = Font(bold=True)
    
        total_observaciones = 0
        total_emitido = 0.0
        total_solventado = 0.0
        total_pendiente = 0.0
        conteo_pdp = 0
        conteo_pendiente = 0
        conteo_solventado = 0
    
        for row in rows:
            monto_emitido = float(row["monto_pdp_emitido"] or 0)
            monto_solventado = float(row["monto_pdp_solventado"] or 0)
            monto_pendiente = float(row["monto_pdp_pendiente"] or 0)
            total_observaciones += 1
            total_emitido += monto_emitido
            total_solventado += monto_solventado
            total_pendiente += monto_pendiente
            if (row["tipo_anexo"] or "") == "PDP":
                conteo_pdp += 1
            if (row["estado"] or "").strip().lower() == "pendiente":
                conteo_pendiente += 1
            if (row["estado"] or "").strip().lower() == "solventado":
                conteo_solventado += 1
    
            sheet.append(
                [
                    row["ente_nombre"] or "—",
                    row["tipo_auditoria"] or "—",
                    row["tipo_anexo"] or "—",
                    row["numero_observacion"] if row["numero_observacion"] is not None else "—",
                    row["estado"] or "—",
                    row["fecha_notificacion"] or "—",
                    row["fuente_financiamiento"] or "—",
                    row["pdp_concepto_irregularidad"] or "—",
                    monto_emitido if (row["tipo_anexo"] or "") == "PDP" else 0,
                    monto_solventado if (row["tipo_anexo"] or "") == "PDP" else 0,
                    monto_pendiente if (row["tipo_anexo"] or "") == "PDP" else 0,
                ]
            )
    
        last_data_row = sheet.max_row
        for row_idx in range(2, last_data_row + 1):
            for col in (9, 10, 11):
                sheet.cell(row=row_idx, column=col).number_format = "#,##0.00"
    
        summary_start = sheet.max_row + 2
        sheet.cell(row=summary_start, column=1, value="Subtotal / Resumen").font = Font(bold=True)
        summary_rows = [
            ("Total observaciones", total_observaciones),
            ("Observaciones PDP", conteo_pdp),
            ("Observaciones pendientes", conteo_pendiente),
            ("Observaciones solventadas", conteo_solventado),
            ("Monto total emitido", total_emitido),
            ("Monto total solventado", total_solventado),
            ("Monto total pendiente", total_pendiente),
        ]
        for offset, (label, value) in enumerate(summary_rows, start=1):
            current_row = summary_start + offset
            sheet.cell(row=current_row, column=1, value=label).font = Font(bold=True)
            sheet.cell(row=current_row, column=2, value=value)
            if "Monto" in label:
                sheet.cell(row=current_row, column=2).number_format = "#,##0.00"
    
        for column_cells in sheet.columns:
            max_len = 0
            for cell in column_cells:
                val = "" if cell.value is None else str(cell.value)
                if len(val) > max_len:
                    max_len = len(val)
            sheet.column_dimensions[column_cells[0].column_letter].width = min(max(12, max_len + 2), 50)
    
        stream = BytesIO()
        workbook.save(stream)
        stream.seek(0)
        filename = f"observaciones_{ejercicio}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(
            stream,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    
    
    @app.get("/observaciones-stats")
    @luis_required
    def observaciones_stats():
        ente_id = normalize_ente_id(request.args.get("ente_id", ""))
        tipo_auditoria = normalize_tipo_auditoria(request.args.get("tipo_auditoria", ""))
        where_clauses = []
        params = []
        if ente_id:
            where_clauses.append(f"{normalize_ente_id_sql('ente_id')} = ?")
            params.append(ente_id)
        if tipo_auditoria:
            where_clauses.append("tipo_auditoria = ?")
            params.append(tipo_auditoria)
    
        where_clause = f"WHERE {' AND '.join(where_clauses)}" if where_clauses else ""
    
        db = get_db()
        totals = db.execute(
            f"""
            SELECT ejercicio, COUNT(*) as total
            FROM observaciones
            {where_clause}
            GROUP BY ejercicio
            ORDER BY ejercicio
            """,
            params,
        ).fetchall()
    
        estados = db.execute(
            f"""
            SELECT ejercicio, estado, COUNT(*) as total
            FROM observaciones
            {where_clause}
            GROUP BY ejercicio, estado
            ORDER BY ejercicio
            """,
            params,
        ).fetchall()
    
        tipos = db.execute(
            f"""
            SELECT ejercicio, tipo_anexo, COUNT(*) as total
            FROM observaciones
            {where_clause}
            GROUP BY ejercicio, tipo_anexo
            ORDER BY ejercicio
            """,
            params,
        ).fetchall()
    
        tipos_estados = db.execute(
            f"""
            SELECT ejercicio, tipo_anexo, estado, COUNT(*) as total
            FROM observaciones
            {where_clause}
            GROUP BY ejercicio, tipo_anexo, estado
            ORDER BY ejercicio
            """,
            params,
        ).fetchall()
    
        fuentes = db.execute(
            f"""
            SELECT ejercicio, fuente_financiamiento, COUNT(*) as total
            FROM observaciones
            {where_clause}
            GROUP BY ejercicio, fuente_financiamiento
            ORDER BY ejercicio
            """,
            params,
        ).fetchall()
    
        pdp_where = "WHERE tipo_anexo = 'PDP'"
        pdp_clauses = []
        pdp_params = []
        if ente_id:
            pdp_clauses.append(f"{normalize_ente_id_sql('ente_id')} = ?")
            pdp_params.append(ente_id)
        if tipo_auditoria:
            pdp_clauses.append("tipo_auditoria = ?")
            pdp_params.append(tipo_auditoria)
        if pdp_clauses:
            pdp_where = f"WHERE {' AND '.join(pdp_clauses)} AND tipo_anexo = 'PDP'"
    
        pdp = db.execute(
            f"""
            SELECT
                ejercicio,
                SUM(COALESCE(monto_pdp_emitido, 0)) as emitido,
                SUM(COALESCE(monto_pdp_solventado, 0)) as solventado,
                SUM(COALESCE(monto_pdp_pendiente, 0)) as pendiente
            FROM observaciones
            {pdp_where}
            GROUP BY ejercicio
            ORDER BY ejercicio
            """,
            pdp_params,
        ).fetchall()
    
        return jsonify(
            {
                "totals": [dict(row) for row in totals],
                "estados": [dict(row) for row in estados],
                "tipos": [dict(row) for row in tipos],
                "tipos_estados": [dict(row) for row in tipos_estados],
                "fuentes": [dict(row) for row in fuentes],
                "pdp": [dict(row) for row in pdp],
            }
        )
    
    
    @app.get("/catalogo-entes")
    @luis_required
    def catalogo_entes():
        ejercicio = request.args.get("ejercicio", "").strip()
        if not ejercicio:
            return jsonify([])
    
        db = get_db()
        rows = db.execute(
            """
            SELECT
                entes_detalle.ente_id,
                entes_detalle.ente_uid,
                entes_detalle.ente_numero,
                entes_detalle.ente_nombre,
                entes_detalle.ejercicio,
                entes_detalle.clasificacion,
                entes_detalle.ramo33,
                entes_detalle.responsable,
                (
                    SELECT ed.ente_nombre
                    FROM entes_detalle AS ed
                    WHERE COALESCE(ed.ente_uid, ed.ente_id)
                      = COALESCE(entes_detalle.ente_uid, entes_detalle.ente_id)
                      AND ed.ejercicio < entes_detalle.ejercicio
                    ORDER BY ed.ejercicio DESC
                    LIMIT 1
                ) AS nombre_anterior,
                (
                    SELECT COUNT(DISTINCT ed.ente_nombre)
                    FROM entes_detalle AS ed
                    WHERE COALESCE(ed.ente_uid, ed.ente_id)
                      = COALESCE(entes_detalle.ente_uid, entes_detalle.ente_id)
                ) AS nombres_distintos
            FROM entes_detalle
            WHERE entes_detalle.ejercicio = ?
            ORDER BY CAST(entes_detalle.ente_numero AS REAL) ASC, entes_detalle.ente_numero ASC
            """,
            (ejercicio,),
        ).fetchall()
    
        return jsonify([dict(row) for row in rows])
    
    
    @app.post("/fuentes")
    @luis_required
    @role_required("editor")
    def fuentes():
        nombre = request.form.get("fuente_nombre", "").strip()
        if not nombre:
            return redirect(url_for("index", notice="fuente_error"))
    
        db = get_db()
        db.execute(
            """
            INSERT OR IGNORE INTO fuentes_financiamiento (nombre, created_at)
            VALUES (?, ?)
            """,
            (nombre, datetime.now().strftime("%Y-%m-%d %H:%M")),
        )
        db.commit()
        return redirect(url_for("index", notice="fuente_saved"))
    
    
    @app.post("/irregularidades")
    @luis_required
    @role_required("editor")
    def irregularidades():
        concepto = request.form.get("irregularidad_concepto", "").strip()
        if not concepto:
            return redirect(url_for("index", notice="irregularidad_error"))
    
        db = get_db()
        db.execute(
            """
            INSERT OR IGNORE INTO catalogo_irregularidades (concepto, created_at)
            VALUES (?, ?)
            """,
            (concepto, datetime.now().strftime("%Y-%m-%d %H:%M")),
        )
        db.commit()
        return redirect(url_for("index", notice="irregularidad_saved"))
    
    
    @app.route("/stats")
    @luis_required
    def stats():
        db = get_db()
        totals = db.execute(
            """
            SELECT ejercicio, COUNT(*) as total
            FROM registros
            GROUP BY ejercicio
            ORDER BY ejercicio
            """
        ).fetchall()
    
        estados = db.execute(
            """
            SELECT ejercicio, estado, COUNT(*) as total
            FROM registros
            GROUP BY ejercicio, estado
            ORDER BY ejercicio
            """
        ).fetchall()
    
        tipos = db.execute(
            """
            SELECT ejercicio, tipo_anexo, COUNT(*) as total
            FROM registros
            GROUP BY ejercicio, tipo_anexo
            ORDER BY ejercicio
            """
        ).fetchall()
    
        return jsonify(
            {
                "totals": [dict(row) for row in totals],
                "estados": [dict(row) for row in estados],
                "tipos": [dict(row) for row in tipos],
            }
        )
    
    
    @app.post("/reclasificar/<int:registro_id>")
    @luis_required
    @role_required("editor")
    def reclasificar(registro_id: int):
        db = get_db()
        row = db.execute(
            """
            SELECT id, tipo_anexo, tipo_anexo_origen
            FROM registros
            WHERE id = ?
            """,
            (registro_id,),
        ).fetchone()
    
        if row is None:
            return redirect(url_for("index", saved="0"))
    
        if row["tipo_anexo"] not in {"PDP", "PRAS"}:
            return redirect(url_for("index"))
    
        nuevo_tipo = "PRAS" if row["tipo_anexo"] == "PDP" else "PDP"
        origen = row["tipo_anexo_origen"] or row["tipo_anexo"]
        db.execute(
            """
            UPDATE registros
            SET tipo_anexo = ?, tipo_anexo_origen = ?
            WHERE id = ?
            """,
            (nuevo_tipo, origen, registro_id),
        )
        db.commit()
        return redirect(url_for("index", saved="1"))
    
    
    @app.get("/")
    @luis_required
    def index():
        user = get_current_user()
        can_edit = user["role"] == "editor" if user else False
        return render_template(
            "index.html",
            user=user,
            can_edit=can_edit,
        )
    
    
