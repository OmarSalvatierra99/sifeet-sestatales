from datetime import datetime
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

    def _tipo_auditoria_options(tipo_auditoria: str) -> list[str]:
        clean = " ".join((tipo_auditoria or "").split())
        if clean == "Financiera y Obra Pública":
            return ["Financiera", "Obra Pública"]
        if clean in {"Financiera", "Obra Pública"}:
            return [clean]
        return []

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
                if tipo_anexo == "PDP":
                    monto_emitido = pdp_amounts[pdp_index] if pdp_index < len(pdp_amounts) else 0.0
                    pdp_index += 1
                db.execute(
                    """
                    INSERT INTO observaciones (
                        ejercicio,
                        ente_id,
                        ente_numero,
                        ente_numero_sort,
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
                        created_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        ejercicio,
                        ente_id,
                        ente_numero,
                        parse_ente_numero_sort(ente_numero),
                        ente_nombre,
                        tipo_auditoria,
                        fuente_nombre,
                        ramo_33,
                        periodo_cedula,
                        periodo_titular or periodo_cedula,
                        oficio,
                        fecha_notificacion,
                        tipo_anexo,
                        numero_observacion,
                        estado,
                        monto_emitido,
                        monto_pdp_solventado if tipo_anexo == "PDP" else None,
                        monto_pdp_pendiente if tipo_anexo == "PDP" else None,
                        now,
                    ),
                )

    @app.get("/carga/entes")
    @gabo_required
    def carga_entes_por_ejercicio():
        ejercicio = (request.args.get("ejercicio") or "").strip()
        if not ejercicio or ejercicio != TITULAR_EJERCICIO_FIJO:
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
        if not ejercicio or ejercicio != TITULAR_EJERCICIO_FIJO or not ente_id:
            return jsonify([])

        db = get_db()
        rows = fuentes_por_ente(db, ejercicio, ente_id, tipo_auditoria=tipo_auditoria)
        return jsonify([dict(row) for row in rows])

    @app.route("/carga", methods=["GET", "POST"])
    @gabo_required
    def carga():
        user = get_current_user()
        db = get_db()
        fuentes_rows = db.execute(
            """
            SELECT id, nombre
            FROM fuentes_financiamiento
            ORDER BY nombre ASC
            """
        ).fetchall()
        fuentes = [dict(row) for row in fuentes_rows]
    
        titular_ejercicios = [TITULAR_EJERCICIO_FIJO]

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
            "manual_tipo_auditoria": "Financiera",
            "manual_tipo_responsable": "Titular",
            "manual_titular_nombre": "",
            "manual_administrativo_nombre": "",
            "manual_numero_oficio": "",
            "manual_asunto": "Notificación de Cédula de Resultados",
            "manual_ejercicio": TITULAR_EJERCICIO_FIJO,
            "manual_fuente_id": "",
            "manual_fuente_nueva": "",
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
            "titular_ejercicio": TITULAR_EJERCICIO_FIJO,
            "titular_ente_id": "",
            "titular_tipo_auditoria": "Financiera",
            "titular_periodo_informe": "",
            "titular_nombre": "",
            "titular_periodo_administrativo": "",
            "titular_administrativo": "",
            "titular_cedula_resultados": "",
        }
    
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
                    "manual_tipo_auditoria": (request.form.get("manual_tipo_auditoria") or "").strip() or "Financiera",
                    "manual_tipo_responsable": (request.form.get("manual_tipo_responsable") or "").strip() or "Titular",
                    "manual_titular_nombre": (request.form.get("manual_titular_nombre") or "").strip(),
                    "manual_administrativo_nombre": (request.form.get("manual_administrativo_nombre") or "").strip(),
                    "manual_numero_oficio": (request.form.get("manual_numero_oficio") or "").strip(),
                    "manual_asunto": (request.form.get("manual_asunto") or "").strip() or "Notificación de Cédula de Resultados",
                    "manual_ejercicio": (request.form.get("manual_ejercicio") or "").strip() or TITULAR_EJERCICIO_FIJO,
                    "manual_fuente_id": (request.form.get("manual_fuente_id") or "").strip(),
                    "manual_fuente_nueva": (request.form.get("manual_fuente_nueva") or "").strip(),
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
                    "titular_ejercicio": TITULAR_EJERCICIO_FIJO,
                    "titular_ente_id": (request.form.get("titular_ente_id") or "").strip(),
                    "titular_tipo_auditoria": (request.form.get("titular_tipo_auditoria") or "").strip() or "Financiera",
                    "titular_periodo_informe": (request.form.get("titular_periodo_informe") or "").strip(),
                    "titular_nombre": (request.form.get("titular_nombre") or "").strip(),
                    "titular_periodo_administrativo": (request.form.get("titular_periodo_administrativo") or "").strip(),
                    "titular_administrativo": (request.form.get("titular_administrativo") or "").strip(),
                    "titular_cedula_resultados": (request.form.get("titular_cedula_resultados") or "").strip(),
                }
            )
    
            try:
                if action == "titular_save":
                    titular_ejercicio = form_data["titular_ejercicio"]
                    titular_ente_id = form_data["titular_ente_id"]
                    titular_tipo_auditoria = form_data["titular_tipo_auditoria"]
                    titular_periodo_informe = " ".join(form_data["titular_periodo_informe"].split())
                    titular_nombre = form_data["titular_nombre"]
                    titular_periodo_administrativo = " ".join(form_data["titular_periodo_administrativo"].split())
                    titular_administrativo = form_data["titular_administrativo"]
                    titular_periodo_informe_key = normalize_periodo_key(
                        titular_ejercicio,
                        titular_periodo_informe,
                        label="periodo informe",
                    )
                    titular_periodo_administrativo_key = normalize_periodo_key(
                        titular_ejercicio,
                        titular_periodo_administrativo,
                        label="periodo administrativo",
                    )
                    titular_cedula_periodos, titular_cedula_keys = parse_cedula_periodos(
                        titular_ejercicio,
                        form_data["titular_cedula_resultados"],
                    )
                    titular_cedula_resultados = " | ".join(titular_cedula_periodos)

                    if not titular_ejercicio:
                        raise ValueError("Titulares: ejercicio requerido.")
                    if titular_ejercicio != TITULAR_EJERCICIO_FIJO:
                        raise ValueError("Titulares: solo se permite ejercicio 2025.")
                    if not titular_ente_id:
                        raise ValueError("Titulares: selecciona un ente.")
                    if titular_tipo_auditoria not in {"Financiera", "Obra Pública"}:
                        raise ValueError("Titulares: tipo de auditoría inválido.")
                    if not titular_periodo_informe:
                        raise ValueError("Titulares: periodo informe requerido.")
                    if not titular_nombre:
                        raise ValueError("Titulares: nombre del titular requerido.")
                    if not titular_periodo_administrativo:
                        raise ValueError("Titulares: periodo administrativo requerido.")
                    if not titular_administrativo:
                        raise ValueError("Titulares: nombre administrativo requerido.")
                    if not titular_cedula_periodos:
                        raise ValueError(
                            "Titulares: cédula de resultados requerida. "
                            "Puedes capturar varias particiones."
                        )

                    ente_row = db.execute(
                        """
                        SELECT ente_nombre
                        FROM entes_detalle
                        WHERE ejercicio = ? AND ente_id = ?
                        LIMIT 1
                        """,
                        (titular_ejercicio, titular_ente_id),
                    ).fetchone()
                    if not ente_row:
                        raise ValueError("Titulares: el ente seleccionado no existe para ese ejercicio.")

                    posibles_duplicados = db.execute(
                        """
                        SELECT
                            id,
                            created_at,
                            periodo_informe,
                            periodo_administrativo,
                            cedula_resultados
                        FROM cargas_titulares
                        WHERE ejercicio = ?
                          AND ente_id = ?
                          AND tipo_auditoria = ?
                        ORDER BY id DESC
                        LIMIT 200
                        """,
                        (
                            titular_ejercicio,
                            titular_ente_id,
                            titular_tipo_auditoria,
                        ),
                    ).fetchall()
                    duplicado = None
                    titular_cedula_key = "|".join(titular_cedula_keys)
                    for row in posibles_duplicados:
                        row_periodo_informe_key = normalize_periodo_key(
                            titular_ejercicio,
                            row["periodo_informe"],
                            label="periodo informe",
                            strict=False,
                        )
                        row_periodo_admin_key = normalize_periodo_key(
                            titular_ejercicio,
                            row["periodo_administrativo"],
                            label="periodo administrativo",
                            strict=False,
                        )
                        _, row_cedula_keys = parse_cedula_periodos(
                            titular_ejercicio,
                            row["cedula_resultados"],
                            strict=False,
                        )
                        row_cedula_key = "|".join(row_cedula_keys)
                        if (
                            row_periodo_informe_key == titular_periodo_informe_key
                            and row_periodo_admin_key == titular_periodo_administrativo_key
                            and row_cedula_key == titular_cedula_key
                        ):
                            duplicado = row
                            break
                    if duplicado:
                        titular_result = {
                            "ok": False,
                            "level": "info",
                            "message": (
                                f"Titulares: ya existe un registro con esos periodos "
                                f"(ID {duplicado['id']}, fecha {duplicado['created_at']})."
                            ),
                        }
                    else:
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
                                (ente_row["ente_nombre"] or "").strip(),
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
                        titular_result = {
                            "ok": True,
                            "level": "success",
                            "message": "Titulares: registro guardado correctamente.",
                        }
                elif action in {"manual_check", "manual_save"}:
                    manual_id_raw = form_data["manual_id"]
                    manual_ente_id = form_data["manual_ente_id"]
                    tipo_auditoria = form_data["manual_tipo_auditoria"]
                    tipo_responsable = "Titular"
                    titular_nombre = ""
                    administrativo_nombre = ""
                    numero_oficio = form_data["manual_numero_oficio"]
                    asunto = form_data["manual_asunto"]
                    ejercicio = form_data["manual_ejercicio"]
                    fuente_id_raw = form_data["manual_fuente_id"]
                    fuente_nueva = " ".join(form_data["manual_fuente_nueva"].split())
                    periodo = form_data["manual_periodo"]
                    periodo_titular = form_data["manual_periodo_titular"]
                    ramo_33 = (form_data["manual_ramo_33"] or "No").strip()
                    estado = (form_data["manual_estado"] or "E").strip().upper()
                    fecha_notificacion = form_data["manual_fecha_notificacion"]
                    raw_montos_pdp = form_data["manual_montos_pdp"]
                    manual_edit_id = None
                    if manual_id_raw:
                        try:
                            manual_edit_id = int(manual_id_raw)
                        except ValueError as exc:
                            raise ValueError("ID de edición manual inválido.") from exc

                    if not manual_ente_id:
                        raise ValueError("Debes seleccionar un ente.")
                    if not tipo_auditoria:
                        raise ValueError("Debes seleccionar el tipo de auditoría.")
                    if tipo_auditoria not in {"Financiera", "Financiera y Obra Pública"}:
                        raise ValueError("Tipo de auditoría inválido.")
                    if not numero_oficio:
                        raise ValueError("Debes capturar el número de oficio.")
                    if asunto not in ASUNTOS_MANUALES:
                        raise ValueError("Debes seleccionar un asunto válido.")
                    if not ejercicio:
                        raise ValueError("Debes capturar el ejercicio.")
                    if ejercicio != TITULAR_EJERCICIO_FIJO:
                        raise ValueError("Solo se permite ejercicio 2025.")
                    if not fuente_id_raw and not fuente_nueva:
                        raise ValueError("Debes seleccionar una fuente o capturar una nueva.")
                    if not periodo:
                        raise ValueError("Debes capturar el periodo.")
                    if ramo_33 not in {"Si", "No"}:
                        raise ValueError("RAMO XXXIII debe ser 'Si' o 'No'.")
                    if estado not in {"E", "R"}:
                        raise ValueError("E/R debe ser 'E' o 'R'.")
                    if not fecha_notificacion:
                        raise ValueError("Debes capturar la fecha de notificación.")
                    try:
                        int(ejercicio)
                    except ValueError as exc:
                        raise ValueError("Ejercicio inválido.") from exc
                    fuente_id = None
                    fuente_nombre = ""
                    if fuente_nueva:
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
                    if asunto == "Notificación de Cédula de Resultados":
                        cantidad_sa = parse_non_negative_int(form_data["manual_cantidad_sa"], "Cantidad SA")
                        cantidad_pdp = parse_non_negative_int(form_data["manual_cantidad_pdp"], "Cantidad PDP")
                        cantidad_pras = parse_non_negative_int(form_data["manual_cantidad_pras"], "Cantidad PRAS")
                        cantidad_pefcf = parse_non_negative_int(form_data["manual_cantidad_pefcf"], "Cantidad PEFCF")
                        cantidad_r = parse_non_negative_int(form_data["manual_cantidad_r"], "Cantidad R")
                    else:
                        cantidad_sa = 0
                        cantidad_pdp = 0
                        cantidad_pras = 0
                        cantidad_pefcf = 0
                        cantidad_r = 0
                    monto_pdp_emitido = parse_non_negative_float(form_data["manual_monto_pdp_emitido"], "Monto PDP emitido")
                    monto_pdp_solventado = parse_non_negative_float(form_data["manual_monto_pdp_solventado"], "Monto PDP solventado")
                    monto_pdp_pendiente = parse_non_negative_float(form_data["manual_monto_pdp_pendiente"], "Monto PDP pendiente")
                    pdp_amounts = parse_pdp_amounts(raw_montos_pdp)
                    if cantidad_pdp == 0 and pdp_amounts:
                        raise ValueError("Capturaste montos PDP pero la cantidad PDP es 0.")
                    if pdp_amounts and len(pdp_amounts) != cantidad_pdp:
                        raise ValueError(
                            "La cantidad de montos PDP no coincide con 'Cantidad PDP'. "
                            "Captura un monto por línea."
                        )
                    if cantidad_pdp > 1 and monto_pdp_emitido > 0 and not pdp_amounts:
                        raise ValueError(
                            "Para PDP con más de una observación, captura el detalle de montos por línea."
                        )
                    if not pdp_amounts and cantidad_pdp == 1:
                        pdp_amounts = [monto_pdp_emitido]

                    tipos_auditoria = (
                        ["Financiera", "Obra Pública"]
                        if tipo_auditoria == "Financiera y Obra Pública"
                        else [tipo_auditoria]
                    )
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
                                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                                        pdp_amounts=pdp_amounts,
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
        manual_fuentes_rows = []
        if manual_ente_id_norm:
            manual_fuentes = fuentes_por_ente(
                db,
                TITULAR_EJERCICIO_FIJO,
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
            form_data=form_data,
            fuentes=fuentes,
            manual_fuentes=manual_fuentes,
            titular_ejercicios=titular_ejercicios,
            titular_entes=[dict(row) for row in titular_entes_rows],
            manual_entes=[dict(row) for row in titular_entes_rows],
            asuntos=[
                "Notificación de Cédula de Resultados",
                "Se emiten resultados de solventación del periodo",
            ],
            tipos_responsable=["Titular", "Administrativo", "Ambos"],
        )
    
