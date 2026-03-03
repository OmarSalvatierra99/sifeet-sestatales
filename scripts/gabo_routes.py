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
    SOLVENTACION_TIPOS = ("PDP", "PEFCF", "PRAS", "R", "SA")

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

    def parse_manual_solventacion_details(raw_value: str) -> dict[str, dict]:
        base = {
            tipo: {
                "cantidad": 0,
                "emitido": 0.0,
                "solventado": 0.0,
                "pendiente": 0.0,
            }
            for tipo in SOLVENTACION_TIPOS
        }
        raw = (raw_value or "").strip()
        if not raw:
            return base
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise ValueError("El detalle de solventación tiene un formato inválido.") from exc
        if not isinstance(payload, dict):
            raise ValueError("El detalle de solventación tiene un formato inválido.")

        for raw_tipo, raw_data in payload.items():
            tipo = str(raw_tipo or "").strip().upper()
            if tipo == "PEFCT":
                tipo = "PEFCF"
            if tipo not in base:
                continue
            data = raw_data if isinstance(raw_data, dict) else {}
            cantidad = parse_non_negative_int(str(data.get("cantidad", "0")), f"Cantidad {tipo}")
            if cantidad <= 0:
                continue
            emitido_raw = str(data.get("emitido", "")).strip()
            solventado_raw = str(data.get("solventado", "")).strip()
            if not emitido_raw:
                raise ValueError(f"Debes capturar monto emitido para {tipo}.")
            if not solventado_raw:
                raise ValueError(f"Debes capturar monto solventado para {tipo}.")
            emitido = parse_non_negative_float(emitido_raw, f"Monto emitido {tipo}")
            solventado = parse_non_negative_float(solventado_raw, f"Monto solventado {tipo}")
            if solventado > emitido:
                raise ValueError(f"En {tipo}, monto solventado no puede ser mayor a emitido.")
            base[tipo] = {
                "cantidad": cantidad,
                "emitido": emitido,
                "solventado": solventado,
                "pendiente": emitido - solventado,
            }
        return base

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
            payload.append(item)
        return jsonify(payload)

    @app.post("/carga/observaciones-cargadas/<int:observacion_id>/actualizar")
    @gabo_required
    def carga_observacion_actualizar(observacion_id: int):
        payload = request.get_json(silent=True) or {}
        estado = _normalize_observacion_estado(payload.get("estado", ""))
        monto_emitido_raw = payload.get("monto_pdp_emitido", "")
        monto_solventado_raw = payload.get("monto_pdp_solventado", "")

        if estado not in OBSERVACION_ESTADOS_VALIDOS:
            return jsonify({"ok": False, "error": "Estado inválido."}), 400

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

    @app.route("/carga", methods=["GET", "POST"])
    @gabo_required
    def carga():
        user = get_current_user()
        db = get_db()
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
            "manual_tipo_auditoria": "",
            "manual_tipo_responsable": "Titular",
            "manual_titular_nombre": "",
            "manual_administrativo_nombre": "",
            "manual_numero_oficio": "",
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
            "manual_solventacion_detalle_json": "",
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
                    "manual_tipo_auditoria": (request.form.get("manual_tipo_auditoria") or "").strip(),
                    "manual_tipo_responsable": (request.form.get("manual_tipo_responsable") or "").strip() or "Titular",
                    "manual_titular_nombre": (request.form.get("manual_titular_nombre") or "").strip(),
                    "manual_administrativo_nombre": (request.form.get("manual_administrativo_nombre") or "").strip(),
                    "manual_numero_oficio": (request.form.get("manual_numero_oficio") or "").strip(),
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
                    "manual_solventacion_detalle_json": (request.form.get("manual_solventacion_detalle_json") or "").strip(),
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
                    fuente_nueva = " ".join(form_data["manual_fuente_nueva"].split()) if fuente_id_raw == "__new__" else ""
                    periodo = form_data["manual_periodo"]
                    periodo_titular = form_data["manual_periodo_titular"]
                    ramo_33 = "No"
                    estado = "E"
                    fecha_notificacion = form_data["manual_fecha_notificacion"]
                    raw_montos_pdp = form_data["manual_montos_pdp"]
                    raw_pdp_detalle_json = form_data["manual_pdp_detalle_json"]
                    raw_solventacion_detalle_json = form_data["manual_solventacion_detalle_json"]
                    fuentes_detalle_rows = parse_manual_fuentes_detalle(form_data["manual_fuentes_detalle_json"])
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
                    if tipo_auditoria not in {"Financiera", "Obra Pública"}:
                        raise ValueError("Tipo de auditoría inválido.")
                    if not numero_oficio:
                        raise ValueError("Debes capturar el número de oficio.")
                    if asunto not in ASUNTOS_MANUALES:
                        raise ValueError("Debes seleccionar un asunto válido.")
                    if not ejercicio:
                        raise ValueError("Debes capturar el ejercicio.")
                    if ejercicio not in manual_ejercicios:
                        raise ValueError("El ejercicio seleccionado no está disponible.")
                    if not periodo:
                        raise ValueError("Debes capturar el periodo.")
                    if not fecha_notificacion:
                        raise ValueError("Debes capturar la fecha de notificación.")
                    try:
                        int(ejercicio)
                    except ValueError as exc:
                        raise ValueError("Ejercicio inválido.") from exc
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
                    else:
                        detalle_solventacion = parse_manual_solventacion_details(raw_solventacion_detalle_json)
                        cantidad_sa = int(detalle_solventacion["SA"]["cantidad"])
                        cantidad_pdp = int(detalle_solventacion["PDP"]["cantidad"])
                        cantidad_pras = int(detalle_solventacion["PRAS"]["cantidad"])
                        cantidad_pefcf = int(detalle_solventacion["PEFCF"]["cantidad"])
                        cantidad_r = int(detalle_solventacion["R"]["cantidad"])
                        monto_pdp_emitido = float(detalle_solventacion["PDP"]["emitido"])
                        monto_pdp_solventado = float(detalle_solventacion["PDP"]["solventado"])
                        monto_pdp_pendiente = float(detalle_solventacion["PDP"]["pendiente"])
                        solventacion_totales_by_anexo = detalle_solventacion
                        pdp_details = []
                        pdp_amounts = [monto_pdp_emitido] + [0.0] * (cantidad_pdp - 1) if cantidad_pdp > 0 else []
                    is_solventacion = asunto == "Se emiten resultados de solventación del periodo"

                    tipos_auditoria = [tipo_auditoria]
                    if is_solventacion:
                        requested_counts = {
                            "SA": cantidad_sa,
                            "PDP": cantidad_pdp,
                            "PRAS": cantidad_pras,
                            "PEFCF": cantidad_pefcf,
                            "R": cantidad_r,
                        }
                        for tipo_item in tipos_auditoria:
                            total_scope, existing_counts = count_existing_observaciones_scope(
                                db,
                                ejercicio=ejercicio,
                                ente_id=manual_ente_id,
                                tipo_auditoria=tipo_item,
                                fuente_nombre=fuente_nombre,
                                periodo=periodo,
                                oficio=numero_oficio,
                            )
                            if total_scope <= 0:
                                raise ValueError(
                                    "Para capturar Oficios de Respuesta a Solventación primero debe existir "
                                    "la Notificación de Cédula de Resultados para la misma clave "
                                    "(ejercicio, ente, tipo, fuente, periodo y oficio)."
                                )
                            for tipo_anexo, requested in requested_counts.items():
                                if requested <= 0:
                                    continue
                                available = int(existing_counts.get(tipo_anexo, 0))
                                if requested > available:
                                    raise ValueError(
                                        f"En solventación, cantidad {tipo_anexo} ({requested}) excede "
                                        f"las observaciones existentes ({available})."
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
                                if not is_solventacion:
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
                                    if not is_solventacion:
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
            form_data=form_data,
            fuentes=fuentes,
            manual_fuentes=manual_fuentes,
            manual_ejercicios=manual_ejercicios,
            titular_ejercicios=titular_ejercicios,
            titular_entes=[dict(row) for row in titular_entes_rows],
            manual_entes=[dict(row) for row in manual_entes_rows],
            asuntos=[
                "Notificación de Cédula de Resultados",
                "Se emiten resultados de solventación del periodo",
            ],
            tipos_responsable=["Titular", "Administrativo", "Ambos"],
        )
    
