from datetime import datetime
import sys

from flask import render_template, request


def register_gabo_routes(app, deps):
    globals().update(deps)
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
    
        script_result = None
        manual_result = None
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
            "manual_tipo_auditoria": "Financiera",
            "manual_tipo_responsable": "Titular",
            "manual_titular_nombre": "",
            "manual_administrativo_nombre": "",
            "manual_numero_oficio": "",
            "manual_asunto": "Notificación de Cédula de Resultados",
            "manual_ejercicio": "",
            "manual_fuente_id": "",
            "manual_periodo": "",
            "manual_cantidad_sa": "0",
            "manual_cantidad_pdp": "0",
            "manual_cantidad_pras": "0",
            "manual_cantidad_pefcf": "0",
            "manual_cantidad_r": "0",
            "manual_monto_pdp_emitido": "0",
            "manual_monto_pdp_solventado": "0",
            "manual_monto_pdp_pendiente": "0",
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
                    "manual_tipo_auditoria": (request.form.get("manual_tipo_auditoria") or "").strip() or "Financiera",
                    "manual_tipo_responsable": (request.form.get("manual_tipo_responsable") or "").strip() or "Titular",
                    "manual_titular_nombre": (request.form.get("manual_titular_nombre") or "").strip(),
                    "manual_administrativo_nombre": (request.form.get("manual_administrativo_nombre") or "").strip(),
                    "manual_numero_oficio": (request.form.get("manual_numero_oficio") or "").strip(),
                    "manual_asunto": (request.form.get("manual_asunto") or "").strip() or "Notificación de Cédula de Resultados",
                    "manual_ejercicio": (request.form.get("manual_ejercicio") or "").strip(),
                    "manual_fuente_id": (request.form.get("manual_fuente_id") or "").strip(),
                    "manual_periodo": (request.form.get("manual_periodo") or "").strip(),
                    "manual_cantidad_sa": (request.form.get("manual_cantidad_sa") or "").strip() or "0",
                    "manual_cantidad_pdp": (request.form.get("manual_cantidad_pdp") or "").strip() or "0",
                    "manual_cantidad_pras": (request.form.get("manual_cantidad_pras") or "").strip() or "0",
                    "manual_cantidad_pefcf": (request.form.get("manual_cantidad_pefcf") or "").strip() or "0",
                    "manual_cantidad_r": (request.form.get("manual_cantidad_r") or "").strip() or "0",
                    "manual_monto_pdp_emitido": (request.form.get("manual_monto_pdp_emitido") or "").strip() or "0",
                    "manual_monto_pdp_solventado": (request.form.get("manual_monto_pdp_solventado") or "").strip() or "0",
                    "manual_monto_pdp_pendiente": (request.form.get("manual_monto_pdp_pendiente") or "").strip() or "0",
                }
            )
    
            try:
                if action == "manual_save":
                    tipo_auditoria = form_data["manual_tipo_auditoria"]
                    tipo_responsable = form_data["manual_tipo_responsable"]
                    titular_nombre = form_data["manual_titular_nombre"]
                    administrativo_nombre = form_data["manual_administrativo_nombre"]
                    numero_oficio = form_data["manual_numero_oficio"]
                    asunto = form_data["manual_asunto"]
                    ejercicio = form_data["manual_ejercicio"]
                    fuente_id_raw = form_data["manual_fuente_id"]
                    periodo = form_data["manual_periodo"]
    
                    if not tipo_auditoria:
                        raise ValueError("Debes seleccionar el tipo de auditoría.")
                    if tipo_responsable not in TIPOS_RESPONSABLE:
                        raise ValueError("Tipo de responsable inválido.")
                    if tipo_responsable in {"Titular", "Ambos"} and not titular_nombre:
                        raise ValueError("Debes capturar el nombre del titular.")
                    if tipo_responsable in {"Administrativo", "Ambos"} and not administrativo_nombre:
                        raise ValueError("Debes capturar el nombre del administrativo.")
                    if not numero_oficio:
                        raise ValueError("Debes capturar el número de oficio.")
                    if asunto not in ASUNTOS_MANUALES:
                        raise ValueError("Debes seleccionar un asunto válido.")
                    if not ejercicio:
                        raise ValueError("Debes capturar el ejercicio.")
                    if not periodo:
                        raise ValueError("Debes capturar el periodo.")
                    try:
                        int(ejercicio)
                    except ValueError as exc:
                        raise ValueError("Ejercicio inválido.") from exc
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
    
                    cantidad_sa = parse_non_negative_int(form_data["manual_cantidad_sa"], "Cantidad SA")
                    cantidad_pdp = parse_non_negative_int(form_data["manual_cantidad_pdp"], "Cantidad PDP")
                    cantidad_pras = parse_non_negative_int(form_data["manual_cantidad_pras"], "Cantidad PRAS")
                    cantidad_pefcf = parse_non_negative_int(form_data["manual_cantidad_pefcf"], "Cantidad PEFCF")
                    cantidad_r = parse_non_negative_int(form_data["manual_cantidad_r"], "Cantidad R")
                    monto_pdp_emitido = parse_non_negative_float(form_data["manual_monto_pdp_emitido"], "Monto PDP emitido")
                    monto_pdp_solventado = parse_non_negative_float(form_data["manual_monto_pdp_solventado"], "Monto PDP solventado")
                    monto_pdp_pendiente = parse_non_negative_float(form_data["manual_monto_pdp_pendiente"], "Monto PDP pendiente")
    
                    db.execute(
                        """
                        INSERT INTO cargas_manuales (
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
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            tipo_auditoria,
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
                    db.commit()
                    manual_result = {
                        "ok": True,
                        "message": "Registro manual guardado correctamente.",
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
    
        manual_rows = db.execute(
            """
            SELECT
                cm.id,
                cm.tipo_auditoria,
                cm.tipo_responsable,
                cm.titular_nombre,
                cm.administrativo_nombre,
                cm.numero_oficio,
                cm.asunto,
                cm.ejercicio,
                cm.periodo,
                cm.cantidad_sa,
                cm.cantidad_pdp,
                cm.cantidad_pras,
                cm.cantidad_pefcf,
                cm.cantidad_r,
                cm.monto_pdp_emitido,
                cm.monto_pdp_solventado,
                cm.monto_pdp_pendiente,
                cm.created_at,
                ff.nombre AS fuente_nombre
            FROM cargas_manuales AS cm
            LEFT JOIN fuentes_financiamiento AS ff ON ff.id = cm.fuente_id
            WHERE cm.created_by = ?
            ORDER BY cm.id DESC
            LIMIT 25
            """,
            (user["username"],),
        ).fetchall()
    
        return render_template(
            "carga.html",
            user=user,
            result=script_result,
            manual_result=manual_result,
            form_data=form_data,
            fuentes=fuentes,
            asuntos=[
                "Notificación de Cédula de Resultados",
                "Se emiten resultados de solventación del periodo",
            ],
            tipos_responsable=["Titular", "Administrativo", "Ambos"],
            manual_rows=[dict(row) for row in manual_rows],
        )
    
