from pathlib import Path
import json
import sqlite3

import pytest

from scripts.parsers import parse_cedula


@pytest.fixture
def client():
    from app import app

    app.config["TESTING"] = True
    with app.test_client() as test_client:
        yield test_client


def test_parse_cedula_detects_obra_convenios_without_duplicate_child_group():
    payload = parse_cedula(Path("examples/1.16.- SI_OFS_0342_2026_Ene-Jun.pdf"))

    auditorias = {item["tipo"]: item for item in payload["auditorias"]}
    obra = auditorias["Obra Pública"]
    convenios_section = auditorias["Convenios"]
    convenios = convenios_section["fuentes"]

    assert auditorias["Financiera"]["totales"] == {
        "SA": 12,
        "PDP": 2,
        "PRAS": 2,
        "PEFCF": 0,
        "R": 8,
        "total_emitidas": 24,
    }
    assert obra["totales"] == {
        "SA": 0,
        "PDP": 79,
        "PRAS": 13,
        "PEFCF": 0,
        "R": 0,
        "total_emitidas": 92,
    }
    assert convenios_section["totales"] == {
        "SA": 1,
        "PDP": 8,
        "PRAS": 1,
        "PEFCF": 0,
        "R": 0,
        "total_emitidas": 10,
    }
    assert len(convenios) == 3
    assert {
        fuente["convenio_ente_nombre"] for fuente in convenios
    } == {
        "TRIBUNAL DE JUSTICIA ADMINISTRATIVA DEL ESTADO DE TLAXCALA",
        "INSTITUTO DE CAPACITACIÓN PARA EL TRABAJO DEL ESTADO DE TLAXCALA",
        "UNIDAD DE SERVICIOS EDUCATIVOS DEL ESTADO DE TLAXCALA",
    }
    assert sum(
        len(registro[tipo])
        for fuente in convenios
        for registro in fuente["registros"]
        for tipo in ("SA", "PDP", "PRAS", "PEFCF", "R")
    ) == 10


def test_parse_cedula_keeps_itife_financial_continuation_before_obra_header():
    payload = parse_cedula(Path("examples/19.- ITIFE_OFS_0328_2026_Ene-Jun.pdf"))

    auditorias = {item["tipo"]: item for item in payload["auditorias"]}
    financiera = auditorias["Financiera"]
    obra = auditorias["Obra Pública"]
    convenios = auditorias["Convenios"]

    assert financiera["totales"] == {
        "SA": 24,
        "PDP": 8,
        "PRAS": 6,
        "PEFCF": 5,
        "R": 11,
        "total_emitidas": 54,
    }
    assert obra["totales"] == {
        "SA": 35,
        "PDP": 66,
        "PRAS": 15,
        "PEFCF": 0,
        "R": 0,
        "total_emitidas": 116,
    }
    assert convenios["totales"] == {
        "SA": 0,
        "PDP": 10,
        "PRAS": 4,
        "PEFCF": 0,
        "R": 0,
        "total_emitidas": 14,
    }


def test_gabo_manual_save_materializes_convenio_scope(client, monkeypatch, tmp_path):
    import app as app_module

    test_db_path = tmp_path / "convenios_save.db"
    monkeypatch.setattr(app_module, "DB_PATH", str(test_db_path))
    app_module.init_db()

    conn = sqlite3.connect(test_db_path)
    try:
        conn.execute(
            """
            INSERT INTO entes_detalle (
                ente_id, ejercicio, ente_numero, ente_nombre,
                responsable, clasificacion, ramo33, ramo28, created_at
            )
            VALUES ('1.16', '2025', '1.16', 'Secretaría de Infraestructura (SI)', '', '', 'No', 'No', 'now')
            """
        )
        conn.execute(
            """
            INSERT INTO entes_detalle (
                ente_id, ejercicio, ente_numero, ente_nombre,
                responsable, clasificacion, ramo33, ramo28, created_at
            )
            VALUES ('33', '2025', '33', 'Tribunal de Justicia Administrativa del Estado de Tlaxcala (TJA)', '', '', 'No', 'No', 'now')
            """
        )
        conn.commit()
    finally:
        conn.close()

    fuente = "RECURSOS FISCALES PROPIOS Y RECURSOS FISCALES ESTATALES"
    detalle = [
        {
            "tipo_auditoria": "Obra Pública",
            "fuente_nombre": fuente,
            "modalidad": "Convenio",
            "convenio_nombre": "CONVENIO DE COLABORACIÓN PARA LA EJECUCIÓN DE OBRA PÚBLICA: TRIBUNAL DE JUSTICIA ADMINISTRATIVA DEL ESTADO DE TLAXCALA",
            "convenio_ente_nombre": "TRIBUNAL DE JUSTICIA ADMINISTRATIVA DEL ESTADO DE TLAXCALA",
            "periodo": "01 de Enero al 15 de Mayo",
            "cantidad_sa": 1,
            "cantidad_pdp": 0,
            "cantidad_pras": 0,
            "cantidad_pefcf": 0,
            "cantidad_r": 0,
        }
    ]

    with client.session_transaction() as session_data:
        session_data["user"] = "gabo"
        session_data["role"] = "loader"

    response = client.post(
        "/carga",
        data={
            "action": "manual_save",
            "manual_ente_id": "1.16",
            "manual_tipo_auditoria": "Obra Pública",
            "manual_numero_oficio": "OFS/TEST/2025",
            "manual_asunto": "Notificación de Cédula de Resultados",
            "manual_ejercicio": "2025",
            "manual_periodo": "01 de Enero al 15 de Mayo",
            "manual_fecha_notificacion": "2025-01-01",
            "manual_fuente_id": "__new__",
            "manual_fuente_nueva": fuente,
            "manual_fuentes_detalle_json": json.dumps(detalle, ensure_ascii=False),
            "manual_pdp_detalle_json": "[]",
        },
    )

    assert response.status_code == 200

    conn = sqlite3.connect(test_db_path)
    conn.row_factory = sqlite3.Row
    try:
        observacion = conn.execute(
            """
            SELECT modalidad, convenio_ente_id, convenio_ente_nombre, COUNT(*) AS total
            FROM observaciones
            GROUP BY modalidad, convenio_ente_id, convenio_ente_nombre
            """
        ).fetchone()
        carga = conn.execute(
            "SELECT modalidad, convenio_ente_id, convenio_ente_nombre FROM cargas_manuales"
        ).fetchone()
    finally:
        conn.close()

    assert dict(observacion) == {
        "modalidad": "Convenio",
        "convenio_ente_id": "33",
        "convenio_ente_nombre": "TRIBUNAL DE JUSTICIA ADMINISTRATIVA DEL ESTADO DE TLAXCALA",
        "total": 1,
    }
    assert dict(carga) == {
        "modalidad": "Convenio",
        "convenio_ente_id": "33",
        "convenio_ente_nombre": "TRIBUNAL DE JUSTICIA ADMINISTRATIVA DEL ESTADO DE TLAXCALA",
    }


def test_gabo_manual_save_replaces_partial_multi_source_oficio(client, monkeypatch, tmp_path):
    import app as app_module

    test_db_path = tmp_path / "replace_partial_save.db"
    monkeypatch.setattr(app_module, "DB_PATH", str(test_db_path))
    app_module.init_db()

    conn = sqlite3.connect(test_db_path)
    try:
        conn.execute(
            """
            INSERT INTO entes_detalle (
                ente_id, ejercicio, ente_numero, ente_nombre,
                responsable, clasificacion, ramo33, ramo28, created_at
            )
            VALUES ('1.3', '2025', '1.3', 'Secretaría de Finanzas (SF)', '', '', 'No', 'No', 'now')
            """
        )
        conn.commit()
    finally:
        conn.close()

    first_row = [
        {
            "tipo_auditoria": "Financiera",
            "fuente_nombre": "PARTICIPACIONES ESTATALES (FONDO GENERAL DE PARTICIPACIONES)",
            "periodo": "01 de Enero al 20 de Marzo",
            "cantidad_sa": 0,
            "cantidad_pdp": 0,
            "cantidad_pras": 0,
            "cantidad_pefcf": 0,
            "cantidad_r": 1,
        }
    ]
    full_rows = first_row + [
        {
            "tipo_auditoria": "Financiera",
            "fuente_nombre": "PARTICIPACIONES ESTATALES (FONDO GENERAL DE PARTICIPACIONES)",
            "periodo": "21 de Marzo al 16 de Septiembre",
            "cantidad_sa": 1,
            "cantidad_pdp": 0,
            "cantidad_pras": 0,
            "cantidad_pefcf": 0,
            "cantidad_r": 0,
        }
    ]

    with client.session_transaction() as session_data:
        session_data["user"] = "gabo"
        session_data["role"] = "loader"

    base_form = {
        "action": "manual_save",
        "manual_ente_id": "1.3",
        "manual_tipo_auditoria": "Financiera",
        "manual_numero_oficio": "OFS/1142/2026",
        "manual_asunto": "Notificación de Cédula de Resultados",
        "manual_ejercicio": "2025",
        "manual_periodo": "01 de Enero al 20 de Marzo",
        "manual_fecha_notificacion": "2026-05-04",
        "manual_fuente_id": "__new__",
        "manual_fuente_nueva": "PARTICIPACIONES ESTATALES (FONDO GENERAL DE PARTICIPACIONES)",
        "manual_pdp_detalle_json": "[]",
    }

    response = client.post(
        "/carga",
        data={**base_form, "manual_fuentes_detalle_json": json.dumps(first_row, ensure_ascii=False)},
    )
    assert response.status_code == 200

    response = client.post(
        "/carga",
        data={**base_form, "manual_fuentes_detalle_json": json.dumps(full_rows, ensure_ascii=False)},
    )
    assert response.status_code == 200

    conn = sqlite3.connect(test_db_path)
    conn.row_factory = sqlite3.Row
    try:
        cargas_count = conn.execute("SELECT COUNT(*) AS total FROM cargas_manuales").fetchone()["total"]
        observaciones = conn.execute(
            """
            SELECT periodo_cedula, tipo_anexo, COUNT(*) AS total
            FROM observaciones
            GROUP BY periodo_cedula, tipo_anexo
            ORDER BY periodo_cedula, tipo_anexo
            """
        ).fetchall()
    finally:
        conn.close()

    assert cargas_count == 2
    assert [dict(row) for row in observaciones] == [
        {
            "periodo_cedula": "01 de Enero al 20 de Marzo",
            "tipo_anexo": "R",
            "total": 1,
        },
        {
            "periodo_cedula": "21 de Marzo al 16 de Septiembre",
            "tipo_anexo": "SA",
            "total": 1,
        },
    ]


def test_gabo_manual_save_uses_fuente_catalog_classification(client, monkeypatch, tmp_path):
    import app as app_module

    test_db_path = tmp_path / "ramo_flags_save.db"
    monkeypatch.setattr(app_module, "DB_PATH", str(test_db_path))
    app_module.init_db()

    conn = sqlite3.connect(test_db_path)
    try:
        conn.execute(
            """
            INSERT INTO entes_detalle (
                ente_id, ejercicio, ente_numero, ente_nombre,
                responsable, clasificacion, ramo33, ramo28, created_at
            )
            VALUES ('1.1', '2025', '1.1', 'Ente de prueba', '', '', 'No', 'No', 'now')
            """
        )
        conn.execute(
            """
            INSERT INTO fuentes_financiamiento (
                nombre, ramo_33, ramo_28, origen_fuente, created_at
            )
            VALUES
              ('Fondo de Aportaciones', 'Si', 'No', 'Remanentes', 'now'),
              ('Participaciones estatales', 'No', 'Si', 'Del Ejercicio', 'now')
            """
        )
        conn.commit()
    finally:
        conn.close()

    detalle = [
        {
            "tipo_auditoria": "Financiera",
            "fuente_nombre": "Fondo de Aportaciones",
            "periodo": "01 de Enero al 31 de Enero",
            "cantidad_sa": 1,
            "cantidad_pdp": 0,
            "cantidad_pras": 0,
            "cantidad_pefcf": 0,
            "cantidad_r": 0,
            "ramo_33": "No",
            "origen_fuente": "Del Ejercicio",
            "ramo_28": "No",
        },
        {
            "tipo_auditoria": "Financiera",
            "fuente_nombre": "Participaciones estatales",
            "periodo": "01 de Febrero al 28 de Febrero",
            "cantidad_sa": 0,
            "cantidad_pdp": 0,
            "cantidad_pras": 1,
            "cantidad_pefcf": 0,
            "cantidad_r": 0,
            "ramo_33": "Si",
            "origen_fuente": "Remanentes",
            "ramo_28": "No",
        },
    ]

    with client.session_transaction() as session_data:
        session_data["user"] = "gabo"
        session_data["role"] = "loader"

    response = client.post(
        "/carga",
        data={
            "action": "manual_save",
            "manual_ente_id": "1.1",
            "manual_tipo_auditoria": "Financiera",
            "manual_numero_oficio": "OFS/RAMOS/2025",
            "manual_asunto": "Notificación de Cédula de Resultados",
            "manual_ejercicio": "2025",
            "manual_periodo": "01 de Enero al 31 de Enero",
            "manual_fecha_notificacion": "2025-02-01",
            "manual_fuente_id": "__new__",
            "manual_fuente_nueva": "Fondo de Aportaciones",
            "manual_fuentes_detalle_json": json.dumps(detalle, ensure_ascii=False),
            "manual_pdp_detalle_json": "[]",
        },
    )

    assert response.status_code == 200

    conn = sqlite3.connect(test_db_path)
    conn.row_factory = sqlite3.Row
    try:
        cargas = conn.execute(
            """
            SELECT fuente_nombre, ramo_33, origen_fuente, ramo_28
            FROM cargas_manuales
            ORDER BY id ASC
            """
        ).fetchall()
        observaciones = conn.execute(
            """
            SELECT fuente_financiamiento, ramo_33, origen_fuente, ramo_28
            FROM observaciones
            ORDER BY id ASC
            """
        ).fetchall()
    finally:
        conn.close()

    assert [dict(row) for row in cargas] == [
        {
            "fuente_nombre": "Fondo de Aportaciones",
            "ramo_33": "Si",
            "origen_fuente": "Remanentes",
            "ramo_28": "No",
        },
        {
            "fuente_nombre": "Participaciones estatales",
            "ramo_33": "No",
            "origen_fuente": "Del Ejercicio",
            "ramo_28": "Si",
        },
    ]
    assert [dict(row) for row in observaciones] == [
        {
            "fuente_financiamiento": "Fondo de Aportaciones",
            "ramo_33": "Si",
            "origen_fuente": "Remanentes",
            "ramo_28": "No",
        },
        {
            "fuente_financiamiento": "Participaciones estatales",
            "ramo_33": "No",
            "origen_fuente": "Del Ejercicio",
            "ramo_28": "Si",
        },
    ]


def test_gabo_fuente_catalog_classification_updates_existing_rows(client, monkeypatch, tmp_path):
    import app as app_module

    test_db_path = tmp_path / "fuentes_catalog_admin.db"
    monkeypatch.setattr(app_module, "DB_PATH", str(test_db_path))
    app_module.init_db()

    conn = sqlite3.connect(test_db_path)
    try:
        conn.execute(
            """
            INSERT INTO fuentes_financiamiento (
                id, nombre, ramo_33, ramo_28, origen_fuente, created_at
            )
            VALUES (1, 'Participaciones estatales', 'No', 'No', 'Del Ejercicio', 'now')
            """
        )
        conn.execute(
            """
            INSERT INTO cargas_manuales (
                ente_id, ente_nombre, tipo_auditoria, tipo_responsable,
                numero_oficio, asunto, ejercicio, fuente_id, fuente_nombre,
                periodo, periodo_titular, fecha_notificacion, ramo_33, ramo_28,
                origen_fuente, estado, cantidad_sa, cantidad_pdp, cantidad_pras,
                cantidad_pefcf, cantidad_r, monto_pdp_emitido, monto_pdp_solventado,
                monto_pdp_pendiente, fuente_detalle_json, pdp_detalle_json,
                created_by, created_at
            )
            VALUES (
                '1.1', 'Ente de prueba', 'Financiera', 'Titular',
                'OFS/FUENTES/2025', 'Notificación de Cédula de Resultados',
                '2025', 1, 'Participaciones estatales',
                '01 de Enero al 31 de Enero', '01 de Enero al 31 de Enero',
                '2025-02-01', 'No', 'No', 'Del Ejercicio', 'Emitido',
                1, 0, 0, 0, 0, 0, 0, 0, '{}', '[]', 'gabo', 'now'
            )
            """
        )
        conn.execute(
            """
            INSERT INTO observaciones (
                ejercicio, ente_id, ente_nombre, tipo_auditoria,
                fuente_financiamiento, ramo_33, ramo_28, origen_fuente,
                periodo_cedula, periodo_titular, periodo, oficio,
                fecha_notificacion, tipo_anexo, numero_observacion, estado,
                created_at
            )
            VALUES (
                '2025', '1.1', 'Ente de prueba', 'Financiera',
                'Participaciones estatales', 'No', 'No', 'Del Ejercicio',
                '01 de Enero al 31 de Enero', '01 de Enero al 31 de Enero',
                '01 de Enero al 31 de Enero', 'OFS/FUENTES/2025',
                '2025-02-01', 'SA', 1, 'Emitido', 'now'
            )
            """
        )
        conn.commit()
    finally:
        conn.close()

    with client.session_transaction() as session_data:
        session_data["user"] = "gabo"
        session_data["role"] = "loader"

    response = client.post(
        "/carga/fuentes-financiamiento/clasificacion",
        json={
            "ejercicio": "2025",
            "fuente_id": "1",
            "fuente_nombre": "Participaciones estatales",
            "ramo_33": "No",
            "ramo_28": "Si",
            "origen_fuente": "Remanentes",
        },
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["ok"] is True
    assert payload["cargas_actualizadas"] == 1
    assert payload["observaciones_actualizadas"] == 1

    conn = sqlite3.connect(test_db_path)
    conn.row_factory = sqlite3.Row
    try:
        catalog_row = conn.execute(
            """
            SELECT ramo_33, ramo_28, origen_fuente
            FROM fuentes_financiamiento
            WHERE id = 1
            """
        ).fetchone()
        carga_row = conn.execute(
            """
            SELECT ramo_33, ramo_28, origen_fuente
            FROM cargas_manuales
            WHERE fuente_id = 1
            """
        ).fetchone()
        obs_row = conn.execute(
            """
            SELECT ramo_33, ramo_28, origen_fuente
            FROM observaciones
            WHERE fuente_financiamiento = 'Participaciones estatales'
            """
        ).fetchone()
    finally:
        conn.close()

    assert dict(catalog_row) == {
        "ramo_33": "No",
        "ramo_28": "Si",
        "origen_fuente": "Remanentes",
    }
    assert dict(carga_row) == dict(catalog_row)
    assert dict(obs_row) == dict(catalog_row)


def test_gabo_oficios_resumen_repair_action_groups_convenios_as_obra(client, monkeypatch, tmp_path):
    import app as app_module

    test_db_path = tmp_path / "convenios_oficios_resumen.db"
    monkeypatch.setattr(app_module, "DB_PATH", str(test_db_path))
    app_module.init_db()

    conn = sqlite3.connect(test_db_path)
    try:
        conn.execute(
            """
            INSERT INTO entes_detalle (
                ente_id, ejercicio, ente_numero, ente_nombre,
                responsable, clasificacion, ramo33, ramo28, created_at
            )
            VALUES ('1.16', '2025', '1.16', 'Secretaría de Infraestructura (SI)', '', '', 'No', 'No', 'now')
            """
        )
        conn.execute(
            "INSERT INTO fuentes_financiamiento (id, nombre, created_at) VALUES (1, 'Fuente estatal', 'now')"
        )
        conn.execute(
            """
            INSERT INTO cargas_manuales (
                tipo_auditoria, tipo_responsable, numero_oficio, asunto, ejercicio,
                fuente_id, periodo, cantidad_sa, cantidad_pdp, cantidad_pras,
                cantidad_pefcf, cantidad_r, monto_pdp_emitido, monto_pdp_solventado,
                monto_pdp_pendiente, created_by, created_at, ente_id, ente_nombre,
                fuente_nombre, periodo_titular, fecha_notificacion, ramo_33, ramo_28,
                estado, fuente_detalle_json, pdp_detalle_json, modalidad,
                convenio_nombre, convenio_ente_nombre, convenio_ente_id
            )
            VALUES
              ('Financiera', 'Titular', 'OFS/0342/2026', 'Notificación de Cédula de Resultados', '2025',
               1, '01 de Enero al 15 de Mayo', 1, 0, 0, 0, 0, 0, 0, 0, 'gabo', 'now',
               '1.16', 'Secretaría de Infraestructura (SI)', 'Fuente estatal',
               '01 de Enero al 15 de Mayo', '2026-02-11', 'No', 'No', 'Emitido',
               '{}', '[]', 'Fuente', '', '', ''),
              ('Obra Pública', 'Titular', 'OFS/0342/2026', 'Notificación de Cédula de Resultados', '2025',
               1, '01 de Enero al 15 de Mayo', 0, 2, 1, 0, 0, 100, 0, 0, 'gabo', 'now',
               '1.16', 'Secretaría de Infraestructura (SI)', 'Fuente estatal',
               '01 de Enero al 15 de Mayo', '2026-02-11', 'No', 'No', 'Emitido',
               '{}', '[{"concepto":"","subconcepto":"","monto":100,"fuente":""},{"concepto":"","subconcepto":"","monto":0,"fuente":""}]',
               'Convenio', 'CONVENIO DE OBRA PÚBLICA', 'TRIBUNAL DE JUSTICIA ADMINISTRATIVA DEL ESTADO DE TLAXCALA', '33')
            """
        )
        conn.commit()
    finally:
        conn.close()

    with client.session_transaction() as session_data:
        session_data["user"] = "gabo"
        session_data["role"] = "loader"

    response = client.get("/carga/oficios-resumen?ejercicio=2025&ente_id=1.16")
    assert response.status_code == 200
    initial_payload = response.get_json()
    assert initial_payload["rows"] == []
    assert initial_payload["repair_candidates"] == 2

    repair_response = client.post(
        "/carga/oficios-resumen/reparar",
        json={"ejercicio": "2025", "ente_id": "1.16"},
    )
    assert repair_response.status_code == 200
    repair_payload = repair_response.get_json()
    assert repair_payload["ok"] is True
    assert repair_payload["repaired"] == 2

    response = client.get("/carga/oficios-resumen?ejercicio=2025&ente_id=1.16")
    assert response.status_code == 200
    rows = response.get_json()["rows"]
    assert [(row["tipos_auditoria"], row["total"]) for row in rows] == [
        ("Financiera", 1),
        ("Obra Pública", 3),
    ]
    obra = rows[1]
    assert obra["pdp"] == 2
    assert obra["pras"] == 1
    assert obra["monto_emitido"] == 100.0
