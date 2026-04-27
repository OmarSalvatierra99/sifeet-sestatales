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

    obra = next(item for item in payload["auditorias"] if item["tipo"] == "Obra Pública")
    convenios = [fuente for fuente in obra["fuentes"] if fuente.get("modalidad") == "Convenio"]

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


def test_gabo_oficios_resumen_repairs_cargas_and_groups_convenios_as_obra(client, monkeypatch, tmp_path):
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
    rows = response.get_json()["rows"]
    assert [(row["tipos_auditoria"], row["total"]) for row in rows] == [
        ("Financiera", 1),
        ("Obra Pública", 3),
    ]
    obra = rows[1]
    assert obra["pdp"] == 2
    assert obra["pras"] == 1
    assert obra["monto_emitido"] == 100.0
