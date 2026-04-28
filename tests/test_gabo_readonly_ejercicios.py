import json
import sqlite3

import pytest


@pytest.fixture
def readonly_client(monkeypatch, tmp_path):
    import app as app_module
    import scripts.gabo_routes as gabo_routes

    test_db_path = tmp_path / "readonly_gabo.db"
    monkeypatch.setattr(app_module, "DB_PATH", str(test_db_path))
    monkeypatch.setattr(gabo_routes, "DB_PATH", str(test_db_path))
    monkeypatch.setattr(gabo_routes, "BASE_DIR", str(tmp_path))
    app_module.init_db()

    with sqlite3.connect(test_db_path) as conn:
        conn.execute(
            """
            INSERT INTO entes_detalle (
                ente_uid, ente_id, ejercicio, ente_numero, ente_nombre,
                responsable, clasificacion, ramo33, ramo28, created_at
            )
            VALUES
                ('ENTE-2024-1', '1.1', '2024', '1.1', 'Ente Histórico', '', '', 'No', 'No', 'now'),
                ('ENTE-2025-1', '1.1', '2025', '1.1', 'Ente Editable', '', '', 'No', 'No', 'now')
            """
        )
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
                periodo, periodo_titular, fecha_notificacion,
                ramo_33, ramo_28, origen_fuente, estado,
                cantidad_sa, cantidad_pdp, cantidad_pras, cantidad_pefcf, cantidad_r,
                monto_pdp_emitido, monto_pdp_solventado, monto_pdp_pendiente,
                fuente_detalle_json, pdp_detalle_json, created_by, created_at
            )
            VALUES (
                '1.1', 'Ente Histórico', 'Financiera', 'Titular',
                'OFS/2024/001', 'Notificación de Cédula de Resultados',
                '2024', 1, 'Participaciones estatales',
                '01 de Enero al 31 de Enero', '01 de Enero al 31 de Enero', '2024-02-01',
                'No', 'No', 'Del Ejercicio', 'Emitido',
                1, 0, 0, 0, 0, 0, 0, 0, '{}', '[]', 'gabo', 'now'
            )
            """
        )
        conn.execute(
            """
            INSERT INTO observaciones (
                id, ejercicio, ente_id, ente_nombre, tipo_auditoria,
                fuente_financiamiento, ramo_33, ramo_28, origen_fuente,
                periodo_cedula, periodo_titular, periodo, oficio,
                fecha_notificacion, tipo_anexo, numero_observacion, estado,
                monto_pdp_emitido, monto_pdp_solventado, monto_pdp_pendiente,
                created_at
            )
            VALUES (
                1, '2024', '1.1', 'Ente Histórico', 'Financiera',
                'Participaciones estatales', 'No', 'No', 'Del Ejercicio',
                '01 de Enero al 31 de Enero', '01 de Enero al 31 de Enero',
                '01 de Enero al 31 de Enero', 'OFS/2024/001',
                '2024-02-01', 'SA', 1, 'Emitido', 0, 0, 0, 'now'
            )
            """
        )
        conn.execute(
            """
            INSERT INTO historial_titulares (
                id, ejercicio, ente_uid, ente, tipo_auditoria,
                nombre, cargo, fecha_inicio, fecha_fin, tipo_registro
            )
            VALUES (
                1, 2024, 'ENTE-2024-1', 'Ente Histórico', 'Financiera',
                'Titular Histórico', 'Titular', '2024-01-01', '2024-12-31', 'titular'
            )
            """
        )

    app_module.app.config["TESTING"] = True
    with app_module.app.test_client() as client:
        with client.session_transaction() as session_data:
            session_data["user"] = "gabo"
            session_data["role"] = "loader"
        yield client, test_db_path


def _fetch_one(db_path, query, params=()):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        return dict(conn.execute(query, params).fetchone())
    finally:
        conn.close()


def test_gabo_cannot_reclassify_sources_for_readonly_years(readonly_client):
    client, db_path = readonly_client

    for ejercicio in ("2023", "2024"):
        response = client.post(
            "/carga/fuentes-financiamiento/clasificacion",
            json={
                "ejercicio": ejercicio,
                "fuente_id": "1",
                "fuente_nombre": "Participaciones estatales",
                "ramo_33": "Si",
                "ramo_28": "Si",
                "origen_fuente": "Remanentes",
            },
        )

        assert response.status_code == 403
        assert "concluido" in response.get_json()["error"]
    assert _fetch_one(
        db_path,
        "SELECT ramo_33, ramo_28, origen_fuente FROM observaciones WHERE id = 1",
    ) == {"ramo_33": "No", "ramo_28": "No", "origen_fuente": "Del Ejercicio"}


def test_gabo_cannot_update_readonly_observation(readonly_client):
    client, db_path = readonly_client

    response = client.post(
        "/carga/observaciones-cargadas/1/actualizar",
        json={
            "estado": "Solventado",
            "scope": {"ejercicio": "2024", "ente_id": "1.1", "oficio": "OFS/2024/001"},
        },
    )

    assert response.status_code == 403
    assert "concluido" in response.get_json()["error"]
    assert _fetch_one(db_path, "SELECT estado FROM observaciones WHERE id = 1") == {
        "estado": "Emitido"
    }


def test_gabo_cannot_move_readonly_titular_to_editable_year(readonly_client):
    client, db_path = readonly_client

    response = client.post(
        "/carga/titulares/historial/1/actualizar",
        json={
            "ejercicio": "2025",
            "ente_id": "1.1",
            "tipo_auditoria": "Financiera",
            "tipo_registro": "titular",
            "nombre": "Titular Actualizado",
            "cargo": "Titular",
            "fecha_inicio": "2025-01-01",
            "fecha_fin": "2025-12-31",
        },
    )

    assert response.status_code == 403
    assert "concluido" in response.get_json()["error"]
    assert _fetch_one(
        db_path,
        "SELECT CAST(ejercicio AS TEXT) AS ejercicio, nombre FROM historial_titulares WHERE id = 1",
    ) == {"ejercicio": "2024", "nombre": "Titular Histórico"}


def test_gabo_titulares_form_blocks_readonly_year_without_fallback(readonly_client):
    client, db_path = readonly_client

    response = client.post(
        "/carga/titulares",
        data={
            "titular_ejercicio": "2024",
            "titular_ente_id": "1.1",
            "titular_tipo_auditoria": "Financiera",
            "titular_nombre": "Nuevo Titular",
            "titular_fecha_inicio": "2024-01-01",
            "titular_fecha_fin": "2024-12-31",
            "titular_administrativo": "Nuevo Administrativo",
            "titular_admin_fecha_inicio": "2024-01-01",
            "titular_admin_fecha_fin": "2024-12-31",
        },
    )

    assert response.status_code == 200
    assert b"concluido" in response.data
    assert _fetch_one(db_path, "SELECT COUNT(*) AS total FROM cargas_titulares") == {"total": 0}
    assert _fetch_one(
        db_path,
        "SELECT COUNT(*) AS total FROM historial_titulares WHERE nombre = 'Nuevo Titular'",
    ) == {"total": 0}


def test_gabo_manual_carga_blocks_readonly_year_before_insert(readonly_client):
    client, db_path = readonly_client
    detalle = [
        {
            "tipo_auditoria": "Financiera",
            "fuente_nombre": "Participaciones estatales",
            "periodo": "01 de Enero al 31 de Enero",
            "cantidad_sa": 1,
            "cantidad_pdp": 0,
            "cantidad_pras": 0,
            "cantidad_pefcf": 0,
            "cantidad_r": 0,
        }
    ]

    response = client.post(
        "/carga",
        data={
            "action": "manual_save",
            "manual_ente_id": "1.1",
            "manual_tipo_auditoria": "Financiera",
            "manual_numero_oficio": "OFS/2024/002",
            "manual_asunto": "Notificación de Cédula de Resultados",
            "manual_ejercicio": "2024",
            "manual_fecha_notificacion": "2024-02-02",
            "manual_periodo": "01 de Enero al 31 de Enero",
            "manual_fuentes_detalle_json": json.dumps(detalle),
            "manual_pdp_detalle_json": "[]",
        },
    )

    assert response.status_code == 200
    assert b"concluido" in response.data
    assert _fetch_one(
        db_path,
        "SELECT COUNT(*) AS total FROM cargas_manuales WHERE numero_oficio = 'OFS/2024/002'",
    ) == {"total": 0}
    assert _fetch_one(
        db_path,
        "SELECT COUNT(*) AS total FROM observaciones WHERE oficio = 'OFS/2024/002'",
    ) == {"total": 0}
