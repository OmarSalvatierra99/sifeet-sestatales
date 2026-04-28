"""Pruebas de regresión para logout y comparativo anual."""

from datetime import UTC, datetime
import sqlite3

import pytest


@pytest.fixture
def client():
    from app import app

    app.config["TESTING"] = True
    with app.test_client() as test_client:
        yield test_client


def _set_logged_user(client, username: str, role: str) -> None:
    with client.session_transaction() as session_data:
        session_data["user"] = username
        session_data["role"] = role


def _insert_observacion(
    conn: sqlite3.Connection,
    *,
    ejercicio: str,
    tipo_anexo: str,
    numero_observacion: int,
    estado: str = "Pendiente",
    oficio: str = "OFS/0001/2025",
    monto_emitido_override: float | None = None,
    monto_solventado_override: float | None = None,
    monto_pendiente_override: float | None = None,
) -> None:
    now = datetime.now(UTC).isoformat(timespec="seconds")
    monto_emitido = 100.0 if tipo_anexo == "PDP" else 0.0
    monto_solventado = 25.0 if tipo_anexo == "PDP" else 0.0
    monto_pendiente = 75.0 if tipo_anexo == "PDP" else 0.0
    if monto_emitido_override is not None:
        monto_emitido = monto_emitido_override
    if monto_solventado_override is not None:
        monto_solventado = monto_solventado_override
    if monto_pendiente_override is not None:
        monto_pendiente = monto_pendiente_override
    conn.execute(
        """
        INSERT INTO observaciones (
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
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            ejercicio,
            "ENTE-001",
            "001",
            "Ente de prueba",
            "Financiera",
            "Ingresos propios",
            "No",
            "01 de enero al 31 de enero",
            "01 de enero al 31 de enero",
            oficio,
            "2025-01-01",
            tipo_anexo,
            numero_observacion,
            estado,
            monto_emitido,
            monto_solventado,
            monto_pendiente,
            now,
        ),
    )


def test_gabo_logout_uses_post_and_clears_session(client):
    """La vista de Gabo debe cerrar sesión vía POST y limpiar la sesión."""
    _set_logged_user(client, "gabo", "loader")

    page = client.get("/carga")
    assert page.status_code == 200
    assert b'id="gaboLogoutForm"' in page.data
    assert b'action="/logout"' in page.data

    response = client.post("/logout")
    assert response.status_code == 302
    assert response.headers["Location"].endswith("/login")

    with client.session_transaction() as session_data:
        assert "user" not in session_data
        assert "role" not in session_data

    protected = client.get("/carga")
    assert protected.status_code == 302
    assert "/login" in protected.headers["Location"]


def test_gabo_oficios_resumen_groups_observaciones_by_oficio(client, monkeypatch, tmp_path):
    """El resumen de Gabo debe agrupar conteos y montos por oficio."""
    import app as app_module

    test_db_path = tmp_path / "oficios_resumen_test.db"
    monkeypatch.setattr(app_module, "DB_PATH", str(test_db_path))
    app_module.init_db()

    conn = sqlite3.connect(test_db_path)
    try:
        _insert_observacion(conn, ejercicio="2025", tipo_anexo="SA", numero_observacion=1)
        _insert_observacion(conn, ejercicio="2025", tipo_anexo="PDP", numero_observacion=2)
        _insert_observacion(conn, ejercicio="2025", tipo_anexo="PRAS", numero_observacion=3)
        _insert_observacion(conn, ejercicio="2025", tipo_anexo="R", numero_observacion=4)
        _insert_observacion(conn, ejercicio="2025", tipo_anexo="PEFCF", numero_observacion=5)
        conn.commit()
    finally:
        conn.close()

    _set_logged_user(client, "gabo", "loader")
    response = client.get("/carga/oficios-resumen?ejercicio=2025&ente_id=ENTE-001")

    assert response.status_code == 200
    payload = response.get_json()
    assert len(payload["rows"]) == 1
    row = payload["rows"][0]
    assert row["oficio"] == "OFS/0001/2025"
    assert row["sa"] == 1
    assert row["pdp"] == 1
    assert row["pras"] == 1
    assert row["r"] == 1
    assert row["pefcf"] == 1
    assert row["monto_emitido"] == 100.0
    assert row["monto_solventado"] == 25.0
    assert row["monto_pendiente"] == 75.0


def test_gabo_oficios_resumen_recalculates_pending_amounts(client, monkeypatch, tmp_path):
    """El pendiente se calcula desde emitido-solventado aunque el dato guardado venga en cero."""
    import app as app_module

    test_db_path = tmp_path / "oficios_resumen_pendiente_test.db"
    monkeypatch.setattr(app_module, "DB_PATH", str(test_db_path))
    app_module.init_db()

    conn = sqlite3.connect(test_db_path)
    try:
        _insert_observacion(
            conn,
            ejercicio="2025",
            tipo_anexo="PDP",
            numero_observacion=1,
            estado="Pendiente",
            monto_emitido_override=21996079.74,
            monto_solventado_override=0.0,
            monto_pendiente_override=0.0,
        )
        _insert_observacion(
            conn,
            ejercicio="2025",
            tipo_anexo="PDP",
            numero_observacion=2,
            estado="Solventado",
            monto_emitido_override=100.0,
            monto_solventado_override=0.0,
            monto_pendiente_override=0.0,
        )
        conn.commit()
    finally:
        conn.close()

    _set_logged_user(client, "gabo", "loader")
    response = client.get("/carga/oficios-resumen?ejercicio=2025&ente_id=ENTE-001")

    assert response.status_code == 200
    row = response.get_json()["rows"][0]
    assert row["monto_emitido"] == pytest.approx(21996179.74)
    assert row["monto_solventado"] == pytest.approx(100.0)
    assert row["monto_pendiente"] == pytest.approx(21996079.74)


def test_gabo_can_delete_observaciones_by_oficio_scope(client, monkeypatch, tmp_path):
    """El borrado por oficio debe eliminar solo el oficio seleccionado."""
    import app as app_module

    test_db_path = tmp_path / "oficios_delete_test.db"
    monkeypatch.setattr(app_module, "DB_PATH", str(test_db_path))
    app_module.init_db()

    conn = sqlite3.connect(test_db_path)
    try:
        _insert_observacion(conn, ejercicio="2025", tipo_anexo="SA", numero_observacion=1)
        _insert_observacion(conn, ejercicio="2025", tipo_anexo="PDP", numero_observacion=2)
        _insert_observacion(
            conn,
            ejercicio="2025",
            tipo_anexo="PRAS",
            numero_observacion=3,
            oficio="OFS/0002/2025",
        )
        conn.commit()
    finally:
        conn.close()

    _set_logged_user(client, "gabo", "loader")
    response = client.post(
        "/carga/observaciones-admin/borrar-todo",
        json={"ejercicio": "2025", "ente_id": "ENTE-001", "oficio": "OFS/0001/2025"},
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["ok"] is True
    assert payload["deleted"] == 2

    conn = sqlite3.connect(test_db_path)
    try:
        remaining_selected = conn.execute(
            "SELECT COUNT(*) FROM observaciones WHERE oficio = 'OFS/0001/2025'"
        ).fetchone()[0]
        remaining_other = conn.execute(
            "SELECT COUNT(*) FROM observaciones WHERE oficio = 'OFS/0002/2025'"
        ).fetchone()[0]
    finally:
        conn.close()

    assert remaining_selected == 0
    assert remaining_other == 1


def test_comparativo_anual_stats_orders_and_normalizes_anexos(
    client,
    monkeypatch,
    tmp_path,
):
    """El comparativo anual debe devolver los 5 anexos canónicos en orden fijo."""
    import app as app_module

    test_db_path = tmp_path / "comparativo_test.db"
    monkeypatch.setattr(app_module, "DB_PATH", str(test_db_path))
    app_module.init_db()

    conn = sqlite3.connect(test_db_path)
    try:
        _insert_observacion(conn, ejercicio="2024", tipo_anexo="R", numero_observacion=1)
        _insert_observacion(conn, ejercicio="2024", tipo_anexo="PEFCT", numero_observacion=2)
        _insert_observacion(conn, ejercicio="2024", tipo_anexo="SA", numero_observacion=3)
        _insert_observacion(conn, ejercicio="2024", tipo_anexo="PRAS", numero_observacion=4, estado="Solventado")
        _insert_observacion(conn, ejercicio="2024", tipo_anexo="PDP", numero_observacion=5)
        conn.commit()
    finally:
        conn.close()

    _set_logged_user(client, "luis", "viewer")
    response = client.get("/comparativo-anual/stats?ejercicio=2024")
    assert response.status_code == 200

    payload = response.get_json()
    anexo_totals = payload["summary"]["anexo_totals_by_year"]
    stacked_rows = payload["summary"]["stacked_by_anexo"]

    assert [row["tipo_anexo"] for row in anexo_totals] == [
        "SA",
        "PDP",
        "PRAS",
        "PEFCF",
        "R",
    ]
    assert [row["tipo_anexo"] for row in stacked_rows] == [
        "SA",
        "PDP",
        "PRAS",
        "PEFCF",
        "R",
    ]
    assert all(row["ejercicio"] == "2024" for row in anexo_totals)
    assert all(row["tipo_anexo"] != "PEFCT" for row in anexo_totals)
