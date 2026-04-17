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
) -> None:
    now = datetime.now(UTC).isoformat(timespec="seconds")
    monto_emitido = 100.0 if tipo_anexo == "PDP" else 0.0
    monto_solventado = 25.0 if tipo_anexo == "PDP" else 0.0
    monto_pendiente = 75.0 if tipo_anexo == "PDP" else 0.0
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
            "OFS/0001/2025",
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
