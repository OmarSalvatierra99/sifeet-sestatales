import io
import sqlite3
from pathlib import Path

import pytest

from scripts.parsers import parse_solventacion


EXAMPLES_DIR = Path(__file__).resolve().parents[1] / "examples"
SOLVENTACION_PDF = EXAMPLES_DIR / "1.16.- SI_OFS_0985_2026_Ene-Jun.pdf"


@pytest.fixture
def solventacion_client(monkeypatch, tmp_path):
    import app as app_module
    import scripts.gabo_routes as gabo_routes

    test_db_path = tmp_path / "solventacion_import.db"
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
                ('ENTE-2025-SI', '1.16', '2025', '1.16', 'Secretaría de Infraestructura (SI)', '', '', 'No', 'No', 'now'),
                ('ENTE-2025-USET', '27', '2025', '27', 'Unidad de Servicios Educativos del Estado de Tlaxcala (USET)', '', '', 'No', 'No', 'now')
            """
        )

    app_module.app.config["TESTING"] = True
    with app_module.app.test_client() as client:
        with client.session_transaction() as session_data:
            session_data["user"] = "gabo"
            session_data["role"] = "loader"
        yield client


def test_parse_solventacion_extracts_periodo_plural_and_destinatario():
    parsed = parse_solventacion(str(SOLVENTACION_PDF))

    assert parsed["oficio"] == "OFS/0985/2026"
    assert parsed["ejercicio"] == "2025"
    assert parsed["oficio_base"] == "OFS/0342/2026"
    assert parsed["periodo"] == "01 de Enero - 15 de Mayo, 16 de Mayo - 30 de Junio"
    assert parsed["destinatario"] == "SECRETARIO DE INFRAESTRUCTURA (SI)"
    assert len(parsed["auditorias"]) >= 1


def test_api_solventacion_accepts_pdf_and_returns_rows(solventacion_client):
    client = solventacion_client

    with SOLVENTACION_PDF.open("rb") as fh:
        response = client.post(
            "/api/solventacion/procesar",
            data={
                "ejercicio": "2025",
                "ente_id": "1.16",
                "file": (io.BytesIO(fh.read()), SOLVENTACION_PDF.name),
            },
            content_type="multipart/form-data",
        )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["ok"] is True
    assert payload["mode"] == "solventacion"
    assert payload["oficio"] == "OFS/0985/2026"
    assert payload["oficio_base"] == "OFS/0342/2026"
    assert payload["destinatario"] == "SECRETARIO DE INFRAESTRUCTURA (SI)"
    assert payload["rows"]


def test_api_solventacion_rejects_wrong_ente(solventacion_client):
    client = solventacion_client

    with SOLVENTACION_PDF.open("rb") as fh:
        response = client.post(
            "/api/solventacion/procesar",
            data={
                "ejercicio": "2025",
                "ente_id": "27",
                "file": (io.BytesIO(fh.read()), SOLVENTACION_PDF.name),
            },
            content_type="multipart/form-data",
        )

    assert response.status_code == 400
    payload = response.get_json()
    assert payload["ok"] is False
    assert "corresponde al ente 1.16" in payload["error"]
