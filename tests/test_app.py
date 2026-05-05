"""Tests para SIFET Estatales (07-sifet-estatales)."""
from io import BytesIO
import os
import pytest
from openpyxl import load_workbook

os.environ.setdefault("SECRET_KEY", "test-secret-key-for-testing-only")
os.environ.setdefault("FLASK_ENV", "testing")


@pytest.fixture
def client():
    from app import app
    app.config["TESTING"] = True
    with app.test_client() as c:
        yield c


def test_health_standard_route(client):
    """GET /api/health debe retornar 200 sin autenticación."""
    r = client.get("/api/health")
    assert r.status_code == 200
    data = r.get_json()
    assert data["status"] == "ok"
    assert data["service"] == "sifet-estatales"


def test_health_compat_route(client):
    """GET /health también debe retornar 200."""
    r = client.get("/health")
    assert r.status_code == 200


def test_login_page_loads(client):
    """GET /login debe cargar el formulario de acceso."""
    r = client.get("/login")
    assert r.status_code == 200


def test_root_redirects_to_login(client):
    """GET / sin sesión debe redirigir al login."""
    r = client.get("/")
    assert r.status_code in (200, 302)


def test_login_with_valid_credentials(client):
    """POST /login con credenciales válidas debe redirigir al dashboard."""
    r = client.post("/login", data={"username": "luis", "password": "luis2025"})
    assert r.status_code == 302


def test_odilia_login_uses_luis_viewer_access(client):
    """Odilia debe entrar con rol de consulta desde el catálogo compartido."""
    login_response = client.post(
        "/login",
        data={"username": "odilia", "password": "odilia2025"},
    )
    assert login_response.status_code == 302

    response = client.get("/")
    assert response.status_code == 200
    assert b'data-luis-view="consulta"' in response.data
    assert "C.P. Odilia Cuamatzi Bautista".encode("utf-8") in response.data

    export_response = client.get("/fuentes-financiamiento-exportar?ejercicio=2025")
    assert export_response.status_code == 200
    assert export_response.headers["Content-Type"].startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@pytest.mark.parametrize(
    ("username", "display_name"),
    [
        ("luis", "C.P Luis Felipe Camilo Fuentes"),
        ("odilia", "C.P. Odilia Cuamatzi Bautista"),
    ],
)
def test_viewer_operational_pages_share_navigation(client, username, display_name):
    """Los usuarios de consulta deben exponer la misma navegación modular."""
    with client.session_transaction() as session_data:
        session_data["user"] = username
        session_data["role"] = "viewer"

    pages = [
        ("/", b'data-luis-view="consulta"', b"Consulta operativa"),
        ("/resumen-general", b'data-luis-view="resumen_general"', b"Resumen general"),
        ("/tipo-auditoria", b'data-luis-view="tipo_auditoria"', b"Tipo de Auditor\xc3\xada"),
        ("/fuente-financiamiento", b'data-luis-view="fuente_financiamiento"', b"Fuente de Financiamiento"),
        ("/graficas", b'data-luis-view="graficas"', b"Gr\xc3\xa1ficas"),
        ("/titulares-administrativos", b'data-luis-view="titulares"', b"Titulares y administrativos"),
        ("/pendientes-periodo", b'data-luis-view="pendientes"', b"Pendientes por periodo"),
        ("/catalogo", b'data-luis-view="catalogo"', b"Cat\xc3\xa1logo de entes"),
    ]
    for path, view_marker, label in pages:
        response = client.get(path)
        assert response.status_code == 200
        assert view_marker in response.data
        assert label in response.data
        assert b"Comparativo anual" in response.data
        assert display_name.encode("utf-8") in response.data
        if path == "/fuente-financiamiento":
            assert b"Exportar Fuentes de Financiamiento" in response.data


def test_login_with_invalid_credentials(client):
    """POST /login con credenciales inválidas debe volver al login."""
    r = client.post("/login", data={"username": "luis", "password": "wrong"})
    assert r.status_code == 200


def test_fuentes_export_requires_and_uses_ejercicio(client):
    """La exportación de fuentes debe estar acotada a un ejercicio."""
    with client.session_transaction() as session_data:
        session_data["user"] = "luis"
        session_data["role"] = "viewer"

    missing = client.get("/fuentes-financiamiento-exportar")
    assert missing.status_code == 400

    response = client.get("/fuentes-financiamiento-exportar?ejercicio=2025")
    assert response.status_code == 200
    assert response.headers["Content-Type"].startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    assert "fuentes_financiamiento_2025_" in response.headers["Content-Disposition"]
    workbook = load_workbook(BytesIO(response.data))
    sheet = workbook.active
    assert sheet["A1"].value == "Fuentes de Financiamiento"
    assert sheet["A2"].value == "Ejercicio fiscal"
    assert sheet["B2"].value == "2025"
    assert sheet["A6"].value == "No."
    assert sheet["B6"].value == "Fuente de Financiamiento"


def test_luis_breakdown_exports(client):
    """Los resúmenes de Luis deben exportar Excel por agrupación."""
    with client.session_transaction() as session_data:
        session_data["user"] = "luis"
        session_data["role"] = "viewer"

    for group_by in ("general", "tipo_auditoria", "fuente_financiamiento"):
        response = client.get(
            f"/observaciones-desglose-exportar?ejercicio=2025&group_by={group_by}"
        )
        assert response.status_code == 200
        assert response.headers["Content-Type"].startswith(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        assert f"resumen_{group_by}_2025_" in response.headers["Content-Disposition"]


def test_gabo_tools_page_only_shows_fuentes(client):
    """La página de herramientas de Gabo debe mostrar solo fuentes de financiamiento."""
    with client.session_transaction() as session_data:
        session_data["user"] = "gabo"
        session_data["role"] = "loader"

    response = client.get("/carga/herramientas?ejercicio=2025")

    assert response.status_code == 200
    assert b"Fuentes de Financiamiento" in response.data
    assert b"Exportar fuentes" not in response.data
    assert b"Exportaciones" not in response.data
    assert b"Accesos r\xc3\xa1pidos" not in response.data


def test_gabo_can_export_observaciones_with_ramo_filter(client):
    """Gabo debe poder exportar reportes de observaciones filtrados por Ramo 33."""
    with client.session_transaction() as session_data:
        session_data["user"] = "gabo"
        session_data["role"] = "loader"

    response = client.get("/observaciones-exportar?ejercicio=2025&ramo_33=Si")

    assert response.status_code == 200
    assert response.headers["Content-Type"].startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    assert "observaciones_2025_" in response.headers["Content-Disposition"]


def test_auth_users_from_catalog():
    """_build_users() debe leer credenciales del catálogo compartido."""
    from app import _build_users
    from werkzeug.security import check_password_hash
    users = _build_users()
    assert "luis" in users
    assert users["luis"]["role"] == "viewer"
    assert check_password_hash(users["luis"]["password_hash"], "luis2025")
    assert "gabo" in users
    assert users["gabo"]["role"] == "loader"
    assert "odilia" in users
    assert users["odilia"]["role"] == "viewer"
    assert check_password_hash(users["odilia"]["password_hash"], "odilia2025")
