"""Tests para SIFET Estatales (07-sifet-estatales)."""
import os
import pytest

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
