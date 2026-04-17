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
