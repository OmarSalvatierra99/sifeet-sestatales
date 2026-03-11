from datetime import datetime
from functools import wraps
from io import BytesIO
import os
from pathlib import Path
import re
import sqlite3
import subprocess
import sys
import unicodedata

from flask import Flask, render_template, request, redirect, url_for, g, jsonify, session, send_file
from openpyxl import Workbook
from openpyxl.styles import Font
from werkzeug.security import check_password_hash, generate_password_hash
from scripts.gabo_routes import register_gabo_routes
from scripts.luis_routes import register_luis_routes

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "sifeet.db")

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "dev-secret-key")
template_reload_env = os.getenv("TEMPLATES_AUTO_RELOAD")
template_auto_reload = (
    template_reload_env.strip().lower() in {"1", "true", "yes", "on"}
    if template_reload_env
    else False
)
app.config["TEMPLATES_AUTO_RELOAD"] = template_auto_reload
app.jinja_env.auto_reload = template_auto_reload

USERS = {
    "luis": {
        "password_hash": generate_password_hash("luis2025"),
        "role": "viewer",
    },
    "gabo": {
        "password_hash": generate_password_hash("gabo2025"),
        "role": "loader",
    },
}
LUIS_USERNAME = "luis"
GABO_USERNAME = "gabo"
ASUNTOS_MANUALES = {
    "Notificación de Cédula de Resultados",
}
TIPOS_RESPONSABLE = {"Titular", "Administrativo", "Ambos"}

UID_PREFIX = "ENT-"
UID_PATTERN = re.compile(rf"^{UID_PREFIX}(\\d+)$")
SIGLA_QUOTE_PATTERN = re.compile(r"[\"“”]([^\"“”]+)[\"“”]")
SIGLA_PAREN_PATTERN = re.compile(r"\\(([^)]+)\\)")

MONTHS_ES = {
    "01": "enero",
    "02": "febrero",
    "03": "marzo",
    "04": "abril",
    "05": "mayo",
    "06": "junio",
    "07": "julio",
    "08": "agosto",
    "09": "septiembre",
    "10": "octubre",
    "11": "noviembre",
    "12": "diciembre",
}
MONTHS_ES_TO_NUM = {normalize_key: number for number, normalize_key in (
    ("01", "enero"),
    ("02", "febrero"),
    ("03", "marzo"),
    ("04", "abril"),
    ("05", "mayo"),
    ("06", "junio"),
    ("07", "julio"),
    ("08", "agosto"),
    ("09", "septiembre"),
    ("10", "octubre"),
    ("11", "noviembre"),
    ("12", "diciembre"),
)}


def periodo_sql(alias: str) -> str:
    start_month = "CASE " + " ".join(
        f"WHEN strftime('%m', {alias}.fecha_inicio) = '{key}' THEN '{value}'"
        for key, value in MONTHS_ES.items()
    ) + " END"
    end_month = "CASE " + " ".join(
        f"WHEN strftime('%m', {alias}.fecha_fin) = '{key}' THEN '{value}'"
        for key, value in MONTHS_ES.items()
    ) + " END"
    start_day = f"printf('%02d', CAST(strftime('%d', {alias}.fecha_inicio) AS INTEGER))"
    end_day = f"printf('%02d', CAST(strftime('%d', {alias}.fecha_fin) AS INTEGER))"
    return (
        f"{start_day} || ' de ' || {start_month} || ' al ' || "
        f"{end_day} || ' de ' || {end_month}"
    )


def normalize_ente_id(value: str) -> str:
    return (value or "").strip().rstrip(".").strip()


def normalize_ente_id_sql(column: str) -> str:
    return f"RTRIM(TRIM(COALESCE({column}, '')), '.')"


def normalize_text_key(value: str) -> str:
    clean = (value or "").strip().lower()
    if not clean:
        return ""
    clean = unicodedata.normalize("NFKD", clean)
    clean = "".join(char for char in clean if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", clean).strip()


def normalize_tipo_auditoria(value: str) -> str:
    clean = (value or "").strip()
    key = normalize_text_key(clean)
    if key in {"auditoria", "auditoria financiera", "financiera", "financiero"}:
        return "Financiera"
    if key in {"obra publica", "obra"}:
        return "Obra Pública"
    if key == "cuenta publica":
        return "Cuenta Pública"
    return clean


def parse_periodo_cedula(ejercicio: str, periodo_cedula: str):
    if not ejercicio or not periodo_cedula:
        return None, None
    match = re.match(
        r"^\s*(\d{1,2})\s+de\s+([a-zA-ZáéíóúÁÉÍÓÚñÑ]+)\s+al\s+(\d{1,2})\s+de\s+([a-zA-ZáéíóúÁÉÍÓÚñÑ]+)\s*$",
        periodo_cedula,
    )
    if not match:
        return None, None

    start_day = int(match.group(1))
    start_month_key = normalize_text_key(match.group(2))
    end_day = int(match.group(3))
    end_month_key = normalize_text_key(match.group(4))

    start_month = MONTHS_ES_TO_NUM.get(start_month_key)
    end_month = MONTHS_ES_TO_NUM.get(end_month_key)
    if not start_month or not end_month:
        return None, None

    try:
        year = int(str(ejercicio).strip())
    except ValueError:
        return None, None

    try:
        start_date = datetime(year, int(start_month), start_day)
        end_date = datetime(year, int(end_month), end_day)
    except ValueError:
        return None, None

    return start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")


def parse_historial_date(value: str):
    raw = (value or "").strip()
    if not raw:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    return None


def parse_non_negative_int(value: str, label: str) -> int:
    raw = (value or "").strip()
    if not raw:
        return 0
    try:
        parsed = int(raw)
    except ValueError as exc:
        raise ValueError(f"{label}: valor inválido.") from exc
    if parsed < 0:
        raise ValueError(f"{label}: no puede ser negativo.")
    return parsed


def parse_non_negative_float(value: str, label: str) -> float:
    raw = (value or "").strip()
    if not raw:
        return 0.0
    cleaned = raw.replace(",", "")
    try:
        parsed = float(cleaned)
    except ValueError as exc:
        raise ValueError(f"{label}: monto inválido.") from exc
    if parsed < 0:
        raise ValueError(f"{label}: no puede ser negativo.")
    return parsed


def parse_ente_numero_sort(value: str) -> float:
    raw = (value or "").strip()
    if not raw:
        return 0.0
    match = re.search(r"-?\d+(?:\.\d+)?", raw.replace(",", ""))
    if not match:
        return 0.0
    try:
        return float(match.group(0))
    except ValueError:
        return 0.0


def get_ente_aliases_by_uid(
    conn: sqlite3.Connection,
    ejercicio: str,
    ente_id_norm: str,
    fallback_names=None,
):
    aliases = []
    for name in (fallback_names or []):
        clean = (name or "").strip()
        if clean and clean not in aliases:
            aliases.append(clean)

    if not ejercicio or not ente_id_norm:
        return aliases

    base_row = conn.execute(
        f"""
        SELECT ente_uid, ente_nombre
        FROM entes_detalle
        WHERE ejercicio = ? AND {normalize_ente_id_sql('ente_id')} = ?
        LIMIT 1
        """,
        (ejercicio, ente_id_norm),
    ).fetchone()
    if not base_row:
        return aliases

    base_name = (base_row["ente_nombre"] or "").strip()
    if base_name and base_name not in aliases:
        aliases.append(base_name)

    ente_uid = (base_row["ente_uid"] or "").strip()
    if not ente_uid:
        return aliases

    uid_rows = conn.execute(
        """
        SELECT DISTINCT ente_nombre
        FROM entes_detalle
        WHERE ente_uid = ?
          AND ente_nombre IS NOT NULL
          AND TRIM(ente_nombre) != ''
        ORDER BY ejercicio
        """,
        (ente_uid,),
    ).fetchall()
    for row in uid_rows:
        name = (row["ente_nombre"] or "").strip()
        if name and name not in aliases:
            aliases.append(name)
    return aliases


def get_ente_uid_by_ejercicio_id(
    conn: sqlite3.Connection,
    ejercicio: str,
    ente_id_norm: str,
) -> str:
    if not ejercicio or not ente_id_norm:
        return ""
    row = conn.execute(
        f"""
        SELECT ente_uid
        FROM entes_detalle
        WHERE ejercicio = ? AND {normalize_ente_id_sql('ente_id')} = ?
        LIMIT 1
        """,
        (ejercicio, ente_id_norm),
    ).fetchone()
    if not row:
        return ""
    return (row["ente_uid"] or "").strip()


def backfill_historial_ente_uids(conn: sqlite3.Connection) -> None:
    pending_rows = conn.execute(
        """
        SELECT id, ejercicio, ente
        FROM historial_titulares
        WHERE ente_uid IS NULL OR TRIM(ente_uid) = ''
        ORDER BY id
        """
    ).fetchall()
    if not pending_rows:
        return

    catalog_rows = conn.execute(
        """
        SELECT ejercicio, ente_uid, ente_nombre
        FROM entes_detalle
        WHERE ente_uid IS NOT NULL AND TRIM(ente_uid) != ''
        """
    ).fetchall()
    name_map = {}
    sigla_map = {}
    for row in catalog_rows:
        ejercicio = str(row["ejercicio"]).strip()
        ente_uid = (row["ente_uid"] or "").strip()
        ente_nombre = (row["ente_nombre"] or "").strip()
        if not ejercicio or not ente_uid or not ente_nombre:
            continue
        name_key = normalize_text(ente_nombre)
        sigla_key = extract_sigla(ente_nombre)
        if name_key:
            name_map.setdefault((ejercicio, name_key), ente_uid)
        if sigla_key:
            sigla_map.setdefault((ejercicio, sigla_key), ente_uid)

    for row in pending_rows:
        row_id = row["id"]
        ejercicio = str(row["ejercicio"]).strip()
        ente_nombre = (row["ente"] or "").strip()
        if not ejercicio or not ente_nombre:
            continue
        ente_uid = ""
        sigla_key = extract_sigla(ente_nombre)
        if sigla_key:
            ente_uid = sigla_map.get((ejercicio, sigla_key), "")
        if not ente_uid:
            ente_uid = name_map.get((ejercicio, normalize_text(ente_nombre)), "")
        if not ente_uid:
            continue
        conn.execute(
            "UPDATE historial_titulares SET ente_uid = ? WHERE id = ?",
            (ente_uid, row_id),
        )



def get_current_user():
    username = session.get("user")
    if not username:
        return None
    user = USERS.get(username)
    if not user:
        return None
    return {"username": username, "role": user["role"]}


def is_luis_user(user: dict | None) -> bool:
    return bool(user and user.get("username") == LUIS_USERNAME)


def is_gabo_user(user: dict | None) -> bool:
    return bool(user and user.get("username") == GABO_USERNAME)


def home_endpoint_for_user(user: dict | None) -> str:
    if is_gabo_user(user):
        return "carga"
    return "index"


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if get_current_user() is None:
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)

    return wrapped


def luis_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        user = get_current_user()
        if user is None:
            return redirect(url_for("login", next=request.path))
        if not is_luis_user(user):
            return redirect(url_for("carga", notice="no_view_permission"))
        return view(*args, **kwargs)

    return wrapped


def gabo_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        user = get_current_user()
        if user is None:
            return redirect(url_for("login", next=request.path))
        if not is_gabo_user(user):
            return redirect(url_for("index", notice="no_load_permission"))
        return view(*args, **kwargs)

    return wrapped


def role_required(role: str):
    def decorator(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            user = get_current_user()
            if user is None:
                return redirect(url_for("login", next=request.path))
            if user["role"] != role:
                return redirect(url_for(home_endpoint_for_user(user), notice="no_permission"))
            return view(*args, **kwargs)

        return wrapped

    return decorator


def normalize_text(value: str) -> str:
    if not value:
        return ""
    normalized = unicodedata.normalize("NFKD", value)
    normalized = "".join(
        char for char in normalized if not unicodedata.combining(char)
    )
    normalized = re.sub(r"[^A-Za-z0-9]+", " ", normalized)
    return normalized.strip().lower()


def normalize_sigla(value: str) -> str:
    if not value:
        return ""
    normalized = unicodedata.normalize("NFKD", value)
    normalized = "".join(
        char for char in normalized if not unicodedata.combining(char)
    )
    normalized = re.sub(r"[^A-Za-z0-9]+", "", normalized)
    return normalized.upper()


def extract_sigla(value: str) -> str:
    if not value:
        return ""
    match = SIGLA_QUOTE_PATTERN.search(value)
    if not match:
        match = SIGLA_PAREN_PATTERN.search(value)
    if not match:
        return ""
    return normalize_sigla(match.group(1))


def next_ente_uid(current_max: int) -> tuple[int, str]:
    new_value = current_max + 1
    return new_value, f"{UID_PREFIX}{new_value:04d}"


def max_ente_uid(conn: sqlite3.Connection) -> int:
    rows = conn.execute(
        "SELECT ente_uid FROM entes_detalle WHERE ente_uid IS NOT NULL"
    ).fetchall()
    max_value = 0
    for row in rows:
        match = UID_PATTERN.match(row[0])
        if match:
            max_value = max(max_value, int(match.group(1)))
    return max_value


def resolve_ente_uid(conn: sqlite3.Connection, ente_nombre: str) -> str:
    sigla_key = extract_sigla(ente_nombre)
    name_key = normalize_text(ente_nombre)
    rows = conn.execute(
        "SELECT ente_uid, ente_nombre FROM entes_detalle WHERE ente_uid IS NOT NULL"
    ).fetchall()
    if sigla_key:
        for row in rows:
            if extract_sigla(row[1]) == sigla_key:
                return row[0]
    if name_key:
        for row in rows:
            if normalize_text(row[1]) == name_key:
                return row[0]
    max_value = max_ente_uid(conn)
    _, new_uid = next_ente_uid(max_value)
    return new_uid


def backfill_ente_uids(conn: sqlite3.Connection) -> None:
    rows = conn.execute(
        """
        SELECT id, ente_nombre, ente_uid
        FROM entes_detalle
        ORDER BY ejercicio, id
        """
    ).fetchall()
    max_value = max_ente_uid(conn)
    sigla_map = {}
    name_map = {}
    for row in rows:
        row_id, ente_nombre, ente_uid = row
        sigla_key = extract_sigla(ente_nombre)
        name_key = normalize_text(ente_nombre)
        if not ente_uid:
            if sigla_key and sigla_key in sigla_map:
                ente_uid = sigla_map[sigla_key]
            elif name_key and name_key in name_map:
                ente_uid = name_map[name_key]
            else:
                max_value, ente_uid = next_ente_uid(max_value)
            conn.execute(
                "UPDATE entes_detalle SET ente_uid = ? WHERE id = ?",
                (ente_uid, row_id),
            )
        if sigla_key:
            sigla_map.setdefault(sigla_key, ente_uid)
        if name_key:
            name_map.setdefault(name_key, ente_uid)


def init_db() -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS registros (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ejercicio TEXT NOT NULL,
                ente TEXT NOT NULL,
                ente_id TEXT,
                responsable TEXT NOT NULL,
                administrador TEXT NOT NULL,
                responsable_hist_id INTEGER,
                administrador_hist_id INTEGER,
                tipo_anexo TEXT NOT NULL,
                tipo_anexo_origen TEXT NOT NULL,
                monto_pdp REAL,
                estado TEXT NOT NULL,
                fuente_id INTEGER,
                irregularidad_id INTEGER,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS historial_titulares (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ejercicio INTEGER NOT NULL,
                ente_uid TEXT,
                ente TEXT NOT NULL,
                tipo_auditoria TEXT NOT NULL,
                nombre TEXT NOT NULL,
                cargo TEXT NOT NULL,
                fecha_inicio DATE NOT NULL,
                fecha_fin DATE NOT NULL,
                tipo_registro TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS entes_detalle (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ente_uid TEXT,
                ente_id TEXT NOT NULL,
                ejercicio TEXT NOT NULL,
                ente_numero TEXT NOT NULL,
                ente_nombre TEXT NOT NULL,
                responsable TEXT NOT NULL,
                clasificacion TEXT NOT NULL,
                ramo33 TEXT NOT NULL,
                created_at TEXT NOT NULL,
                UNIQUE(ente_id, ejercicio)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS fuentes_financiamiento (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL UNIQUE,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS entes_fuentes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ejercicio TEXT NOT NULL,
                ente_id TEXT NOT NULL,
                fuente_id INTEGER NOT NULL,
                tipo_auditoria TEXT NOT NULL,
                created_by TEXT,
                created_at TEXT NOT NULL,
                UNIQUE(ejercicio, ente_id, fuente_id, tipo_auditoria)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS catalogo_irregularidades (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                concepto TEXT NOT NULL UNIQUE,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS entes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ejercicio TEXT NOT NULL,
                nombre TEXT NOT NULL,
                created_at TEXT NOT NULL,
                UNIQUE(ejercicio, nombre)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS oficios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ente_id INTEGER NOT NULL,
                ejercicio TEXT NOT NULL,
                oficio TEXT NOT NULL,
                tipo_auditoria TEXT NOT NULL,
                fecha_notificacion TEXT NOT NULL,
                observaciones TEXT NOT NULL,
                fuente_id INTEGER,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS observaciones (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ejercicio TEXT NOT NULL,
                ente_id TEXT NOT NULL,
                ente_numero TEXT,
                ente_numero_sort REAL DEFAULT 0,
                ente_nombre TEXT NOT NULL,
                tipo_auditoria TEXT NOT NULL,
                fuente_financiamiento TEXT NOT NULL,
                ramo_33 TEXT NOT NULL,
                periodo_cedula TEXT,
                periodo_titular TEXT,
                oficio TEXT,
                fecha_notificacion TEXT,
                tipo_anexo TEXT NOT NULL,
                numero_observacion INTEGER NOT NULL,
                estado TEXT NOT NULL,
                reclasificada INTEGER NOT NULL DEFAULT 0,
                monto_pdp_emitido REAL,
                monto_pdp_solventado REAL,
                monto_pdp_pendiente REAL,
                pdp_no_irregularidad TEXT,
                pdp_concepto_irregularidad TEXT,
                pdp_subconcepto_irregularidad TEXT,
                pdp_irregularidad TEXT,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS cargas_manuales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ente_id TEXT NOT NULL,
                ente_nombre TEXT NOT NULL,
                tipo_auditoria TEXT NOT NULL,
                tipo_responsable TEXT NOT NULL,
                titular_nombre TEXT,
                administrativo_nombre TEXT,
                numero_oficio TEXT NOT NULL,
                asunto TEXT NOT NULL,
                ejercicio TEXT NOT NULL,
                fuente_id INTEGER NOT NULL,
                fuente_nombre TEXT,
                periodo TEXT NOT NULL,
                periodo_titular TEXT,
                fecha_notificacion TEXT,
                ramo_33 TEXT NOT NULL DEFAULT 'No',
                estado TEXT NOT NULL DEFAULT 'E',
                cantidad_sa INTEGER NOT NULL DEFAULT 0,
                cantidad_pdp INTEGER NOT NULL DEFAULT 0,
                cantidad_pras INTEGER NOT NULL DEFAULT 0,
                cantidad_pefcf INTEGER NOT NULL DEFAULT 0,
                cantidad_r INTEGER NOT NULL DEFAULT 0,
                monto_pdp_emitido REAL NOT NULL DEFAULT 0,
                monto_pdp_solventado REAL NOT NULL DEFAULT 0,
                monto_pdp_pendiente REAL NOT NULL DEFAULT 0,
                fuente_detalle_json TEXT,
                pdp_detalle_json TEXT,
                created_by TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS cargas_titulares (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ejercicio TEXT NOT NULL,
                ente_id TEXT NOT NULL,
                ente_nombre TEXT NOT NULL,
                tipo_auditoria TEXT NOT NULL,
                periodo_informe TEXT NOT NULL,
                titular TEXT NOT NULL,
                periodo_administrativo TEXT NOT NULL,
                administrativo TEXT NOT NULL,
                cedula_resultados TEXT NOT NULL,
                created_by TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )
        existing_columns = {
            row[1]
            for row in conn.execute("PRAGMA table_info(registros)").fetchall()
        }

        historial_columns = {
            row[1]
            for row in conn.execute("PRAGMA table_info(historial_titulares)").fetchall()
        }
        if "tipo_auditoria" not in historial_columns:
            # Needed to store separate titular/admin history per audit type
            # (e.g. "Financiera" vs "Obra Pública") without overwriting.
            conn.execute(
                "ALTER TABLE historial_titulares ADD COLUMN tipo_auditoria TEXT NOT NULL DEFAULT 'Financiera'"
            )
        if "ente_uid" not in historial_columns:
            conn.execute("ALTER TABLE historial_titulares ADD COLUMN ente_uid TEXT")
        cargas_manuales_columns = {
            row[1]
            for row in conn.execute("PRAGMA table_info(cargas_manuales)").fetchall()
        }
        if "ente_id" not in cargas_manuales_columns:
            conn.execute("ALTER TABLE cargas_manuales ADD COLUMN ente_id TEXT")
            conn.execute(
                """
                UPDATE cargas_manuales
                SET ente_id = ''
                WHERE ente_id IS NULL
                """
            )
        if "ente_nombre" not in cargas_manuales_columns:
            conn.execute("ALTER TABLE cargas_manuales ADD COLUMN ente_nombre TEXT")
            conn.execute(
                """
                UPDATE cargas_manuales
                SET ente_nombre = ''
                WHERE ente_nombre IS NULL
                """
            )
        if "fuente_nombre" not in cargas_manuales_columns:
            conn.execute("ALTER TABLE cargas_manuales ADD COLUMN fuente_nombre TEXT")
            conn.execute(
                """
                UPDATE cargas_manuales
                SET fuente_nombre = (
                    SELECT ff.nombre
                    FROM fuentes_financiamiento AS ff
                    WHERE ff.id = cargas_manuales.fuente_id
                    LIMIT 1
                )
                WHERE TRIM(COALESCE(fuente_nombre, '')) = ''
                """
            )
        if "periodo_titular" not in cargas_manuales_columns:
            conn.execute("ALTER TABLE cargas_manuales ADD COLUMN periodo_titular TEXT")
            conn.execute(
                """
                UPDATE cargas_manuales
                SET periodo_titular = ''
                WHERE periodo_titular IS NULL
                """
            )
        if "fecha_notificacion" not in cargas_manuales_columns:
            conn.execute("ALTER TABLE cargas_manuales ADD COLUMN fecha_notificacion TEXT")
            conn.execute(
                """
                UPDATE cargas_manuales
                SET fecha_notificacion = ''
                WHERE fecha_notificacion IS NULL
                """
            )
        if "ramo_33" not in cargas_manuales_columns:
            conn.execute("ALTER TABLE cargas_manuales ADD COLUMN ramo_33 TEXT NOT NULL DEFAULT 'No'")
        conn.execute(
            """
            UPDATE cargas_manuales
            SET ramo_33 = 'No'
            WHERE TRIM(COALESCE(ramo_33, '')) = ''
            """
        )
        if "estado" not in cargas_manuales_columns:
            conn.execute("ALTER TABLE cargas_manuales ADD COLUMN estado TEXT NOT NULL DEFAULT 'E'")
        conn.execute(
            """
            UPDATE cargas_manuales
            SET estado = 'E'
            WHERE TRIM(COALESCE(estado, '')) = ''
            """
        )
        if "fuente_detalle_json" not in cargas_manuales_columns:
            conn.execute("ALTER TABLE cargas_manuales ADD COLUMN fuente_detalle_json TEXT")
        if "pdp_detalle_json" not in cargas_manuales_columns:
            conn.execute("ALTER TABLE cargas_manuales ADD COLUMN pdp_detalle_json TEXT")
        if "tipo_anexo_origen" not in existing_columns:
            conn.execute(
                "ALTER TABLE registros ADD COLUMN tipo_anexo_origen TEXT"
            )
            conn.execute(
                """
                UPDATE registros
                SET tipo_anexo_origen = tipo_anexo
                WHERE tipo_anexo_origen IS NULL
                """
            )
        if "responsable" not in existing_columns:
            conn.execute("ALTER TABLE registros ADD COLUMN responsable TEXT")
            conn.execute(
                """
                UPDATE registros
                SET responsable = ''
                WHERE responsable IS NULL
                """
            )
        if "monto_pdp" not in existing_columns:
            conn.execute("ALTER TABLE registros ADD COLUMN monto_pdp REAL")
        if "estado" not in existing_columns:
            conn.execute("ALTER TABLE registros ADD COLUMN estado TEXT")
            conn.execute(
                """
                UPDATE registros
                SET estado = 'Pendiente'
                WHERE estado IS NULL
                """
            )
        if "administrador" not in existing_columns:
            conn.execute("ALTER TABLE registros ADD COLUMN administrador TEXT")
            conn.execute(
                """
                UPDATE registros
                SET administrador = ''
                WHERE administrador IS NULL
                """
            )
        if "ente_id" not in existing_columns:
            conn.execute("ALTER TABLE registros ADD COLUMN ente_id TEXT")
        if "fuente_id" not in existing_columns:
            conn.execute("ALTER TABLE registros ADD COLUMN fuente_id INTEGER")
        if "irregularidad_id" not in existing_columns:
            conn.execute("ALTER TABLE registros ADD COLUMN irregularidad_id INTEGER")
        if "responsable_hist_id" not in existing_columns:
            conn.execute("ALTER TABLE registros ADD COLUMN responsable_hist_id INTEGER")
        if "administrador_hist_id" not in existing_columns:
            conn.execute("ALTER TABLE registros ADD COLUMN administrador_hist_id INTEGER")
        oficios_columns = {
            row[1]
            for row in conn.execute("PRAGMA table_info(oficios)").fetchall()
        }
        if "oficio" not in oficios_columns:
            conn.execute("ALTER TABLE oficios ADD COLUMN oficio TEXT")
        if "tipo_auditoria" not in oficios_columns:
            conn.execute("ALTER TABLE oficios ADD COLUMN tipo_auditoria TEXT")
        if "fuente_id" not in oficios_columns:
            conn.execute("ALTER TABLE oficios ADD COLUMN fuente_id INTEGER")
        if "fecha_notificacion" not in oficios_columns:
            conn.execute("ALTER TABLE oficios ADD COLUMN fecha_notificacion TEXT")
        observaciones_columns = {
            row[1]
            for row in conn.execute("PRAGMA table_info(observaciones)").fetchall()
        }
        if "ente_numero" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN ente_numero TEXT")
        if "ente_numero_sort" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN ente_numero_sort REAL DEFAULT 0")
        conn.execute(
            """
            UPDATE observaciones
            SET ente_numero_sort = COALESCE(CAST(NULLIF(TRIM(ente_numero), '') AS REAL), 0)
            WHERE ente_numero_sort IS NULL OR ente_numero_sort = 0
            """
        )
        if "ente_nombre" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN ente_nombre TEXT")
            conn.execute(
                """
                UPDATE observaciones
                SET ente_nombre = ''
                WHERE ente_nombre IS NULL
                """
            )
        if "tipo_auditoria" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN tipo_auditoria TEXT")
        if "fuente_financiamiento" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN fuente_financiamiento TEXT")
        if "ramo_33" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN ramo_33 TEXT")
        if "periodo_cedula" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN periodo_cedula TEXT")
        if "periodo_titular" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN periodo_titular TEXT")
        if "oficio" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN oficio TEXT")
        if "fecha_notificacion" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN fecha_notificacion TEXT")
        if "tipo_anexo" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN tipo_anexo TEXT")
        if "numero_observacion" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN numero_observacion INTEGER")
        if "estado" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN estado TEXT")
            if "estatus" in observaciones_columns:
                conn.execute(
                    """
                    UPDATE observaciones
                    SET estado = estatus
                    WHERE estado IS NULL AND estatus IS NOT NULL
                    """
                )
        if "reclasificada" not in observaciones_columns:
            conn.execute(
                "ALTER TABLE observaciones ADD COLUMN reclasificada INTEGER NOT NULL DEFAULT 0"
            )
        if "monto_pdp_emitido" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN monto_pdp_emitido REAL")
        if "monto_pdp_solventado" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN monto_pdp_solventado REAL")
        if "monto_pdp_pendiente" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN monto_pdp_pendiente REAL")
        if "pdp_no_irregularidad" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN pdp_no_irregularidad TEXT")
        if "pdp_concepto_irregularidad" not in observaciones_columns:
            conn.execute(
                "ALTER TABLE observaciones ADD COLUMN pdp_concepto_irregularidad TEXT"
            )
        if "pdp_subconcepto_irregularidad" not in observaciones_columns:
            conn.execute(
                "ALTER TABLE observaciones ADD COLUMN pdp_subconcepto_irregularidad TEXT"
            )
        if "pdp_irregularidad" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN pdp_irregularidad TEXT")
        entes_detalle_columns = {
            row[1]
            for row in conn.execute("PRAGMA table_info(entes_detalle)").fetchall()
        }
        if "ente_uid" not in entes_detalle_columns:
            conn.execute("ALTER TABLE entes_detalle ADD COLUMN ente_uid TEXT")
        missing_uid = conn.execute(
            "SELECT COUNT(*) FROM entes_detalle WHERE ente_uid IS NULL"
        ).fetchone()[0]
        if missing_uid:
            backfill_ente_uids(conn)
        backfill_historial_ente_uids(conn)
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_obs_ejercicio
            ON observaciones (ejercicio)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_obs_ejercicio_ente
            ON observaciones (ejercicio, ente_id)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_obs_ejercicio_auditoria
            ON observaciones (ejercicio, tipo_auditoria)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_obs_ejercicio_anexo_estado
            ON observaciones (ejercicio, tipo_anexo, estado)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_obs_ejercicio_filtros
            ON observaciones (ejercicio, fuente_financiamiento, ramo_33, periodo_cedula)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_obs_ejercicio_full_scope
            ON observaciones (
                ejercicio,
                ente_id,
                tipo_auditoria,
                tipo_anexo,
                estado,
                fuente_financiamiento,
                ramo_33,
                periodo_cedula
            )
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_obs_ejercicio_sort
            ON observaciones (
                ejercicio,
                ente_numero_sort,
                ente_numero,
                ente_id,
                tipo_anexo,
                numero_observacion
            )
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_obs_pdp_concepto
            ON observaciones (pdp_concepto_irregularidad)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_obs_pdp_subconcepto
            ON observaciones (pdp_subconcepto_irregularidad)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_entes_ejercicio_id
            ON entes_detalle (ejercicio, ente_id)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_entes_uid
            ON entes_detalle (ente_uid)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_entes_fuentes_scope
            ON entes_fuentes (ejercicio, ente_id, tipo_auditoria, fuente_id)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_historial_scope
            ON historial_titulares (ejercicio, tipo_auditoria, tipo_registro)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_historial_uid
            ON historial_titulares (ente_uid, ejercicio)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_historial_ente
            ON historial_titulares (ente, ejercicio)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_cargas_manuales_usuario_fecha
            ON cargas_manuales (created_by, id DESC)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_cargas_titulares_usuario_fecha
            ON cargas_titulares (created_by, id DESC)
            """
        )
        conn.commit()


init_db()


def get_db() -> sqlite3.Connection:
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA busy_timeout = 5000")
        g.db.execute("PRAGMA temp_store = MEMORY")
    return g.db


@app.teardown_appcontext
def close_db(_exception: Exception | None) -> None:
    db = g.pop("db", None)
    if db is not None:
        db.close()


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip().lower()
        password = request.form.get("password", "").strip()
        next_url = (request.form.get("next") or "").strip()
        user = USERS.get(username)
        if not user or not check_password_hash(user["password_hash"], password):
            error = "Usuario o contraseña incorrectos."
        else:
            session.clear()
            session["user"] = username
            session["role"] = user["role"]
            user_payload = {"username": username, "role": user["role"]}
            if next_url:
                return redirect(next_url)
            return redirect(url_for(home_endpoint_for_user(user_payload)))
    else:
        current_user = get_current_user()
        if current_user is not None:
            return redirect(url_for(home_endpoint_for_user(current_user)))
        next_url = request.args.get("next", "")
    return render_template("login.html", error=error, next_url=next_url)


@app.get("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


def resolve_project_path(raw_path: str, *, must_exist: bool) -> Path:
    clean = (raw_path or "").strip()
    if not clean:
        raise ValueError("Debes indicar una ruta de archivo.")
    path = Path(clean)
    if not path.is_absolute():
        path = Path(BASE_DIR) / path
    resolved = path.resolve()
    base = Path(BASE_DIR).resolve()
    if resolved != base and base not in resolved.parents:
        raise ValueError("La ruta debe estar dentro del proyecto.")
    if must_exist and not resolved.exists():
        raise ValueError(f"No existe el archivo: {resolved}")
    return resolved


def run_loader_command(command: list[str]) -> dict:
    result = subprocess.run(
        command,
        cwd=BASE_DIR,
        capture_output=True,
        text=True,
        check=False,
    )
    return {
        "ok": result.returncode == 0,
        "returncode": result.returncode,
        "command": " ".join(command),
        "stdout": (result.stdout or "").strip(),
        "stderr": (result.stderr or "").strip(),
    }



ROUTE_DEPS = {
    "luis_required": luis_required,
    "gabo_required": gabo_required,
    "role_required": role_required,
    "get_current_user": get_current_user,
    "get_db": get_db,
    "normalize_ente_id": normalize_ente_id,
    "normalize_ente_id_sql": normalize_ente_id_sql,
    "normalize_tipo_auditoria": normalize_tipo_auditoria,
    "periodo_sql": periodo_sql,
    "parse_periodo_cedula": parse_periodo_cedula,
    "parse_historial_date": parse_historial_date,
    "get_ente_aliases_by_uid": get_ente_aliases_by_uid,
    "get_ente_uid_by_ejercicio_id": get_ente_uid_by_ejercicio_id,
    "parse_non_negative_int": parse_non_negative_int,
    "parse_non_negative_float": parse_non_negative_float,
    "parse_ente_numero_sort": parse_ente_numero_sort,
    "resolve_ente_uid": resolve_ente_uid,
    "resolve_project_path": resolve_project_path,
    "run_loader_command": run_loader_command,
    "home_endpoint_for_user": home_endpoint_for_user,
    "ASUNTOS_MANUALES": ASUNTOS_MANUALES,
    "TIPOS_RESPONSABLE": TIPOS_RESPONSABLE,
    "DB_PATH": DB_PATH,
    "BASE_DIR": BASE_DIR,
    "datetime": datetime,
    "sys": sys,
}

init_db()
register_gabo_routes(app, ROUTE_DEPS)
register_luis_routes(app, ROUTE_DEPS)

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5008)
