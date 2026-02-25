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


def user_can_view_data(user: dict | None) -> bool:
    return bool(user and user.get("role") in {"viewer", "editor"})


def user_can_load_data(user: dict | None) -> bool:
    return bool(user and user.get("role") in {"loader", "editor"})


def ensure_data_view_access():
    user = get_current_user()
    if not user_can_view_data(user):
        return redirect(url_for("carga", notice="no_view_permission"))
    return None


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if get_current_user() is None:
            return redirect(url_for("login", next=request.path))
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
                return redirect(url_for("index", notice="no_permission"))
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
        next_url = request.form.get("next") or url_for("index")
        user = USERS.get(username)
        if not user or not check_password_hash(user["password_hash"], password):
            error = "Usuario o contraseña incorrectos."
        else:
            session.clear()
            session["user"] = username
            session["role"] = user["role"]
            if user["role"] == "loader":
                return redirect(url_for("carga"))
            return redirect(next_url)
    else:
        if get_current_user() is not None:
            if get_current_user()["role"] == "loader":
                return redirect(url_for("carga"))
            return redirect(url_for("index"))
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


@app.route("/carga", methods=["GET", "POST"])
@login_required
def carga():
    user = get_current_user()
    if not user_can_load_data(user):
        return redirect(url_for("index", notice="no_permission"))

    result = None
    form_data = {
        "template_ejercicio": "",
        "template_out": "bases/historial_titulares_template.csv",
        "template_tipo_auditoria": "Financiera",
        "csv_path": "",
        "csv_ejercicio": "",
        "csv_replace": "0",
        "json_path": "",
        "json_ejercicio": "",
        "json_tipo_auditoria": "Financiera",
    }

    if request.method == "POST":
        action = (request.form.get("action") or "").strip()
        form_data.update(
            {
                "template_ejercicio": (request.form.get("template_ejercicio") or "").strip(),
                "template_out": (request.form.get("template_out") or "").strip() or "bases/historial_titulares_template.csv",
                "template_tipo_auditoria": (request.form.get("template_tipo_auditoria") or "").strip() or "Financiera",
                "csv_path": (request.form.get("csv_path") or "").strip(),
                "csv_ejercicio": (request.form.get("csv_ejercicio") or "").strip(),
                "csv_replace": "1" if request.form.get("csv_replace") else "0",
                "json_path": (request.form.get("json_path") or "").strip(),
                "json_ejercicio": (request.form.get("json_ejercicio") or "").strip(),
                "json_tipo_auditoria": (request.form.get("json_tipo_auditoria") or "").strip() or "Financiera",
            }
        )

        try:
            command = [sys.executable]
            if action == "template_generate":
                if not form_data["template_ejercicio"]:
                    raise ValueError("Debes indicar el ejercicio para generar plantilla.")
                out_path = resolve_project_path(form_data["template_out"], must_exist=False)
                out_path.parent.mkdir(parents=True, exist_ok=True)
                command.extend(
                    [
                        "scripts/make_historial_titulares_template.py",
                        "--db",
                        DB_PATH,
                        "--ejercicio",
                        str(int(form_data["template_ejercicio"])),
                        "--tipo-auditoria",
                        form_data["template_tipo_auditoria"],
                        "--out",
                        str(out_path),
                    ]
                )
            elif action in {"csv_validate", "csv_import"}:
                csv_path = resolve_project_path(form_data["csv_path"], must_exist=True)
                command.extend(
                    [
                        "scripts/import_historial_titulares.py",
                        "--db",
                        DB_PATH,
                        "--csv",
                        str(csv_path),
                    ]
                )
                if form_data["csv_ejercicio"]:
                    command.extend(["--ejercicio", str(int(form_data["csv_ejercicio"]))])
                if form_data["csv_replace"] == "1":
                    command.append("--replace")
                if action == "csv_validate":
                    command.append("--dry-run")
            elif action in {"json_validate", "json_import"}:
                json_path = resolve_project_path(form_data["json_path"], must_exist=True)
                command.extend(
                    [
                        "scripts/import_historial_titulares_json.py",
                        "--db",
                        DB_PATH,
                        "--json",
                        str(json_path),
                        "--tipo-auditoria",
                        form_data["json_tipo_auditoria"],
                    ]
                )
                if form_data["json_ejercicio"]:
                    command.extend(["--ejercicio", str(int(form_data["json_ejercicio"]))])
                if action == "json_validate":
                    command.append("--dry-run")
            else:
                raise ValueError("Acción de carga no soportada.")
            result = run_loader_command(command)
        except ValueError as exc:
            result = {
                "ok": False,
                "returncode": 1,
                "command": "",
                "stdout": "",
                "stderr": str(exc),
            }

    return render_template(
        "carga.html",
        user=user,
        result=result,
        form_data=form_data,
    )


@app.route("/entes", methods=["GET", "POST"])
@login_required
def entes():
    denied = ensure_data_view_access()
    if denied:
        return denied

    if request.method == "POST":
        if get_current_user()["role"] != "editor":
            return redirect(url_for("index", notice="no_permission"))
        ejercicio = request.form.get("ente_ejercicio", "").strip()
        ente_id = request.form.get("ente_id", "").strip()
        ente_numero = request.form.get("ente_numero", "").strip()
        ente_nombre = request.form.get("ente_nombre", "").strip()
        responsable = request.form.get("ente_responsable", "").strip()
        clasificacion = request.form.get("ente_clasificacion", "").strip()
        ramo33 = request.form.get("ente_ramo33", "").strip()

        if not all([ejercicio, ente_id, ente_numero, ente_nombre]):
            return redirect(url_for("index", notice="ente_error"))

        db = get_db()
        ente_uid = resolve_ente_uid(db, ente_nombre)
        db.execute(
            """
            INSERT INTO entes_detalle (
                ente_uid, ente_id, ejercicio, ente_numero, ente_nombre,
                responsable, clasificacion, ramo33, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(ente_id, ejercicio) DO UPDATE SET
                ente_uid = COALESCE(entes_detalle.ente_uid, excluded.ente_uid),
                ente_numero = excluded.ente_numero,
                ente_nombre = excluded.ente_nombre,
                responsable = excluded.responsable,
                clasificacion = excluded.clasificacion,
                ramo33 = excluded.ramo33
            """,
            (
                ente_uid,
                ente_id,
                ejercicio,
                ente_numero,
                ente_nombre,
                responsable or "",
                clasificacion or "",
                ramo33 or "",
                datetime.now().strftime("%Y-%m-%d %H:%M"),
            ),
        )
        db.commit()
        return redirect(url_for("index", notice="ente_saved"))

    ejercicio = request.args.get("ejercicio", "").strip()
    if not ejercicio:
        return jsonify([])

    db = get_db()
    rows = db.execute(
        """
        SELECT ente_id, ente_numero, ente_nombre, responsable, clasificacion, ramo33
        FROM entes_detalle
        WHERE ejercicio = ?
        ORDER BY CAST(ente_numero AS REAL) ASC, ente_numero ASC
        """,
        (ejercicio,),
    ).fetchall()
    return jsonify([dict(row) for row in rows])


@app.route("/historial", methods=["GET", "POST"])
@login_required
def historial():
    denied = ensure_data_view_access()
    if denied:
        return denied

    if request.method == "POST":
        if get_current_user()["role"] != "editor":
            return redirect(url_for("index", notice="no_permission"))
        ejercicio = request.form.get("historial_ejercicio", "").strip()
        ente_id = normalize_ente_id(request.form.get("historial_ente_id", ""))
        nombre = request.form.get("historial_nombre", "").strip()
        cargo = request.form.get("historial_cargo", "").strip()
        fecha_inicio = request.form.get("historial_fecha_inicio", "").strip()
        fecha_fin = request.form.get("historial_fecha_fin", "").strip()
        tipo_registro = request.form.get("historial_tipo_registro", "").strip()
        tipo_auditoria = request.form.get("historial_tipo_auditoria", "").strip() or "Financiera"

        if not all([ejercicio, ente_id, nombre, cargo, fecha_inicio, fecha_fin, tipo_registro]):
            return redirect(url_for("index", notice="historial_error"))

        db = get_db()
        ente_row = db.execute(
            f"""
            SELECT ente_nombre, ente_uid
            FROM entes_detalle
            WHERE ejercicio = ? AND {normalize_ente_id_sql('ente_id')} = ?
            """,
            (ejercicio, ente_id),
        ).fetchone()
        if ente_row is None:
            return redirect(url_for("index", notice="historial_error"))
        ente_nombre = ente_row["ente_nombre"]
        ente_uid = (ente_row["ente_uid"] or "").strip()
        db.execute(
            """
            INSERT INTO historial_titulares (
                ejercicio, ente_uid, ente, tipo_auditoria, nombre, cargo, fecha_inicio, fecha_fin, tipo_registro
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                int(ejercicio),
                ente_uid or None,
                ente_nombre,
                tipo_auditoria,
                nombre,
                cargo,
                fecha_inicio,
                fecha_fin,
                tipo_registro,
            ),
        )
        db.commit()
        return redirect(url_for("index", notice="historial_saved"))

    ejercicio = request.args.get("ejercicio", "").strip()
    ente_id = normalize_ente_id(request.args.get("ente_id", ""))
    if not ejercicio or not ente_id:
        return jsonify([])

    db = get_db()
    ente_info = db.execute(
        f"""
        SELECT ente_uid, ente_nombre
        FROM entes_detalle
        WHERE ejercicio = ? AND {normalize_ente_id_sql('ente_id')} = ?
        LIMIT 1
        """,
        (ejercicio, ente_id),
    ).fetchone()
    if not ente_info:
        return jsonify([])

    ente_uid = (ente_info["ente_uid"] or "").strip()
    ente_aliases = get_ente_aliases_by_uid(
        db,
        ejercicio,
        ente_id,
        fallback_names=[ente_info["ente_nombre"]],
    )

    filter_clause = ""
    filter_params = []
    if ente_uid and ente_aliases:
        placeholders = ", ".join(["?"] * len(ente_aliases))
        filter_clause = (
            f"AND (TRIM(COALESCE(ente_uid, '')) = ? OR TRIM(COALESCE(ente, '')) IN ({placeholders}))"
        )
        filter_params.extend([ente_uid, *ente_aliases])
    elif ente_uid:
        filter_clause = "AND TRIM(COALESCE(ente_uid, '')) = ?"
        filter_params.append(ente_uid)
    elif ente_aliases:
        placeholders = ", ".join(["?"] * len(ente_aliases))
        filter_clause = f"AND TRIM(COALESCE(ente, '')) IN ({placeholders})"
        filter_params.extend(ente_aliases)

    rows = db.execute(
        f"""
        SELECT id, nombre, cargo, fecha_inicio, fecha_fin, tipo_registro, tipo_auditoria
        FROM historial_titulares
        WHERE ejercicio = ?
        {filter_clause}
        ORDER BY id DESC
        """,
        [ejercicio, *filter_params],
    ).fetchall()
    return jsonify([dict(row) for row in rows])


@app.post("/oficios")
@login_required
def oficios():
    user = get_current_user()
    if not user or user["username"] != "omar":
        return redirect(url_for("index", notice="no_permission"))

    ejercicio = request.form.get("oficio_ejercicio", "").strip()
    ente_id = request.form.get("oficio_ente_id", "").strip()
    oficio_numero = request.form.get("oficio_numero", "").strip()
    tipo_auditoria = request.form.get("tipo_auditoria", "").strip()
    fecha_notificacion = request.form.get("fecha_notificacion", "").strip()
    fuente_id = request.form.get("oficio_fuente_id", "").strip()

    if not all(
        [
            ejercicio,
            ente_id,
            oficio_numero,
            tipo_auditoria,
            fecha_notificacion,
            fuente_id,
        ]
    ):
        return redirect(url_for("index", notice="oficio_error"))

    db = get_db()
    ente_row = db.execute(
        """
        SELECT ente_id
        FROM entes_detalle
        WHERE ente_id = ? AND ejercicio = ?
        """,
        (ente_id, ejercicio),
    ).fetchone()

    if ente_row is None:
        return redirect(url_for("index", notice="oficio_error"))

    fuente_row = db.execute(
        """
        SELECT 1
        FROM registros
        WHERE ejercicio = ? AND ente_id = ? AND fuente_id = ?
        LIMIT 1
        """,
        (ejercicio, ente_id, fuente_id),
    ).fetchone()
    if fuente_row is None:
        return redirect(url_for("index", notice="oficio_error"))

    db.execute(
        """
        INSERT INTO oficios (
            ente_id, ejercicio, oficio, tipo_auditoria, fecha_notificacion, observaciones,
            fuente_id, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            ente_id,
            ejercicio,
            oficio_numero,
            tipo_auditoria,
            fecha_notificacion,
            "",
            int(fuente_id),
            datetime.now().strftime("%Y-%m-%d %H:%M"),
        ),
    )
    db.commit()
    return redirect(url_for("index", notice="oficio_saved"))


@app.get("/fuentes-ente")
@login_required
def fuentes_ente():
    denied = ensure_data_view_access()
    if denied:
        return denied

    ejercicio = request.args.get("ejercicio", "").strip()
    ente_id = request.args.get("ente_id", "").strip()
    if not ejercicio or not ente_id:
        return jsonify([])

    db = get_db()
    rows = db.execute(
        """
        SELECT DISTINCT fuentes_financiamiento.id, fuentes_financiamiento.nombre
        FROM registros
        JOIN fuentes_financiamiento
            ON registros.fuente_id = fuentes_financiamiento.id
        WHERE registros.ejercicio = ? AND registros.ente_id = ?
        ORDER BY fuentes_financiamiento.nombre ASC
        """,
        (ejercicio, ente_id),
    ).fetchall()
    return jsonify([dict(row) for row in rows])


@app.get("/ejercicios-disponibles")
@login_required
def ejercicios_disponibles():
    denied = ensure_data_view_access()
    if denied:
        return denied

    db = get_db()
    observaciones_rows = db.execute(
        """
        SELECT ejercicio, COUNT(*) AS total
        FROM observaciones
        GROUP BY ejercicio
        """
    ).fetchall()
    entes_rows = db.execute(
        """
        SELECT ejercicio, COUNT(*) AS total
        FROM entes_detalle
        GROUP BY ejercicio
        """
    ).fetchall()

    resumen = {}
    for row in observaciones_rows:
        ejercicio = row["ejercicio"]
        resumen.setdefault(
            ejercicio,
            {"ejercicio": ejercicio, "total_observaciones": 0, "total_entes": 0},
        )
        resumen[ejercicio]["total_observaciones"] = row["total"]
    for row in entes_rows:
        ejercicio = row["ejercicio"]
        resumen.setdefault(
            ejercicio,
            {"ejercicio": ejercicio, "total_observaciones": 0, "total_entes": 0},
        )
        resumen[ejercicio]["total_entes"] = row["total"]

    ordered = sorted(resumen.values(), key=lambda item: item["ejercicio"], reverse=True)
    return jsonify(ordered)


@app.get("/observaciones")
@login_required
def observaciones_api():
    denied = ensure_data_view_access()
    if denied:
        return denied

    ejercicio = request.args.get("ejercicio", "").strip()
    ente_id = normalize_ente_id(request.args.get("ente_id", ""))
    tipo_anexo = request.args.get("tipo_anexo", "").strip()
    tipo_auditoria = normalize_tipo_auditoria(request.args.get("tipo_auditoria", ""))
    estado = request.args.get("estado", "").strip()
    fuente = request.args.get("fuente_financiamiento", "").strip()
    ramo_33 = request.args.get("ramo_33", "").strip()
    concepto_irregularidad = request.args.get("concepto_irregularidad", "").strip()
    periodo_informe = request.args.get("periodo_informe", "").strip()
    titular = request.args.get("titular", "").strip()
    periodo_admin = request.args.get("periodo_admin", "").strip()
    administrativo = request.args.get("administrativo", "").strip()
    periodo_cedula = request.args.get("periodo_cedula", "").strip()
    search = request.args.get("search", "").strip()
    if not ejercicio:
        return jsonify([])

    db = get_db()
    params = [ejercicio]
    filter_clauses = []
    if ente_id:
        filter_clauses.append(f"{normalize_ente_id_sql('observaciones.ente_id')} = ?")
        params.append(ente_id)
    if tipo_anexo:
        filter_clauses.append("observaciones.tipo_anexo = ?")
        params.append(tipo_anexo)
    if tipo_auditoria:
        filter_clauses.append("observaciones.tipo_auditoria = ?")
        params.append(tipo_auditoria)
    if estado:
        filter_clauses.append("observaciones.estado = ?")
        params.append(estado)
    if fuente:
        filter_clauses.append("observaciones.fuente_financiamiento = ?")
        params.append(fuente)
    if ramo_33:
        filter_clauses.append("observaciones.ramo_33 = ?")
        params.append(ramo_33)
    if concepto_irregularidad:
        filter_clauses.append(
            "(observaciones.pdp_concepto_irregularidad = ? OR observaciones.pdp_subconcepto_irregularidad = ?)"
        )
        params.extend([concepto_irregularidad, concepto_irregularidad])
    if periodo_cedula:
        filter_clauses.append("observaciones.periodo_cedula = ?")
        params.append(periodo_cedula)
    if search:
        filter_clauses.append(
            """
            (
                observaciones.ente_nombre LIKE ?
                OR observaciones.oficio LIKE ?
                OR observaciones.fuente_financiamiento LIKE ?
                OR observaciones.pdp_concepto_irregularidad LIKE ?
                OR observaciones.pdp_subconcepto_irregularidad LIKE ?
                OR CAST(observaciones.numero_observacion AS TEXT) LIKE ?
            )
            """
        )
        search_term = f"%{search}%"
        params.extend([search_term] * 6)

    filter_sql = ""
    if filter_clauses:
        filter_sql = " AND " + " AND ".join(filter_clauses)
    needs_historial_join = bool(periodo_informe or titular or periodo_admin or administrativo)

    if needs_historial_join:
        if periodo_informe:
            filter_clauses.append(f"{periodo_sql('resp')} = ?")
            params.append(periodo_informe)
        if titular:
            filter_clauses.append("resp.nombre = ?")
            params.append(titular)
        if periodo_admin:
            filter_clauses.append(f"{periodo_sql('admin')} = ?")
            params.append(periodo_admin)
        if administrativo:
            filter_clauses.append("admin.nombre = ?")
            params.append(administrativo)

        join_filter_sql = " AND " + " AND ".join(filter_clauses) if filter_clauses else ""
        rows = db.execute(
            f"""
            SELECT
                observaciones.id,
                observaciones.ejercicio,
                observaciones.ente_id,
                observaciones.ente_numero,
                observaciones.ente_nombre,
                observaciones.tipo_auditoria,
                observaciones.fuente_financiamiento,
                observaciones.ramo_33,
                observaciones.periodo_cedula,
                observaciones.periodo_titular,
                observaciones.oficio,
                observaciones.fecha_notificacion,
                observaciones.tipo_anexo,
                observaciones.numero_observacion,
                observaciones.estado,
                observaciones.monto_pdp_emitido,
                observaciones.monto_pdp_solventado,
                observaciones.monto_pdp_pendiente,
                observaciones.pdp_no_irregularidad,
                observaciones.pdp_concepto_irregularidad,
                observaciones.pdp_subconcepto_irregularidad
            FROM observaciones
            LEFT JOIN entes_detalle
                ON {normalize_ente_id_sql("observaciones.ente_id")} = {normalize_ente_id_sql("entes_detalle.ente_id")}
                AND observaciones.ejercicio = entes_detalle.ejercicio
            LEFT JOIN historial_titulares AS resp
                ON resp.id = (
                    SELECT id
                    FROM historial_titulares
                    WHERE ejercicio = observaciones.ejercicio
                      AND tipo_auditoria = observaciones.tipo_auditoria
                      AND (
                          (
                              TRIM(COALESCE(entes_detalle.ente_uid, '')) != ''
                              AND TRIM(COALESCE(historial_titulares.ente_uid, '')) = TRIM(entes_detalle.ente_uid)
                          )
                          OR ente = COALESCE(observaciones.ente_nombre, entes_detalle.ente_nombre)
                      )
                      AND tipo_registro = 'titular'
                    ORDER BY id DESC
                    LIMIT 1
                )
            LEFT JOIN historial_titulares AS admin
                ON admin.id = (
                    SELECT id
                    FROM historial_titulares
                    WHERE ejercicio = observaciones.ejercicio
                      AND tipo_auditoria = observaciones.tipo_auditoria
                      AND (
                          (
                              TRIM(COALESCE(entes_detalle.ente_uid, '')) != ''
                              AND TRIM(COALESCE(historial_titulares.ente_uid, '')) = TRIM(entes_detalle.ente_uid)
                          )
                          OR ente = COALESCE(observaciones.ente_nombre, entes_detalle.ente_nombre)
                      )
                      AND tipo_registro = 'director_administrativo'
                    ORDER BY id DESC
                    LIMIT 1
                )
            WHERE observaciones.ejercicio = ?
            {join_filter_sql}
            ORDER BY
                CAST(COALESCE(observaciones.ente_numero, '0') AS REAL) ASC,
                observaciones.ente_numero ASC,
                observaciones.ente_id ASC,
                observaciones.tipo_anexo ASC,
                observaciones.numero_observacion ASC
            """,
            params,
        ).fetchall()
    else:
        rows = db.execute(
            f"""
            SELECT
                observaciones.id,
                observaciones.ejercicio,
                observaciones.ente_id,
                observaciones.ente_numero,
                observaciones.ente_nombre,
                observaciones.tipo_auditoria,
                observaciones.fuente_financiamiento,
                observaciones.ramo_33,
                observaciones.periodo_cedula,
                observaciones.periodo_titular,
                observaciones.oficio,
                observaciones.fecha_notificacion,
                observaciones.tipo_anexo,
                observaciones.numero_observacion,
                observaciones.estado,
                observaciones.monto_pdp_emitido,
                observaciones.monto_pdp_solventado,
                observaciones.monto_pdp_pendiente,
                observaciones.pdp_no_irregularidad,
                observaciones.pdp_concepto_irregularidad,
                observaciones.pdp_subconcepto_irregularidad
            FROM observaciones
            WHERE observaciones.ejercicio = ?
            {filter_sql}
            ORDER BY
                CAST(COALESCE(observaciones.ente_numero, '0') AS REAL) ASC,
                observaciones.ente_numero ASC,
                observaciones.ente_id ASC,
                observaciones.tipo_anexo ASC,
                observaciones.numero_observacion ASC
            """,
            params,
        ).fetchall()

    return jsonify([dict(row) for row in rows])


@app.get("/observaciones-filtros")
@login_required
def observaciones_filtros():
    denied = ensure_data_view_access()
    if denied:
        return denied

    ejercicio = request.args.get("ejercicio", "").strip()
    ente_id = normalize_ente_id(request.args.get("ente_id", ""))
    tipo_auditoria = normalize_tipo_auditoria(request.args.get("tipo_auditoria", ""))
    tipo_anexo = request.args.get("tipo_anexo", "").strip()
    estado = request.args.get("estado", "").strip()
    fuente = request.args.get("fuente_financiamiento", "").strip()
    ramo_33 = request.args.get("ramo_33", "").strip()
    concepto_irregularidad = request.args.get("concepto_irregularidad", "").strip()
    periodo_cedula = request.args.get("periodo_cedula", "").strip()
    titular_seleccionado = request.args.get("titular", "").strip()
    administrativo_seleccionado = request.args.get("administrativo", "").strip()
    include_historial = request.args.get("include_historial", "").strip() == "1"
    if not ejercicio:
        return jsonify({})

    db = get_db()
    filtros = {}
    selected = {
        "tipo_auditoria": tipo_auditoria,
        "tipo_anexo": tipo_anexo,
        "estado": estado,
        "fuente_financiamiento": fuente,
        "ramo_33": ramo_33,
        "concepto_irregularidad": concepto_irregularidad,
        "periodo_cedula": periodo_cedula,
    }

    def build_observaciones_scope(exclude_key: str = "", include_ente: bool = True):
        clauses = ["ejercicio = ?"]
        params = [ejercicio]
        if include_ente and ente_id:
            clauses.append(f"{normalize_ente_id_sql('ente_id')} = ?")
            params.append(ente_id)

        for key, value in selected.items():
            if key == exclude_key or not value:
                continue
            if key == "tipo_auditoria":
                clauses.append("tipo_auditoria = ?")
                params.append(value)
            elif key == "tipo_anexo":
                clauses.append("tipo_anexo = ?")
                params.append(value)
            elif key == "estado":
                clauses.append("estado = ?")
                params.append(value)
            elif key == "fuente_financiamiento":
                clauses.append("fuente_financiamiento = ?")
                params.append(value)
            elif key == "ramo_33":
                clauses.append("ramo_33 = ?")
                params.append(value)
            elif key == "periodo_cedula":
                clauses.append("periodo_cedula = ?")
                params.append(value)
            elif key == "concepto_irregularidad":
                clauses.append("(pdp_concepto_irregularidad = ? OR pdp_subconcepto_irregularidad = ?)")
                params.extend([value, value])
        return " AND ".join(clauses), params

    def query_distinct(column: str, exclude_key: str):
        where_sql, where_params = build_observaciones_scope(exclude_key)
        return db.execute(
            f"""
            SELECT DISTINCT {column} AS value
            FROM observaciones
            WHERE {where_sql}
              AND {column} IS NOT NULL AND TRIM({column}) != ''
            ORDER BY value
            """,
            where_params,
        ).fetchall()

    auditorias = query_distinct("tipo_auditoria", "tipo_auditoria")
    tipos = query_distinct("tipo_anexo", "tipo_anexo")
    estados = query_distinct("estado", "estado")
    fuentes = query_distinct("fuente_financiamiento", "fuente_financiamiento")
    ramos = query_distinct("ramo_33", "ramo_33")
    cedulas = query_distinct("periodo_cedula", "periodo_cedula")
    concepto_where, concepto_params = build_observaciones_scope("concepto_irregularidad")
    conceptos = db.execute(
        f"""
        SELECT DISTINCT concepto
        FROM (
            SELECT pdp_concepto_irregularidad AS concepto
            FROM observaciones
            WHERE {concepto_where}
              AND pdp_concepto_irregularidad IS NOT NULL AND TRIM(pdp_concepto_irregularidad) != ''
            UNION
            SELECT pdp_subconcepto_irregularidad AS concepto
            FROM observaciones
            WHERE {concepto_where}
              AND pdp_subconcepto_irregularidad IS NOT NULL AND TRIM(pdp_subconcepto_irregularidad) != ''
        )
        ORDER BY concepto
        """,
        concepto_params + concepto_params,
    ).fetchall()
    entes_where, entes_params = build_observaciones_scope(include_ente=False)
    entes = db.execute(
        f"""
        SELECT DISTINCT
            TRIM(COALESCE(ente_id, '')) AS ente_id,
            TRIM(COALESCE(ente_numero, '')) AS ente_numero,
            TRIM(COALESCE(ente_nombre, '')) AS ente_nombre
        FROM observaciones
        WHERE {entes_where}
          AND TRIM(COALESCE(ente_id, '')) != ''
        ORDER BY CAST(COALESCE(NULLIF(ente_numero, ''), '0') AS REAL), ente_numero, ente_nombre
        """,
        entes_params,
    ).fetchall()
    periodos_informe = []
    titulares = []
    periodos_admin = []
    administrativos = []
    if include_historial:
        ente_aliases = get_ente_aliases_by_uid(db, ejercicio, ente_id)
        ente_uid = get_ente_uid_by_ejercicio_id(db, ejercicio, ente_id)

        titular_params = [ejercicio]
        titular_clause = ""
        if ente_uid and ente_aliases:
            placeholders = ", ".join(["?"] * len(ente_aliases))
            titular_clause = (
                f"AND (TRIM(COALESCE(ente_uid, '')) = ? OR TRIM(COALESCE(ente, '')) IN ({placeholders}))"
            )
            titular_params.extend([ente_uid, *ente_aliases])
        elif ente_uid:
            titular_clause = "AND TRIM(COALESCE(ente_uid, '')) = ?"
            titular_params.append(ente_uid)
        elif ente_aliases:
            placeholders = ", ".join(["?"] * len(ente_aliases))
            titular_clause = f"AND TRIM(COALESCE(ente, '')) IN ({placeholders})"
            titular_params.extend(ente_aliases)

        periodos_params = titular_params.copy()
        periodo_titular_clause = ""
        if titular_seleccionado:
            periodo_titular_clause = "AND nombre = ?"
            periodos_params.append(titular_seleccionado)

        titular_tipo_clause = ""
        titular_tipo_params = []
        if tipo_auditoria:
            titular_tipo_clause = "AND tipo_auditoria = ?"
            titular_tipo_params.append(tipo_auditoria)

        periodos_informe = db.execute(
            f"""
            SELECT DISTINCT {periodo_sql("historial_titulares")} AS periodo
            FROM historial_titulares
            WHERE ejercicio = ? {titular_clause} {periodo_titular_clause}
              {titular_tipo_clause}
              AND tipo_registro = 'titular'
              AND fecha_inicio IS NOT NULL AND fecha_fin IS NOT NULL
            ORDER BY fecha_inicio
            """,
            periodos_params + titular_tipo_params,
        ).fetchall()
        titulares = db.execute(
            f"""
            SELECT DISTINCT nombre
            FROM historial_titulares
            WHERE ejercicio = ? {titular_clause}
              {titular_tipo_clause}
              AND tipo_registro = 'titular'
              AND nombre IS NOT NULL AND nombre != ''
            ORDER BY nombre
            """,
            titular_params + titular_tipo_params,
        ).fetchall()

        admin_params = [ejercicio]
        admin_clause = ""
        if ente_uid and ente_aliases:
            placeholders = ", ".join(["?"] * len(ente_aliases))
            admin_clause = (
                f"AND (TRIM(COALESCE(ente_uid, '')) = ? OR TRIM(COALESCE(ente, '')) IN ({placeholders}))"
            )
            admin_params.extend([ente_uid, *ente_aliases])
        elif ente_uid:
            admin_clause = "AND TRIM(COALESCE(ente_uid, '')) = ?"
            admin_params.append(ente_uid)
        elif ente_aliases:
            placeholders = ", ".join(["?"] * len(ente_aliases))
            admin_clause = f"AND TRIM(COALESCE(ente, '')) IN ({placeholders})"
            admin_params.extend(ente_aliases)

        admin_periodos_params = admin_params.copy()
        periodo_admin_clause = ""
        if administrativo_seleccionado:
            periodo_admin_clause = "AND nombre = ?"
            admin_periodos_params.append(administrativo_seleccionado)

        admin_tipo_clause = ""
        admin_tipo_params = []
        if tipo_auditoria:
            admin_tipo_clause = "AND tipo_auditoria = ?"
            admin_tipo_params.append(tipo_auditoria)

        periodos_admin = db.execute(
            f"""
            SELECT DISTINCT {periodo_sql("historial_titulares")} AS periodo
            FROM historial_titulares
            WHERE ejercicio = ? {admin_clause} {periodo_admin_clause}
              {admin_tipo_clause}
              AND tipo_registro = 'director_administrativo'
              AND fecha_inicio IS NOT NULL AND fecha_fin IS NOT NULL
            ORDER BY fecha_inicio
            """,
            admin_periodos_params + admin_tipo_params,
        ).fetchall()
        administrativos = db.execute(
            f"""
            SELECT DISTINCT nombre
            FROM historial_titulares
            WHERE ejercicio = ? {admin_clause}
              {admin_tipo_clause}
              AND tipo_registro = 'director_administrativo'
              AND nombre IS NOT NULL AND nombre != ''
            ORDER BY nombre
            """,
            admin_params + admin_tipo_params,
        ).fetchall()
    filtros["tipo_anexo"] = [row[0] for row in tipos]
    filtros["tipo_auditoria"] = [row[0] for row in auditorias]
    filtros["entes"] = [dict(row) for row in entes]
    filtros["estado"] = [row[0] for row in estados]
    filtros["fuente_financiamiento"] = [row[0] for row in fuentes]
    filtros["ramo_33"] = [row[0] for row in ramos]
    filtros["conceptos_irregularidad"] = [row[0] for row in conceptos]
    filtros["periodo_informe"] = [row[0] for row in periodos_informe]
    filtros["titulares"] = [row[0] for row in titulares]
    filtros["periodo_admin"] = [row[0] for row in periodos_admin]
    filtros["administrativos"] = [row[0] for row in administrativos]
    filtros["cedulas"] = [row[0] for row in cedulas]

    return jsonify(filtros)


@app.get("/observaciones-responsables")
@login_required
def observaciones_responsables():
    denied = ensure_data_view_access()
    if denied:
        return denied

    ejercicio = request.args.get("ejercicio", "").strip()
    ente_id = normalize_ente_id(request.args.get("ente_id", ""))
    tipo_auditoria = normalize_tipo_auditoria(request.args.get("tipo_auditoria", ""))
    estado = request.args.get("estado", "").strip()
    fuente = request.args.get("fuente_financiamiento", "").strip()
    ramo_33 = request.args.get("ramo_33", "").strip()
    periodo_cedula = request.args.get("periodo_cedula", "").strip()
    if not ejercicio or not periodo_cedula:
        return jsonify([])

    db = get_db()
    filter_clauses = ["o.ejercicio = ?"]
    params = [ejercicio]
    if ente_id:
        filter_clauses.append(f"{normalize_ente_id_sql('o.ente_id')} = ?")
        params.append(ente_id)
    if tipo_auditoria:
        filter_clauses.append("o.tipo_auditoria = ?")
        params.append(tipo_auditoria)
    if estado:
        filter_clauses.append("o.estado = ?")
        params.append(estado)
    if fuente:
        filter_clauses.append("o.fuente_financiamiento = ?")
        params.append(fuente)
    if ramo_33:
        filter_clauses.append("o.ramo_33 = ?")
        params.append(ramo_33)
    if periodo_cedula:
        filter_clauses.append("o.periodo_cedula = ?")
        params.append(periodo_cedula)

    where_sql = " AND ".join(filter_clauses)
    observaciones_rows = db.execute(
        f"""
        SELECT DISTINCT
            o.ejercicio,
            o.ente_id,
            o.ente_nombre,
            ed.ente_nombre AS ente_detalle_nombre,
            ed.ente_uid AS ente_uid,
            o.tipo_auditoria,
            o.periodo_cedula
        FROM observaciones AS o
        LEFT JOIN entes_detalle AS ed
            ON {normalize_ente_id_sql("o.ente_id")} = {normalize_ente_id_sql("ed.ente_id")}
            AND o.ejercicio = ed.ejercicio
        WHERE {where_sql}
        ORDER BY o.ente_nombre ASC, o.tipo_auditoria ASC
        """,
        params,
    ).fetchall()

    resultado = []
    for row in observaciones_rows:
        cedula_inicio, cedula_fin = parse_periodo_cedula(row["ejercicio"], row["periodo_cedula"])
        if not cedula_inicio or not cedula_fin:
            continue
        cedula_inicio_date = parse_historial_date(cedula_inicio)
        cedula_fin_date = parse_historial_date(cedula_fin)
        if not cedula_inicio_date or not cedula_fin_date:
            continue

        ente_id_norm = normalize_ente_id(row["ente_id"])
        nombres_ente = get_ente_aliases_by_uid(
            db,
            row["ejercicio"],
            ente_id_norm,
            fallback_names=[row["ente_nombre"], row["ente_detalle_nombre"]],
        )
        ente_uid = (
            get_ente_uid_by_ejercicio_id(db, row["ejercicio"], ente_id_norm)
            or (row["ente_uid"] or "").strip()
        )
        scope_clause = ""
        scope_params = []
        if ente_uid and nombres_ente:
            placeholders = ", ".join(["?"] * len(nombres_ente))
            scope_clause = (
                f"AND (TRIM(COALESCE(h.ente_uid, '')) = ? OR TRIM(COALESCE(h.ente, '')) IN ({placeholders}))"
            )
            scope_params.extend([ente_uid, *nombres_ente])
        elif ente_uid:
            scope_clause = "AND TRIM(COALESCE(h.ente_uid, '')) = ?"
            scope_params.append(ente_uid)
        elif nombres_ente:
            placeholders = ", ".join(["?"] * len(nombres_ente))
            scope_clause = f"AND TRIM(COALESCE(h.ente, '')) IN ({placeholders})"
            scope_params.extend(nombres_ente)
        else:
            continue

        historial_rows = db.execute(
            f"""
            SELECT
                h.tipo_registro,
                h.tipo_auditoria,
                h.nombre,
                h.fecha_inicio,
                h.fecha_fin,
                {periodo_sql("h")} AS periodo
            FROM historial_titulares AS h
            WHERE h.ejercicio = ?
              {scope_clause}
              AND h.tipo_registro IN ('titular', 'director_administrativo')
              AND h.nombre IS NOT NULL AND h.nombre != ''
            ORDER BY h.tipo_registro ASC, h.fecha_inicio ASC, h.nombre ASC
            """,
            [row["ejercicio"], *scope_params],
        ).fetchall()

        titulares = []
        administrativos = []
        titulares_seen = set()
        administrativos_seen = set()
        for item in historial_rows:
            if normalize_tipo_auditoria(item["tipo_auditoria"] or "") != normalize_tipo_auditoria(row["tipo_auditoria"] or ""):
                continue
            inicio = parse_historial_date(item["fecha_inicio"])
            fin = parse_historial_date(item["fecha_fin"])
            if not inicio or not fin:
                continue
            # Inclusive overlap between [inicio, fin] and cedula range
            if inicio > cedula_fin_date or fin < cedula_inicio_date:
                continue
            payload = {
                "nombre": item["nombre"],
                "periodo": item["periodo"],
            }
            key = (payload["nombre"], payload["periodo"])
            if item["tipo_registro"] == "titular":
                if key in titulares_seen:
                    continue
                titulares_seen.add(key)
                titulares.append(payload)
            elif item["tipo_registro"] == "director_administrativo":
                if key in administrativos_seen:
                    continue
                administrativos_seen.add(key)
                administrativos.append(payload)

        resultado.append(
            {
                "ejercicio": row["ejercicio"],
                "ente_id": row["ente_id"],
                "ente_nombre": row["ente_nombre"] or row["ente_detalle_nombre"] or "—",
                "tipo_auditoria": row["tipo_auditoria"],
                "periodo_cedula": row["periodo_cedula"],
                "titulares": titulares,
                "administrativos": administrativos,
            }
        )

    return jsonify(resultado)


@app.get("/observaciones-exportar")
@login_required
def observaciones_exportar():
    denied = ensure_data_view_access()
    if denied:
        return denied

    ejercicio = request.args.get("ejercicio", "").strip()
    ente_id = normalize_ente_id(request.args.get("ente_id", ""))
    tipo_anexo = request.args.get("tipo_anexo", "").strip()
    tipo_auditoria = normalize_tipo_auditoria(request.args.get("tipo_auditoria", ""))
    estado = request.args.get("estado", "").strip()
    fuente = request.args.get("fuente_financiamiento", "").strip()
    ramo_33 = request.args.get("ramo_33", "").strip()
    concepto_irregularidad = request.args.get("concepto_irregularidad", "").strip()
    periodo_cedula = request.args.get("periodo_cedula", "").strip()
    if not ejercicio:
        return jsonify({"error": "ejercicio requerido"}), 400

    db = get_db()
    params = [ejercicio]
    filter_clauses = []
    if ente_id:
        filter_clauses.append(f"{normalize_ente_id_sql('observaciones.ente_id')} = ?")
        params.append(ente_id)
    if tipo_anexo:
        filter_clauses.append("observaciones.tipo_anexo = ?")
        params.append(tipo_anexo)
    if tipo_auditoria:
        filter_clauses.append("observaciones.tipo_auditoria = ?")
        params.append(tipo_auditoria)
    if estado:
        filter_clauses.append("observaciones.estado = ?")
        params.append(estado)
    if fuente:
        filter_clauses.append("observaciones.fuente_financiamiento = ?")
        params.append(fuente)
    if ramo_33:
        filter_clauses.append("observaciones.ramo_33 = ?")
        params.append(ramo_33)
    if concepto_irregularidad:
        filter_clauses.append(
            "(observaciones.pdp_concepto_irregularidad = ? OR observaciones.pdp_subconcepto_irregularidad = ?)"
        )
        params.extend([concepto_irregularidad, concepto_irregularidad])
    if periodo_cedula:
        filter_clauses.append("observaciones.periodo_cedula = ?")
        params.append(periodo_cedula)

    filter_sql = ""
    if filter_clauses:
        filter_sql = " AND " + " AND ".join(filter_clauses)

    rows = db.execute(
        f"""
        SELECT
            observaciones.ente_nombre,
            observaciones.tipo_anexo,
            observaciones.numero_observacion,
            observaciones.estado,
            observaciones.fecha_notificacion,
            observaciones.fuente_financiamiento,
            observaciones.pdp_concepto_irregularidad,
            observaciones.monto_pdp_emitido,
            observaciones.monto_pdp_solventado,
            observaciones.monto_pdp_pendiente,
            observaciones.tipo_auditoria
        FROM observaciones
        LEFT JOIN entes_detalle
            ON {normalize_ente_id_sql("observaciones.ente_id")} = {normalize_ente_id_sql("entes_detalle.ente_id")}
            AND observaciones.ejercicio = entes_detalle.ejercicio
        WHERE observaciones.ejercicio = ?
        {filter_sql}
        ORDER BY
            CAST(entes_detalle.ente_numero AS REAL) ASC,
            entes_detalle.ente_numero ASC,
            observaciones.ente_id ASC,
            observaciones.tipo_anexo ASC,
            observaciones.numero_observacion ASC
        """,
        params,
    ).fetchall()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Observaciones"

    headers = [
        "Ente",
        "Tipo auditoria",
        "Anexo",
        "No. Obs",
        "Estado",
        "Fecha",
        "Fuente",
        "Concepto de Irregularidad",
        "Monto emitido",
        "Monto solventado",
        "Monto pendiente",
    ]
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)

    total_observaciones = 0
    total_emitido = 0.0
    total_solventado = 0.0
    total_pendiente = 0.0
    conteo_pdp = 0
    conteo_pendiente = 0
    conteo_solventado = 0

    for row in rows:
        monto_emitido = float(row["monto_pdp_emitido"] or 0)
        monto_solventado = float(row["monto_pdp_solventado"] or 0)
        monto_pendiente = float(row["monto_pdp_pendiente"] or 0)
        total_observaciones += 1
        total_emitido += monto_emitido
        total_solventado += monto_solventado
        total_pendiente += monto_pendiente
        if (row["tipo_anexo"] or "") == "PDP":
            conteo_pdp += 1
        if (row["estado"] or "").strip().lower() == "pendiente":
            conteo_pendiente += 1
        if (row["estado"] or "").strip().lower() == "solventado":
            conteo_solventado += 1

        sheet.append(
            [
                row["ente_nombre"] or "—",
                row["tipo_auditoria"] or "—",
                row["tipo_anexo"] or "—",
                row["numero_observacion"] if row["numero_observacion"] is not None else "—",
                row["estado"] or "—",
                row["fecha_notificacion"] or "—",
                row["fuente_financiamiento"] or "—",
                row["pdp_concepto_irregularidad"] or "—",
                monto_emitido if (row["tipo_anexo"] or "") == "PDP" else 0,
                monto_solventado if (row["tipo_anexo"] or "") == "PDP" else 0,
                monto_pendiente if (row["tipo_anexo"] or "") == "PDP" else 0,
            ]
        )

    last_data_row = sheet.max_row
    for row_idx in range(2, last_data_row + 1):
        for col in (9, 10, 11):
            sheet.cell(row=row_idx, column=col).number_format = "#,##0.00"

    summary_start = sheet.max_row + 2
    sheet.cell(row=summary_start, column=1, value="Subtotal / Resumen").font = Font(bold=True)
    summary_rows = [
        ("Total observaciones", total_observaciones),
        ("Observaciones PDP", conteo_pdp),
        ("Observaciones pendientes", conteo_pendiente),
        ("Observaciones solventadas", conteo_solventado),
        ("Monto total emitido", total_emitido),
        ("Monto total solventado", total_solventado),
        ("Monto total pendiente", total_pendiente),
    ]
    for offset, (label, value) in enumerate(summary_rows, start=1):
        current_row = summary_start + offset
        sheet.cell(row=current_row, column=1, value=label).font = Font(bold=True)
        sheet.cell(row=current_row, column=2, value=value)
        if "Monto" in label:
            sheet.cell(row=current_row, column=2).number_format = "#,##0.00"

    for column_cells in sheet.columns:
        max_len = 0
        for cell in column_cells:
            val = "" if cell.value is None else str(cell.value)
            if len(val) > max_len:
                max_len = len(val)
        sheet.column_dimensions[column_cells[0].column_letter].width = min(max(12, max_len + 2), 50)

    stream = BytesIO()
    workbook.save(stream)
    stream.seek(0)
    filename = f"observaciones_{ejercicio}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/observaciones-stats")
@login_required
def observaciones_stats():
    denied = ensure_data_view_access()
    if denied:
        return denied

    ente_id = normalize_ente_id(request.args.get("ente_id", ""))
    tipo_auditoria = normalize_tipo_auditoria(request.args.get("tipo_auditoria", ""))
    where_clauses = []
    params = []
    if ente_id:
        where_clauses.append(f"{normalize_ente_id_sql('ente_id')} = ?")
        params.append(ente_id)
    if tipo_auditoria:
        where_clauses.append("tipo_auditoria = ?")
        params.append(tipo_auditoria)

    where_clause = f"WHERE {' AND '.join(where_clauses)}" if where_clauses else ""

    db = get_db()
    totals = db.execute(
        f"""
        SELECT ejercicio, COUNT(*) as total
        FROM observaciones
        {where_clause}
        GROUP BY ejercicio
        ORDER BY ejercicio
        """,
        params,
    ).fetchall()

    estados = db.execute(
        f"""
        SELECT ejercicio, estado, COUNT(*) as total
        FROM observaciones
        {where_clause}
        GROUP BY ejercicio, estado
        ORDER BY ejercicio
        """,
        params,
    ).fetchall()

    tipos = db.execute(
        f"""
        SELECT ejercicio, tipo_anexo, COUNT(*) as total
        FROM observaciones
        {where_clause}
        GROUP BY ejercicio, tipo_anexo
        ORDER BY ejercicio
        """,
        params,
    ).fetchall()

    tipos_estados = db.execute(
        f"""
        SELECT ejercicio, tipo_anexo, estado, COUNT(*) as total
        FROM observaciones
        {where_clause}
        GROUP BY ejercicio, tipo_anexo, estado
        ORDER BY ejercicio
        """,
        params,
    ).fetchall()

    fuentes = db.execute(
        f"""
        SELECT ejercicio, fuente_financiamiento, COUNT(*) as total
        FROM observaciones
        {where_clause}
        GROUP BY ejercicio, fuente_financiamiento
        ORDER BY ejercicio
        """,
        params,
    ).fetchall()

    pdp_where = "WHERE tipo_anexo = 'PDP'"
    pdp_clauses = []
    pdp_params = []
    if ente_id:
        pdp_clauses.append(f"{normalize_ente_id_sql('ente_id')} = ?")
        pdp_params.append(ente_id)
    if tipo_auditoria:
        pdp_clauses.append("tipo_auditoria = ?")
        pdp_params.append(tipo_auditoria)
    if pdp_clauses:
        pdp_where = f"WHERE {' AND '.join(pdp_clauses)} AND tipo_anexo = 'PDP'"

    pdp = db.execute(
        f"""
        SELECT
            ejercicio,
            SUM(COALESCE(monto_pdp_emitido, 0)) as emitido,
            SUM(COALESCE(monto_pdp_solventado, 0)) as solventado,
            SUM(COALESCE(monto_pdp_pendiente, 0)) as pendiente
        FROM observaciones
        {pdp_where}
        GROUP BY ejercicio
        ORDER BY ejercicio
        """,
        pdp_params,
    ).fetchall()

    return jsonify(
        {
            "totals": [dict(row) for row in totals],
            "estados": [dict(row) for row in estados],
            "tipos": [dict(row) for row in tipos],
            "tipos_estados": [dict(row) for row in tipos_estados],
            "fuentes": [dict(row) for row in fuentes],
            "pdp": [dict(row) for row in pdp],
        }
    )


@app.get("/catalogo-entes")
@login_required
def catalogo_entes():
    denied = ensure_data_view_access()
    if denied:
        return denied

    ejercicio = request.args.get("ejercicio", "").strip()
    if not ejercicio:
        return jsonify([])

    db = get_db()
    rows = db.execute(
        """
        SELECT
            entes_detalle.ente_id,
            entes_detalle.ente_uid,
            entes_detalle.ente_numero,
            entes_detalle.ente_nombre,
            entes_detalle.ejercicio,
            entes_detalle.clasificacion,
            entes_detalle.ramo33,
            entes_detalle.responsable,
            (
                SELECT ed.ente_nombre
                FROM entes_detalle AS ed
                WHERE COALESCE(ed.ente_uid, ed.ente_id)
                  = COALESCE(entes_detalle.ente_uid, entes_detalle.ente_id)
                  AND ed.ejercicio < entes_detalle.ejercicio
                ORDER BY ed.ejercicio DESC
                LIMIT 1
            ) AS nombre_anterior,
            (
                SELECT COUNT(DISTINCT ed.ente_nombre)
                FROM entes_detalle AS ed
                WHERE COALESCE(ed.ente_uid, ed.ente_id)
                  = COALESCE(entes_detalle.ente_uid, entes_detalle.ente_id)
            ) AS nombres_distintos
        FROM entes_detalle
        WHERE entes_detalle.ejercicio = ?
        ORDER BY CAST(entes_detalle.ente_numero AS REAL) ASC, entes_detalle.ente_numero ASC
        """,
        (ejercicio,),
    ).fetchall()

    return jsonify([dict(row) for row in rows])


@app.post("/fuentes")
@role_required("editor")
def fuentes():
    nombre = request.form.get("fuente_nombre", "").strip()
    if not nombre:
        return redirect(url_for("index", notice="fuente_error"))

    db = get_db()
    db.execute(
        """
        INSERT OR IGNORE INTO fuentes_financiamiento (nombre, created_at)
        VALUES (?, ?)
        """,
        (nombre, datetime.now().strftime("%Y-%m-%d %H:%M")),
    )
    db.commit()
    return redirect(url_for("index", notice="fuente_saved"))


@app.post("/irregularidades")
@role_required("editor")
def irregularidades():
    concepto = request.form.get("irregularidad_concepto", "").strip()
    if not concepto:
        return redirect(url_for("index", notice="irregularidad_error"))

    db = get_db()
    db.execute(
        """
        INSERT OR IGNORE INTO catalogo_irregularidades (concepto, created_at)
        VALUES (?, ?)
        """,
        (concepto, datetime.now().strftime("%Y-%m-%d %H:%M")),
    )
    db.commit()
    return redirect(url_for("index", notice="irregularidad_saved"))


@app.route("/stats")
@login_required
def stats():
    denied = ensure_data_view_access()
    if denied:
        return denied

    db = get_db()
    totals = db.execute(
        """
        SELECT ejercicio, COUNT(*) as total
        FROM registros
        GROUP BY ejercicio
        ORDER BY ejercicio
        """
    ).fetchall()

    estados = db.execute(
        """
        SELECT ejercicio, estado, COUNT(*) as total
        FROM registros
        GROUP BY ejercicio, estado
        ORDER BY ejercicio
        """
    ).fetchall()

    tipos = db.execute(
        """
        SELECT ejercicio, tipo_anexo, COUNT(*) as total
        FROM registros
        GROUP BY ejercicio, tipo_anexo
        ORDER BY ejercicio
        """
    ).fetchall()

    return jsonify(
        {
            "totals": [dict(row) for row in totals],
            "estados": [dict(row) for row in estados],
            "tipos": [dict(row) for row in tipos],
        }
    )


@app.post("/reclasificar/<int:registro_id>")
@role_required("editor")
def reclasificar(registro_id: int):
    db = get_db()
    row = db.execute(
        """
        SELECT id, tipo_anexo, tipo_anexo_origen
        FROM registros
        WHERE id = ?
        """,
        (registro_id,),
    ).fetchone()

    if row is None:
        return redirect(url_for("index", saved="0"))

    if row["tipo_anexo"] not in {"PDP", "PRAS"}:
        return redirect(url_for("index"))

    nuevo_tipo = "PRAS" if row["tipo_anexo"] == "PDP" else "PDP"
    origen = row["tipo_anexo_origen"] or row["tipo_anexo"]
    db.execute(
        """
        UPDATE registros
        SET tipo_anexo = ?, tipo_anexo_origen = ?
        WHERE id = ?
        """,
        (nuevo_tipo, origen, registro_id),
    )
    db.commit()
    return redirect(url_for("index", saved="1"))


@app.get("/")
@login_required
def index():
    user = get_current_user()
    if user and user.get("role") == "loader":
        return redirect(url_for("carga"))
    can_edit = user["role"] == "editor" if user else False
    return render_template(
        "index.html",
        user=user,
        can_edit=can_edit,
    )


if __name__ == "__main__":
    init_db()
    app.run(debug=True, host="0.0.0.0", port=5008)
