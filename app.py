from datetime import datetime
from functools import wraps
import os
import re
import sqlite3
import unicodedata

from flask import Flask, render_template, request, redirect, url_for, g, jsonify, session
from werkzeug.security import check_password_hash, generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "sifeet.db")

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "dev-secret-key")

USERS = {
    "luis": {
        "password_hash": generate_password_hash("luis2025"),
        "role": "viewer",
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



def get_current_user():
    username = session.get("user")
    if not username:
        return None
    user = USERS.get(username)
    if not user:
        return None
    return {"username": username, "role": user["role"]}


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
        conn.commit()


init_db()


def get_db() -> sqlite3.Connection:
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
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
            return redirect(next_url)
    else:
        if get_current_user() is not None:
            return redirect(url_for("index"))
        next_url = request.args.get("next", "")
    return render_template("login.html", error=error, next_url=next_url)


@app.get("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/entes", methods=["GET", "POST"])
@login_required
def entes():
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
    if request.method == "POST":
        if get_current_user()["role"] != "editor":
            return redirect(url_for("index", notice="no_permission"))
        ejercicio = request.form.get("historial_ejercicio", "").strip()
        ente_id = request.form.get("historial_ente_id", "").strip()
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
            """
            SELECT ente_nombre
            FROM entes_detalle
            WHERE ejercicio = ? AND ente_id = ?
            """,
            (ejercicio, ente_id),
        ).fetchone()
        if ente_row is None:
            return redirect(url_for("index", notice="historial_error"))
        ente_nombre = ente_row[0]
        db.execute(
            """
            INSERT INTO historial_titulares (
                ejercicio, ente, tipo_auditoria, nombre, cargo, fecha_inicio, fecha_fin, tipo_registro
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                int(ejercicio),
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
    ente_id = request.args.get("ente_id", "").strip()
    if not ejercicio or not ente_id:
        return jsonify([])

    db = get_db()
    rows = db.execute(
        """
        SELECT id, nombre, cargo, fecha_inicio, fecha_fin, tipo_registro, tipo_auditoria
        FROM historial_titulares
        WHERE ejercicio = ? AND ente = (
            SELECT ente_nombre FROM entes_detalle WHERE ejercicio = ? AND ente_id = ?
        )
        ORDER BY id DESC
        """,
        (ejercicio, ejercicio, ente_id),
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
    ejercicio = request.args.get("ejercicio", "").strip()
    ente_id = normalize_ente_id(request.args.get("ente_id", ""))
    tipo_anexo = request.args.get("tipo_anexo", "").strip()
    tipo_auditoria = request.args.get("tipo_auditoria", "").strip()
    estado = request.args.get("estado", "").strip()
    fuente = request.args.get("fuente_financiamiento", "").strip()
    ramo_33 = request.args.get("ramo_33", "").strip()
    oficio = request.args.get("oficio", "").strip()
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
    if oficio:
        filter_clauses.append("observaciones.oficio = ?")
        params.append(oficio)
    if periodo_cedula:
        filter_clauses.append("observaciones.periodo_cedula = ?")
        params.append(periodo_cedula)
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
            observaciones.pdp_subconcepto_irregularidad,
            entes_detalle.clasificacion,
            entes_detalle.responsable AS ente_responsable,
            resp.nombre AS responsable_nombre,
            {periodo_sql("resp")} AS responsable_periodo,
            admin.nombre AS administrador_nombre,
            {periodo_sql("admin")} AS administrador_periodo
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
                  AND ente = COALESCE(observaciones.ente_nombre, entes_detalle.ente_nombre)
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
                  AND ente = COALESCE(observaciones.ente_nombre, entes_detalle.ente_nombre)
                  AND tipo_registro = 'director_administrativo'
                ORDER BY id DESC
                LIMIT 1
            )
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

    return jsonify([dict(row) for row in rows])


@app.get("/observaciones-filtros")
@login_required
def observaciones_filtros():
    ejercicio = request.args.get("ejercicio", "").strip()
    ente_id = normalize_ente_id(request.args.get("ente_id", ""))
    tipo_auditoria = request.args.get("tipo_auditoria", "").strip()
    tipo_anexo = request.args.get("tipo_anexo", "").strip()
    estado = request.args.get("estado", "").strip()
    fuente = request.args.get("fuente_financiamiento", "").strip()
    ramo_33 = request.args.get("ramo_33", "").strip()
    oficio = request.args.get("oficio", "").strip()
    periodo_cedula = request.args.get("periodo_cedula", "").strip()
    titular_seleccionado = request.args.get("titular", "").strip()
    administrativo_seleccionado = request.args.get("administrativo", "").strip()
    if not ejercicio:
        return jsonify({})

    db = get_db()
    filtros = {}
    base_clauses = ["ejercicio = ?"]
    base_params = [ejercicio]
    if ente_id:
        base_clauses.append(f"{normalize_ente_id_sql('ente_id')} = ?")
        base_params.append(ente_id)
    if tipo_auditoria:
        base_clauses.append("tipo_auditoria = ?")
        base_params.append(tipo_auditoria)

    scoped_clauses = list(base_clauses)
    scoped_params = list(base_params)
    if tipo_anexo:
        scoped_clauses.append("tipo_anexo = ?")
        scoped_params.append(tipo_anexo)
    if estado:
        scoped_clauses.append("estado = ?")
        scoped_params.append(estado)
    if fuente:
        scoped_clauses.append("fuente_financiamiento = ?")
        scoped_params.append(fuente)
    if ramo_33:
        scoped_clauses.append("ramo_33 = ?")
        scoped_params.append(ramo_33)
    if oficio:
        scoped_clauses.append("oficio = ?")
        scoped_params.append(oficio)
    if periodo_cedula:
        scoped_clauses.append("periodo_cedula = ?")
        scoped_params.append(periodo_cedula)

    base_where = " AND ".join(base_clauses)
    scoped_where = " AND ".join(scoped_clauses)

    tipos = db.execute(
        f"""
        SELECT DISTINCT tipo_anexo
        FROM observaciones
        WHERE {base_where}
          AND tipo_anexo IS NOT NULL AND tipo_anexo != ''
        ORDER BY tipo_anexo
        """,
        base_params,
    ).fetchall()
    auditorias = db.execute(
        f"""
        SELECT DISTINCT tipo_auditoria
        FROM observaciones
        WHERE ejercicio = ?
          {('AND ' + normalize_ente_id_sql('ente_id') + ' = ?' if ente_id else '')}
          AND tipo_auditoria IS NOT NULL AND tipo_auditoria != ''
        ORDER BY tipo_auditoria
        """,
        [ejercicio, ente_id] if ente_id else [ejercicio],
    ).fetchall()
    estados = db.execute(
        f"""
        SELECT DISTINCT estado
        FROM observaciones
        WHERE {scoped_where}
          AND estado IS NOT NULL AND estado != ''
        ORDER BY estado
        """,
        scoped_params,
    ).fetchall()
    fuentes = db.execute(
        f"""
        SELECT DISTINCT fuente_financiamiento
        FROM observaciones
        WHERE {scoped_where}
          AND fuente_financiamiento IS NOT NULL AND fuente_financiamiento != ''
        ORDER BY fuente_financiamiento
        """,
        scoped_params,
    ).fetchall()
    ramos = db.execute(
        f"""
        SELECT DISTINCT ramo_33
        FROM observaciones
        WHERE {scoped_where}
          AND ramo_33 IS NOT NULL AND ramo_33 != ''
        ORDER BY ramo_33
        """,
        scoped_params,
    ).fetchall()
    oficios = db.execute(
        f"""
        SELECT DISTINCT oficio
        FROM observaciones
        WHERE {scoped_where}
          AND oficio IS NOT NULL AND oficio != ''
        ORDER BY oficio
        """,
        scoped_params,
    ).fetchall()
    ente_nombre = None
    if ente_id:
        ente_row = db.execute(
            """
            SELECT ente_nombre
            FROM entes_detalle
            WHERE ejercicio = ? AND ente_id = ?
            """,
            (ejercicio, ente_id),
        ).fetchone()
        if ente_row:
            ente_nombre = ente_row[0]

    titular_params = [ejercicio]
    titular_clause = ""
    if ente_nombre:
        titular_clause = "AND ente = ?"
        titular_params.append(ente_nombre)

    periodos_params = titular_params.copy()
    periodo_titular_clause = ""
    if titular_seleccionado:
        periodo_titular_clause = "AND nombre = ?"
        periodos_params.append(titular_seleccionado)

    periodos_informe = db.execute(
        f"""
        SELECT DISTINCT {periodo_sql("historial_titulares")} AS periodo
        FROM historial_titulares
        WHERE ejercicio = ? {titular_clause} {periodo_titular_clause}
          AND tipo_auditoria = ?
          AND tipo_registro = 'titular'
          AND fecha_inicio IS NOT NULL AND fecha_fin IS NOT NULL
        ORDER BY fecha_inicio
        """,
        periodos_params + [tipo_auditoria or "Financiera"],
    ).fetchall()
    titulares = db.execute(
        f"""
        SELECT DISTINCT nombre
        FROM historial_titulares
        WHERE ejercicio = ? {titular_clause}
          AND tipo_auditoria = ?
          AND tipo_registro = 'titular'
          AND nombre IS NOT NULL AND nombre != ''
        ORDER BY nombre
        """,
        titular_params + [tipo_auditoria or "Financiera"],
    ).fetchall()

    admin_params = [ejercicio]
    admin_clause = ""
    if ente_nombre:
        admin_clause = "AND ente = ?"
        admin_params.append(ente_nombre)

    admin_periodos_params = admin_params.copy()
    periodo_admin_clause = ""
    if administrativo_seleccionado:
        periodo_admin_clause = "AND nombre = ?"
        admin_periodos_params.append(administrativo_seleccionado)

    periodos_admin = db.execute(
        f"""
        SELECT DISTINCT {periodo_sql("historial_titulares")} AS periodo
        FROM historial_titulares
        WHERE ejercicio = ? {admin_clause} {periodo_admin_clause}
          AND tipo_auditoria = ?
          AND tipo_registro = 'director_administrativo'
          AND fecha_inicio IS NOT NULL AND fecha_fin IS NOT NULL
        ORDER BY fecha_inicio
        """,
        admin_periodos_params + [tipo_auditoria or "Financiera"],
    ).fetchall()
    administrativos = db.execute(
        f"""
        SELECT DISTINCT nombre
        FROM historial_titulares
        WHERE ejercicio = ? {admin_clause}
          AND tipo_auditoria = ?
          AND tipo_registro = 'director_administrativo'
          AND nombre IS NOT NULL AND nombre != ''
        ORDER BY nombre
        """,
        admin_params + [tipo_auditoria or "Financiera"],
    ).fetchall()
    cedulas = db.execute(
        f"""
        SELECT DISTINCT periodo_cedula
        FROM observaciones
        WHERE {scoped_where}
          AND periodo_cedula IS NOT NULL AND periodo_cedula != ''
        ORDER BY periodo_cedula
        """,
        scoped_params,
    ).fetchall()

    filtros["tipo_anexo"] = [row[0] for row in tipos]
    filtros["tipo_auditoria"] = [row[0] for row in auditorias]
    filtros["estado"] = [row[0] for row in estados]
    filtros["fuente_financiamiento"] = [row[0] for row in fuentes]
    filtros["ramo_33"] = [row[0] for row in ramos]
    filtros["oficios"] = [row[0] for row in oficios]
    filtros["periodo_informe"] = [row[0] for row in periodos_informe]
    filtros["titulares"] = [row[0] for row in titulares]
    filtros["periodo_admin"] = [row[0] for row in periodos_admin]
    filtros["administrativos"] = [row[0] for row in administrativos]
    filtros["cedulas"] = [row[0] for row in cedulas]

    return jsonify(filtros)


@app.get("/observaciones-stats")
@login_required
def observaciones_stats():
    ente_id = normalize_ente_id(request.args.get("ente_id", ""))
    where_clause = ""
    params = []
    if ente_id:
        where_clause = f"WHERE {normalize_ente_id_sql('ente_id')} = ?"
        params.append(ente_id)

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
    pdp_params = []
    if ente_id:
        pdp_where = f"WHERE {normalize_ente_id_sql('ente_id')} = ? AND tipo_anexo = 'PDP'"
        pdp_params.append(ente_id)

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
    can_edit = user["role"] == "editor" if user else False
    return render_template(
        "index.html",
        user=user,
        can_edit=can_edit,
    )


if __name__ == "__main__":
    init_db()
    app.run(debug=True, host="0.0.0.0", port=5008)
