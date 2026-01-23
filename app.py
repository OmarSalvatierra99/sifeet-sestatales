from datetime import datetime
from functools import wraps
import os
import sqlite3

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
    "omar": {
        "password_hash": generate_password_hash("omar2025"),
        "role": "editor",
    },
}


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
                ente_id TEXT NOT NULL,
                ejercicio TEXT NOT NULL,
                rol TEXT NOT NULL,
                nombre TEXT NOT NULL,
                periodo TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS entes_detalle (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
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
        existing_columns = {
            row[1]
            for row in conn.execute("PRAGMA table_info(registros)").fetchall()
        }
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
        db.execute(
            """
            INSERT INTO entes_detalle (
                ente_id, ejercicio, ente_numero, ente_nombre,
                responsable, clasificacion, ramo33, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(ente_id, ejercicio) DO UPDATE SET
                ente_numero = excluded.ente_numero,
                ente_nombre = excluded.ente_nombre,
                responsable = excluded.responsable,
                clasificacion = excluded.clasificacion,
                ramo33 = excluded.ramo33
            """,
            (
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
        rol = request.form.get("historial_rol", "").strip()
        nombre = request.form.get("historial_nombre", "").strip()
        periodo = request.form.get("historial_periodo", "").strip()

        if not all([ejercicio, ente_id, rol, nombre, periodo]):
            return redirect(url_for("index", notice="historial_error"))

        db = get_db()
        db.execute(
            """
            INSERT INTO historial_titulares (
                ente_id, ejercicio, rol, nombre, periodo, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                ente_id,
                ejercicio,
                rol,
                nombre,
                periodo,
                datetime.now().strftime("%Y-%m-%d %H:%M"),
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
        SELECT id, rol, nombre, periodo
        FROM historial_titulares
        WHERE ejercicio = ? AND ente_id = ?
        ORDER BY id DESC
        """,
        (ejercicio, ente_id),
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


@app.route("/", methods=["GET", "POST"])
@login_required
def index():
    message = None
    message_type = "info"
    user = get_current_user()
    can_edit = user["role"] == "editor" if user else False
    is_omar = user["username"] == "omar" if user else False

    if request.method == "POST":
        if not can_edit:
            return redirect(url_for("index", notice="no_permission"))
        ejercicio = request.form.get("ejercicio", "").strip()
        ente_id = request.form.get("ente_id", "").strip()
        responsable_hist_id = request.form.get("responsable_hist_id", "").strip()
        administrador_hist_id = request.form.get("administrador_hist_id", "").strip()
        tipo_anexo = request.form.get("tipo_anexo", "").strip()
        monto_pdp = request.form.get("monto_pdp", "").strip()
        estado = request.form.get("estado", "").strip()
        fuente_nombre = request.form.get("fuente_nombre", "").strip()
        irregularidad_concepto = request.form.get("irregularidad_concepto", "").strip()

        if not ejercicio or not ente_id or not responsable_hist_id or not administrador_hist_id or not tipo_anexo or not estado:
            message = "Completa todos los campos para guardar el registro."
            message_type = "error"
        else:
            db = get_db()
            ente_row = db.execute(
                """
                SELECT ente_nombre
                FROM entes_detalle
                WHERE ente_id = ? AND ejercicio = ?
                """,
                (ente_id, ejercicio),
            ).fetchone()

            if ente_row is None:
                message = "El ID de ente no existe para el ejercicio seleccionado."
                message_type = "error"

            responsable_row = db.execute(
                """
                SELECT nombre, periodo
                FROM historial_titulares
                WHERE id = ? AND rol = 'Responsable' AND ente_id = ? AND ejercicio = ?
                """,
                (responsable_hist_id, ente_id, ejercicio),
            ).fetchone()

            administrador_row = db.execute(
                """
                SELECT nombre, periodo
                FROM historial_titulares
                WHERE id = ? AND rol = 'Administrador' AND ente_id = ? AND ejercicio = ?
                """,
                (administrador_hist_id, ente_id, ejercicio),
            ).fetchone()

            if responsable_row is None or administrador_row is None:
                message = "Selecciona responsable y administrador válidos para el ente."
                message_type = "error"

            fuente_id = None
            if fuente_nombre:
                db.execute(
                    """
                    INSERT OR IGNORE INTO fuentes_financiamiento (nombre, created_at)
                    VALUES (?, ?)
                    """,
                    (fuente_nombre, datetime.now().strftime("%Y-%m-%d %H:%M")),
                )
                fuente_id = db.execute(
                    "SELECT id FROM fuentes_financiamiento WHERE nombre = ?",
                    (fuente_nombre,),
                ).fetchone()["id"]

            irregularidad_id = None
            if tipo_anexo == "PDP" and irregularidad_concepto:
                db.execute(
                    """
                    INSERT OR IGNORE INTO catalogo_irregularidades (concepto, created_at)
                    VALUES (?, ?)
                    """,
                    (irregularidad_concepto, datetime.now().strftime("%Y-%m-%d %H:%M")),
                )
                irregularidad_id = db.execute(
                    "SELECT id FROM catalogo_irregularidades WHERE concepto = ?",
                    (irregularidad_concepto,),
                ).fetchone()["id"]

            monto_valor = None
            if tipo_anexo == "PDP":
                try:
                    monto_valor = float(monto_pdp.replace(",", "")) if monto_pdp else 0.0
                except ValueError:
                    message = "El monto PDP debe ser numérico."
                    message_type = "error"
                    monto_valor = None

            if message_type != "error":
                db.execute(
                    """
                    INSERT INTO registros (
                        ejercicio, ente, ente_id, responsable, administrador,
                        responsable_hist_id, administrador_hist_id, tipo_anexo,
                        tipo_anexo_origen, monto_pdp, estado, fuente_id,
                        irregularidad_id, created_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        ejercicio,
                        ente_row["ente_nombre"],
                        ente_id,
                        responsable_row["nombre"],
                        administrador_row["nombre"],
                        int(responsable_hist_id),
                        int(administrador_hist_id),
                        tipo_anexo,
                        tipo_anexo,
                        monto_valor,
                        estado,
                        fuente_id,
                        irregularidad_id,
                        datetime.now().strftime("%Y-%m-%d %H:%M"),
                    ),
                )
                db.execute(
                    """
                    INSERT OR IGNORE INTO entes (ejercicio, nombre, created_at)
                    VALUES (?, ?, ?)
                    """,
                    (
                    ejercicio,
                    ente_row["ente_nombre"],
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                ),
            )
            db.commit()
            return redirect(url_for("index", saved="1"))

    if request.args.get("saved") == "1":
        message = "Registro guardado correctamente."
        message_type = "success"
    elif request.args.get("notice") == "ente_saved":
        message = "Ente guardado correctamente."
        message_type = "success"
    elif request.args.get("notice") == "ente_error":
        message = "Completa todos los campos del ente."
        message_type = "error"
    elif request.args.get("notice") == "fuente_saved":
        message = "Fuente de financiamiento guardada."
        message_type = "success"
    elif request.args.get("notice") == "fuente_error":
        message = "Captura el nombre de la fuente."
        message_type = "error"
    elif request.args.get("notice") == "irregularidad_saved":
        message = "Concepto de irregularidad guardado."
        message_type = "success"
    elif request.args.get("notice") == "irregularidad_error":
        message = "Captura el concepto de irregularidad."
        message_type = "error"
    elif request.args.get("notice") == "historial_saved":
        message = "Titular guardado en historial."
        message_type = "success"
    elif request.args.get("notice") == "historial_error":
        message = "Completa todos los datos del historial."
        message_type = "error"
    elif request.args.get("notice") == "oficio_saved":
        message = "Oficio guardado correctamente."
        message_type = "success"
    elif request.args.get("notice") == "oficio_error":
        message = "Completa ejercicio, ente, oficio, tipo de auditoria, fecha y fuente."
        message_type = "error"
    elif request.args.get("notice") == "no_permission":
        message = "No tienes permisos para registrar informacion."
        message_type = "error"

    db = get_db()
    records = db.execute(
        """
        SELECT
            registros.id,
            registros.ejercicio,
            registros.ente,
            registros.ente_id,
            registros.responsable,
            registros.administrador,
            registros.tipo_anexo,
            registros.tipo_anexo_origen,
            registros.monto_pdp,
            registros.estado,
            registros.created_at,
            fuentes_financiamiento.nombre AS fuente_nombre,
            catalogo_irregularidades.concepto AS irregularidad_concepto,
            resp.periodo AS responsable_periodo,
            admin.periodo AS administrador_periodo
        FROM registros
        LEFT JOIN fuentes_financiamiento
            ON registros.fuente_id = fuentes_financiamiento.id
        LEFT JOIN catalogo_irregularidades
            ON registros.irregularidad_id = catalogo_irregularidades.id
        LEFT JOIN historial_titulares AS resp
            ON registros.responsable_hist_id = resp.id
        LEFT JOIN historial_titulares AS admin
            ON registros.administrador_hist_id = admin.id
        ORDER BY registros.id DESC
        """
    ).fetchall()

    fuentes = db.execute(
        """
        SELECT id, nombre
        FROM fuentes_financiamiento
        ORDER BY nombre ASC
        """
    ).fetchall()

    irregularidades_list = db.execute(
        """
        SELECT id, concepto
        FROM catalogo_irregularidades
        ORDER BY concepto ASC
        """
    ).fetchall()

    oficios = db.execute(
        """
        SELECT
            oficios.id,
            oficios.ejercicio,
            oficios.oficio,
            oficios.tipo_auditoria,
            oficios.fecha_notificacion,
            oficios.created_at,
            entes_detalle.ente_nombre AS ente_nombre,
            fuentes_financiamiento.nombre AS fuente_nombre
        FROM oficios
        LEFT JOIN entes_detalle
            ON oficios.ente_id = entes_detalle.ente_id
            AND oficios.ejercicio = entes_detalle.ejercicio
        LEFT JOIN fuentes_financiamiento
            ON oficios.fuente_id = fuentes_financiamiento.id
        ORDER BY oficios.id DESC
        """
    ).fetchall()

    return render_template(
        "index.html",
        records=records,
        fuentes=fuentes,
        irregularidades=irregularidades_list,
        oficios=oficios,
        message=message,
        message_type=message_type,
        user=user,
        can_edit=can_edit,
        is_omar=is_omar,
    )


if __name__ == "__main__":
    init_db()
    app.run(debug=True, host="0.0.0.0", port=5008)
