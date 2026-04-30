from datetime import datetime
from functools import wraps
from io import BytesIO
import json
import os

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass
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
WORKSPACE_ROOT = Path(__file__).resolve().parent.parent
if str(WORKSPACE_ROOT) not in sys.path:
    sys.path.insert(0, str(WORKSPACE_ROOT))

from shared_user_catalog import build_user_map, get_display_name, ordered_users

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY")
template_reload_env = os.getenv("TEMPLATES_AUTO_RELOAD")
template_auto_reload = (
    template_reload_env.strip().lower() in {"1", "true", "yes", "on"}
    if template_reload_env
    else False
)
app.config["TEMPLATES_AUTO_RELOAD"] = template_auto_reload
app.jinja_env.auto_reload = template_auto_reload

PROJECT_KEY = "07-sifet-estatales"
LUIS_USERNAME = "luis"
GABO_USERNAME = "gabo"

def _build_users() -> dict:
    users: dict = {}
    for username, payload in build_user_map(PROJECT_KEY).items():
        role = payload["role"] or ("loader" if username == GABO_USERNAME else "viewer")
        users[username] = {
            "password_hash": generate_password_hash(payload["password"]),
            "role": role,
            "display_name": payload["display_name"],
        }
    return users

USERS = _build_users()
ASUNTOS_MANUALES = {
    "Notificación de Cédula de Resultados",
    "Resultados de Solventación",
}
TIPOS_RESPONSABLE = {"Titular", "Administrativo", "Ambos"}

UID_PREFIX = "ENT-"
UID_PATTERN = re.compile(rf"^{UID_PREFIX}(\\d+)$")
SIGLA_QUOTE_PATTERN = re.compile(r"[\"“”]([^\"“”]+)[\"“”]")
SIGLA_PAREN_PATTERN = re.compile(r"\\(([^)]+)\\)")
FUENTE_STOPWORDS = {
    "a",
    "al",
    "con",
    "de",
    "del",
    "el",
    "en",
    "la",
    "las",
    "los",
    "o",
    "para",
    "por",
    "sin",
    "y",
}
FUENTE_ACRONYMS = {
    "ASF",
    "CC",
    "CECYTE",
    "COBACH",
    "CONACYT",
    "CONADE",
    "CONASAMA",
    "CRESCA",
    "EMSAD",
    "FAM",
    "IB",
    "IMSS",
    "INEA",
    "INSABI",
    "ISR",
    "PAIBIM",
    "REA",
    "SANAS",
    "S200",
    "TLAX",
}
IRREGULARIDAD_ACRONYMS = {
    "CFDI",
}
IRREGULARIDAD_CANONICAL_RAW = tuple(
    line.strip()
    for line in """
DEUDORES DIVERSOS
PAGO A PROVEEDORES, PRESTADORES DE SERVICIOS Y/O CONTRATISTAS SIN ACREDITAR LA RECEPCIÓN DEL BIEN, SERVICIO U OBRA
OMISIÓN EN LA ENTREGA DE INFORMACIÓN
FALTA DE DISPOSICIONES ADMINISTRATIVAS
INCUMPLIMIENTO A LOS LINEAMIENTOS DE AUSTERIDAD DEL GASTO PÚBLICO DE LA GESTIÓN ADMINISTRATIVA
INCUMPLIMIENTO A LOS LINEAMIENTOS EN MATERIA DE ADQUISICIÓN DE BIENES Y CONTRATACIÓN DE SERVICIOS Y PAGOS RESPECTIVOS
OMISIÓN DE COMPROBANTE FISCAL DIGITAL POR INTERNET
INCORRECTO CONTROL DE REGISTROS CONTABLES Y PRESUPUESTALES
DEFICIENCIAS EN LA PUBLICACIÓN DE LAS OBLIGACIONES DE TRANSPARENCIA
INCONSISTENCIAS EN TRÁMITES, PROCESOS, SISTEMAS O EXPEDIENTES
INCONSISTENCIAS EN MATERIA DE SERVICIOS PERSONALES
COMPROBACIÓN CANCELADA CON LA LEYENDA OPERADO U OTRO
FALTA DE DOCUMENTACIÓN JUSTIFICATIVA
ANTICIPO A PROVEEDORES, PRESTADORES DE SERVICIOS, CONTRATISTAS POR OBRAS PÚBLICAS
OBLIGACIONES FINANCIERAS GENERADAS EN EJERCICIOS ANTERIORES SIN SER PAGADAS
INCONSISTENCIA EN LA VERIFICACIÓN Y RESGUARDO DE BIENES MUEBLES E INMUEBLES
DEFICIENTE CONTROL DE CUENTAS BANCARIAS
FALTA DE DISTRIBUCIÓN Y/O APLICACIÓN DE SUPERÁVIT
ACLARACIÓN DE PROCESOS ESPECÍFICOS
DIFERENCIAS EN EL ENTERO DE IMPUESTOS, CUOTAS O APORTACIONES
DIFERENCIAS EN LA MINISTRACIÓN DE RECURSOS
GASTOS PAGADOS SIN DOCUMENTACIÓN COMPROBATORIA
PAGO DE GASTOS IMPROCEDENTES
PAGO DE GASTOS EN EXCESO
PAGO DE BIENES Y/O SERVICIOS SIN ACREDITAR SU RECEPCIÓN Y/O APLICACIÓN
INGRESOS RECAUDADOS NO DEPOSITADOS
BIENES O APOYOS A PERSONAS O INSTITUCIONES NO PROPORCIONADOS
FALTANTE DE BIENES MUEBLES
OBLIGACIONES FINANCIERAS CONTRAÍDAS SIN LIQUIDEZ PARA PAGARLAS POR TERMINO DE ADMINISTRACIÓN
PAGO POR CONCEPTOS DE OBRA, INSUMOS, BIENES O SERVICIOS A PRECIOS SUPERIORES AL DE MERCADO
VOLÚMENES DE OBRA PAGADOS NO EJECUTADOS
CONCEPTOS DE OBRA PAGADOS NO EJECUTADOS
PROCESOS CONSTRUCTIVOS DEFICIENTES QUE CAUSAN AFECTACIONES FÍSICAS EN LAS OBRAS PÚBLICAS
PAGO DE OBRAS SIN ACREDITAR SU EXISTENCIA FÍSICA
PAGO DE CONCEPTOS QUE NO CUMPLEN CON LAS ESPECIFICACIONES TÉCNICAS CONTRATADAS
IMPUESTOS, CUOTAS Y DERECHOS RETENIDOS NO ENTERADOS SIN LIQUIDEZ
PENALIZACIÓN POR ATRASO EN LA EJECUCIÓN DE LOS TRABAJOS CON BASE A LAS FECHAS CONTRATADAS
INCUMPLIMIENTO A LA NORMATIVA
ADQUISICIONES A PRECIOS SUPERIORES A LOS DE MERCADO
INCONSISTENCIAS EN INFORMACIÓN FINANCIERA CONTABLE Y PRESUPUESTAL
APLICACIÓN DE RECURSOS NO AUTORIZADOS DE REMANENTES DE EJERCICIOS ANTERIORES
SOBREGIROS Y/O SUBEJERCICIOS PRESUPUESTALES
INCUMPLIMIENTO DE LA NORMATIVA EN MATERIA DE SERVICIOS PERSONALES
INCUMPLIMIENTO A CONDICIONES CONTRACTUALES
INADECUADA INTEGRACIÓN DE EXPEDIENTES
OBLIGACIONES FINANCIERAS CONTRAÍDAS SIN LIQUIDEZ PARA SU PAGO
OMISIÓN E INCONSISTENCIAS EN LOS CONTRATOS DE ADQUISICIONES, ARRENDAMIENTOS Y SERVICIOS/OBRA PÚBLICA
INCUMPLIMIENTO A LA NORMA POR EL ALTA Y BAJA DE BIENES MUEBLES E INMUEBLES
INCUMPLIMIENTO E INCONSISTENCIAS AL PROCEDIMIENTO DE ADJUDICACIÓN
FALTA DE CAPACIDAD FINANCIERA
INCONSISTENCIAS EN LOS EXPEDIENTES DE OBRA PÚBLICA
GASTOS SUPERIORES A LOS INGRESOS
OMISIÓN EN EL CUMPLIMIENTO DE LA LEY DE DISCIPLINA FINANCIERA
GASTO DEVENGADO SIN ACREDITAR BIENES Y SERVICIOS
COMPROBANTES FISCALES DIGITALES POR INTERNET CANCELADOS Y/O ALTERADOS
PROVEEDORES Y PRESTADORES DE SERVICIOS QUE NO CUENTAN CON LA ACTIVIDAD ECONÓMICA
COMPROBANTE QUE NO REÚNE LOS REQUISITOS FISCALES
DEFICIENTE CONTROL DE REGISTROS CONTABLES Y PRESUPUESTALES
DEFICIENCIAS DETECTADAS MEDIANTE APLICACIÓN DE CUESTIONARIO DE CONTROL INTERNO
OBLIGACIONES FINANCIERAS CONTRAÍDAS CON LIQUIDEZ PARA SU PAGO
INCONSISTENCIA EN LA IDENTIFICACIÓN DE BIENES MUEBLES EN EL INVENTARIO
FALTA DE SISTEMAS DE CONTROL INTERNO
PROCESOS DE CONTROL INTERNO DEFICIENTE / NO ACTUALIZADO
RESULTADOS DE EVALUACIÓN DE CONTROL INTERNO DEFICIENTES
FALTA DE EVALUACIÓN DE INDICADORES / METAS - OBJETIVOS
PAGO DE SUELDOS Y REMUNERACIONES POR SERVICIOS PERSONALES NO RECIBIDOS
OMISIÓN A REQUERIMIENTOS DE INFORMACIÓN
OMISIÓN AL PAGO DE OBLIGACIONES FINANCIERAS Y/O FISCALES
OMISIÓN EN LA ACTUALIZACIÓN Y CONCILIACIÓN FÍSICO-CONTABLE DE INVENTARIOS
INCOMPATIBILIDAD EN EL DESEMPEÑO DE EMPLEO, CARGO O COMISIÓN
OBRAS Y/O CONCEPTOS PAGADOS NO FISCALIZADOS POR OCULTAMIENTO DE DOCUMENTACIÓN COMPROBATORIA DE SU EJECUCIÓN
FORTALECIMIENTO EN MATERIA DE IGUALDAD, PERSPECTIVA DE GÉNERO, PREVENCIÓN Y ERRADICACIÓN DE LA VIOLENCIA EN CONTRA DE LAS MUJERES.
INCUMPLIMIENTO NORMATIVO EN MATERIA DE IGUALDAD, PERSPECTIVA DE GÉNERO, PREVENCIÓN Y ERRADICACIÓN DE LA VIOLENCIA EN CONTRA DE LAS MUJERES
""".strip().splitlines()
    if line.strip()
)
FUENTE_CANONICAL_OVERRIDES_RAW = {
    "CRÉDITOS OTORGADOS AUTORIZADOS": "Créditos Otorgados Autorizados",
    "Créditos Otorgados Autorizados": "Créditos Otorgados Autorizados",
    "REA: RECURSOS RECAUDADOS Y PARTICIPACIONES ESTATALES": "REA: Recursos Recaudados y Participaciones Estatales",
    "REA: Recursos Recaudados y Participaciones Estatales": "REA: Recursos Recaudados y Participaciones Estatales",
    "Participaciones estatales del (Fondo General de Participaciones)": "Participaciones Estatales (Fondo General de Participaciones)",
    "Participaciones Estatales (Ingresos Derivados de Fuentes Locales)": "Participaciones Estatales (Ingresos derivados de fuentes locales)",
    "Participaciones Estatales (Ingresos derivados de fuentes locales)": "Participaciones Estatales (Ingresos derivados de fuentes locales)",
    "Recursos Recaudados, Participaciones Estatales y Subsidio Federal para Organismos Descentralizados (EMSAD)": "Recursos recaudados, participaciones estatales y subsidio federal para organismos descentralizados (EMSAD)",
    "Recursos recaudados, participaciones estatales y subsidio federal para organismos descentralizados (EMSAD)": "Recursos recaudados, participaciones estatales y subsidio federal para organismos descentralizados (EMSAD)",
    "Remanentes de Ejercicios Anteriores": "Remanentes de ejercicios anteriores",
    "Remanentes de ejercicios anteriores": "Remanentes de ejercicios anteriores",
    "Convenio De Apoyo Financiero Para El Evento Olimpiada Nacional CONADE": "Convenio de Apoyo Financiero para el Evento Olimpiada Nacional CONADE",
    "Convenio Específico De Colaboración Para Operar El Programa Denominado Educación Para Adultos (INEA)": "Convenio Específico de Colaboración para Operar el Programa Denominado Educación para Adultos (INEA)",
    "Recursos Recaudados Y Convenio De Colaboración Para La Transferencia De Recursos Con El Organismo Público Descentralizado Salud De Tlaxcala": "Recursos Recaudados y Convenio de Colaboración para la Transferencia de Recursos con el Organismo Público Descentralizado Salud de Tlaxcala",
    "Remanentes De Ejercicios Anteriores: Recursos Recaudados": "Remanentes de Ejercicios Anteriores: Recursos Recaudados",
    "Remanentes De Ejercicios Anteriores: Recursos Recaudados Y Convenio De Colaboración Para La Transferencia De Recursos Con El Organismo Público Descentralizado Salud De Tlaxcala": "Remanentes de Ejercicios Anteriores: Recursos Recaudados y Convenio de Colaboración para la Transferencia de Recursos con el Organismo Público Descentralizado Salud de Tlaxcala",
    "Programa de Atención Integral para el Bienestar de las Mujeres (PAIBIM).": "Programa de Atención Integral para el Bienestar de las Mujeres (PAIBIM)",
}
FUENTE_WORD_PATTERN = re.compile(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9]+")

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


def ente_numero_sort_sql(column: str) -> str:
    clean = f"TRIM(COALESCE({column}, ''))"
    return (
        "CASE "
        f"WHEN {clean} = '' THEN 0 "
        f"WHEN INSTR({clean}, '.') > 0 THEN "
        f"CAST(SUBSTR({clean}, 1, INSTR({clean}, '.') - 1) AS REAL) * 1000 "
        f"+ CAST(SUBSTR({clean}, INSTR({clean}, '.') + 1) AS REAL) "
        f"ELSE CAST({clean} AS REAL) * 1000 "
        "END"
    )


def normalize_text_key(value: str) -> str:
    clean = (value or "").strip().lower()
    if not clean:
        return ""
    clean = unicodedata.normalize("NFKD", clean)
    clean = "".join(char for char in clean if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", clean).strip()


FUENTE_CANONICAL_OVERRIDES = {
    normalize_text_key(raw): canonical
    for raw, canonical in FUENTE_CANONICAL_OVERRIDES_RAW.items()
}


def _fuente_is_all_caps(value: str) -> bool:
    letters = [char for char in (value or "") if char.isalpha()]
    return bool(letters) and all(char == char.upper() for char in letters)


def _fuente_has_capitalized_stopwords(value: str) -> bool:
    capitalized = 0
    for word in re.findall(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+", value or ""):
        key = normalize_text_key(word)
        if key in FUENTE_STOPWORDS and word[:1].isupper():
            capitalized += 1
            if capitalized >= 2:
                return True
    return False


def _capitalize_fuente_token(word: str, *, capitalize_next: bool) -> str:
    if not word:
        return ""
    key = normalize_text_key(word)
    if not key:
        return word
    if word.isdigit():
        return word
    if word.upper() in FUENTE_ACRONYMS:
        return word.upper()
    if not capitalize_next and key in FUENTE_STOPWORDS:
        return word.lower()
    return word[:1].upper() + word[1:].lower()


def _smart_capitalize_fuente(value: str) -> str:
    clean = " ".join((value or "").replace("—", "-").replace("–", "-").split())
    if not clean:
        return ""
    parts: list[str] = []
    last_end = 0
    capitalize_next = True
    for match in FUENTE_WORD_PATTERN.finditer(clean):
        between = clean[last_end:match.start()]
        if any(marker in between for marker in ":([{"):
            capitalize_next = True
        parts.append(between)
        parts.append(
            _capitalize_fuente_token(
                match.group(0),
                capitalize_next=capitalize_next,
            )
        )
        last_end = match.end()
        capitalize_next = False
    parts.append(clean[last_end:])
    return "".join(parts)


def normalize_fuente_financiamiento(value: str) -> str:
    clean = " ".join((value or "").replace("—", "-").replace("–", "-").split())
    if not clean:
        return ""
    override = FUENTE_CANONICAL_OVERRIDES.get(normalize_text_key(clean))
    if override:
        return override
    if _fuente_is_all_caps(clean) or _fuente_has_capitalized_stopwords(clean):
        clean = _smart_capitalize_fuente(clean)
    return FUENTE_CANONICAL_OVERRIDES.get(normalize_text_key(clean), clean)


def is_remanente_fuente(value: str) -> bool:
    key = normalize_text_key(value)
    return (
        key.startswith("remanente")
        or key.startswith("remanentes")
        or key.startswith("rea:")
        or key.startswith("seguimiento")
    )


def normalize_origen_fuente(value: str) -> str:
    clean = " ".join((value or "").split())
    key = normalize_text_key(clean)
    if key in {"remanente", "remanentes"}:
        return "Remanentes"
    if key in {"del ejercicio", "ejercicio", "del_ejercicio"}:
        return "Del Ejercicio"
    return clean


def infer_origen_fuente(fuente_nombre: str, origen_fuente: str = "") -> str:
    origen = normalize_origen_fuente(origen_fuente)
    if origen in {"Del Ejercicio", "Remanentes"}:
        return origen
    return "Remanentes" if is_remanente_fuente(fuente_nombre) else "Del Ejercicio"


def _sentence_case_irregularidad(value: str) -> str:
    clean = " ".join((value or "").replace("—", "-").replace("–", "-").split()).rstrip(".").strip()
    if not clean:
        return ""
    sentence = clean.lower()
    first_alpha_index = next((index for index, char in enumerate(sentence) if char.isalpha()), None)
    if first_alpha_index is not None:
        sentence = (
            sentence[:first_alpha_index]
            + sentence[first_alpha_index].upper()
            + sentence[first_alpha_index + 1:]
        )
    for acronym in IRREGULARIDAD_ACRONYMS:
        sentence = re.sub(
            rf"\b{re.escape(acronym.lower())}\b",
            acronym,
            sentence,
            flags=re.IGNORECASE,
        )
    return sentence


def normalize_irregularidad_key(value: str) -> str:
    clean = normalize_text_key(value)
    if not clean:
        return ""
    clean = re.sub(r"^\d+\s*[-.)]?\s*", "", clean)
    clean = clean.rstrip(".").strip()
    clean = re.sub(r"[^a-z0-9]+", " ", clean)
    return re.sub(r"\s+", " ", clean).strip()


IRREGULARIDAD_CANONICAL = tuple(
    _sentence_case_irregularidad(item)
    for item in IRREGULARIDAD_CANONICAL_RAW
)
IRREGULARIDAD_CANONICAL_MAP = {
    normalize_irregularidad_key(item): item
    for item in IRREGULARIDAD_CANONICAL
}
IRREGULARIDAD_CANONICAL_ITEMS = tuple(
    sorted(
        IRREGULARIDAD_CANONICAL_MAP.items(),
        key=lambda item: len(item[0]),
        reverse=True,
    )
)


def _irregularidad_canonical(label: str) -> str:
    return IRREGULARIDAD_CANONICAL_MAP[normalize_irregularidad_key(label)]


def _irregularidad_has_any(key: str, *tokens: str) -> bool:
    return any(token in key for token in tokens)


def _irregularidad_match_family(key: str) -> str | None:
    for canonical_key, canonical_label in IRREGULARIDAD_CANONICAL_ITEMS:
        if key == canonical_key or key.startswith(canonical_key + " "):
            return canonical_label
    if "deudores diversos" in key:
        return _irregularidad_canonical("Deudores diversos")
    if "gastos a comprobar" in key or "recursos publicos faltantes" in key:
        return _irregularidad_canonical("Deudores diversos")
    if "requerimiento" in key and "informacion" in key:
        return _irregularidad_canonical("Omisión a requerimientos de información")
    if "entrega de informacion" in key:
        return _irregularidad_canonical("Omisión en la entrega de información")
    if "cfdi" in key and _irregularidad_has_any(key, "cancelad", "alterad"):
        return _irregularidad_canonical(
            "Comprobantes fiscales digitales por internet cancelados y/o alterados"
        )
    if "cfdi" in key or "comprobante fiscal digital por internet" in key:
        return _irregularidad_canonical("Omisión de comprobante fiscal digital por internet")
    if "requisitos fiscales" in key:
        return _irregularidad_canonical("Comprobante que no reúne los requisitos fiscales")
    if "gastos improcedentes" in key or "pago improcedente" in key or "pago duplicado" in key:
        return _irregularidad_canonical("Pago de gastos improcedentes")
    if "improcedente" in key and _irregularidad_has_any(key, "gasto", "pago", "servicio", "bienes y o servicios"):
        return _irregularidad_canonical("Pago de gastos improcedentes")
    if _irregularidad_has_any(
        key,
        "pago en exceso",
        "pagos en exceso",
        "pago de gastos en exceso",
        "superior a lo estipulado",
    ):
        return _irregularidad_canonical("Pago de gastos en exceso")
    if "mercado" in key:
        if "adquisicion" in key:
            return _irregularidad_canonical("Adquisiciones a precios superiores a los de mercado")
        if _irregularidad_has_any(key, "obra", "insumo", "bien", "servicio"):
            return _irregularidad_canonical(
                "Pago por conceptos de obra, insumos, bienes o servicios a precios superiores al de mercado"
            )
    if (
        _irregularidad_has_any(
            key,
            "diferencias volumetricas",
            "volumen faltante",
            "volumenes excedentes",
            "volumenes faltantes",
            "volumen pagado no ejecutado",
            "medidas superiores",
            "medidas menores",
        )
        or ("volumenes de obra" in key and _irregularidad_has_any(key, "faltante", "diferencias"))
    ):
        return _irregularidad_canonical("Volúmenes de obra pagados no ejecutados")
    if _irregularidad_has_any(
        key,
        "concepto no ejecutado",
        "conceptos no ejecutados",
        "trabajos no ejecutados",
        "medidas superiores no ejecutadas",
        "concepto de obra pagado no ejecutado",
        "concepto pagado no ejecutado",
        "concepto sin evidencia",
    ) or ("conceptos de obra" in key and "no ejecutados" in key):
        return _irregularidad_canonical("Conceptos de obra pagados no ejecutados")
    if "conceptos de obra" in key and _irregularidad_has_any(key, "faltantes", "sin evidencia de ejecucion"):
        return _irregularidad_canonical("Conceptos de obra pagados no ejecutados")
    if "procesos constructivos deficientes" in key:
        return _irregularidad_canonical(
            "Procesos constructivos deficientes que causan afectaciones físicas en las obras públicas"
        )
    if "mala calidad" in key:
        return _irregularidad_canonical(
            "Procesos constructivos deficientes que causan afectaciones físicas en las obras públicas"
        )
    if "pago de obras" in key and "sin acreditar" in key:
        return _irregularidad_canonical("Pago de obras sin acreditar su existencia física")
    if "especificaciones tecnicas" in key or "incumplimiento de especificaciones" in key:
        return _irregularidad_canonical(
            "Pago de conceptos que no cumplen con las especificaciones técnicas contratadas"
        )
    if "conceptos que no cumplen especificaciones" in key:
        return _irregularidad_canonical(
            "Pago de conceptos que no cumplen con las especificaciones técnicas contratadas"
        )
    if (
        _irregularidad_has_any(key, "bienes y o servicios", "pago de bienes", "pago de servicios")
        and _irregularidad_has_any(
            key,
            "sin acreditar",
            "sin evidencia",
            "sin justificar",
            "sin documentacion",
            "uso y destino",
        )
    ):
        return _irregularidad_canonical(
            "Pago de bienes y/o servicios sin acreditar su recepción y/o aplicación"
        )
    if _irregularidad_has_any(key, "pagos sin acreditar", "pago sin acreditar"):
        return _irregularidad_canonical(
            "Pago de bienes y/o servicios sin acreditar su recepción y/o aplicación"
        )
    if (
        _irregularidad_has_any(
            key,
            "adquisicion de",
            "compra de",
            "arrendamiento de",
            "renta de",
            "servicio de",
            "servicios de",
            "consumo de",
            "entrega no acreditada",
        )
        and _irregularidad_has_any(
            key,
            "sin acreditar",
            "sin evidencia",
            "sin documentacion",
            "no acreditada",
            "no acreditado",
            "sin presentar evidencia",
            "sin recepcion",
            "sin acreditar entrega",
        )
    ):
        return _irregularidad_canonical(
            "Pago de bienes y/o servicios sin acreditar su recepción y/o aplicación"
        )
    if "combustible" in key and _irregularidad_has_any(
        key,
        "bitacora",
        "bitacoras",
        "no se acredita suministro",
        "sin acreditar",
        "sin evidencia",
        "sin documentacion",
    ):
        return _irregularidad_canonical(
            "Pago de bienes y/o servicios sin acreditar su recepción y/o aplicación"
        )
    if (
        _irregularidad_has_any(key, "prestadores de servicios", "contratistas", "proveedores")
        and "sin acreditar" in key
        and _irregularidad_has_any(key, "bien servicio u obra", "recepcion del bien", "recepcion del servicio")
    ):
        return _irregularidad_canonical(
            "Pago a proveedores, prestadores de servicios y/o contratistas sin acreditar la recepción del bien, servicio u obra"
        )
    if _irregularidad_has_any(
        key,
        "sin documentacion justificativa",
        "sin documentacion comprobatoria",
        "sin evidencia",
    ):
        return _irregularidad_canonical("Falta de documentación justificativa")
    if _irregularidad_has_any(
        key,
        "omision de documentacion",
        "falta de documentacion",
        "sin presentar evidencia",
        "no justificados",
        "sin contrato",
    ):
        return _irregularidad_canonical("Falta de documentación justificativa")
    if "apoyo economico" in key and _irregularidad_has_any(
        key,
        "sin documentacion",
        "sin evidencia",
    ):
        return _irregularidad_canonical("Falta de documentación justificativa")
    if "gastos pagados" in key and "documentacion comprobatoria" in key:
        return _irregularidad_canonical("Gastos pagados sin documentación comprobatoria")
    if "bienes o apoyos" in key and _irregularidad_has_any(key, "no proporcionados", "no entregados"):
        return _irregularidad_canonical("Bienes o apoyos a personas o instituciones no proporcionados")
    if "uniformes" in key and _irregularidad_has_any(key, "no acreditada", "no proporcionados", "no entregados"):
        return _irregularidad_canonical("Bienes o apoyos a personas o instituciones no proporcionados")
    if "faltante de bienes muebles" in key:
        return _irregularidad_canonical("Faltante de bienes muebles")
    if "ingresos recaudados" in key and "no depositados" in key:
        return _irregularidad_canonical("Ingresos recaudados no depositados")
    if "ingresos no registrados" in key:
        return _irregularidad_canonical("Inconsistencias en información financiera contable y presupuestal")
    if "sueldos y remuneraciones" in key and "no recibidos" in key:
        return _irregularidad_canonical(
            "Pago de sueldos y remuneraciones por servicios personales no recibidos"
        )
    if "inasistencias" in key and _irregularidad_has_any(key, "sueldos", "remuneraciones"):
        return _irregularidad_canonical(
            "Pago de sueldos y remuneraciones por servicios personales no recibidos"
        )
    if "incompatibilidad" in key and _irregularidad_has_any(key, "empleo", "cargo", "comision"):
        return _irregularidad_canonical("Incompatibilidad en el desempeño de empleo, cargo o comisión")
    if "contrato" in key and _irregularidad_has_any(
        key,
        "prestacion de servicios",
        "servicios",
        "arrendamiento",
        "obra publica",
    ):
        return _irregularidad_canonical(
            "Omisión e inconsistencias en los contratos de adquisiciones, arrendamientos y servicios/obra pública"
        )
    if "obligaciones financieras" in key:
        if "con liquidez" in key and "pago" in key:
            return _irregularidad_canonical("Obligaciones financieras contraídas con liquidez para su pago")
        if "sin liquidez" in key and "termino de administracion" in key:
            return _irregularidad_canonical(
                "Obligaciones financieras contraídas sin liquidez para pagarlas por termino de administración"
            )
        if "sin liquidez" in key and "pago" in key:
            return _irregularidad_canonical("Obligaciones financieras contraídas sin liquidez para su pago")
    if "control interno" in key:
        if "cuestionario" in key:
            return _irregularidad_canonical(
                "Deficiencias detectadas mediante aplicación de cuestionario de control interno"
            )
        if "evaluacion" in key and "deficient" in key:
            return _irregularidad_canonical("Resultados de evaluación de control interno deficientes")
        if "sistemas" in key and "falta" in key:
            return _irregularidad_canonical("Falta de sistemas de control interno")
        if _irregularidad_has_any(key, "deficiente", "no actualizado"):
            return _irregularidad_canonical("Procesos de control interno deficiente / no actualizado")
    if "identificacion de bienes muebles" in key and "inventario" in key:
        return _irregularidad_canonical(
            "Inconsistencia en la identificación de bienes muebles en el inventario"
        )
    if "actualizacion y conciliacion fisico contable" in key and "inventario" in key:
        return _irregularidad_canonical(
            "Omisión en la actualización y conciliación físico-contable de inventarios"
        )
    if "verificacion" in key and "resguardo" in key and "bienes muebles e inmuebles" in key:
        return _irregularidad_canonical(
            "Inconsistencia en la verificación y resguardo de bienes muebles e inmuebles"
        )
    if "cuentas bancarias" in key:
        return _irregularidad_canonical("Deficiente control de cuentas bancarias")
    if "registros contables" in key and "presupuestales" in key:
        if "deficiente" in key:
            return _irregularidad_canonical("Deficiente control de registros contables y presupuestales")
        return _irregularidad_canonical("Incorrecto control de registros contables y presupuestales")
    if "informacion financiera" in key and "contable" in key and "presupuestal" in key:
        return _irregularidad_canonical("Inconsistencias en información financiera contable y presupuestal")
    if "remanentes de ejercicios anteriores" in key and "no autorizados" in key:
        return _irregularidad_canonical(
            "Aplicación de recursos no autorizados de remanentes de ejercicios anteriores"
        )
    if _irregularidad_has_any(key, "sobregiro", "subejercicio"):
        return _irregularidad_canonical("Sobregiros y/o subejercicios presupuestales")
    if _irregularidad_has_any(
        key,
        "personal no autorizado",
        "prestacion no autorizada",
        "bonos por proceso electoral",
        "bonos",
        "apoyo bimestral",
        "apoyo trimestral",
        "complemento de sueldo",
        "compensacion garantizada",
        "medidas del bienestar",
    ):
        return _irregularidad_canonical("Incumplimiento de la normativa en materia de servicios personales")
    if key in {
        "arrendamiento de pension vehicular",
        "servicio de arrendamiento de impresoras",
    }:
        return _irregularidad_canonical("Aclaración de procesos específicos")
    return None


def normalize_irregularidad_concepto(
    value: str,
    *,
    allow_blank: bool = False,
    strict: bool = False,
) -> str:
    clean = " ".join((value or "").replace("—", "-").replace("–", "-").split())
    clean = re.sub(r"^\s*\d+\s*[-.)]?\s*", "", clean).rstrip(".").strip()
    if not clean:
        if allow_blank:
            return ""
        if strict:
            raise ValueError("Debes capturar un concepto de irregularidad.")
        return ""
    key = normalize_irregularidad_key(clean)
    if key in {"0", "-"}:
        if allow_blank:
            return ""
        if strict:
            raise ValueError("Debes capturar un concepto de irregularidad.")
        return ""
    canonical = IRREGULARIDAD_CANONICAL_MAP.get(key) or _irregularidad_match_family(key)
    if canonical:
        return canonical
    if strict:
        raise ValueError(
            f"Concepto de irregularidad no homologado: '{clean}'. Usa el catálogo canónico."
        )
    return _sentence_case_irregularidad(clean)


def normalize_irregularidad_subconcepto(value: str) -> str:
    clean = " ".join((value or "").replace("—", "-").replace("–", "-").split())
    clean = re.sub(r"^\s*\d+\s*[-.)]?\s*", "", clean).rstrip(".").strip()
    if normalize_irregularidad_key(clean) in {"", "0", "-", "na", "n a", "s n"}:
        return ""
    return clean


def _normalize_fuente_snapshot_json(raw_value: str, *, fields: tuple[str, ...]) -> str:
    clean = (raw_value or "").strip()
    if not clean:
        return raw_value
    try:
        payload = json.loads(clean)
    except json.JSONDecodeError:
        return raw_value

    changed = False
    items = payload if isinstance(payload, list) else [payload] if isinstance(payload, dict) else []
    for item in items:
        if not isinstance(item, dict):
            continue
        for field in fields:
            current = " ".join(str(item.get(field) or "").split())
            normalized = normalize_fuente_financiamiento(current)
            if current != normalized:
                item[field] = normalized
                changed = True
    if not changed:
        return raw_value
    return json.dumps(payload, ensure_ascii=False)


def normalize_fuentes_data(conn: sqlite3.Connection) -> None:
    original_row_factory = conn.row_factory
    conn.row_factory = sqlite3.Row
    try:
        catalog_rows = conn.execute(
            """
            SELECT id, TRIM(COALESCE(nombre, '')) AS nombre
            FROM fuentes_financiamiento
            ORDER BY id ASC
            """
        ).fetchall()
        catalog_groups: dict[str, dict[str, object]] = {}
        replacement_map: dict[int, int] = {}

        for row in catalog_rows:
            fuente_id = int(row["id"])
            canonical_name = normalize_fuente_financiamiento(row["nombre"] or "")
            if not canonical_name:
                continue
            group_key = normalize_text_key(canonical_name)
            group = catalog_groups.setdefault(
                group_key,
                {"canonical": canonical_name, "rows": []},
            )
            group["rows"].append(
                {
                    "id": fuente_id,
                    "nombre": (row["nombre"] or "").strip(),
                }
            )

        for group in catalog_groups.values():
            canonical_name = str(group["canonical"])
            rows = list(group["rows"])
            exact_matches = [
                row for row in rows
                if " ".join((row["nombre"] or "").split()) == canonical_name
            ]
            survivor = min(exact_matches or rows, key=lambda item: int(item["id"]))
            survivor_id = int(survivor["id"])

            if (survivor["nombre"] or "").strip() != canonical_name:
                conn.execute(
                    "UPDATE fuentes_financiamiento SET nombre = ? WHERE id = ?",
                    (canonical_name, survivor_id),
                )

            for row in rows:
                row_id = int(row["id"])
                if row_id == survivor_id:
                    continue
                replacement_map[row_id] = survivor_id

        for source_id, target_id in replacement_map.items():
            conn.execute(
                """
                INSERT OR IGNORE INTO entes_fuentes (
                    ejercicio,
                    ente_id,
                    fuente_id,
                    tipo_auditoria,
                    created_by,
                    created_at
                )
                SELECT
                    ejercicio,
                    ente_id,
                    ?,
                    tipo_auditoria,
                    created_by,
                    created_at
                FROM entes_fuentes
                WHERE fuente_id = ?
                """,
                (target_id, source_id),
            )
            conn.execute("DELETE FROM entes_fuentes WHERE fuente_id = ?", (source_id,))
            for table_name in ("registros", "oficios", "cargas_manuales"):
                conn.execute(
                    f"UPDATE {table_name} SET fuente_id = ? WHERE fuente_id = ?",
                    (target_id, source_id),
                )
            conn.execute("DELETE FROM fuentes_financiamiento WHERE id = ?", (source_id,))

        catalog_lookup = {
            int(row["id"]): normalize_fuente_financiamiento(row["nombre"] or "")
            for row in conn.execute(
                """
                SELECT id, TRIM(COALESCE(nombre, '')) AS nombre
                FROM fuentes_financiamiento
                ORDER BY id ASC
                """
            ).fetchall()
        }

        cargas_rows = conn.execute(
            """
            SELECT
                id,
                fuente_id,
                TRIM(COALESCE(fuente_nombre, '')) AS fuente_nombre,
                TRIM(COALESCE(estado, '')) AS estado,
                fuente_detalle_json,
                pdp_detalle_json
            FROM cargas_manuales
            ORDER BY id ASC
            """
        ).fetchall()
        for row in cargas_rows:
            carga_id = int(row["id"])
            fuente_id = row["fuente_id"]
            fuente_id_int = int(fuente_id) if fuente_id is not None else None
            canonical_name = (
                catalog_lookup.get(fuente_id_int)
                if fuente_id_int is not None and fuente_id_int > 0
                else normalize_fuente_financiamiento(row["fuente_nombre"] or "")
            ) or normalize_fuente_financiamiento(row["fuente_nombre"] or "")
            fuente_detalle_json = _normalize_fuente_snapshot_json(
                row["fuente_detalle_json"] or "",
                fields=("fuente_nombre",),
            )
            pdp_detalle_json = _normalize_fuente_snapshot_json(
                row["pdp_detalle_json"] or "",
                fields=("fuente",),
            )
            if (
                " ".join((row["fuente_nombre"] or "").split()) != canonical_name
                or fuente_detalle_json != (row["fuente_detalle_json"] or "")
                or pdp_detalle_json != (row["pdp_detalle_json"] or "")
            ):
                conn.execute(
                    """
                    UPDATE cargas_manuales
                    SET fuente_nombre = ?,
                        fuente_detalle_json = ?,
                        pdp_detalle_json = ?
                    WHERE id = ?
                    """,
                    (
                        canonical_name,
                        fuente_detalle_json,
                        pdp_detalle_json,
                        carga_id,
                    ),
                )

        observacion_fuentes = conn.execute(
            """
            SELECT DISTINCT TRIM(COALESCE(fuente_financiamiento, '')) AS fuente
            FROM observaciones
            WHERE TRIM(COALESCE(fuente_financiamiento, '')) != ''
            """
        ).fetchall()
        for row in observacion_fuentes:
            fuente_actual = (row["fuente"] or "").strip()
            canonical_name = normalize_fuente_financiamiento(fuente_actual)
            if not canonical_name or canonical_name == fuente_actual:
                continue
            conn.execute(
                """
                UPDATE observaciones
                SET fuente_financiamiento = ?
                WHERE TRIM(COALESCE(fuente_financiamiento, '')) = ?
                """,
                (canonical_name, fuente_actual),
            )
    finally:
        conn.row_factory = original_row_factory


def _normalize_irregularidad_snapshot_json(raw_value: str) -> str:
    clean = (raw_value or "").strip()
    if not clean:
        return raw_value
    try:
        payload = json.loads(clean)
    except json.JSONDecodeError:
        return raw_value
    if not isinstance(payload, list):
        return raw_value

    changed = False
    for item in payload:
        if not isinstance(item, dict):
            continue
        concepto_actual = " ".join(str(item.get("concepto") or "").split())
        subconcepto_actual = " ".join(str(item.get("subconcepto") or "").split())
        concepto_normalizado = normalize_irregularidad_concepto(
            concepto_actual,
            allow_blank=True,
            strict=bool(concepto_actual),
        )
        subconcepto_normalizado = normalize_irregularidad_subconcepto(subconcepto_actual)
        if concepto_actual != concepto_normalizado:
            item["concepto"] = concepto_normalizado
            changed = True
        if subconcepto_actual != subconcepto_normalizado:
            item["subconcepto"] = subconcepto_normalizado
            changed = True
    if not changed:
        return raw_value
    return json.dumps(payload, ensure_ascii=False)


def normalize_irregularidades_data(conn: sqlite3.Connection) -> None:
    original_row_factory = conn.row_factory
    conn.row_factory = sqlite3.Row
    try:
        catalog_rows = conn.execute(
            """
            SELECT id, TRIM(COALESCE(concepto, '')) AS concepto
            FROM catalogo_irregularidades
            ORDER BY id ASC
            """
        ).fetchall()
        catalog_groups: dict[str, dict[str, object]] = {}
        replacement_map: dict[int, int] = {}

        for row in catalog_rows:
            irregularidad_id = int(row["id"])
            canonical_name = normalize_irregularidad_concepto(
                row["concepto"] or "",
                allow_blank=True,
                strict=bool((row["concepto"] or "").strip()),
            )
            if not canonical_name:
                continue
            group_key = normalize_irregularidad_key(canonical_name)
            group = catalog_groups.setdefault(
                group_key,
                {"canonical": canonical_name, "rows": []},
            )
            group["rows"].append(
                {
                    "id": irregularidad_id,
                    "concepto": (row["concepto"] or "").strip(),
                }
            )

        for group in catalog_groups.values():
            canonical_name = str(group["canonical"])
            rows = list(group["rows"])
            exact_matches = [
                row for row in rows
                if " ".join((row["concepto"] or "").split()) == canonical_name
            ]
            survivor = min(exact_matches or rows, key=lambda item: int(item["id"]))
            survivor_id = int(survivor["id"])

            if (survivor["concepto"] or "").strip() != canonical_name:
                conn.execute(
                    "UPDATE catalogo_irregularidades SET concepto = ? WHERE id = ?",
                    (canonical_name, survivor_id),
                )

            for row in rows:
                row_id = int(row["id"])
                if row_id == survivor_id:
                    continue
                replacement_map[row_id] = survivor_id

        for source_id, target_id in replacement_map.items():
            conn.execute(
                "UPDATE registros SET irregularidad_id = ? WHERE irregularidad_id = ?",
                (target_id, source_id),
            )
            conn.execute(
                "DELETE FROM catalogo_irregularidades WHERE id = ?",
                (source_id,),
            )

        observaciones_rows = conn.execute(
            """
            SELECT
                id,
                TRIM(COALESCE(pdp_concepto_irregularidad, '')) AS concepto,
                TRIM(COALESCE(pdp_subconcepto_irregularidad, '')) AS subconcepto
            FROM observaciones
            WHERE TRIM(COALESCE(pdp_concepto_irregularidad, '')) != ''
               OR TRIM(COALESCE(pdp_subconcepto_irregularidad, '')) != ''
            ORDER BY id ASC
            """
        ).fetchall()
        for row in observaciones_rows:
            concepto_actual = (row["concepto"] or "").strip()
            subconcepto_actual = (row["subconcepto"] or "").strip()
            concepto_normalizado = normalize_irregularidad_concepto(
                concepto_actual,
                allow_blank=True,
                strict=bool(concepto_actual),
            )
            subconcepto_normalizado = normalize_irregularidad_subconcepto(subconcepto_actual)
            if (
                concepto_actual != concepto_normalizado
                or subconcepto_actual != subconcepto_normalizado
            ):
                conn.execute(
                    """
                    UPDATE observaciones
                    SET pdp_concepto_irregularidad = ?,
                        pdp_subconcepto_irregularidad = ?
                    WHERE id = ?
                    """,
                    (
                        concepto_normalizado,
                        subconcepto_normalizado,
                        int(row["id"]),
                    ),
                )

        cargas_rows = conn.execute(
            """
            SELECT id, pdp_detalle_json
            FROM cargas_manuales
            WHERE TRIM(COALESCE(pdp_detalle_json, '')) != ''
            ORDER BY id ASC
            """
        ).fetchall()
        for row in cargas_rows:
            pdp_detalle_json = _normalize_irregularidad_snapshot_json(
                row["pdp_detalle_json"] or ""
            )
            if pdp_detalle_json != (row["pdp_detalle_json"] or ""):
                conn.execute(
                    """
                    UPDATE cargas_manuales
                    SET pdp_detalle_json = ?
                    WHERE id = ?
                    """,
                    (pdp_detalle_json, int(row["id"])),
                )
    finally:
        conn.row_factory = original_row_factory


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
    match = re.search(r"(-?\d+)(?:\.(\d+))?", raw.replace(",", ""))
    if not match:
        return 0.0
    try:
        major = int(match.group(1))
        minor = int(match.group(2) or "0")
        return float((major * 1000) + minor)
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
                ramo28 TEXT NOT NULL DEFAULT 'No',
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
                ramo_33 TEXT NOT NULL DEFAULT 'No',
                ramo_28 TEXT NOT NULL DEFAULT 'No',
                origen_fuente TEXT NOT NULL DEFAULT 'Del Ejercicio',
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
                auditoria TEXT,
                ente_numero TEXT,
                ente_numero_sort REAL DEFAULT 0,
                ente_nombre TEXT NOT NULL,
                tipo_auditoria TEXT NOT NULL,
                modalidad TEXT NOT NULL DEFAULT 'Fuente',
                fuente_financiamiento TEXT NOT NULL,
                convenio_nombre TEXT,
                convenio_ente_nombre TEXT,
                convenio_ente_id TEXT,
                ramo_33 TEXT NOT NULL,
                ramo_28 TEXT NOT NULL DEFAULT 'No',
                origen_fuente TEXT NOT NULL DEFAULT 'Del Ejercicio',
                periodo_cedula TEXT,
                periodo_titular TEXT,
                periodo TEXT,
                oficio TEXT,
                fecha_notificacion TEXT,
                tipo_anexo TEXT NOT NULL,
                numero_observacion INTEGER NOT NULL,
                estado TEXT NOT NULL,
                estatus TEXT,
                reclasificada INTEGER NOT NULL DEFAULT 0,
                monto REAL,
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
                modalidad TEXT NOT NULL DEFAULT 'Fuente',
                convenio_nombre TEXT,
                convenio_ente_nombre TEXT,
                convenio_ente_id TEXT,
                periodo TEXT NOT NULL,
                periodo_titular TEXT,
                fecha_notificacion TEXT,
                ramo_33 TEXT NOT NULL DEFAULT 'No',
                ramo_28 TEXT NOT NULL DEFAULT 'No',
                origen_fuente TEXT NOT NULL DEFAULT 'Del Ejercicio',
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

        fuentes_columns = {
            row[1]
            for row in conn.execute("PRAGMA table_info(fuentes_financiamiento)").fetchall()
        }
        if "ramo_33" not in fuentes_columns:
            conn.execute("ALTER TABLE fuentes_financiamiento ADD COLUMN ramo_33 TEXT NOT NULL DEFAULT 'No'")
        if "ramo_28" not in fuentes_columns:
            conn.execute("ALTER TABLE fuentes_financiamiento ADD COLUMN ramo_28 TEXT NOT NULL DEFAULT 'No'")
        if "origen_fuente" not in fuentes_columns:
            conn.execute(
                "ALTER TABLE fuentes_financiamiento ADD COLUMN origen_fuente TEXT NOT NULL DEFAULT 'Del Ejercicio'"
            )
        conn.execute(
            """
            UPDATE fuentes_financiamiento
            SET ramo_33 = 'No'
            WHERE TRIM(COALESCE(ramo_33, '')) = ''
            """
        )
        conn.execute(
            """
            UPDATE fuentes_financiamiento
            SET ramo_28 = 'No'
            WHERE TRIM(COALESCE(ramo_28, '')) = ''
            """
        )
        conn.execute(
            """
            UPDATE fuentes_financiamiento
            SET origen_fuente = CASE
                WHEN LOWER(TRIM(COALESCE(nombre, ''))) LIKE 'remanente%' THEN 'Remanentes'
                WHEN LOWER(TRIM(COALESCE(nombre, ''))) LIKE 'rea:%' THEN 'Remanentes'
                WHEN LOWER(TRIM(COALESCE(nombre, ''))) LIKE 'seguimiento%' THEN 'Remanentes'
                ELSE 'Del Ejercicio'
            END
            WHERE TRIM(COALESCE(origen_fuente, '')) = ''
            """
        )

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
        if "ramo_28" not in cargas_manuales_columns:
            conn.execute("ALTER TABLE cargas_manuales ADD COLUMN ramo_28 TEXT NOT NULL DEFAULT 'No'")
        if "origen_fuente" not in cargas_manuales_columns:
            conn.execute(
                "ALTER TABLE cargas_manuales ADD COLUMN origen_fuente TEXT NOT NULL DEFAULT 'Del Ejercicio'"
            )
        conn.execute(
            """
            UPDATE cargas_manuales
            SET ramo_33 = 'No'
            WHERE TRIM(COALESCE(ramo_33, '')) = ''
            """
        )
        conn.execute(
            """
            UPDATE cargas_manuales
            SET ramo_28 = 'No'
            WHERE TRIM(COALESCE(ramo_28, '')) = ''
            """
        )
        conn.execute(
            """
            UPDATE cargas_manuales
            SET origen_fuente = CASE
                WHEN LOWER(TRIM(COALESCE(fuente_nombre, ''))) LIKE 'remanente%' THEN 'Remanentes'
                WHEN LOWER(TRIM(COALESCE(fuente_nombre, ''))) LIKE 'rea:%' THEN 'Remanentes'
                WHEN LOWER(TRIM(COALESCE(fuente_nombre, ''))) LIKE 'seguimiento%' THEN 'Remanentes'
                ELSE 'Del Ejercicio'
            END
            WHERE TRIM(COALESCE(origen_fuente, '')) = ''
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
        if "modalidad" not in cargas_manuales_columns:
            conn.execute("ALTER TABLE cargas_manuales ADD COLUMN modalidad TEXT NOT NULL DEFAULT 'Fuente'")
        if "convenio_nombre" not in cargas_manuales_columns:
            conn.execute("ALTER TABLE cargas_manuales ADD COLUMN convenio_nombre TEXT")
        if "convenio_ente_nombre" not in cargas_manuales_columns:
            conn.execute("ALTER TABLE cargas_manuales ADD COLUMN convenio_ente_nombre TEXT")
        if "convenio_ente_id" not in cargas_manuales_columns:
            conn.execute("ALTER TABLE cargas_manuales ADD COLUMN convenio_ente_id TEXT")
        conn.execute(
            """
            UPDATE cargas_manuales
            SET modalidad = 'Fuente'
            WHERE TRIM(COALESCE(modalidad, '')) = ''
            """
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
        conn.execute(
            """
            UPDATE observaciones
            SET ente_id = TRIM(RTRIM(COALESCE(ente_id, ''), '.'))
            WHERE COALESCE(ente_id, '') != TRIM(RTRIM(COALESCE(ente_id, ''), '.'))
            """
        )
        if "ente_numero" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN ente_numero TEXT")
        if "ente_numero_sort" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN ente_numero_sort REAL DEFAULT 0")
        ente_sort_expr = ente_numero_sort_sql("ente_numero")
        conn.execute(
            f"""
            UPDATE observaciones
            SET ente_numero_sort = {ente_sort_expr}
            WHERE ente_numero_sort IS NULL
               OR ABS(COALESCE(ente_numero_sort, 0) - ({ente_sort_expr})) > 0.001
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
        if "auditoria" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN auditoria TEXT")
        if "tipo_auditoria" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN tipo_auditoria TEXT")
        if "fuente_financiamiento" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN fuente_financiamiento TEXT")
        if "modalidad" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN modalidad TEXT NOT NULL DEFAULT 'Fuente'")
        if "convenio_nombre" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN convenio_nombre TEXT")
        if "convenio_ente_nombre" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN convenio_ente_nombre TEXT")
        if "convenio_ente_id" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN convenio_ente_id TEXT")
        conn.execute(
            """
            UPDATE observaciones
            SET modalidad = 'Fuente'
            WHERE TRIM(COALESCE(modalidad, '')) = ''
            """
        )
        if "ramo_33" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN ramo_33 TEXT")
        if "ramo_28" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN ramo_28 TEXT NOT NULL DEFAULT 'No'")
        if "origen_fuente" not in observaciones_columns:
            conn.execute(
                "ALTER TABLE observaciones ADD COLUMN origen_fuente TEXT NOT NULL DEFAULT 'Del Ejercicio'"
            )
        conn.execute(
            """
            UPDATE observaciones
            SET ramo_28 = 'No'
            WHERE TRIM(COALESCE(ramo_28, '')) = ''
            """
        )
        conn.execute(
            """
            UPDATE observaciones
            SET origen_fuente = CASE
                WHEN LOWER(TRIM(COALESCE(fuente_financiamiento, ''))) LIKE 'remanente%' THEN 'Remanentes'
                WHEN LOWER(TRIM(COALESCE(fuente_financiamiento, ''))) LIKE 'rea:%' THEN 'Remanentes'
                WHEN LOWER(TRIM(COALESCE(fuente_financiamiento, ''))) LIKE 'seguimiento%' THEN 'Remanentes'
                ELSE 'Del Ejercicio'
            END
            WHERE TRIM(COALESCE(origen_fuente, '')) = ''
            """
        )
        if "periodo_cedula" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN periodo_cedula TEXT")
        if "periodo_titular" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN periodo_titular TEXT")
        if "periodo" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN periodo TEXT")
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
        if "estatus" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN estatus TEXT")
        if "reclasificada" not in observaciones_columns:
            conn.execute(
                "ALTER TABLE observaciones ADD COLUMN reclasificada INTEGER NOT NULL DEFAULT 0"
            )
        if "monto" not in observaciones_columns:
            conn.execute("ALTER TABLE observaciones ADD COLUMN monto REAL")
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
        if "ramo28" not in entes_detalle_columns:
            conn.execute("ALTER TABLE entes_detalle ADD COLUMN ramo28 TEXT NOT NULL DEFAULT 'No'")
        conn.execute(
            """
            UPDATE entes_detalle
            SET ramo28 = 'No'
            WHERE TRIM(COALESCE(ramo28, '')) = ''
            """
        )
        missing_uid = conn.execute(
            "SELECT COUNT(*) FROM entes_detalle WHERE ente_uid IS NULL"
        ).fetchone()[0]
        if missing_uid:
            backfill_ente_uids(conn)
        backfill_historial_ente_uids(conn)
        normalize_fuentes_data(conn)
        normalize_irregularidades_data(conn)
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
            CREATE INDEX IF NOT EXISTS idx_obs_ejercicio_filtros_ramo28
            ON observaciones (ejercicio, fuente_financiamiento, ramo_33, ramo_28, periodo_cedula)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_obs_ejercicio_origen_filtros
            ON observaciones (ejercicio, origen_fuente, fuente_financiamiento, ramo_33, ramo_28, periodo_cedula)
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
                ramo_28,
                periodo_cedula
            )
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_obs_ejercicio_convenios
            ON observaciones (
                ejercicio,
                ente_id,
                tipo_auditoria,
                modalidad,
                convenio_ente_id,
                convenio_nombre,
                fuente_financiamiento,
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
    selected_username = (request.values.get("username") or "").strip().lower()
    if request.method == "POST":
        username = selected_username
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
    usuarios_activos = ordered_users(PROJECT_KEY, priority={"luis": 0, "gabo": 1})
    selected_display = get_display_name(selected_username, fallback="usuario")
    return render_template(
        "login.html",
        error=error,
        next_url=next_url,
        usuarios_activos=usuarios_activos,
        selected_username=selected_username,
        selected_display=selected_display,
    )


@app.route("/logout", methods=["GET", "POST"])
def logout():
    session.clear()
    response = redirect(url_for("login"))
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    return response


@app.route("/api/health")
@app.route("/health")
def health_check():
    return jsonify({"status": "ok", "service": "sifet-estatales"}), 200


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
    "normalize_text_key": normalize_text_key,
    "normalize_fuente_financiamiento": normalize_fuente_financiamiento,
    "normalize_origen_fuente": normalize_origen_fuente,
    "infer_origen_fuente": infer_origen_fuente,
    "normalize_irregularidad_concepto": normalize_irregularidad_concepto,
    "normalize_irregularidad_subconcepto": normalize_irregularidad_subconcepto,
    "normalize_tipo_auditoria": normalize_tipo_auditoria,
    "is_remanente_fuente": is_remanente_fuente,
    "periodo_sql": periodo_sql,
    "ente_numero_sort_sql": ente_numero_sort_sql,
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
