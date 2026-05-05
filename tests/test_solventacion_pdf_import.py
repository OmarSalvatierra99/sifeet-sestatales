import io
import sqlite3
from datetime import UTC, datetime
from pathlib import Path

import pytest
from openpyxl import Workbook

from scripts.parsers import parse_solventacion


EXAMPLES_DIR = Path(__file__).resolve().parents[1] / "examples"
SOLVENTACION_PDF = EXAMPLES_DIR / "1.16.- SI_OFS_0985_2026_Ene-Jun.pdf"
SOLVENTACION_OPD_PDF = EXAMPLES_DIR / "22.- OPD_OFS_1094_2026_Ene-Jun.pdf"
SOLVENTACION_ZIP = EXAMPLES_DIR / "2.- OFICIOS SOLVENTACIÓN 2025.zip"
SOLVENTACION_PJET_MEMBER = "2.- OFICIOS SOLVENTACIÓN 2025/3.- PJET_OFS_0340_2026_Ene-Jun.pdf"
SOLVENTACION_SEPE_MEMBER = "2.- OFICIOS SOLVENTACIÓN 2025/1.4.- SEPE_OFS_0949_2026_Ene-Jun.pdf"
SOLVENTACION_CONVENIO_MEMBER = "2.- OFICIOS SOLVENTACIÓN 2025/23.- CRI-ESCUELA_OFS_0948_2026_Ene-May_Convenio ITIFE.pdf"


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
    assert [auditoria["tipo"] for auditoria in parsed["auditorias"]] == ["Financiera", "Obra Pública", "Convenios"]


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


def test_parse_solventacion_sepe_extracts_all_anexos_and_progressives(tmp_path):
    import zipfile

    target_pdf = tmp_path / "sepe_0949.pdf"
    with zipfile.ZipFile(SOLVENTACION_ZIP) as zf:
        with zf.open(SOLVENTACION_SEPE_MEMBER) as source, target_pdf.open("wb") as target:
            target.write(source.read())

    parsed = parse_solventacion(str(target_pdf))

    financiera = next(aud for aud in parsed["auditorias"] if aud["tipo"] == "Financiera")
    fuente = next(item for item in financiera["fuentes"] if item["nombre"] == "PARTICIPACIONES ESTATALES")
    registro = next(item for item in fuente["registros"] if item["periodo"] == "01 ene - 30 jun")

    assert len(registro["SA"]) == 20
    assert len(registro["PDP"]) == 14
    assert len(registro["PRAS"]) == 0
    assert len(registro["PEFCF"]) == 0
    assert len(registro["R"]) == 2

    assert registro["solventacion"]["SA"]["solventadas_indices"] == [3, 5, 6, 7, 8, 10, 11, 12, 13, 20]
    assert registro["solventacion"]["SA"]["pendientes_indices"] == [1, 2, 4, 9, 14, 15, 16, 17, 18, 19]
    assert registro["solventacion"]["PDP"]["solventadas_indices"] == [1, 2, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14]
    assert registro["solventacion"]["PDP"]["pendientes_indices"] == [3, 10]
    assert registro["solventacion"]["R"]["solventadas_indices"] == [1, 2]
    assert registro["solventacion"]["R"]["pendientes_indices"] == []


def test_parse_solventacion_convenio_pdf_stays_in_convenios_block(tmp_path):
    import zipfile

    target_pdf = tmp_path / "cri_escuela_0948.pdf"
    with zipfile.ZipFile(SOLVENTACION_ZIP) as zf:
        with zf.open(SOLVENTACION_CONVENIO_MEMBER) as source, target_pdf.open("wb") as target:
            target.write(source.read())

    parsed = parse_solventacion(str(target_pdf))

    assert [auditoria["tipo"] for auditoria in parsed["auditorias"]] == ["Convenios"]
    convenio = parsed["auditorias"][0]["fuentes"][0]

    assert convenio["modalidad"] == "Convenio"
    assert "INSTITUTO TLAXCALTECA" in convenio["convenio_nombre"]
    assert "FÍSICA EDUCATIVA" in convenio["convenio_nombre"]


def test_parse_solventacion_opd_keeps_financial_continuation_before_obra_header():
    parsed = parse_solventacion(str(SOLVENTACION_OPD_PDF))

    auditorias = {auditoria["tipo"]: auditoria for auditoria in parsed["auditorias"]}
    financiera = auditorias["Financiera"]
    obra = auditorias["Obra Pública"]

    assert parsed["oficio"] == "OFS/1094/2026"
    assert parsed["oficio_base"] == "OFS/0567/2026"
    assert financiera["totales"]["total_emitidas"] == 99
    assert obra["totales"]["total_emitidas"] == 6
    assert financiera["totales"]["SA"]["emitidas"] == 42
    assert financiera["totales"]["PDP"]["emitidas"] == 17
    assert financiera["totales"]["PRAS"]["emitidas"] == 33
    assert financiera["totales"]["PEFCF"]["emitidas"] == 5
    assert financiera["totales"]["R"]["emitidas"] == 2
    assert obra["totales"]["PDP"]["emitidas"] == 5
    assert obra["totales"]["PRAS"]["emitidas"] == 1

    s200_fuente = next(
        fuente
        for fuente in financiera["fuentes"]
        if "CONVENIO IB-CC-S200" in fuente["nombre"]
    )
    assert s200_fuente["registros"][0]["total_emitidas"] == 1
    assert all("CONVENIO IB-CC-S200" not in fuente["nombre"] for fuente in obra["fuentes"])


def test_parse_solventacion_pjet_splits_two_auditorias_on_same_page(tmp_path):
    import zipfile

    target_pdf = tmp_path / "pjet_0340.pdf"
    with zipfile.ZipFile(SOLVENTACION_ZIP) as zf:
        with zf.open(SOLVENTACION_PJET_MEMBER) as source, target_pdf.open("wb") as target:
            target.write(source.read())

    parsed = parse_solventacion(str(target_pdf))

    auditorias = {auditoria["tipo"]: auditoria for auditoria in parsed["auditorias"]}
    financiera = auditorias["Financiera"]
    obra = auditorias["Obra Pública"]

    assert parsed["oficio"] == "OFS/0340/2026"
    assert parsed["oficio_base"] == "OFS/3008/2025"
    assert financiera["totales"]["total_emitidas"] == 52
    assert financiera["totales"]["total_solventadas"] == 50
    assert financiera["totales"]["total_pendientes"] == 2
    assert obra["totales"]["total_emitidas"] == 14
    assert obra["totales"]["total_solventadas"] == 11
    assert obra["totales"]["total_pendientes"] == 3


def test_solventacion_import_ignores_zero_blocks_without_matching_observaciones(solventacion_client, tmp_path):
    client = solventacion_client
    db_path = tmp_path / "solventacion_import.db"
    now = datetime.now(UTC).isoformat(timespec="seconds")

    with sqlite3.connect(db_path) as conn:
        conn.executemany(
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
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    "2025",
                    "1.4",
                    "1.4",
                    "Secretaría de Educación Pública (SEPE)",
                    "Financiera",
                    "Participaciones Estatales",
                    "No",
                    "01 de Enero al 30 de Junio",
                    "01 de Enero al 30 de Junio",
                    "OFS/0341/2026",
                    "2026-02-10",
                    "SA",
                    1,
                    "Pendiente",
                    now,
                ),
                (
                    "2025",
                    "1.4",
                    "1.4",
                    "Secretaría de Educación Pública (SEPE)",
                    "Financiera",
                    "Participaciones Estatales",
                    "No",
                    "01 de Enero al 30 de Junio",
                    "01 de Enero al 30 de Junio",
                    "OFS/0341/2026",
                    "2026-02-10",
                    "SA",
                    2,
                    "Pendiente",
                    now,
                ),
                (
                    "2025",
                    "1.4",
                    "1.4",
                    "Secretaría de Educación Pública (SEPE)",
                    "Financiera",
                    "Participaciones Estatales",
                    "No",
                    "01 de Enero al 30 de Junio",
                    "01 de Enero al 30 de Junio",
                    "OFS/0341/2026",
                    "2026-02-10",
                    "SA",
                    3,
                    "Pendiente",
                    now,
                ),
            ],
        )

    response = client.post(
        "/carga/observaciones-admin/solventacion-importar",
        json={
            "scope": {
                "ejercicio": "2025",
                "ente_id": "1.4",
                "oficio": "OFS/0341/2026",
                "tipo_auditoria": "Financiera",
            },
            "rows": [
                {
                    "tipo_auditoria": "Financiera",
                    "fuente_financiamiento": "Participaciones Estatales",
                    "periodo": "01 ene - 30 jun",
                    "tipo_anexo": "SA",
                    "emitidas": 3,
                    "solventadas_indices": [2],
                    "pendientes_indices": [1, 3],
                },
                {
                    "tipo_auditoria": "Financiera",
                    "fuente_financiamiento": "Participaciones Estatales",
                    "periodo": "01 ene - 30 jun",
                    "tipo_anexo": "PRAS",
                    "emitidas": 0,
                    "solventadas_indices": [],
                    "pendientes_indices": [],
                },
            ],
        },
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["ok"] is True
    assert payload["updated"] == 3

    with sqlite3.connect(db_path) as conn:
        rows = conn.execute(
            """
            SELECT numero_observacion, estado
            FROM observaciones
            WHERE ejercicio = '2025' AND ente_id = '1.4' AND oficio = 'OFS/0341/2026'
            ORDER BY numero_observacion
            """
        ).fetchall()

    assert rows == [(1, "Pendiente"), (2, "Solventado"), (3, "Pendiente")]


def test_solventacion_import_normalizes_pdp_amounts_for_solventadas(solventacion_client, tmp_path):
    client = solventacion_client
    db_path = tmp_path / "solventacion_import.db"
    now = datetime.now(UTC).isoformat(timespec="seconds")

    with sqlite3.connect(db_path) as conn:
        conn.executemany(
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
                monto,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    "2025",
                    "27",
                    "27",
                    "Unidad de Servicios Educativos del Estado de Tlaxcala (USET)",
                    "Financiera",
                    "Participaciones Estatales",
                    "No",
                    "01 de Enero al 31 de Marzo",
                    "01 de Enero al 31 de Marzo",
                    "OFS/0646/2026",
                    "2026-02-10",
                    "PDP",
                    1,
                    "Emitido",
                    120.50,
                    0.0,
                    0.0,
                    120.50,
                    now,
                ),
                (
                    "2025",
                    "27",
                    "27",
                    "Unidad de Servicios Educativos del Estado de Tlaxcala (USET)",
                    "Financiera",
                    "Participaciones Estatales",
                    "No",
                    "01 de Enero al 31 de Marzo",
                    "01 de Enero al 31 de Marzo",
                    "OFS/0646/2026",
                    "2026-02-10",
                    "PDP",
                    2,
                    "Emitido",
                    80.0,
                    15.0,
                    65.0,
                    80.0,
                    now,
                ),
            ],
        )

    response = client.post(
        "/carga/observaciones-admin/solventacion-importar",
        json={
            "scope": {
                "ejercicio": "2025",
                "ente_id": "27",
                "oficio": "OFS/0646/2026",
                "tipo_auditoria": "Financiera",
            },
            "rows": [
                {
                    "tipo_auditoria": "Financiera",
                    "fuente_financiamiento": "Participaciones Estatales",
                    "periodo": "01 de Enero al 31 de Marzo",
                    "tipo_anexo": "PDP",
                    "emitidas": 2,
                    "solventadas_indices": [1],
                    "pendientes_indices": [2],
                },
            ],
        },
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["ok"] is True
    assert payload["updated"] == 2

    with sqlite3.connect(db_path) as conn:
        rows = conn.execute(
            """
            SELECT numero_observacion, estado, monto_pdp_emitido, monto_pdp_solventado, monto_pdp_pendiente
            FROM observaciones
            WHERE ejercicio = '2025' AND ente_id = '27' AND oficio = 'OFS/0646/2026'
            ORDER BY numero_observacion
            """
        ).fetchall()

    assert rows == [
        (1, "Solventado", 120.5, 120.5, 0.0),
        (2, "Pendiente", 80.0, 15.0, 65.0),
    ]


def _build_pdp_solventado_workbook() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Montos PDP"
    sheet.append(
        [
            "TIPO DE FUENTE",
            "F.F",
            "PERIODO",
            "SUBTIPO DE AUDITORIA",
            "NUMERAL",
            "CONCEPTO PDP",
            "MONTO PDP",
        ]
    )
    sheet.append(
        [
            "Del Ejercicio",
            "Participaciones Estatales",
            "01 DE ENERO AL 30 DE JUNIO",
            "Obra Pública",
            1,
            "Volúmenes de obra pagados no ejecutados",
            40,
        ]
    )
    sheet.append(
        [
            "Del Ejercicio",
            "Participaciones Estatales",
            "01 DE ENERO AL 30 DE JUNIO",
            "Obra Pública",
            2,
            "Volúmenes de obra pagados no ejecutados",
            75.25,
        ]
    )
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output.read()


def test_solventacion_pdp_solventado_excel_preview_and_apply(solventacion_client, tmp_path):
    client = solventacion_client
    db_path = tmp_path / "solventacion_import.db"
    now = datetime.now(UTC).isoformat(timespec="seconds")

    with sqlite3.connect(db_path) as conn:
        conn.executemany(
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
                monto,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    "2025",
                    "1.16",
                    "1.16",
                    "Secretaría de Infraestructura (SI)",
                    "Obra Pública",
                    "Participaciones Estatales",
                    "No",
                    "01 de Enero al 30 de Junio",
                    "01 de Enero al 30 de Junio",
                    "OFS/0342/2026",
                    "2026-02-10",
                    "PDP",
                    1,
                    "Pendiente",
                    100.0,
                    0.0,
                    100.0,
                    100.0,
                    now,
                ),
                (
                    "2025",
                    "1.16",
                    "1.16",
                    "Secretaría de Infraestructura (SI)",
                    "Obra Pública",
                    "Participaciones Estatales",
                    "No",
                    "01 de Enero al 30 de Junio",
                    "01 de Enero al 30 de Junio",
                    "OFS/0342/2026",
                    "2026-02-10",
                    "PDP",
                    2,
                    "Pendiente",
                    80.0,
                    10.0,
                    70.0,
                    80.0,
                    now,
                ),
            ],
        )

    preview = client.post(
        "/carga/observaciones-admin/pdp-solventado-excel/preview",
        data={
            "ejercicio": "2025",
            "ente_id": "1.16",
            "oficio": "OFS/0342/2026",
            "pdp_file": (io.BytesIO(_build_pdp_solventado_workbook()), "montos_pdp.xlsx"),
        },
        content_type="multipart/form-data",
    )

    assert preview.status_code == 200
    preview_payload = preview.get_json()
    assert preview_payload["ok"] is True
    assert preview_payload["summary"]["matched_rows"] == 2
    assert preview_payload["summary"]["changed_rows"] == 2
    assert [item["monto_pdp_solventado_after"] for item in preview_payload["updates"]] == [40.0, 75.25]

    apply_response = client.post(
        "/carga/observaciones-admin/pdp-solventado-excel/aplicar",
        json={
            "scope": {
                "ejercicio": "2025",
                "ente_id": "1.16",
                "oficio": "OFS/0342/2026",
            },
            "updates": [
                {
                    "id": item["id"],
                    "monto_pdp_solventado": item["monto_pdp_solventado_after"],
                }
                for item in preview_payload["updates"]
            ],
        },
    )

    assert apply_response.status_code == 200
    apply_payload = apply_response.get_json()
    assert apply_payload["ok"] is True
    assert apply_payload["updated"] == 2

    with sqlite3.connect(db_path) as conn:
        rows = conn.execute(
            """
            SELECT numero_observacion, monto_pdp_emitido, monto_pdp_solventado, monto_pdp_pendiente
            FROM observaciones
            WHERE ejercicio = '2025' AND ente_id = '1.16' AND oficio = 'OFS/0342/2026'
            ORDER BY numero_observacion
            """
        ).fetchall()

    assert rows == [
        (1, 100.0, 40.0, 60.0),
        (2, 80.0, 75.25, 4.75),
    ]
