from io import BytesIO

import pytest
from openpyxl import Workbook


@pytest.fixture
def client():
    from app import app

    app.config["TESTING"] = True
    with app.test_client() as test_client:
        yield test_client


def _build_pdp_detail_workbook() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Detalle PDP"
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
            "ConvenioSeguimientoRemanente 2024",
            (
                "CONVENIO: CONVENIO DE COLABORACION PARA LA EJECUCION DE OBRA PUBLICA: "
                "UNIVERSIDAD POLITECNICA DE TLAXCALA\n"
                "REMANENTES DE EJERCICIOS ANTERIORES: FONDO DE APORTACIONES MULTIPLES 2024"
            ),
            "01 DE ENERO AL 27 DE MAYO",
            "Obra Pública",
            2,
            "Volúmenes de obra pagados no ejecutados",
            "$797,710.01",
        ]
    )
    sheet.append(
        [
            "ConvenioDel Ejercicio",
            (
                "CONVENIO: CONVENIO DE COLABORACION PARA LA EJECUCION DE OBRA PUBLICA: "
                "UNIVERSIDAD POLITECNICA DE TLAXCALA\n"
                "FONDO DE APORTACIONES MULTIPLES (FAM)"
            ),
            "16 DE AGOSTO AL 31 DE DICIEMBRE",
            "Obra Pública",
            1,
            "Volúmenes de obra pagados no ejecutados",
            "$367,778.86",
        ]
    )
    sheet.append(
        [
            "ConvenioRemanente 2024",
            (
                "CONVENIO: CONVENIO DE COLABORACION PARA LA EJECUCION DE OBRA PUBLICA: "
                "UNIVERSIDAD INTERCULTURAL DE TLAXCALA\n"
                "REMANENTES DE EJERCICIOS ANTERIORES: FONDO DE APORTACIONES MULTIPLES 2024 (FAM)"
            ),
            "01 DE ENERO AL 27 DE MAYO",
            "Obra Pública",
            "",
            "Conceptos de obra pagados no ejecutados",
            "$1,205,920.52",
        ]
    )
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output.read()


def test_gabo_pdp_detail_excel_preview_normalizes_tipo_fuente(client):
    with client.session_transaction() as session_data:
        session_data["user"] = "gabo"
        session_data["role"] = "loader"

    response = client.post(
        "/carga/pdp-detalle-excel/preview",
        data={
            "pdp_file": (
                BytesIO(_build_pdp_detail_workbook()),
                "detalle_pdp.xlsx",
            ),
        },
        content_type="multipart/form-data",
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["ok"] is True
    assert payload["summary"]["entries"] == 2
    assert payload["summary"]["warnings"] == 1
    assert "NUMERAL requerido" in payload["warnings"][0]
    entry = payload["entries"][0]
    assert entry["modalidad"] == "Convenio"
    assert entry["origen_fuente"] == "Remanentes"
    assert entry["tipo_fuente_clave"] == "convenio_seguimiento_remanente_2024"
    assert entry["ejercicio_fuente"] == "2024"
    assert entry["es_seguimiento"] is True
    assert entry["convenio_nombre"] == (
        "CONVENIO DE COLABORACION PARA LA EJECUCION DE OBRA PUBLICA: "
        "UNIVERSIDAD POLITECNICA DE TLAXCALA"
    )
    assert entry["fuente_nombre"] == (
        "Remanentes de Ejercicios Anteriores: "
        "Fondo de Aportaciones Multiples 2024"
    )
    assert entry["tipo_auditoria"] == "Obra Pública"
    assert entry["numeral"] == 2
    assert entry["monto"] == 797710.01

    convenio_del_ejercicio = payload["entries"][1]
    assert convenio_del_ejercicio["modalidad"] == "Convenio"
    assert convenio_del_ejercicio["origen_fuente"] == "Del Ejercicio"
    assert convenio_del_ejercicio["tipo_fuente_clave"] == "convenio_del_ejercicio"
    assert convenio_del_ejercicio["es_seguimiento"] is False
    assert convenio_del_ejercicio["fuente_nombre"] == "Fondo de Aportaciones Multiples (FAM)"
