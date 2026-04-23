"""
Parser para Cédulas de Resultados — extraído de 19-sifeet-herramientas.
"""

import re
import pdfplumber


ACCIONES = ['SA', 'PDP', 'PRAS', 'PEFCF', 'R']

_MESES_ES = {
    'enero': '01', 'febrero': '02', 'marzo': '03', 'abril': '04',
    'mayo': '05', 'junio': '06', 'julio': '07', 'agosto': '08',
    'septiembre': '09', 'octubre': '10', 'noviembre': '11', 'diciembre': '12',
}


def _fecha_es_to_iso(text):
    """Convierte '15 de enero de 2025' a '2025-01-15'."""
    if not text:
        return ''
    m = re.search(r'(\d{1,2})\s+de\s+(\w+)\s+(?:de[l]?\s+)?(\d{4})', text.lower())
    if not m:
        return ''
    dia, mes_nombre, anio = m.group(1), m.group(2), m.group(3)
    mes = _MESES_ES.get(mes_nombre, '')
    if not mes:
        return ''
    return f'{anio}-{mes}-{dia.zfill(2)}'


def _clean(text):
    if not text:
        return ''
    return text.replace('\n', ' ').strip()


def _parse_num_list(cell):
    """'1, 2, 3' o '0' → [1, 2, 3] o []"""
    if not cell:
        return []
    s = _clean(cell)
    if s in ('0', ''):
        return []
    return [int(x) for x in re.findall(r'\d+', s)]


def _parse_int(cell):
    if not cell:
        return 0
    nums = re.findall(r'\d+', str(cell))
    return int(nums[0]) if nums else 0


def _extract_header(full_text):
    """Extrae número de oficio y fecha del texto completo del documento."""
    oficio = re.search(r'[Oo][Ff][Ii][Cc][Ii][Oo]\s+No\.\s+(OFS/\d+/\d+)', full_text)
    fecha = re.search(r'Tlaxcala,\s+Tlax\.,\s+(?:a\s+)?(.+?\d{4})\.', full_text)
    fecha_raw = _clean(fecha.group(1)) if fecha else None
    fecha_iso = _fecha_es_to_iso(fecha_raw) if fecha_raw else None
    return (
        oficio.group(1) if oficio else None,
        fecha_iso or fecha_raw,
    )


def _detect_audit_type(page_text):
    """Detecta el tipo de auditoría a partir del texto de la página."""
    t = page_text or ''
    if re.search(r'[A-Z]\)\s*Auditor[íi]a\s+(?:de\s+Cumplimiento\s+de\s+)?Obra\s+[Pp][úu]blica', t):
        return 'Obra Pública'
    if re.search(r'[A-Z]\)\s*Auditor[íi]a\s+(?:de\s+Cumplimiento\s+)?Financier[ao]', t):
        return 'Financiera'
    if re.search(r'[A-Z]\)\s*Auditor[íi]a\s+(?:de\s+)?Desempe[ñn]o', t):
        return 'Desempeño'
    if re.search(r'Obra\s+[Pp][úu]blica', t):
        return 'Obra Pública'
    if re.search(r'[Ff]inancier[ao]', t):
        return 'Financiera'
    return 'Sin clasificar'


def _is_cedula_table(table):
    if not table or len(table) < 2:
        return False
    return any('rogresivo' in str(c) for c in table[0] if c)


def _parse_cedula_table(table):
    """
    Convierte las filas de una tabla de cédula en listas estructuradas.
    Retorna (fuentes_list, totales).
    """
    data = table[2:]
    fuentes = {}
    current_fuente = None
    totales = {}

    for row in data:
        if len(row) < 8:
            continue
        fname, periodo, sa, pdp, pras, pefcf, r, total = row

        if fname and 'Total de las Observaciones' in fname:
            totales = {
                'SA':             _parse_int(sa),
                'PDP':            _parse_int(pdp),
                'PRAS':           _parse_int(pras),
                'PEFCF':          _parse_int(pefcf),
                'R':              _parse_int(r),
                'total_emitidas': _parse_int(total),
            }
            continue

        if fname:
            current_fuente = _clean(fname)

        if current_fuente and periodo:
            registro = {
                'periodo':        _clean(periodo),
                'SA':             _parse_num_list(sa),
                'PDP':            _parse_num_list(pdp),
                'PRAS':           _parse_num_list(pras),
                'PEFCF':          _parse_num_list(pefcf),
                'R':              _parse_num_list(r),
                'total_emitidas': _parse_int(total),
            }
            fuentes.setdefault(current_fuente, []).append(registro)

    fuentes_list = [{'nombre': k, 'registros': v} for k, v in fuentes.items()]
    return fuentes_list, totales


def parse_cedula(pdf_path):
    """
    Parsea un PDF de Cédula de Resultados.
    Retorna dict con: oficio, fecha, auditorias[]
    """
    result = {'oficio': None, 'fecha': None, 'auditorias': []}
    full_text = ''
    seen_tipos = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ''
            full_text += page_text + '\n'
            tables = page.extract_tables()

            for table in tables:
                if not _is_cedula_table(table):
                    continue
                tipo = _detect_audit_type(page_text)
                fuentes, totales = _parse_cedula_table(table)
                if not fuentes:
                    continue

                if tipo in seen_tipos:
                    seen_tipos[tipo]['fuentes'].extend(fuentes)
                else:
                    entry = {'tipo': tipo, 'fuentes': fuentes, 'totales': totales}
                    seen_tipos[tipo] = entry
                    result['auditorias'].append(entry)

    oficio, fecha = _extract_header(full_text)
    result['oficio'] = oficio
    result['fecha'] = fecha
    return result
