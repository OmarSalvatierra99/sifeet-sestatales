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


def _clean_multiline(text):
    return re.sub(r'\s+', ' ', _clean(text)).strip()


def _clean_convenio(text):
    clean = _clean_multiline(text)
    clean = re.sub(r'\bCONVENI\s+O\b', 'CONVENIO', clean, flags=re.IGNORECASE)
    clean = re.sub(r'\s*:\s*', ': ', clean)
    return clean


def _infer_convenio_ente(convenio_nombre):
    clean = _clean_convenio(convenio_nombre)
    if not clean:
        return ''
    parts = [part.strip() for part in clean.split(':') if part.strip()]
    return parts[-1] if len(parts) > 1 else clean


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


def _extract_solventacion_header(full_text):
    oficio, fecha = _extract_header(full_text)
    ejercicio_match = re.search(r'ejercicio\s+fiscal\s+(\d{4})', full_text, re.IGNORECASE)
    periodo_match = re.search(
        r'Asunto:\s*Se emiten resultados de solventaci[oó]n\s+de(?:l| los)\s+periodo(?:s)?\s+(.+?)\s+del ejercicio',
        full_text,
        re.IGNORECASE | re.DOTALL,
    )
    oficio_base_matches = re.findall(
        r'mediante\s+el\s+oficio\s+(OFS/\d+/\d+)',
        full_text,
        re.IGNORECASE,
    )
    destinatario_match = re.search(
        r'\n\s*([A-ZÁÉÍÓÚÜÑ][A-ZÁÉÍÓÚÜÑ\s"().,-]+?\([A-Z0-9.\-]+\))\s*\n\s*P\s*R\s*E\s*S\s*E\s*N\s*T\s*E',
        full_text,
        re.IGNORECASE,
    )
    destinatario = None
    if destinatario_match:
        destinatario_raw = destinatario_match.group(1)
        destinatario_parts = [part.strip() for part in re.split(r'\r?\n', destinatario_raw) if part.strip()]
        destinatario = _clean_multiline(destinatario_parts[-1] if destinatario_parts else destinatario_raw)
    oficio_base = oficio_base_matches[-1] if oficio_base_matches else None
    periodo = _clean_multiline(periodo_match.group(1)) if periodo_match else None
    return {
        'oficio': oficio,
        'fecha': fecha,
        'ejercicio': ejercicio_match.group(1) if ejercicio_match else None,
        'periodo': periodo,
        'oficio_base': oficio_base,
        'destinatario': destinatario,
    }


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


def _is_convenios_table(table):
    if not table or len(table) < 2:
        return False
    header_text = ' '.join(_clean_multiline(c).lower() for row in table[:2] for c in row if c)
    return 'nombre del convenio' in header_text and 'fuente' in header_text


def _is_solventacion_table(table):
    if not table or len(table) < 3:
        return False
    header_text = ' '.join(_clean_multiline(c).lower() for row in table[:2] for c in row if c)
    return (
        'acciones emitidas' in header_text
        and 'solventadas' in header_text
        and 'no solventadas' in header_text
        and 'anexo' in header_text
    )


def _parse_convenios_table(table):
    """
    Convierte la tabla C) Convenios en la misma estructura de fuentes,
    marcando la modalidad para evitar duplicarla como fuente normal del ente padre.
    """
    data = table[2:]
    fuentes = {}
    current_fuente = None
    current_convenio = None
    totales = {}

    header = [_clean_multiline(c).lower() for c in (table[0] if table else [])]
    convenio_idx = 1
    fuente_idx = 0
    for idx, label in enumerate(header):
        if 'convenio' in label:
            convenio_idx = idx
        elif 'fuente' in label or 'financiamiento' in label:
            fuente_idx = idx

    for row in data:
        if len(row) < 9:
            continue
        cells = list(row) + [''] * max(0, 9 - len(row))
        fuente_raw = cells[fuente_idx]
        convenio_raw = cells[convenio_idx]
        periodo = cells[2]
        sa, pdp, pras, pefcf, r, total = cells[3:9]

        if fuente_raw and 'Total de las Observaciones' in str(fuente_raw):
            totales = {
                'SA':             _parse_int(sa),
                'PDP':            _parse_int(pdp),
                'PRAS':           _parse_int(pras),
                'PEFCF':          _parse_int(pefcf),
                'R':              _parse_int(r),
                'total_emitidas': _parse_int(total),
            }
            continue
        if convenio_raw and 'Total de las Observaciones' in str(convenio_raw):
            totales = {
                'SA':             _parse_int(sa),
                'PDP':            _parse_int(pdp),
                'PRAS':           _parse_int(pras),
                'PEFCF':          _parse_int(pefcf),
                'R':              _parse_int(r),
                'total_emitidas': _parse_int(total),
            }
            continue

        if fuente_raw:
            current_fuente = _clean_multiline(fuente_raw)
        if convenio_raw:
            current_convenio = _clean_convenio(convenio_raw)

        if current_fuente and current_convenio and periodo:
            key = (current_fuente, current_convenio)
            registro = {
                'periodo':        _clean_multiline(periodo),
                'SA':             _parse_num_list(sa),
                'PDP':            _parse_num_list(pdp),
                'PRAS':           _parse_num_list(pras),
                'PEFCF':          _parse_num_list(pefcf),
                'R':              _parse_num_list(r),
                'total_emitidas': _parse_int(total),
            }
            fuentes.setdefault(key, []).append(registro)

    fuentes_list = []
    for (fuente, convenio), registros in fuentes.items():
        fuentes_list.append({
            'nombre': fuente,
            'modalidad': 'Convenio',
            'convenio_nombre': convenio,
            'convenio_ente_nombre': _infer_convenio_ente(convenio),
            'registros': registros,
        })
    return fuentes_list, totales


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


def _build_emitidas_list(total):
    safe_total = _parse_int(total)
    return list(range(1, safe_total + 1)) if safe_total > 0 else []


def _build_solventacion_totales(emitidas, solventadas, pendientes):
    return {
        'emitidas': len(emitidas),
        'solventadas': len(solventadas),
        'pendientes': len(pendientes),
        'solventadas_indices': solventadas,
        'pendientes_indices': pendientes,
    }


def _parse_solventacion_regular_table(table):
    data = table[2:]
    fuentes = {}
    current_fuente = None
    totals = {
        accion: {'emitidas': 0, 'solventadas': 0, 'pendientes': 0}
        for accion in ACCIONES
    }
    totals.update({'total_emitidas': 0, 'total_solventadas': 0, 'total_pendientes': 0})

    for row in data:
        if len(row) < 8:
            continue
        cells = list(row) + [''] * max(0, 8 - len(row))
        fuente_raw, periodo_raw, anexo_raw, emitidas_raw, solv_total_raw, solv_idx_raw, pend_total_raw, pend_idx_raw = cells[:8]
        row_title = _clean_multiline(anexo_raw or fuente_raw)
        if row_title.upper() in {'SUBTOTAL', 'TOTAL'}:
            continue

        if fuente_raw:
            current_fuente = _clean_multiline(fuente_raw)
        periodo = _clean_multiline(periodo_raw)
        anexo = _clean_multiline(anexo_raw).upper()
        if not current_fuente or not periodo or anexo not in ACCIONES:
            continue

        emitidas = _build_emitidas_list(emitidas_raw)
        solventadas = _parse_num_list(solv_idx_raw)
        pendientes = _parse_num_list(pend_idx_raw)
        registro = {
            'periodo': periodo,
            anexo: emitidas,
            'solventacion': {
                anexo: _build_solventacion_totales(emitidas, solventadas, pendientes)
            },
            'total_emitidas': len(emitidas),
            'total_solventadas': len(solventadas),
            'total_pendientes': len(pendientes),
        }
        fuentes.setdefault(current_fuente, []).append(registro)
        totals[anexo] = _build_solventacion_totales(emitidas, solventadas, pendientes)
        totals['total_emitidas'] += len(emitidas)
        totals['total_solventadas'] += len(solventadas)
        totals['total_pendientes'] += len(pendientes)

    fuentes_list = []
    for fuente_nombre, registros in fuentes.items():
        merged = {}
        for registro in registros:
            periodo = registro['periodo']
            target = merged.setdefault(
                periodo,
                {
                    'periodo': periodo,
                    'SA': [],
                    'PDP': [],
                    'PRAS': [],
                    'PEFCF': [],
                    'R': [],
                    'solventacion': {},
                    'total_emitidas': 0,
                    'total_solventadas': 0,
                    'total_pendientes': 0,
                },
            )
            for accion in ACCIONES:
                if accion in registro:
                    target[accion] = registro[accion]
                    target['solventacion'][accion] = registro['solventacion'][accion]
            target['total_emitidas'] += int(registro.get('total_emitidas') or 0)
            target['total_solventadas'] += int(registro.get('total_solventadas') or 0)
            target['total_pendientes'] += int(registro.get('total_pendientes') or 0)
        fuentes_list.append({'nombre': fuente_nombre, 'registros': list(merged.values())})
    return fuentes_list, totals


def _parse_solventacion_convenios_table(table):
    data = table[2:]
    fuentes = {}
    current_fuente = None
    current_convenio = None
    totals = {
        accion: {'emitidas': 0, 'solventadas': 0, 'pendientes': 0}
        for accion in ACCIONES
    }
    totals.update({'total_emitidas': 0, 'total_solventadas': 0, 'total_pendientes': 0})

    for row in data:
        if len(row) < 9:
            continue
        cells = list(row) + [''] * max(0, 9 - len(row))
        convenio_raw, fuente_raw, periodo_raw, anexo_raw, emitidas_raw, solv_total_raw, solv_idx_raw, pend_total_raw, pend_idx_raw = cells[:9]
        row_title = _clean_multiline(anexo_raw or convenio_raw or fuente_raw)
        if row_title.upper() in {'SUBTOTAL', 'TOTAL'}:
            continue

        if convenio_raw:
            current_convenio = _clean_convenio(convenio_raw)
        if fuente_raw:
            current_fuente = _clean_multiline(fuente_raw)
        periodo = _clean_multiline(periodo_raw)
        anexo = _clean_multiline(anexo_raw).upper()
        if not current_fuente or not current_convenio or not periodo or anexo not in ACCIONES:
            continue

        emitidas = _build_emitidas_list(emitidas_raw)
        solventadas = _parse_num_list(solv_idx_raw)
        pendientes = _parse_num_list(pend_idx_raw)
        key = (current_fuente, current_convenio)
        registro = {
            'periodo': periodo,
            anexo: emitidas,
            'solventacion': {
                anexo: _build_solventacion_totales(emitidas, solventadas, pendientes)
            },
            'total_emitidas': len(emitidas),
            'total_solventadas': len(solventadas),
            'total_pendientes': len(pendientes),
        }
        fuentes.setdefault(key, []).append(registro)
        totals[anexo] = _build_solventacion_totales(emitidas, solventadas, pendientes)
        totals['total_emitidas'] += len(emitidas)
        totals['total_solventadas'] += len(solventadas)
        totals['total_pendientes'] += len(pendientes)

    fuentes_list = []
    for (fuente_nombre, convenio_nombre), registros in fuentes.items():
        merged = {}
        for registro in registros:
            periodo = registro['periodo']
            target = merged.setdefault(
                periodo,
                {
                    'periodo': periodo,
                    'SA': [],
                    'PDP': [],
                    'PRAS': [],
                    'PEFCF': [],
                    'R': [],
                    'solventacion': {},
                    'total_emitidas': 0,
                    'total_solventadas': 0,
                    'total_pendientes': 0,
                },
            )
            for accion in ACCIONES:
                if accion in registro:
                    target[accion] = registro[accion]
                    target['solventacion'][accion] = registro['solventacion'][accion]
            target['total_emitidas'] += int(registro.get('total_emitidas') or 0)
            target['total_solventadas'] += int(registro.get('total_solventadas') or 0)
            target['total_pendientes'] += int(registro.get('total_pendientes') or 0)
        fuentes_list.append(
            {
                'nombre': fuente_nombre,
                'modalidad': 'Convenio',
                'convenio_nombre': convenio_nombre,
                'convenio_ente_nombre': _infer_convenio_ente(convenio_nombre),
                'registros': list(merged.values()),
            }
        )
    return fuentes_list, totals


def _fuente_merge_key(fuente):
    return (
        _clean_multiline(fuente.get('modalidad') or 'Fuente').lower(),
        _clean_multiline(fuente.get('nombre') or '').lower(),
        _clean_convenio(fuente.get('convenio_nombre') or '').lower(),
    )


def _merge_fuentes(entry, fuentes):
    existing = {_fuente_merge_key(fuente): fuente for fuente in entry.get('fuentes', [])}
    for fuente in fuentes:
        key = _fuente_merge_key(fuente)
        if key in existing:
            existing[key].setdefault('registros', []).extend(fuente.get('registros') or [])
            continue
        entry.setdefault('fuentes', []).append(fuente)
        existing[key] = fuente


def parse_cedula(pdf_path):
    """
    Parsea un PDF de Cédula de Resultados.
    Retorna dict con: oficio, fecha, auditorias[]
    """
    result = {'oficio': None, 'fecha': None, 'auditorias': []}
    full_text = ''
    seen_tipos = {}
    last_tipo = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ''
            full_text += page_text + '\n'
            tables = page.extract_tables()
            page_tipo = _detect_audit_type(page_text)
            if page_tipo != 'Sin clasificar':
                last_tipo = page_tipo

            for table in tables:
                if not _is_cedula_table(table):
                    continue
                if _is_convenios_table(table):
                    tipo = 'Obra Pública'
                    fuentes, totales = _parse_convenios_table(table)
                else:
                    tipo = page_tipo if page_tipo != 'Sin clasificar' else (last_tipo or page_tipo)
                    fuentes, totales = _parse_cedula_table(table)
                if not fuentes:
                    continue

                if tipo in seen_tipos:
                    _merge_fuentes(seen_tipos[tipo], fuentes)
                else:
                    entry = {'tipo': tipo, 'fuentes': fuentes, 'totales': totales}
                    seen_tipos[tipo] = entry
                    result['auditorias'].append(entry)

    oficio, fecha = _extract_header(full_text)
    result['oficio'] = oficio
    result['fecha'] = fecha
    return result


def parse_solventacion(pdf_path):
    """
    Parsea un PDF de resultados de solventación.
    Retorna dict con metadatos del oficio y auditorías/fuentes.
    """
    result = {
        'oficio': None,
        'fecha': None,
        'ejercicio': None,
        'periodo': None,
        'oficio_base': None,
        'destinatario': None,
        'auditorias': [],
    }
    full_text = ''
    seen_tipos = {}
    last_tipo = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ''
            full_text += page_text + '\n'
            tables = page.extract_tables()
            page_tipo = _detect_audit_type(page_text)
            if page_tipo != 'Sin clasificar':
                last_tipo = page_tipo

            for table in tables:
                if not _is_solventacion_table(table):
                    continue
                if _is_convenios_table(table):
                    tipo = 'Obra Pública'
                    fuentes, totales = _parse_solventacion_convenios_table(table)
                else:
                    tipo = page_tipo if page_tipo != 'Sin clasificar' else (last_tipo or page_tipo)
                    fuentes, totales = _parse_solventacion_regular_table(table)
                if not fuentes:
                    continue
                if tipo in seen_tipos:
                    _merge_fuentes(seen_tipos[tipo], fuentes)
                else:
                    entry = {'tipo': tipo, 'fuentes': fuentes, 'totales': totales}
                    seen_tipos[tipo] = entry
                    result['auditorias'].append(entry)

    result.update(_extract_solventacion_header(full_text))
    return result
