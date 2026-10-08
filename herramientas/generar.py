# -*- coding: utf-8 -*-
import sys, csv, os
from datetime import datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from datos import procs, people
from aprobacion import APROBACION, VERSION
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

AZUL   = PatternFill('solid', fgColor='FF1F4E78')
VERDE  = PatternFill('solid', fgColor='FFC6EFCE')
ROJO   = PatternFill('solid', fgColor='FFFFC7CE')
AMAR   = PatternFill('solid', fgColor='FFFFEB9C')
GRIS   = PatternFill('solid', fgColor='FFD9D9D9')
F      = lambda **k: Font(name='Arial', **{'size': 10, **k})
FB     = lambda **k: Font(name='Arial', bold=True, **{'size': 10, **k})
BLANCO = FB(color='FFFFFFFF')
TH     = Side(style='thin', color='FFB0B0B0')
BORDE  = Border(left=TH, right=TH, top=TH, bottom=TH)
CEN    = Alignment(horizontal='center', vertical='center', wrap_text=True)
IZQ    = Alignment(horizontal='left', vertical='center', wrap_text=True)

EMPRESA = 'COMPAÑÍA MINERA DEL PACIFICO [CMP]'
CONF    = 'POR CONFIRMAR'
LIC     = 'LICENCIA MÉDICA'

# Trabajadores ausentes con licencia medica: sus pendientes no son incumplimiento,
# quedan supeditados a la reincorporacion. Se marcan en amarillo, no en rojo.
SITUACION = {
    'p10': 'Licencia médica',   # SANTIBAÑEZ BAHAMONDES, CRISTOPHER NICOLA
    'p28': 'Licencia médica',   # SAAVEDRA, LEONARDO
}
REINC = 'A su reincorporación'

# El "Plan de Gestion de Riesgos de Desastres" NO se incluye: corresponde al
# "Plan de Respuesta a Emergencia Local", que ya figura como 3.8.3.12 y esta
# difundido y evaluado. Incluirlo seria duplicarlo.
NUEVOS = [
    ('Instructivo Llenado de VAT (Verificación y Autorización de Trabajo)', CONF),
]
RANGO_DIF = '21/09/2026 al 25/09/2026'
RANGO_EVA = '28/09/2026 al 09/10/2026'

# ── datos reales exportados del dashboard: (pid, code) -> (difusion, evaluacion)
CSV = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'export_dashboard_20260916.csv')
registro = {}
with open(CSV, encoding='utf-8-sig') as fh:
    for r in csv.DictReader(fh):
        registro[(r['pid'], r['code'])] = (r['fecha_difusion'], r['fecha_evaluacion_nota'])
if not registro:
    raise SystemExit('El CSV de la Hoja 2 vino vacio')

# Evaluaciones pendientes que se dan por rendidas con nota 100, con la fecha
# que ya usa cada turno (la mas frecuente en sus evaluaciones existentes).
FECHA_TURNO = {'G1': '30/08/2026', 'G2': '30/08/2026',
               'G3': '20/08/2026', 'G4': '01/09/2026'}

def completar_evaluaciones():
    """DESACTIVADO. La app es la fuente de verdad: el libro refleja lo que
    esta registrado en el dashboard, sin dar por rendida ninguna evaluacion
    que alli figure pendiente. Se conserva la funcion por si se necesita
    reactivarla, pero no se invoca."""
    n = 0
    for p in people:
        if p['id'] in SITUACION:
            continue
        fecha = FECHA_TURNO.get(p['g'])
        if not fecha:
            continue
        for pr in procs:
            clave = (p['id'], pr['code'])
            dif, ev = registro.get(clave, ('PENDIENTE', 'PENDIENTE'))
            if ev != 'PENDIENTE':
                continue
            ap = APROBACION.get(pr['code'])
            if ap and datetime.strptime(fecha, '%d/%m/%Y') < datetime.strptime(ap, '%d/%m/%Y'):
                raise SystemExit(
                    'La fecha %s del turno %s es anterior a la aprobacion de %s (%s)'
                    % (fecha, p['g'], pr['code'], ap))
            registro[clave] = (dif, '%s / 100%%' % fecha)
            n += 1
    return n


def aprob(code):
    return APROBACION.get(code, CONF)

def es_previa(fecha_txt, code):
    """True si la fecha es anterior a la aprobacion del documento."""
    ap = APROBACION.get(code)
    if not ap or fecha_txt == 'PENDIENTE':
        return False
    f = datetime.strptime(fecha_txt[:10], '%d/%m/%Y').date()
    return f < datetime.strptime(ap, '%d/%m/%Y').date()

def nombre_completo(p):
    return ('%s %s' % (p['nombres'], p['apellidos'])).strip()

N_COMPLETADAS = 0   # el libro refleja la app tal cual; no se rellena nada

wb = openpyxl.Workbook()

# ───────────────────────── Hoja 0 ─────────────────────────
ws = wb.active
ws.title = '0. Portada y Leyenda'
ws.column_dimensions['A'].width = 66.4
ws.column_dimensions['B'].width = 46.0
ws.column_dimensions['C'].width = 95.0
portada = [
    ('SUBSANACIÓN HALLAZGO SNGM — ART. 28', None, None, 'titulo'),
    ('Falta de claridad respecto de las necesidades de capacitación', None, None, 'sub'),
    (None, None, None, None),
    ('Empresa mandante', 'Compañía Minera del Pacífico S.A. (CMP)', None, 'par'),
    ('Faena / Instalación', 'Puerto Punta Totoralillo', None, 'par'),
    ('Proceso', '3.8 Operaciones Puerto Punta Totoralillo — 3.8.1 Descarga, Recepción y Apilamiento / '
                '3.8.2 Embarque / 3.8.3 Filtrado', None, 'par'),
    ('Medida correctiva', 'Capacitar a los trabajadores sobre el método y procedimiento para ejecutar '
                          'correctamente su trabajo', None, 'par'),
    ('Fecha de emisión de esta planilla', '16-09-2026', None, 'par'),
    ('Elaborado por', 'Héctor Garrido Pérez', None, 'par'),
    (None, None, None, None),
    ('CONTENIDO DEL ARCHIVO', None, None, 'titulo'),
    ('Hoja "1. Aplicabilidad (X · N-A)"',
     'Punto 1: matriz de qué documentos debe tener cada trabajador, con "X" o "N/A". Sin celdas en blanco.',
     None, 'par'),
    ('Hoja "2. Difusión y Evaluación"',
     'Punto 2: copia de la matriz anterior con fecha de difusión, y fecha y nota de la evaluación.',
     None, 'par'),
    ('Hoja "3. Programa de Capacitación"',
     'Punto 3: trabajadores pendientes, documento a difundir y fecha comprometida.',
     None, 'par'),
    (None, None, None, None),
    ('CÓDIGO DE COLORES (aplica a todas las hojas)', None, None, 'titulo'),
    ('Color', 'Significado', 'Detalle', 'cab'),
    (None, 'VERDE', 'Realizado y con registro documental disponible (difusión efectuada o evaluación '
                    'rendida y aprobada), o bien situación regularizada sin observaciones, como la '
                    'de los trabajadores ausentes con licencia médica, identificados en la propia '
                    'celda. Se utiliza sólo en las Hojas 2 y 3.', 'verde'),
    (None, 'ROJO', 'Pendiente: no realizado, sin registro de cumplimiento, o realizado en una versión '
                   'anterior del documento. Debe quedar incluido en el Programa de Capacitación (Hoja 3).', 'rojo'),
    (None, 'AMARILLO', 'No aplica (N/A) al trabajador según su cargo y contrato; difusión informativa '
                       'según el criterio indicado más abajo; o dato por completar/confirmar.', 'amar'),
    ('Nota: en la Hoja 1 el color no indica cumplimiento: la "X" señala únicamente que el documento '
     'aplica al trabajador y el amarillo señala "N/A".', None, None, 'nota'),
    (None, None, None, None),
    ('TRABAJADORES AUSENTES CON LICENCIA MÉDICA', None, None, 'titulo'),
    ('Los trabajadores que se encuentran con licencia médica figuran en la Hoja 2 con la leyenda '
     '"LICENCIA MÉDICA" en amarillo, en lugar de "PENDIENTE" en rojo. Su difusión y evaluación no '
     'constituyen incumplimiento: quedan supeditadas a su reincorporación, y así se consignan en la '
     'Hoja 3 como fecha comprometida.', None, None, 'nota'),
    (None, None, None, None),
    ('CRITERIO DE DIFUSIÓN INFORMATIVA', None, None, 'titulo'),
    ('Difusión informativa: para trabajadores que reciben una difusión exclusivamente informativa, pero '
     'que no ejecutan, participan ni supervisan la tarea regulada por el documento, se registra la fecha '
     'de difusión y "N/A" en evaluación. Esta difusión informativa no modifica la condición de no '
     'aplicabilidad definida en la Planilla 1.', None, None, 'nota'),
]
for i, (a, b, c, tipo) in enumerate(portada, start=1):
    if a is not None: ws.cell(i, 1, a)
    if b is not None: ws.cell(i, 2, b)
    if c is not None: ws.cell(i, 3, c)
    for col in (1, 2, 3):
        cel = ws.cell(i, col)
        cel.font = F(); cel.alignment = IZQ
    if tipo == 'titulo':
        ws.cell(i, 1).font = FB(size=12)
    elif tipo == 'sub':
        ws.cell(i, 1).font = F(italic=True)
    elif tipo == 'par':
        ws.cell(i, 1).font = FB()
    elif tipo == 'cab':
        for col in (1, 2, 3):
            ws.cell(i, col).fill = AZUL; ws.cell(i, col).font = BLANCO
    elif tipo in ('verde', 'rojo', 'amar'):
        ws.cell(i, 1).fill = {'verde': VERDE, 'rojo': ROJO, 'amar': AMAR}[tipo]
        ws.cell(i, 2).font = FB()
    elif tipo == 'nota':
        ws.cell(i, 1).font = F(size=9, italic=True)

# ─────────────── cabecera comun de las hojas 1 y 2 ───────────────
def cabecera(ws, con_actividad):
    filas = [('Documento (nombre según carátula)', 'nombre'),
             ('Código único de identificación', 'code'),
             ('N° versión / revisión', 'version'),
             ('Fecha de aprobación (dd/mm/aaaa)', 'aprob')]
    if con_actividad:
        filas.append(('Código de actividad asociada', 'code'))
    return filas

COLS_FIJAS = [('N°', 5), ('RUT', 15), ('Nombre completo', 34),
              ('Cargo según contrato vigente', 30), ('Empresa', 20)]

def escribe_fijas(ws, fila_cab, primera_fila):
    for j, (tit, ancho) in enumerate(COLS_FIJAS, start=1):
        c = ws.cell(fila_cab, j, tit)
        c.fill = AZUL; c.font = BLANCO; c.alignment = CEN; c.border = BORDE
        ws.column_dimensions[get_column_letter(j)].width = ancho
    for i, p in enumerate(people):
        r = primera_fila + i
        vals = [i + 1, p['rut'], nombre_completo(p), p['cargo'], EMPRESA]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(r, j, v)
            c.font = F(); c.border = BORDE
            c.alignment = CEN if j in (1, 2) else IZQ
        if p['rut'] == CONF:
            ws.cell(r, 2).fill = AMAR

# ───────────────────────── Hoja 1 ─────────────────────────
ws = wb.create_sheet('1. Aplicabilidad (X · N-A)')
ws['A1'] = 'PLANILLA 1 — APLICABILIDAD DE DOCUMENTOS POR TRABAJADOR'
ws['A1'].font = FB(size=12)
ws['A2'] = ('Criterio: "X" = el trabajador debe estar capacitado y evaluado en ese documento.  '
            '"N/A" = no aplica a su cargo y contrato.  No se dejan celdas en blanco.')
ws['A2'].font = F(size=9, italic=True)
for fila, (etq, _) in enumerate(cabecera(ws, False), start=4):
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=4)
    c = ws.cell(fila, 1, etq); c.font = FB(); c.fill = GRIS; c.alignment = IZQ; c.border = BORDE
    ws.cell(fila, 5).fill = GRIS; ws.cell(fila, 5).border = BORDE
for k, pr in enumerate(procs):
    col = 6 + k
    ws.column_dimensions[get_column_letter(col)].width = 13
    ap = aprob(pr['code'])
    vals = [(4, pr['nombre'], AZUL, BLANCO), (5, pr['code'], AMAR, FB()),
            (6, VERSION, AMAR if VERSION == CONF else None, F()),
            (7, ap, AMAR if ap == CONF else None, F())]
    for fila, v, fill, font in vals:
        c = ws.cell(fila, col, v)
        if fill is not None: c.fill = fill
        c.font = font; c.alignment = CEN; c.border = BORDE
ws.row_dimensions[4].height = 60
escribe_fijas(ws, 8, 9)
for k, pr in enumerate(procs):
    col = 6 + k
    c = ws.cell(8, col, pr['code'])
    c.fill = AZUL; c.font = BLANCO; c.alignment = CEN; c.border = BORDE
    for i in range(len(people)):
        c = ws.cell(9 + i, col, 'X')
        c.font = F(); c.alignment = CEN; c.border = BORDE
ws.freeze_panes = 'F9'

# ───────────────────────── Hoja 2 ─────────────────────────
ws = wb.create_sheet('2. Difusión y Evaluación')
ws['A1'] = 'PLANILLA 2 — REGISTRO DE DIFUSIÓN Y EVALUACIÓN (fecha y nota)'
ws['A1'].font = FB(size=12)
ws['A2'] = ('Criterio: se indica la fecha (dd/mm/aaaa) de la difusión y la fecha y nota de la evaluación. '
            '"N/A" = no aplica al cargo. "PENDIENTE" = sin registro de cumplimiento.')
ws['A2'].font = F(size=9, italic=True)
ws['A3'] = ('Datos tomados del dashboard de control (exportación del 16/09/2026). '
            'VERDE = registrado; ROJO = pendiente.')
ws['A3'].font = F(size=9, italic=True)
for fila, (etq, _) in enumerate(cabecera(ws, True), start=5):
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=4)
    c = ws.cell(fila, 1, etq); c.font = FB(); c.fill = GRIS; c.alignment = IZQ; c.border = BORDE
    ws.cell(fila, 5).fill = GRIS; ws.cell(fila, 5).border = BORDE
for k, pr in enumerate(procs):
    col = 6 + k * 2
    ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col + 1)
    c = ws.cell(5, col, pr['nombre']); c.fill = AZUL; c.font = BLANCO; c.alignment = CEN; c.border = BORDE
    ws.cell(5, col + 1).fill = AZUL; ws.cell(5, col + 1).border = BORDE
    ap = aprob(pr['code'])
    for fila, v, fill in ((6, pr['code'], AMAR), (7, VERSION, None),
                          (8, ap, AMAR if ap == CONF else None), (9, pr['code'], AMAR)):
        ws.merge_cells(start_row=fila, start_column=col, end_row=fila, end_column=col + 1)
        c = ws.cell(fila, col, v); c.alignment = CEN; c.border = BORDE
        c.font = FB() if fila in (6, 9) else F()
        if fill is not None: c.fill = fill; ws.cell(fila, col + 1).fill = fill
        ws.cell(fila, col + 1).border = BORDE
    for off, tit in ((0, 'Fecha difusión'), (1, 'Fecha evaluación / nota')):
        c = ws.cell(10, col + off, tit)
        c.fill = AZUL; c.font = BLANCO; c.alignment = CEN; c.border = BORDE
        ws.column_dimensions[get_column_letter(col + off)].width = 16
ws.row_dimensions[5].height = 60
escribe_fijas(ws, 10, 11)
faltantes_csv = []
previas = []
for i, p in enumerate(people):
    for k, pr in enumerate(procs):
        par = registro.get((p['id'], pr['code']))
        if par is None:
            faltantes_csv.append((p['id'], pr['code']))
            par = ('PENDIENTE', 'PENDIENTE')
        for off, val in enumerate(par):
            previa = off == 0 and es_previa(val, pr['code'])
            if previa:
                val = val + ' (previa a aprobación)'
                previas.append((p['id'], pr['code']))
            ausente = p['id'] in SITUACION and val == 'PENDIENTE'
            if ausente:
                val = LIC
            c = ws.cell(11 + i, 6 + k * 2 + off, val)
            c.font = F(); c.alignment = CEN; c.border = BORDE
            # Las de licencia medica van en verde, como el resto; el texto de la
            # celda es el que informa la situacion.
            c.fill = ROJO if (val == 'PENDIENTE' or previa) else VERDE
if faltantes_csv:
    raise SystemExit('Faltan %d combinaciones en el CSV, p.ej. %s' % (len(faltantes_csv), faltantes_csv[:3]))
if previas:
    ws['A4'] = ('Atención: %d difusiones están fechadas antes de la aprobación del documento '
                'respectivo y se marcan en rojo; corresponden a re-instrucción en la versión '
                'vigente.' % len(previas))
    ws['A4'].font = FB(size=9, color='FF9C0006')
ws.freeze_panes = 'F11'

# ───────────────────────── Hoja 3 ─────────────────────────
ws = wb.create_sheet('3. Programa de Capacitación')
ws['A1'] = 'PLANILLA 3 — PROGRAMA DE CAPACITACIÓN'
ws['A1'].font = FB(size=12)
ws['A2'] = ('Incluye (a) las evaluaciones que la Planilla 2 registra como pendientes, cuya difusión ya '
            'está realizada, y (b) el documento nuevo, que no cuenta con registro previo en ningún '
            'trabajador y por eso ingresa como "Difusión y Evaluación".')
ws['A2'].font = F(size=9, italic=True)
cabs = ['N°', 'RUT', 'Nombre completo', 'Cargo según contrato vigente', 'Empresa', 'Turno',
        'Actividad pendiente', 'Documento a difundir / evaluar', 'Código único',
        'Fecha comprometida difusión', 'Fecha comprometida evaluación', 'Responsable de ejecución']
anchos = [5, 15, 34, 30, 20, 8, 19, 46, 14, 22, 22, 26]
for j, (tit, an) in enumerate(zip(cabs, anchos), start=1):
    c = ws.cell(4, j, tit); c.fill = AZUL; c.font = BLANCO; c.alignment = CEN; c.border = BORDE
    ws.column_dimensions[get_column_letter(j)].width = an
ws.row_dimensions[4].height = 30
# Filas del programa: primero los pendientes que arrastra la Hoja 2 (criterio
# de la propia planilla: todo lo rojo entra al programa), luego el documento
# nuevo. La difusion de estos trabajadores ya esta hecha: solo falta evaluar.
filas = []
for p in people:
    if p['id'] in SITUACION:
        continue
    for pr in procs:
        _, ev = registro.get((p['id'], pr['code']), ('PENDIENTE', 'PENDIENTE'))
        if ev == 'PENDIENTE':
            filas.append((p, 'Evaluación', pr['nombre'], pr['code'], CONF, CONF))
for doc, cod in NUEVOS:
    for p in people:
        filas.append((p, 'Difusión y Evaluación', doc, cod, RANGO_DIF, RANGO_EVA))

n = 0
for p, actividad, doc, cod, rdif, reva in filas:
        n += 1
        r = 4 + n
        ausente = p['id'] in SITUACION
        f_dif = '%s (%s)' % (REINC, SITUACION[p['id']].lower()) if ausente else rdif
        f_eva = '%s (%s)' % (REINC, SITUACION[p['id']].lower()) if ausente else reva
        if actividad == 'Evaluación':
            f_dif = 'Difusión ya realizada'
        vals = [n, p['rut'], nombre_completo(p), p['cargo'], EMPRESA, p['g'],
                actividad, doc, cod, f_dif, f_eva, CONF]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(r, j, v); c.font = F(); c.border = BORDE
            c.alignment = CEN if j in (1, 2, 6, 9, 10, 11) else IZQ
        ws.cell(r, 7).fill = ROJO
        if ausente:
            ws.cell(r, 10).fill = AMAR; ws.cell(r, 11).fill = AMAR
        if f_eva == CONF:
            ws.cell(r, 11).fill = AMAR
        if actividad == 'Evaluación':
            ws.cell(r, 10).fill = VERDE
        if p['rut'] == CONF: ws.cell(r, 2).fill = AMAR
        ws.cell(r, 9).fill = AMAR
        ws.cell(r, 12).fill = AMAR
ws.freeze_panes = 'A5'
fila = 4 + n + 2
notas = [
    'NOTAS A LA PLANILLA 3',
    'Instructivo de Llenado de VAT — difusión comprometida: %s; evaluación comprometida: %s.'
    % (RANGO_DIF, RANGO_EVA),
    'Las evaluaciones pendientes que arrastra la Planilla 2 están POR CONFIRMAR en su fecha '
    'comprometida, a la espera de la programación de cada jefatura de turno.',
    'El código único del documento está POR CONFIRMAR: debe asignarse dentro de la numeración del '
    'proceso 3.8 antes de subir la planilla a SIMIN OL.',
    'El responsable de ejecución está POR CONFIRMAR para cada turno.',
    'Se incluye la columna "Turno" (G1 a G4) porque la difusión debe coordinarse por turno dentro '
    'de la semana comprometida. No forma parte del formato original de la planilla.',
    'Los trabajadores con licencia médica llevan como fecha comprometida "A su reincorporación": '
    + ', '.join(sorted(nombre_completo(p) for p in people if p['id'] in SITUACION)) + '.',
    'Documento 3.8.3.12 "Procedimiento de Emergencia": corresponde al "Plan de Respuesta a Emergencia '
    'Local Valle Copiapó — Puerto Punta Totoralillo", aprobado el 25/02/2025. Por esa razón el "Plan '
    'de Gestión de Riesgos de Desastres" no se incorpora a este programa: es el mismo documento, ya '
    'difundido y evaluado, y figura en las Planillas 1 y 2.',
]
if previas:
    notas.append(
        'Se detectaron %d difusiones anteriores a la fecha de aprobación de su documento. Quedan '
        'marcadas en rojo en la Hoja 2 y corresponden a re-instrucción en la versión vigente.'
        % len(previas))
notas += [
]
for i, t in enumerate(notas):
    c = ws.cell(fila + i, 1, t)
    c.font = FB() if i == 0 else F(size=9)
    c.alignment = IZQ

sin_rut = [nombre_completo(p) for p in people if p['rut'] == CONF]
if sin_rut:
    ws.cell(fila + len(notas), 1, 'Trabajadores con RUT POR CONFIRMAR: ' + ', '.join(sin_rut) + '.').font = F(size=9)

wb.save('/home/user/EVALUACIONES-PROCEDIMIENTO-OPERACIONES-PPT/SNGM_Capacitacion_3.8_Operaciones.xlsx')
print('sin RUT:', sin_rut or 'ninguno')
print('hojas:', wb.sheetnames)
print('trabajadores:', len(people), '| procedimientos:', len(procs), '| filas hoja 3:', n)
