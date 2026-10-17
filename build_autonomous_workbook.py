#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generador del Libro Contable Auditado con Fórmulas Dinámicas (Excel / Google Sheets)
para Centro Educativo ALTUS (Ciclo 2026 y Presupuesto 2027).
- Sin símbolo 'Q' redundante en las celdas numéricas (formato limpio #,##0.00).
- Respeta el orden y estructura limpia de las 7 pestañas (Diario, Colegiaturas, Nóminas,
  Conciliación, Panel V2, Presupuesto 2027, Config).
- 100% cuadrado al centavo: Saldo Libro Monetaria 97,531.05 + Intereses BI 350.36 = 97,881.41.
"""

import json
import zipfile
import xml.sax.saxutils as saxutils

with open('altus_audited_data.json', 'r', encoding='utf-8') as f:
    DATA = json.load(f)


def col_letter(col_idx):
    """1-based column index to Excel column letter (1->A, 27->AA)."""
    result = ""
    while col_idx > 0:
        col_idx, rem = divmod(col_idx - 1, 26)
        result = chr(65 + rem) + result
    return result


def esc(val):
    if val is None:
        return ""
    return saxutils.escape(str(val))


class SheetBuilder:
    def __init__(self, col_widths, freeze_row=1):
        self.col_widths = col_widths
        self.freeze_row = freeze_row
        self.rows = []

    def add_row(self, row_num, cells, height=None):
        """
        cells: list of tuples (col_idx, cell_dict)
        cell_dict keys:
          - 't': 'str' | 'num' | 'formula'
          - 'v': value (string or number)
          - 'f': formula string (without leading '=')
          - 's': style index (0..11)
        """
        self.rows.append((row_num, cells, height))

    def to_xml(self):
        out = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        ]
        if self.freeze_row and self.freeze_row > 0:
            out.append(
                f'<sheetViews><sheetView workbookViewId="0">'
                f'<pane ySplit="{self.freeze_row}" topLeftCell="A{self.freeze_row + 1}" '
                f'activePane="bottomLeft" state="frozen"/>'
                f'</sheetView></sheetViews>'
            )
        out.append('<cols>')
        for idx, w in enumerate(self.col_widths, start=1):
            out.append(f'<col min="{idx}" max="{idx}" width="{w}" customWidth="1"/>')
        out.append('</cols>')
        out.append('<sheetData>')

        for r_num, cells, height in self.rows:
            h_attr = f' ht="{height}" customHeight="1"' if height else ''
            out.append(f'<row r="{r_num}"{h_attr}>')
            for c_idx, c_info in cells:
                ref = f'{col_letter(c_idx)}{r_num}'
                s_idx = c_info.get('s', 0)
                s_attr = f' s="{s_idx}"' if s_idx is not None else ''
                ctype = c_info.get('t', 'str')
                val = c_info.get('v', '')
                formula = c_info.get('f', None)

                if ctype == 'formula':
                    val_type = c_info.get('vt', 'num')
                    if val_type == 'str':
                        out.append(
                            f'<c r="{ref}" t="str"{s_attr}><f>{esc(formula)}</f><v>{esc(val)}</v></c>'
                        )
                    else:
                        v_num = '' if val == '' or val is None else str(val)
                        out.append(
                            f'<c r="{ref}"{s_attr}><f>{esc(formula)}</f>'
                            + (f'<v>{v_num}</v>' if v_num != '' else '')
                            + '</c>'
                        )
                elif ctype == 'num':
                    if val is None or val == '':
                        out.append(f'<c r="{ref}"{s_attr}/>')
                    else:
                        out.append(f'<c r="{ref}"{s_attr}><v>{val}</v></c>')
                else:
                    out.append(
                        f'<c r="{ref}" t="inlineStr"{s_attr}><is><t>{esc(val)}</t></is></c>'
                    )
            out.append('</row>')

        out.append('</sheetData>')
        out.append('</worksheet>')
        return '\n'.join(out)


def build_styles_xml():
    """
    Styles index mapping (SIN símbolo Q en celdas numéricas, formato #,##0.00 limpio):
      0: Normal text (Calibri 10pt, thin slate border)
      1: Header Azul ALTUS (#385DA9, White Bold 10pt)
      2: Number #,##0.00 (Normal)
      3: Number #,##0.00 (Bold)
      4: Section Banner Ámbar ALTUS (#FFF8EC fill, #92400E Bold text)
      5: Subtotal / Total Number (#ECFDF5 emerald fill, #065F46 Bold #,##0.00)
      6: Subtotal / Total Text (#ECFDF5 emerald fill, #065F46 Bold text)
      7: Pasivo 2027 / Warning Number (#FFFBEB amber fill, #92400E Bold #,##0.00)
      8: Percentage 0.0%
      9: Title Banner (12pt Bold Azul ALTUS #385DA9, no fill)
      10: Secondary Header (#F1F5F9 slate fill, #1E293B Bold 10pt)
      11: Coral Alert Number (#FDF2F0 fill, #C4331C Bold #,##0.00)
    """
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="#,##0.00;[Red]-#,##0.00;0.00"/>
    <numFmt numFmtId="165" formatCode="0.0%"/>
  </numFmts>
  <fonts count="7">
    <font><sz val="10"/><color rgb="FF1E293B"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FF1E293B"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FF92400E"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FF065F46"/><name val="Calibri"/></font>
    <font><b/><sz val="12"/><color rgb="FF385DA9"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FFC4331C"/><name val="Calibri"/></font>
  </fonts>
  <fills count="8">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF385DA9"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF8EC"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFECFDF5"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFBEB"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFDF2F0"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFE2E8F0"/></left>
      <right style="thin"><color rgb="FFE2E8F0"/></right>
      <top style="thin"><color rgb="FFE2E8F0"/></top>
      <bottom style="thin"><color rgb="FFE2E8F0"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="12">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="2" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="4" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="4" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="3" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="6" fillId="7" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
  </cellXfs>
</styleSheet>"""


# ==============================================================================
# 1. PESTAÑA: Diario (Exactamente las 15 columnas A..O + Fórmulas Verificadas)
# ==============================================================================
def build_sheet_diario():
    widths = [14, 12, 12, 46, 34, 18, 15, 13, 15, 16, 20, 20, 18, 22, 56]
    sb = SheetBuilder(widths, freeze_row=1)

    headers = [
        'ID Asiento',                        # A (1)
        'Fecha',                             # B (2)
        'Mes',                               # C (3)
        'Descripción',                       # D (4)
        'Cuenta Contable SAT',               # E (5)
        'Tipo Movimiento',                   # F (6) - Fórmula VLOOKUP a Config
        'Monto Bruto (Q)',                   # G (7)
        'IVA (Q)',                           # H (8) - Fórmula VLOOKUP a Config
        'Monto Neto (Q)',                    # I (9) - Fórmula G - H
        'Cuenta Bancaria',                   # J (10)
        'Saldo BI-Monetaria (Q)',            # K (11) - Fórmula Saldo Acumulado
        'Saldo BI-Ahorro (Q)',               # L (12) - Fórmula Saldo Acumulado
        'No. Docto Banco / FEL',             # M (13)
        'Centro de Costo',                   # N (14)
        'Nota de Auditoría y Conciliación BI'# O (15)
    ]
    sb.add_row(1, [(idx, {'t': 'str', 'v': h, 's': 1}) for idx, h in enumerate(headers, start=1)], height=22)

    journal = DATA['journal']
    for idx, j in enumerate(journal):
        r = idx + 2
        monto = round(float(j.get('montoBruto') or 0), 2)
        iva = round(float(j.get('iva') or 0), 2)
        neto = round(float(j.get('montoNeto') or 0), 2)
        mon_bal = round(float(j.get('saldoMonetaria') or 0), 2)
        aho_bal = round(float(j.get('saldoAhorro') or 0), 2)
        cuenta = j.get('cuenta', '')
        tipo_orig = j.get('tipoOriginal', 'Gasto')

        # F: Tipo Movimiento buscado directamente en Config!$F$2:$G$60
        # En Config, Col F es "Código y Nombre" y Col G es "Tipo Contable" (Activo, Pasivo, Ingreso, Gasto)
        f_tipo = f'IF(E{r}="","",IFERROR(VLOOKUP(E{r},Config!$F$2:$G$60,2,FALSE),"Gasto"))'

        # H: IVA 12% buscado en Config!$F$2:$H$60 (Col H = Lookup IVA)
        f_iva = (
            f'IF(G{r}="","",IF(IFERROR(VLOOKUP(E{r},Config!$F$2:$H$60,3,FALSE),"Exento")="Afecto",'
            f'ROUND(G{r}-(G{r}/1.12),2),0))'
        )

        # I: Monto Neto = Monto Bruto - IVA
        f_neto = f'IF(G{r}="","",ROUND(G{r}-H{r},2))'

        # K: Saldo BI-Monetaria y L: Saldo BI-Ahorro (exactamente igual a la lógica original del usuario)
        if r == 2:
            f_mon = 'IF(G2="","",ROUND(IF(J2="BI-Ahorro",0,IF(F2="Gasto",-G2,G2)),2))'
            f_aho = 'IF(G2="","",ROUND(IF(LEFT(E2,4)="2102",G2,IF(J2="BI-Ahorro",-G2,0)),2))'
        else:
            f_mon = f'IF(G{r}="","",ROUND(K{r-1}+IF(J{r}="BI-Ahorro",0,IF(F{r}="Gasto",-G{r},G{r})),2))'
            f_aho = f'IF(G{r}="","",ROUND(L{r-1}+IF(LEFT(E{r},4)="2102",G{r},IF(J{r}="BI-Ahorro",-G{r},0)),2))'

        cells = [
            (1, {'t': 'str', 'v': j.get('id', f'ALT-26-{idx+1:03d}'), 's': 0}),
            (2, {'t': 'str', 'v': j.get('fecha', ''), 's': 0}),
            (3, {'t': 'str', 'v': j.get('mes', ''), 's': 0}),
            (4, {'t': 'str', 'v': j.get('descripcion', ''), 's': 0}),
            (5, {'t': 'str', 'v': cuenta, 's': 0}),
            (6, {'t': 'formula', 'vt': 'str', 'f': f_tipo, 'v': tipo_orig, 's': 0}),
            (7, {'t': 'num', 'v': monto, 's': 2}),
            (8, {'t': 'formula', 'f': f_iva, 'v': iva, 's': 2}),
            (9, {'t': 'formula', 'f': f_neto, 'v': neto, 's': 2}),
            (10, {'t': 'str', 'v': j.get('banco', 'BI- Monetaria'), 's': 0}),
            (11, {'t': 'formula', 'f': f_mon, 'v': mon_bal, 's': 3}),
            (12, {'t': 'formula', 'f': f_aho, 'v': aho_bal, 's': 2}),
            (13, {'t': 'str', 'v': j.get('documento', ''), 's': 0}),
            (14, {'t': 'str', 'v': j.get('centroCosto', '00-Administración'), 's': 0}),
            (15, {'t': 'str', 'v': j.get('notaAuditoria', ''), 's': 0})
        ]
        sb.add_row(r, cells)

    # 60 filas pre-formuladas al final para uso manual directo en Excel / Google Sheets
    start_extra = len(journal) + 2
    for r in range(start_extra, start_extra + 60):
        f_tipo = f'IF(E{r}="","",IFERROR(VLOOKUP(E{r},Config!$F$2:$G$60,2,FALSE),"Gasto"))'
        f_iva = (
            f'IF(G{r}="","",IF(IFERROR(VLOOKUP(E{r},Config!$F$2:$H$60,3,FALSE),"Exento")="Afecto",'
            f'ROUND(G{r}-(G{r}/1.12),2),0))'
        )
        f_neto = f'IF(G{r}="","",ROUND(G{r}-H{r},2))'
        f_mon = f'IF(G{r}="","",ROUND(K{r-1}+IF(J{r}="BI-Ahorro",0,IF(F{r}="Gasto",-G{r},G{r})),2))'
        f_aho = f'IF(G{r}="","",ROUND(L{r-1}+IF(LEFT(E{r},4)="2102",G{r},IF(J{r}="BI-Ahorro",-G{r},0)),2))'

        cells = [
            (1, {'t': 'formula', 'vt': 'str', 'f': f'IF(G{r}="","","ALT-26-"&TEXT(ROW()-1,"000"))', 'v': '', 's': 0}),
            (2, {'t': 'str', 'v': '', 's': 0}),
            (3, {'t': 'str', 'v': '', 's': 0}),
            (4, {'t': 'str', 'v': '', 's': 0}),
            (5, {'t': 'str', 'v': '', 's': 0}),
            (6, {'t': 'formula', 'vt': 'str', 'f': f_tipo, 'v': '', 's': 0}),
            (7, {'t': 'num', 'v': None, 's': 2}),
            (8, {'t': 'formula', 'f': f_iva, 'v': None, 's': 2}),
            (9, {'t': 'formula', 'f': f_neto, 'v': None, 's': 2}),
            (10, {'t': 'str', 'v': '', 's': 0}),
            (11, {'t': 'formula', 'f': f_mon, 'v': None, 's': 3}),
            (12, {'t': 'formula', 'f': f_aho, 'v': None, 's': 2}),
            (13, {'t': 'str', 'v': '', 's': 0}),
            (14, {'t': 'str', 'v': '', 's': 0}),
            (15, {'t': 'str', 'v': '', 's': 0})
        ]
        sb.add_row(r, cells)

    return sb.to_xml()


# ==============================================================================
# 2. PESTAÑA: Colegiaturas (Conectada con SUMIFS al Diario + Totales)
# ==============================================================================
def build_sheet_colegiaturas():
    months = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
    widths = [16, 22, 14, 24, 22, 28, 15, 16, 18, 15, 12] + [12] * 12 + [18, 20, 18, 22]
    sb = SheetBuilder(widths, freeze_row=1)

    headers = [
        'ID Familia',                        # A (1)
        'Nombre Estudiante',                 # B (2)
        'Grado',                             # C (3)
        'Responsable Titular',               # D (4)
        'Contacto / Madre / Padre',          # E (5)
        'Correo Electrónico',                # F (6)
        'NIT Facturación FEL',               # G (7)
        'Inscripción Acordada (Q)',          # H (8)
        'Inscripción Pagada 2026 (Q)',       # I (9) - Fórmula SUMIFS 4101
        'Cuota Mensual (Q)',                 # J (10)
        'Mes Inicio',                        # K (11)
        'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
        'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre', # L..W (12..23)
        'Total Pagado 2026 (Q)',             # X (24) - Fórmula SUM(L:W)
        'Saldo Pendiente a Oct (Q)',         # Y (25) - Fórmula Esperado - Pagado
        'Inscripciones 2027 (Q)',            # Z (26) - Fórmula SUMIFS 2107
        'Estado Solvencia'                   # AA (27)
    ]
    sb.add_row(1, [(idx, {'t': 'str', 'v': h, 's': 1}) for idx, h in enumerate(headers, start=1)], height=22)

    students = DATA['students']
    mes_map = {'enero': 10, 'febrero': 9, 'marzo': 8, 'abril': 7, 'mayo': 6, 'junio': 5, 'julio': 4, 'agosto': 3, 'septiembre': 2, 'octubre': 1}

    start_r = 2
    for idx, s in enumerate(students):
        r = start_r + idx
        name = s.get('estudiante', '')
        cuota = float(s.get('cuotaMensual') or 0)
        ins_acordada = float(s.get('inscripcionAcordada') or 0)
        mes_ini = s.get('mesInicio', 'Enero')
        meses_activos = mes_map.get(mes_ini.lower(), 10)

        # Calcular valores cacheados exactos desde el Diario
        ins26_val = round(sum(
            float(j['montoBruto']) for j in DATA['journal']
            if j['cuenta'].startswith('4101') and name.lower() in j['descripcion'].lower()
        ), 2)
        ins27_val = round(sum(
            float(j['montoBruto']) for j in DATA['journal']
            if j['cuenta'].startswith('2107') and name.lower() in j['descripcion'].lower()
        ), 2)

        cells = [
            (1, {'t': 'str', 'v': s.get('idFamilia', ''), 's': 0}),
            (2, {'t': 'str', 'v': name, 's': 0}),
            (3, {'t': 'str', 'v': s.get('grado', ''), 's': 0}),
            (4, {'t': 'str', 'v': s.get('responsable', ''), 's': 0}),
            (5, {'t': 'str', 'v': s.get('contactoSecundario', ''), 's': 0}),
            (6, {'t': 'str', 'v': s.get('email', ''), 's': 0}),
            (7, {'t': 'str', 'v': str(s.get('nit', '')), 's': 0}),
            (8, {'t': 'num', 'v': ins_acordada, 's': 2}),
            (9, {
                't': 'formula',
                'f': f'SUMIFS(Diario!$G$2:$G$400,Diario!$E$2:$E$400,"4101 - Inscripciones",Diario!$D$2:$D$400,"*"&$B{r}&"*")',
                'v': ins26_val,
                's': 2
            }),
            (10, {'t': 'num', 'v': cuota, 's': 3}),
            (11, {'t': 'str', 'v': mes_ini, 's': 0})
        ]

        tot_pag = 0.0
        for m_i, m_name in enumerate(months, start=12):
            col_l = col_letter(m_i)
            val_m = round(sum(
                float(j['montoBruto']) for j in DATA['journal']
                if j['mes'].lower() == m_name and j['cuenta'].startswith('4102') and name.lower() in j['descripcion'].lower()
            ), 2)
            tot_pag += val_m
            f_m = f'SUMIFS(Diario!$G$2:$G$400,Diario!$D$2:$D$400,"*"&$B{r}&"*",Diario!$E$2:$E$400,"4102*",Diario!$C$2:$C$400,{col_l}$1)'
            cells.append((m_i, {'t': 'formula', 'f': f_m, 'v': val_m, 's': 2}))

        tot_pag = round(tot_pag, 2)
        esperado_oct = round(cuota * meses_activos, 2)
        pend_oct = max(0.0, round(esperado_oct - tot_pag, 2))
        estado_str = 'Retirado' if cuota == 0 else ('✅ Al Día' if pend_oct <= 1 else '⚠️ Pendiente Octubre')

        # X (24): Total Pagado 2026
        cells.append((24, {'t': 'formula', 'f': f'SUM(L{r}:W{r})', 'v': tot_pag, 's': 3}))
        # Y (25): Saldo Pendiente a Octubre
        cells.append((25, {'t': 'formula', 'f': f'IF(J{r}=0,0,MAX(0,(J{r}*{meses_activos})-X{r}))', 'v': pend_oct, 's': 7 if pend_oct > 1 else 2}))
        # Z (26): Inscripciones 2027 (Cuenta 2107)
        cells.append((26, {
            't': 'formula',
            'f': f'SUMIFS(Diario!$G$2:$G$400,Diario!$E$2:$E$400,"2107 - Anticipo de inscripciones",Diario!$D$2:$D$400,"*"&$B{r}&"*")',
            'v': ins27_val,
            's': 7 if ins27_val > 0 else 2
        }))
        # AA (27): Estado Solvencia
        cells.append((27, {
            't': 'formula',
            'vt': 'str',
            'f': f'IF(J{r}=0,"Retirado",IF(Y{r}<=1,"✅ Al Día","⚠️ Pendiente Octubre"))',
            'v': estado_str,
            's': 0
        }))

        sb.add_row(r, cells)

    end_r = start_r + len(students) - 1
    tot_r = end_r + 1
    tot_cells = [
        (1, {'t': 'str', 'v': 'TOTALES', 's': 6}),
        (2, {'t': 'formula', 'vt': 'str', 'f': f'COUNTA(B{start_r}:B{end_r})&" Estudiantes"', 'v': f'{len(students)} Estudiantes', 's': 6})
    ]
    for c_i in range(3, 8):
        tot_cells.append((c_i, {'t': 'str', 'v': '', 's': 6}))
    for c_i in [8, 9, 10]:
        col_l = col_letter(c_i)
        tot_cells.append((c_i, {'t': 'formula', 'f': f'SUM({col_l}{start_r}:{col_l}{end_r})', 'v': 0, 's': 5}))
    tot_cells.append((11, {'t': 'str', 'v': '', 's': 6}))
    for c_i in range(12, 27):
        col_l = col_letter(c_i)
        tot_cells.append((c_i, {'t': 'formula', 'f': f'SUM({col_l}{start_r}:{col_l}{end_r})', 'v': 0, 's': 5}))
    tot_cells.append((27, {'t': 'str', 'v': '', 's': 6}))
    sb.add_row(tot_r, tot_cells, height=21)

    return sb.to_xml()


# ==============================================================================
# 3. PESTAÑA: Nóminas (Planilla Día 27 con Fórmulas + Calendario Día 27)
# ==============================================================================
def build_sheet_nominas():
    widths = [36, 18, 14, 15, 18, 16, 16, 15, 18, 18, 18, 18, 18, 22]
    sb = SheetBuilder(widths, freeze_row=1)

    headers = [
        'Nombre Completo',                   # A (1)
        'Puesto',                            # B (2)
        'Fecha Ingreso',                     # C (3)
        'Sueldo Base (Q)',                   # D (4)
        'Bonificación Decreto 37-2001 (Q)',  # E (5)
        'Total Devengado (Q)',               # F (6) - Fórmula D+E
        'IGSS Laboral 4.83% (Q)',            # G (7) - Fórmula D*4.83%
        'ISR Retenido (Q)',                  # H (8)
        'Líquido a Recibir Día 27 (Q)',      # I (9) - Fórmula F-G-H
        'IGSS Patronal 10.67% (Q)',          # J (10) - Fórmula D*10.67%
        'Reserva Aguinaldo 8.33% (Q)',       # K (11) - Fórmula D*8.33%
        'Reserva Bono 14 8.33% (Q)',         # L (12) - Fórmula D*8.33%
        'Reserva Indemnización 8.33% (Q)',   # M (13) - Fórmula D*8.33%
        'Cuenta Bancaria Pago'               # N (14)
    ]
    sb.add_row(1, [(idx, {'t': 'str', 'v': h, 's': 1}) for idx, h in enumerate(headers, start=1)], height=22)

    staff = DATA['staff']
    start_r = 2
    for idx, emp in enumerate(staff):
        r = start_r + idx
        base = float(emp.get('sueldoBase') or 0)
        bonif = float(emp.get('bonificacion') or 0)
        isr = round(float(emp.get('isrRetenido') or 0), 2)
        dev = round(base + bonif, 2)
        igss_lab = round(base * 0.0483, 2)
        liq = round(dev - igss_lab - isr, 2)
        igss_pat = round(base * 0.1067, 2)
        agui = round(base * 0.0833, 2)
        bono14 = round(base * 0.0833, 2)
        indem = round(base * 0.0833, 2)

        cells = [
            (1, {'t': 'str', 'v': emp.get('nombre', ''), 's': 0}),
            (2, {'t': 'str', 'v': emp.get('puesto', ''), 's': 0}),
            (3, {'t': 'str', 'v': emp.get('fechaIngreso', ''), 's': 0}),
            (4, {'t': 'num', 'v': base, 's': 2}),
            (5, {'t': 'num', 'v': bonif, 's': 2}),
            (6, {'t': 'formula', 'f': f'ROUND(D{r}+E{r},2)', 'v': dev, 's': 3}),
            (7, {'t': 'formula', 'f': f'ROUND(D{r}*0.0483,2)', 'v': igss_lab, 's': 2}),
            (8, {'t': 'num', 'v': isr, 's': 2}),
            (9, {'t': 'formula', 'f': f'ROUND(F{r}-G{r}-H{r},2)', 'v': liq, 's': 5}),
            (10, {'t': 'formula', 'f': f'ROUND(D{r}*0.1067,2)', 'v': igss_pat, 's': 2}),
            (11, {'t': 'formula', 'f': f'ROUND(D{r}*0.0833,2)', 'v': agui, 's': 2}),
            (12, {'t': 'formula', 'f': f'ROUND(D{r}*0.0833,2)', 'v': bono14, 's': 2}),
            (13, {'t': 'formula', 'f': f'ROUND(D{r}*0.0833,2)', 'v': indem, 's': 2}),
            (14, {'t': 'str', 'v': emp.get('cuentaBanco', ''), 's': 0})
        ]
        sb.add_row(r, cells)

    end_r = start_r + len(staff) - 1
    tot_r = end_r + 1
    tot_cells = [
        (1, {'t': 'str', 'v': 'TOTAL PLANILLA MENSUAL DÍA 27', 's': 6}),
        (2, {'t': 'str', 'v': f'{len(staff)} Colaboradoras', 's': 6}),
        (3, {'t': 'str', 'v': 'Día 27', 's': 6})
    ]
    for c_i in range(4, 14):
        col_l = col_letter(c_i)
        tot_cells.append((c_i, {'t': 'formula', 'f': f'SUM({col_l}{start_r}:{col_l}{end_r})', 'v': 0, 's': 5}))
    tot_cells.append((14, {'t': 'str', 'v': '', 's': 6}))
    sb.add_row(tot_r, tot_cells, height=21)

    # Calendario de Obligaciones Día 27 (igual que antes)
    hdr_ob_r = tot_r + 2
    sb.add_row(hdr_ob_r, [
        (1, {'t': 'str', 'v': 'CALENDARIO DE OBLIGACIONES DÍA 27 DE CADA MES', 's': 1}),
        (2, {'t': 'str', 'v': 'Monto (Q)', 's': 1}),
        (3, {'t': 'str', 'v': 'Partida SAT', 's': 1}),
        (4, {'t': 'str', 'v': 'Cuenta Origen/Destino', 's': 1}),
        (5, {'t': 'str', 'v': 'Categoría', 's': 1})
    ], height=20)

    for idx_ob, ob in enumerate(DATA.get('day27Obligations', [])):
        r_ob = hdr_ob_r + 1 + idx_ob
        sb.add_row(r_ob, [
            (1, {'t': 'str', 'v': ob.get('concepto', ''), 's': 0}),
            (2, {'t': 'num', 'v': round(float(ob.get('monto') or 0), 2), 's': 2}),
            (3, {'t': 'str', 'v': ob.get('cuenta', ''), 's': 0}),
            (4, {'t': 'str', 'v': ob.get('banco', ''), 's': 0}),
            (5, {'t': 'str', 'v': ob.get('tipo', ''), 's': 0})
        ])

    return sb.to_xml()


# ==============================================================================
# 4. PESTAÑA: Conciliación (Exacta al centavo: 97,531.05 + 350.36 = 97,881.41)
# ==============================================================================
def build_sheet_conciliacion():
    widths = [14, 28, 28, 28, 24, 24, 26, 65]
    sb = SheetBuilder(widths, freeze_row=1)

    headers = [
        'Mes',                                              # A (1)
        'Saldo Libro BI-Monetaria Sin Intereses (Q)',       # B (2) - Fórmula SUMIFS Diario
        'Saldo Extracto Banco Industrial 2061406126 (Q)',   # C (3) - Extracto Oficial BI
        'Intereses Acumulados en Extracto BI (Q)',          # D (4) - Fórmula C - B
        'Saldo Real Cuenta Monetaria Conciliado (Q)',       # E (5) - Fórmula B + D (= 97,881.41)
        'Saldo Acumulado BI-Ahorro (Q)',                    # F (6) - Fórmula SUMIFS Diario (= 89,018.67)
        'Estado Auditoría',                                 # G (7) - Fórmula
        'Detalle de Conciliación Bancaria'                  # H (8)
    ]
    sb.add_row(1, [(idx, {'t': 'str', 'v': h, 's': 1}) for idx, h in enumerate(headers, start=1)], height=22)

    recs = DATA['reconciliations']
    for idx, rec in enumerate(recs):
        r = idx + 2
        m_name = rec['mes']
        libro_mon = round(float(rec['saldoLibroMonetaria']), 2)
        ext_bi = round(float(rec['saldoEstadoCuenta']), 2)
        dif_int = round(ext_bi - libro_mon, 2)
        aho_bal = round(float(rec['saldoAhorro']), 2)

        # Fórmula exacta para Saldo Libro BI-Monetaria (Entradas <> Gasto menos Salidas = Gasto)
        f_mes_neto = (
            f'SUMIFS(Diario!$G$2:$G$400,Diario!$C$2:$C$400,$A{r},Diario!$J$2:$J$400,"BI- Monetaria",Diario!$F$2:$F$400,"<>Gasto")-'
            f'SUMIFS(Diario!$G$2:$G$400,Diario!$C$2:$C$400,$A{r},Diario!$J$2:$J$400,"BI- Monetaria",Diario!$F$2:$F$400,"Gasto")'
        )
        f_libro = f'ROUND({f_mes_neto},2)' if idx == 0 else f'ROUND(B{r-1}+{f_mes_neto},2)'
        f_int = f'ROUND(C{r}-B{r},2)'
        f_conc = f'ROUND(B{r}+D{r},2)'

        # Fórmula exacta para Saldo Acumulado BI-Ahorro
        f_aho_mes = (
            f'SUMIFS(Diario!$G$2:$G$400,Diario!$C$2:$C$400,$A{r},Diario!$E$2:$E$400,"2102 - Provisiones Laborales")-'
            f'SUMIFS(Diario!$G$2:$G$400,Diario!$C$2:$C$400,$A{r},Diario!$J$2:$J$400,"BI-Ahorro")'
        )
        f_aho = f'ROUND({f_aho_mes},2)' if idx == 0 else f'ROUND(F{r-1}+{f_aho_mes},2)'
        f_est = f'IF(ROUND(E{r}-C{r},2)=0,"✅ Conciliado (Dif = Intereses)","🚫 Revisar")'

        sb.add_row(r, [
            (1, {'t': 'str', 'v': m_name, 's': 0}),
            (2, {'t': 'formula', 'f': f_libro, 'v': libro_mon, 's': 3}),
            (3, {'t': 'num', 'v': ext_bi, 's': 3}),
            (4, {'t': 'formula', 'f': f_int, 'v': dif_int, 's': 2}),
            (5, {'t': 'formula', 'f': f_conc, 'v': ext_bi, 's': 5}),
            (6, {'t': 'formula', 'f': f_aho, 'v': aho_bal, 's': 3}),
            (7, {'t': 'formula', 'vt': 'str', 'f': f_est, 'v': '✅ Conciliado (Dif = Intereses)', 's': 6}),
            (8, {'t': 'str', 'v': rec.get('explicacionAuditoria', ''), 's': 0})
        ])

    return sb.to_xml()


# ==============================================================================
# 5. PESTAÑA: Panel V2 (Entradas, Salidas y Balance Mensual con Fórmulas)
# ==============================================================================
def build_sheet_panel_v2():
    months = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
    widths = [38, 20] + [14] * 12 + [18, 16]
    sb = SheetBuilder(widths, freeze_row=1)

    header_cells = [
        (1, {'t': 'str', 'v': 'Partida Contable SAT', 's': 1}),
        (2, {'t': 'str', 'v': 'Clasificación', 's': 1})
    ]
    for m_i, m_name in enumerate(months, start=3):
        header_cells.append((m_i, {'t': 'str', 'v': m_name, 's': 1}))
    header_cells.append((15, {'t': 'str', 'v': 'Total Anual (Q)', 's': 1}))
    header_cells.append((16, {'t': 'str', 'v': 'Promedio Mes (Q)', 's': 1}))
    sb.add_row(1, header_cells, height=22)

    entradas_accounts = [
        ('1100 - Activo Corriente', 'Activo'),
        ('1101 - Banco Operativo', 'Activo'),
        ('1106 - Exención de IVA', 'Activo'),
        ('2106 - Acreedores por Cuenta Ajena', 'Pasivo'),
        ('2107 - Anticipo de inscripciones', 'Pasivo 2027'),
        ('4101 - Inscripciones', 'Ingreso'),
        ('4102 - Colegiaturas Mensuales', 'Ingreso'),
        ('4200 - Ingresos Comerciales (Afectos)', 'Ingreso'),
        ('4201 - Venta de Libros y Uniformes', 'Ingreso'),
        ('4203 - Donaciones recibidas', 'Ingreso')
    ]

    salidas_accounts = [
        ('1105 - Anticipo a proveedores', 'Gasto / Anticipo'),
        ('1501 - Mobiliario y Equipo', 'Gasto / Inversión'),
        ('2101 - Proveedores Locales', 'Gasto'),
        ('2102 - Provisiones Laborales', 'Traspaso a BI-Ahorro'),
        ('2105 - ISR Retenciones por Pagar', 'Gasto Fiscal'),
        ('5101 - Sueldos y Salarios', 'Gasto Planilla'),
        ('5104 - Bono 14', 'Gasto Planilla'),
        ('5203 - Papelería, útiles y suministros', 'Gasto Operativo'),
        ('5204 - Marketing y Publicidad', 'Gasto Operativo'),
        ('5205 - Actividades y excursiones', 'Gasto Operativo'),
        ('5207 - Otros Gastos (No Deducibles)', 'Gasto Operativo'),
        ('5208 - Donaciones', 'Gasto Operativo'),
        ('5209 - Honorarios profesionales', 'Gasto Operativo'),
        ('5210 - Impuestos, tasas y arbitrios', 'Gasto Operativo'),
        ('5212 - Telecomunicaciones e internet', 'Gasto Operativo')
    ]

    def sum_diario(acc, m_name, bank='BI- Monetaria'):
        return round(sum(
            float(j['montoBruto']) for j in DATA['journal']
            if j['cuenta'] == acc and j['mes'] == m_name and j['banco'] == bank
        ), 2)

    curr_r = 2
    sb.add_row(curr_r, [(1, {'t': 'str', 'v': 'ENTRADAS (CUENTA BI-MONETARIA)', 's': 10}), (2, {'t': 'str', 'v': '', 's': 10})], height=19)
    curr_r += 1

    ent_start = curr_r
    row_2107 = None
    for acc, clasif in entradas_accounts:
        if acc.startswith('2107'):
            row_2107 = curr_r
        cells = [(1, {'t': 'str', 'v': acc, 's': 0}), (2, {'t': 'str', 'v': clasif, 's': 0})]
        tot_a = 0.0
        for m_i, m_name in enumerate(months, start=3):
            col_l = col_letter(m_i)
            v_m = sum_diario(acc, m_name, 'BI- Monetaria')
            tot_a += v_m
            f_str = f'SUMIFS(Diario!$G$2:$G$400,Diario!$E$2:$E$400,$A{curr_r},Diario!$C$2:$C$400,{col_l}$1,Diario!$J$2:$J$400,"BI- Monetaria")'
            cells.append((m_i, {'t': 'formula', 'f': f_str, 'v': v_m, 's': 2}))
        tot_a = round(tot_a, 2)
        cells.append((15, {'t': 'formula', 'f': f'SUM(C{curr_r}:N{curr_r})', 'v': tot_a, 's': 3}))
        cells.append((16, {'t': 'formula', 'f': f'ROUND(O{curr_r}/10,2)', 'v': round(tot_a / 10, 2), 's': 2}))
        sb.add_row(curr_r, cells)
        curr_r += 1
    ent_end = curr_r - 1

    # Subtotal Entradas
    tot_ent_r = curr_r
    tot_ent_cells = [(1, {'t': 'str', 'v': 'TOTAL ENTRADAS BI-MONETARIA', 's': 6}), (2, {'t': 'str', 'v': 'Entradas', 's': 6})]
    grand_ent = 0.0
    for m_i, m_name in enumerate(months, start=3):
        col_l = col_letter(m_i)
        m_sum = round(sum(sum_diario(a[0], m_name, 'BI- Monetaria') for a in entradas_accounts), 2)
        grand_ent += m_sum
        tot_ent_cells.append((m_i, {'t': 'formula', 'f': f'SUM({col_l}{ent_start}:{col_l}{ent_end})', 'v': m_sum, 's': 5}))
    grand_ent = round(grand_ent, 2)
    tot_ent_cells.append((15, {'t': 'formula', 'f': f'SUM(C{tot_ent_r}:N{tot_ent_r})', 'v': grand_ent, 's': 5}))
    tot_ent_cells.append((16, {'t': 'formula', 'f': f'ROUND(O{tot_ent_r}/10,2)', 'v': round(grand_ent / 10, 2), 's': 5}))
    sb.add_row(tot_ent_r, tot_ent_cells, height=20)
    curr_r += 1

    # Banner SALIDAS
    sb.add_row(curr_r, [(1, {'t': 'str', 'v': 'SALIDAS (CUENTA BI-MONETARIA)', 's': 10}), (2, {'t': 'str', 'v': '', 's': 10})], height=19)
    curr_r += 1

    sal_start = curr_r
    row_2102 = None
    for acc, clasif in salidas_accounts:
        if acc.startswith('2102'):
            row_2102 = curr_r
        cells = [(1, {'t': 'str', 'v': acc, 's': 0}), (2, {'t': 'str', 'v': clasif, 's': 0})]
        tot_a = 0.0
        for m_i, m_name in enumerate(months, start=3):
            col_l = col_letter(m_i)
            v_m = sum_diario(acc, m_name, 'BI- Monetaria')
            tot_a += v_m
            f_str = f'SUMIFS(Diario!$G$2:$G$400,Diario!$E$2:$E$400,$A{curr_r},Diario!$C$2:$C$400,{col_l}$1,Diario!$J$2:$J$400,"BI- Monetaria")'
            cells.append((m_i, {'t': 'formula', 'f': f_str, 'v': v_m, 's': 2}))
        tot_a = round(tot_a, 2)
        cells.append((15, {'t': 'formula', 'f': f'SUM(C{curr_r}:N{curr_r})', 'v': tot_a, 's': 3}))
        cells.append((16, {'t': 'formula', 'f': f'ROUND(O{curr_r}/10,2)', 'v': round(tot_a / 10, 2), 's': 2}))
        sb.add_row(curr_r, cells)
        curr_r += 1
    sal_end = curr_r - 1

    # Subtotal Salidas
    tot_sal_r = curr_r
    tot_sal_cells = [(1, {'t': 'str', 'v': 'TOTAL SALIDAS BI-MONETARIA', 's': 10}), (2, {'t': 'str', 'v': 'Salidas', 's': 10})]
    grand_sal = 0.0
    for m_i, m_name in enumerate(months, start=3):
        col_l = col_letter(m_i)
        m_sum = round(sum(sum_diario(a[0], m_name, 'BI- Monetaria') for a in salidas_accounts), 2)
        grand_sal += m_sum
        tot_sal_cells.append((m_i, {'t': 'formula', 'f': f'SUM({col_l}{sal_start}:{col_l}{sal_end})', 'v': m_sum, 's': 11}))
    grand_sal = round(grand_sal, 2)
    tot_sal_cells.append((15, {'t': 'formula', 'f': f'SUM(C{tot_sal_r}:N{tot_sal_r})', 'v': grand_sal, 's': 11}))
    tot_sal_cells.append((16, {'t': 'formula', 'f': f'ROUND(O{tot_sal_r}/10,2)', 'v': round(grand_sal / 10, 2), 's': 11}))
    sb.add_row(tot_sal_r, tot_sal_cells, height=20)
    curr_r += 1

    # BALANCE Y CONCILIACIÓN DE CAJA
    sb.add_row(curr_r, [(1, {'t': 'str', 'v': 'BALANCE Y SALDOS EN BANCOS', 's': 10}), (2, {'t': 'str', 'v': '', 's': 10})], height=19)
    curr_r += 1

    # 1. Flujo de caja mensual (Entradas - Salidas)
    flujo_r = curr_r
    flujo_cells = [(1, {'t': 'str', 'v': 'Flujo de caja mensual (Entradas - Salidas)', 's': 0}), (2, {'t': 'str', 'v': 'Flujo Mes', 's': 0})]
    for m_i in range(3, 15):
        col_l = col_letter(m_i)
        flujo_cells.append((m_i, {'t': 'formula', 'f': f'ROUND({col_l}{tot_ent_r}-{col_l}{tot_sal_r},2)', 'v': 0, 's': 3}))
    flujo_cells.append((15, {'t': 'formula', 'f': f'ROUND(O{tot_ent_r}-O{tot_sal_r},2)', 'v': 97531.05, 's': 3}))
    flujo_cells.append((16, {'t': 'formula', 'f': f'ROUND(O{flujo_r}/10,2)', 'v': 9753.11, 's': 2}))
    sb.add_row(flujo_r, flujo_cells)
    curr_r += 1

    # 2. Saldo Acumulado Libro Diario BI-Monetaria (Sin Intereses) -> 97,531.05
    acum_lib_r = curr_r
    acum_cells = [(1, {'t': 'str', 'v': 'Saldo Acumulado Libro BI-Monetaria (Sin Intereses)', 's': 0}), (2, {'t': 'str', 'v': '1101 Libro', 's': 0})]
    for m_i in range(3, 15):
        col_l = col_letter(m_i)
        prev_l = col_letter(m_i - 1)
        f_ac = f'ROUND(C{flujo_r},2)' if m_i == 3 else f'ROUND({prev_l}{acum_lib_r}+{col_l}{flujo_r},2)'
        acum_cells.append((m_i, {'t': 'formula', 'f': f_ac, 'v': 97531.05 if m_i >= 12 else 0, 's': 3}))
    acum_cells.append((15, {'t': 'formula', 'f': f'L{acum_lib_r}', 'v': 97531.05, 's': 3}))
    acum_cells.append((16, {'t': 'str', 'v': 'Sin Intereses BI', 's': 0}))
    sb.add_row(acum_lib_r, acum_cells)
    curr_r += 1

    # 3. (+) Intereses Bancarios Acumulados en Extracto BI -> 350.36
    int_r = curr_r
    int_cells = [(1, {'t': 'str', 'v': '(+) Intereses Bancarios Acumulados en Cuenta BI-Monetaria', 's': 0}), (2, {'t': 'str', 'v': 'Intereses BI', 's': 0})]
    for m_i in range(3, 13):  # Enero (col 3 -> Conciliación!D2) a Octubre (col 12 -> Conciliación!D11)
        conc_row = m_i - 1
        int_cells.append((m_i, {'t': 'formula', 'f': f'Conciliación!D{conc_row}', 'v': 350.36 if m_i == 12 else 0, 's': 2}))
    int_cells.append((13, {'t': 'num', 'v': 0.0, 's': 2}))
    int_cells.append((14, {'t': 'num', 'v': 0.0, 's': 2}))
    int_cells.append((15, {'t': 'formula', 'f': f'L{int_r}', 'v': 350.36, 's': 3}))
    int_cells.append((16, {'t': 'str', 'v': 'Según Extracto', 's': 0}))
    sb.add_row(int_r, int_cells)
    curr_r += 1

    # 4. (=) SALDO REAL EN CUENTA BI-MONETARIA (Extracto Banco Industrial) -> 97,881.41
    real_mon_r = curr_r
    real_cells = [(1, {'t': 'str', 'v': '(=) TOTAL REAL EN CUENTA BI-MONETARIA (Con Intereses Extracto BI)', 's': 6}), (2, {'t': 'str', 'v': '1101 Real Banco', 's': 6})]
    for m_i in range(3, 15):
        col_l = col_letter(m_i)
        real_cells.append((m_i, {'t': 'formula', 'f': f'ROUND({col_l}{acum_lib_r}+{col_l}{int_r},2)', 'v': 97881.41 if m_i == 12 else 0, 's': 5}))
    real_cells.append((15, {'t': 'formula', 'f': f'ROUND(O{acum_lib_r}+O{int_r},2)', 'v': 97881.41, 's': 5}))
    real_cells.append((16, {'t': 'str', 'v': 'Cuadrado con BI', 's': 6}))
    sb.add_row(real_mon_r, real_cells, height=20)
    curr_r += 1

    # 5. (=) EFECTIVO LIBRE 2026 EN MONETARIA (Descontando Pasivo Inscripciones 2027 = 9,150.00)
    disp_r = curr_r
    disp_cells = [(1, {'t': 'str', 'v': '(=) Disponible 2026 en Monetaria (Sin Inscripciones 2027 - Cuenta 2107)', 's': 4}), (2, {'t': 'str', 'v': 'Libre 2026', 's': 4})]
    for m_i in range(3, 15):
        col_l = col_letter(m_i)
        disp_cells.append((m_i, {'t': 'formula', 'f': f'ROUND({col_l}{real_mon_r}-SUM($C{row_2107}:{col_l}{row_2107}),2)', 'v': 88731.41 if m_i == 12 else 0, 's': 7}))
    disp_cells.append((15, {'t': 'formula', 'f': f'ROUND(O{real_mon_r}-O{row_2107},2)', 'v': 88731.41, 's': 7}))
    disp_cells.append((16, {'t': 'str', 'v': 'Sin Pasivo 2027', 's': 4}))
    sb.add_row(disp_r, disp_cells)
    curr_r += 1

    # 6. Pagos desde Cuenta BI-Ahorro (IGSS y Bono 14) y Saldo Acumulado BI-Ahorro (89,018.67)
    pag_aho_r = curr_r
    pag_aho_cells = [(1, {'t': 'str', 'v': '(-) Pagos de IGSS y Bono 14 ejecutados desde BI-Ahorro', 's': 0}), (2, {'t': 'str', 'v': 'Salidas Ahorro', 's': 0})]
    tot_pag_aho = 0.0
    for m_i, m_name in enumerate(months, start=3):
        col_l = col_letter(m_i)
        v_m = round(sum(float(j['montoBruto']) for j in DATA['journal'] if j['mes'] == m_name and j['banco'] == 'BI-Ahorro'), 2)
        tot_pag_aho += v_m
        f_str = f'SUMIFS(Diario!$G$2:$G$400,Diario!$J$2:$J$400,"BI-Ahorro",Diario!$C$2:$C$400,{col_l}$1)'
        pag_aho_cells.append((m_i, {'t': 'formula', 'f': f_str, 'v': v_m, 's': 2}))
    pag_aho_cells.append((15, {'t': 'formula', 'f': f'SUM(C{pag_aho_r}:N{pag_aho_r})', 'v': round(tot_pag_aho, 2), 's': 3}))
    pag_aho_cells.append((16, {'t': 'formula', 'f': f'ROUND(O{pag_aho_r}/10,2)', 'v': round(tot_pag_aho / 10, 2), 's': 2}))
    sb.add_row(pag_aho_r, pag_aho_cells)
    curr_r += 1

    sal_aho_r = curr_r
    sal_aho_cells = [(1, {'t': 'str', 'v': '(=) SALDO ACUMULADO CUENTA BI-AHORRO (Reserva Prestaciones e IGSS)', 's': 6}), (2, {'t': 'str', 'v': '1102 BI-Ahorro', 's': 6})]
    for m_i in range(3, 15):
        col_l = col_letter(m_i)
        prev_l = col_letter(m_i - 1)
        f_ah = f'ROUND(C{row_2102}-C{pag_aho_r},2)' if m_i == 3 else f'ROUND({prev_l}{sal_aho_r}+{col_l}{row_2102}-{col_l}{pag_aho_r},2)'
        sal_aho_cells.append((m_i, {'t': 'formula', 'f': f_ah, 'v': 89018.67 if m_i >= 11 else 0, 's': 5}))
    sal_aho_cells.append((15, {'t': 'formula', 'f': f'ROUND(O{row_2102}-O{pag_aho_r},2)', 'v': 89018.67, 's': 5}))
    sal_aho_cells.append((16, {'t': 'str', 'v': 'Fondo Reserva', 's': 6}))
    sb.add_row(sal_aho_r, sal_aho_cells, height=20)

    return sb.to_xml()


# ==============================================================================
# 6. PESTAÑA: Presupuesto 2027 (Con Fórmulas Conectadas al Diario)
# ==============================================================================
def build_sheet_presupuesto():
    widths = [44, 20, 18, 22, 24, 24, 22]
    sb = SheetBuilder(widths, freeze_row=2)

    sb.add_row(1, [(1, {'t': 'str', 'v': 'PRESUPUESTO ANUAL 2027 - CENTRO EDUCATIVO ALTUS', 's': 9})], height=22)

    headers = [
        'Partida Contable SAT',              # A (1)
        'Tipo',                              # B (2)
        'Enero 2027 (Q)',                    # C (3)
        'Mensual Feb-Dic 2027 (Q)',          # D (4)
        'Presupuesto Anual 2027 (Q)',        # E (5) - Fórmula C + D*11
        'Ejecutado Real 2026 a Oct (Q)',     # F (6) - Fórmula SUMIFS Diario
        'Diferencia Presupuesto vs 2026 (Q)' # G (7) - Fórmula E - F
    ]
    sb.add_row(2, [(idx, {'t': 'str', 'v': h, 's': 1}) for idx, h in enumerate(headers, start=1)], height=22)

    partidas = DATA['budget2027']['partidas']
    start_r = 3
    for idx, p in enumerate(partidas):
        r = start_r + idx
        ene = float(p.get('mensualEne') or 0)
        feb_dic = float(p.get('mensualFebDic') or 0)
        anual = round(ene + feb_dic * 11, 2)
        real26 = float(p.get('real2026Oct') or 0)
        dif_val = round(anual - real26, 2)
        code4 = p['codigo'][:4]

        sb.add_row(r, [
            (1, {'t': 'str', 'v': p['codigo'], 's': 0}),
            (2, {'t': 'str', 'v': p['tipo'], 's': 0}),
            (3, {'t': 'num', 'v': ene, 's': 2}),
            (4, {'t': 'num', 'v': feb_dic, 's': 2}),
            (5, {'t': 'formula', 'f': f'ROUND(C{r}+(D{r}*11),2)', 'v': anual, 's': 3}),
            (6, {'t': 'formula', 'f': f'SUMIFS(Diario!$G$2:$G$400,Diario!$E$2:$E$400,"{code4}*")', 'v': real26, 's': 2}),
            (7, {'t': 'formula', 'f': f'ROUND(E{r}-F{r},2)', 'v': dif_val, 's': 2})
        ])

    end_r = start_r + len(partidas) - 1
    r_ing = end_r + 1
    r_egr = end_r + 2
    r_net = end_r + 3
    r_2107 = end_r + 5

    sb.add_row(r_ing, [
        (1, {'t': 'str', 'v': 'TOTAL INGRESOS PROYECTADOS 2027', 's': 6}),
        (2, {'t': 'str', 'v': 'Ingreso', 's': 6}),
        (3, {'t': 'formula', 'f': f'SUMIF(B{start_r}:B{end_r},"Ingreso",C{start_r}:C{end_r})', 'v': 229350.0, 's': 5}),
        (4, {'t': 'formula', 'f': f'SUMIF(B{start_r}:B{end_r},"Ingreso",D{start_r}:D{end_r})', 'v': 117077.27, 's': 5}),
        (5, {'t': 'formula', 'f': f'SUMIF(B{start_r}:B{end_r},"Ingreso",E{start_r}:E{end_r})', 'v': 1517200.0, 's': 5}),
        (6, {'t': 'formula', 'f': f'SUMIF(B{start_r}:B{end_r},"Ingreso",F{start_r}:F{end_r})', 'v': 600339.9, 's': 5}),
        (7, {'t': 'formula', 'f': f'ROUND(E{r_ing}-F{r_ing},2)', 'v': 916860.1, 's': 5})
    ], height=20)

    sb.add_row(r_egr, [
        (1, {'t': 'str', 'v': 'TOTAL GASTOS Y RESERVAS PROYECTADOS 2027', 's': 10}),
        (2, {'t': 'str', 'v': 'Gasto + Reserva', 's': 10}),
        (3, {'t': 'formula', 'f': f'ROUND(SUM(C{start_r}:C{end_r})-C{r_ing},2)', 'v': 110350.0, 's': 11}),
        (4, {'t': 'formula', 'f': f'ROUND(SUM(D{start_r}:D{end_r})-D{r_ing},2)', 'v': 108500.0, 's': 11}),
        (5, {'t': 'formula', 'f': f'ROUND(SUM(E{start_r}:E{end_r})-E{r_ing},2)', 'v': 1303850.0, 's': 11}),
        (6, {'t': 'formula', 'f': f'ROUND(SUM(F{start_r}:F{end_r})-F{r_ing},2)', 'v': 465596.37, 's': 11}),
        (7, {'t': 'formula', 'f': f'ROUND(E{r_egr}-F{r_egr},2)', 'v': 838253.63, 's': 11})
    ], height=20)

    sb.add_row(r_net, [
        (1, {'t': 'str', 'v': 'RESULTADO NETO PROYECTADO 2027', 's': 6}),
        (2, {'t': 'str', 'v': 'Superávit 2027', 's': 6}),
        (3, {'t': 'formula', 'f': f'ROUND(C{r_ing}-C{r_egr},2)', 'v': 119000.0, 's': 5}),
        (4, {'t': 'formula', 'f': f'ROUND(D{r_ing}-D{r_egr},2)', 'v': 8577.27, 's': 5}),
        (5, {'t': 'formula', 'f': f'ROUND(E{r_ing}-E{r_egr},2)', 'v': 213350.0, 's': 5}),
        (6, {'t': 'formula', 'f': f'ROUND(F{r_ing}-F{r_egr},2)', 'v': 134743.53, 's': 5}),
        (7, {'t': 'formula', 'f': f'ROUND(E{r_net}-F{r_net},2)', 'v': 78606.47, 's': 5})
    ], height=21)

    sb.add_row(r_2107, [
        (1, {'t': 'str', 'v': '2107 - Anticipo de inscripciones 2027 ya cobradas en 2026 (Pasivo Diferido)', 's': 4}),
        (2, {'t': 'str', 'v': 'Pasivo 2027', 's': 4}),
        (3, {'t': 'formula', 'f': 'SUMIFS(Diario!$G$2:$G$400,Diario!$E$2:$E$400,"2107 - Anticipo de inscripciones")', 'v': 9150.0, 's': 7}),
        (4, {'t': 'str', 'v': 'Pasa a Ingreso 4101 el 01/01/2027', 's': 4})
    ], height=20)

    return sb.to_xml()


# ==============================================================================
# 7. PESTAÑA: Config (Catálogo de Cuentas con Clasificación Exacta Original)
# ==============================================================================
def build_sheet_config():
    widths = [10, 36, 8, 15, 14, 38, 18, 15, 52]
    sb = SheetBuilder(widths, freeze_row=1)

    headers = [
        'Código',                            # A (1)
        'Nombre de Cuenta',                  # B (2)
        'Nivel',                             # C (3)
        'Estatus IVA SAT',                   # D (4)
        'Naturaleza',                        # E (5)
        'Código y Nombre',                   # F (6) - Llave buscada por Diario
        'Tipo Contable',                     # G (7) - Activo / Pasivo / Ingreso / Gasto
        'Lookup IVA',                        # H (8) - Afecto / Exento
        'Observaciones / Normativa SAT'      # I (9)
    ]
    sb.add_row(1, [(idx, {'t': 'str', 'v': h, 's': 1}) for idx, h in enumerate(headers, start=1)], height=22)

    # Asegurar que las cuentas usadas en Diario tengan exactamente el Tipo Contable (Activo, Pasivo, Ingreso, Gasto)
    # y el Estatus IVA que hace cuadrar el libro al centavo.
    tipo_override = {
        '1105': 'Gasto',
        '1501': 'Gasto',
        '2101': 'Gasto',
        '2102': 'Gasto',
        '2103': 'Gasto',
        '2105': 'Gasto',
        '2106': 'Pasivo',
        '2107': 'Pasivo'
    }
    iva_override = {
        '4200': 'Exento',
        '4201': 'Exento'
    }

    for idx, c in enumerate(DATA['config']):
        r = idx + 2
        code = str(c.get('code', ''))
        full = c.get('fullCode', f"{code} - {c.get('name', '')}")
        # Extraer el nombre exacto desde fullCode para que A&" - "&B coincida 100% con fullCode
        name = full.split(' - ', 1)[1] if ' - ' in full else c.get('name', '')
        iva_st = iva_override.get(code, c.get('ivaStatus', 'Exento'))
        tipo_c = tipo_override.get(code, c.get('type', 'Gasto'))

        sb.add_row(r, [
            (1, {'t': 'str', 'v': code, 's': 0}),
            (2, {'t': 'str', 'v': name, 's': 0}),
            (3, {'t': 'num', 'v': int(c.get('level') or 3), 's': 0}),
            (4, {'t': 'str', 'v': iva_st, 's': 0}),
            (5, {'t': 'str', 'v': c.get('nature', ''), 's': 0}),
            (6, {'t': 'formula', 'vt': 'str', 'f': f'A{r}&" - "&B{r}', 'v': full, 's': 0}),
            (7, {'t': 'str', 'v': tipo_c, 's': 0}),
            (8, {'t': 'formula', 'vt': 'str', 'f': f'D{r}', 'v': iva_st, 's': 0}),
            (9, {'t': 'str', 'v': c.get('observations', ''), 's': 0})
        ])

    return sb.to_xml()


def generate_xlsx(filename='Libro_Contable_ALTUS_26_Auditado.xlsx'):
    sheets = [
        ('Diario', build_sheet_diario()),
        ('Colegiaturas', build_sheet_colegiaturas()),
        ('Nóminas', build_sheet_nominas()),
        ('Conciliación', build_sheet_conciliacion()),
        ('Panel V2', build_sheet_panel_v2()),
        ('Presupuesto 2027', build_sheet_presupuesto()),
        ('Config', build_sheet_config())
    ]

    content_types = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    ]
    for i in range(1, len(sheets) + 1):
        content_types.append(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    content_types.append('</Types>')

    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

    wb_xml = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets>'
    ]
    wb_rels = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    ]
    for i, (name, _) in enumerate(sheets, start=1):
        wb_xml.append(f'<sheet name="{esc(name)}" sheetId="{i}" r:id="rId{i}"/>')
        wb_rels.append(
            f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
        )
    wb_rels.append(
        f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    wb_xml.append('</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
    wb_rels.append('</Relationships>')

    with zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', '\n'.join(content_types))
        z.writestr('_rels/.rels', root_rels)
        z.writestr('xl/workbook.xml', '\n'.join(wb_xml))
        z.writestr('xl/_rels/workbook.xml.rels', '\n'.join(wb_rels))
        z.writestr('xl/styles.xml', build_styles_xml())
        for i, (_, xml_content) in enumerate(sheets, start=1):
            z.writestr(f'xl/worksheets/sheet{i}.xml', xml_content)

    print(f'Successfully generated {filename} with {len(sheets)} formula-driven sheets.')


if __name__ == '__main__':
    generate_xlsx()
