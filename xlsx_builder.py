import io
import zipfile
import datetime
from xml.sax.saxutils import escape


def _col_letter(col_idx: int) -> str:
    """Convierte índice 1-based (1 -> A, 2 -> B, ...) a letra de columna Excel."""
    result = ""
    while col_idx > 0:
        col_idx, rem = divmod(col_idx - 1, 26)
        result = chr(65 + rem) + result
    return result


def _quarter_from_date(date_str: str) -> str:
    """Devuelve el trimestre fiscal (1T, 2T, 3T, 4T) a partir de YYYY-MM-DD o DD/MM/YYYY."""
    try:
        parts = date_str.replace("/", "-").split("-")
        month = int(parts[1])
        q = (month - 1) // 3 + 1
        year = parts[0] if len(parts[0]) == 4 else parts[2]
        return f"{q}T {year}"
    except Exception:
        return "4T 2026"


def _format_date_es(date_str: str) -> str:
    """Convierte YYYY-MM-DD a DD/MM/YYYY para visualización estándar en España."""
    try:
        parts = date_str.replace("/", "-").split("-")
        if len(parts[0]) == 4:
            return f"{parts[2]}/{parts[1]}/{parts[0]}"
        return date_str
    except Exception:
        return date_str


def build_aeat_xlsx(invoices: list) -> bytes:
    """
    Genera un archivo Excel nativo (.xlsx OpenXML) sin dependencias externas,
    con 2 hojas profesionales, anchos de columna ajustados, formato moneda (€),
    fórmulas reales =SUM() y cuadro de liquidación para Modelo 303 y Modelo 130.
    Estilos (s):
      0 = Normal
      1 = Título Principal (14pt Bold Blanco sobre Azul Oscuro)
      2 = Subtítulo / Metadatos (10pt Gris Oscuro)
      3 = Cabecera de Tabla (10pt Bold Blanco sobre Verde Esmeralda / Slate)
      4 = Celda Texto Normal con bordes
      5 = Celda Texto Centrado (Fecha / Trimestre / NIF) con bordes
      6 = Celda Moneda (€ #,##0.00) con bordes
      7 = Celda Porcentaje / Entero centrado con bordes
      8 = Fila TOTAL Texto Bold con fondo gris claro y doble borde
      9 = Fila TOTAL Moneda Bold (€ #,##0.00) con fondo verde suave y doble borde
    """
    # 1. Calcular desglose por tipos de IVA para la Hoja 2 (Resumen Modelo 303 y 130)
    by_vat = {21: {"base": 0.0, "cuota": 0.0}, 10: {"base": 0.0, "cuota": 0.0}, 4: {"base": 0.0, "cuota": 0.0}, 0: {"base": 0.0, "cuota": 0.0}}
    total_base = 0.0
    total_iva = 0.0
    total_irpf = 0.0
    total_base_irpf = 0.0
    total_gross = 0.0

    for inv in invoices:
        b = float(inv.get("base_imponible", 0.0))
        t_iva = int(inv.get("tipo_iva", 21))
        c_iva = float(inv.get("cuota_iva", 0.0))
        c_irpf = float(inv.get("cuota_irpf", 0.0))
        tot = float(inv.get("total", 0.0))

        bucket = t_iva if t_iva in by_vat else 21
        by_vat[bucket]["base"] += b
        by_vat[bucket]["cuota"] += c_iva

        total_base += b
        total_iva += c_iva
        total_irpf += c_irpf
        if c_irpf > 0 or int(inv.get("tipo_irpf", 0)) > 0:
            total_base_irpf += b
        total_gross += tot

    def cell_str(row_idx: int, col_idx: int, val: str, style: int = 4) -> str:
        ref = f"{_col_letter(col_idx)}{row_idx}"
        return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{escape(str(val))}</t></is></c>'

    def cell_num(row_idx: int, col_idx: int, val: float, style: int = 6) -> str:
        ref = f"{_col_letter(col_idx)}{row_idx}"
        return f'<c r="{ref}" s="{style}" t="n"><v>{val:.2f}</v></c>'

    def cell_int(row_idx: int, col_idx: int, val: int, style: int = 7) -> str:
        ref = f"{_col_letter(col_idx)}{row_idx}"
        return f'<c r="{ref}" s="{style}" t="n"><v>{int(val)}</v></c>'

    def cell_formula(row_idx: int, col_idx: int, formula: str, cached_val: float, style: int = 9) -> str:
        ref = f"{_col_letter(col_idx)}{row_idx}"
        return f'<c r="{ref}" s="{style}" t="n"><f>{escape(formula)}</f><v>{cached_val:.2f}</v></c>'

    # --- CONSTRUIR HOJA 1: LIBRO REGISTRO DE FACTURAS RECIBIDAS (AEAT) ---
    sheet1_rows = []
    # Fila 1: Banner Título
    sheet1_rows.append(
        '<row r="1" ht="28" customHeight="1">'
        + cell_str(1, 1, "LIBRO REGISTRO OFICIAL DE FACTURAS Y GASTOS (ART. 64 RIVA - AEAT)", style=1)
        + "".join(cell_str(1, c, "", style=1) for c in range(2, 15))
        + "</row>"
    )
    # Fila 2: Subtítulo
    today_es = datetime.date.today().strftime("%d/%m/%Y")
    sheet1_rows.append(
        '<row r="2" ht="20" customHeight="1">'
        + cell_str(2, 1, f"Generado por VeriFast.es el {today_es} · Liquidación Modelo 303 (IVA) y Modelo 130 (IRPF) · Cumplimiento VeriFactu SHA-256", style=2)
        + "</row>"
    )
    # Fila 3: Vacía separadora
    sheet1_rows.append('<row r="3" ht="10" customHeight="1"/>')

    # Fila 4: Cabeceras
    headers = [
        "Nº Orden",
        "Trimestre",
        "Fecha Expedición",
        "Nº Factura",
        "NIF / CIF Emisor",
        "Nombre o Razón Social Emisor",
        "Concepto / Cuenta Contable",
        "Base Imponible (€)",
        "% IVA",
        "Cuota IVA Deducible (€)",
        "% IRPF",
        "Retención IRPF (€)",
        "Total Factura (€)",
        "Huella VeriFactu SHA-256",
    ]
    sheet1_rows.append(
        '<row r="4" ht="24" customHeight="1">'
        + "".join(cell_str(4, idx + 1, h, style=3) for idx, h in enumerate(headers))
        + "</row>"
    )

    start_data_row = 5
    for i, inv in enumerate(invoices):
        r = start_data_row + i
        fecha_raw = str(inv.get("fecha", ""))
        row_xml = f'<row r="{r}" ht="20" customHeight="1">'
        row_xml += cell_int(r, 1, i + 1, style=7)
        row_xml += cell_str(r, 2, _quarter_from_date(fecha_raw), style=5)
        row_xml += cell_str(r, 3, _format_date_es(fecha_raw), style=5)
        row_xml += cell_str(r, 4, inv.get("num_factura", ""), style=5)
        row_xml += cell_str(r, 5, inv.get("nif", ""), style=5)
        row_xml += cell_str(r, 6, inv.get("proveedor", ""), style=4)
        row_xml += cell_str(r, 7, inv.get("concepto", "Gastos de explotación (629)"), style=4)
        row_xml += cell_num(r, 8, float(inv.get("base_imponible", 0.0)), style=6)
        row_xml += cell_int(r, 9, int(inv.get("tipo_iva", 21)), style=7)
        row_xml += cell_num(r, 10, float(inv.get("cuota_iva", 0.0)), style=6)
        row_xml += cell_int(r, 11, int(inv.get("tipo_irpf", 0)), style=7)
        row_xml += cell_num(r, 12, float(inv.get("cuota_irpf", 0.0)), style=6)
        row_xml += cell_num(r, 13, float(inv.get("total", 0.0)), style=6)
        row_xml += cell_str(r, 14, inv.get("huella_verifactu", ""), style=4)
        row_xml += "</row>"
        sheet1_rows.append(row_xml)

    end_data_row = max(start_data_row, start_data_row + len(invoices) - 1)
    tot_row = end_data_row + 1

    if invoices:
        tot_xml = f'<row r="{tot_row}" ht="24" customHeight="1">'
        for c in range(1, 7):
            tot_xml += cell_str(tot_row, c, "", style=8)
        tot_xml += cell_str(tot_row, 7, "TOTALES ACUMULADOS TRIMESTRE:", style=8)
        tot_xml += cell_formula(tot_row, 8, f"SUM(H{start_data_row}:H{end_data_row})", total_base, style=9)
        tot_xml += cell_str(tot_row, 9, "", style=8)
        tot_xml += cell_formula(tot_row, 10, f"SUM(J{start_data_row}:J{end_data_row})", total_iva, style=9)
        tot_xml += cell_str(tot_row, 11, "", style=8)
        tot_xml += cell_formula(tot_row, 12, f"SUM(L{start_data_row}:L{end_data_row})", total_irpf, style=9)
        tot_xml += cell_formula(tot_row, 13, f"SUM(M{start_data_row}:M{end_data_row})", total_gross, style=9)
        tot_xml += cell_str(tot_row, 14, f"{len(invoices)} facturas verificadas", style=8)
        tot_xml += "</row>"
        sheet1_rows.append(tot_xml)

    sheet1_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="1" width="10" customWidth="1"/>
    <col min="2" max="2" width="12" customWidth="1"/>
    <col min="3" max="3" width="16" customWidth="1"/>
    <col min="4" max="4" width="18" customWidth="1"/>
    <col min="5" max="5" width="16" customWidth="1"/>
    <col min="6" max="6" width="42" customWidth="1"/>
    <col min="7" max="7" width="46" customWidth="1"/>
    <col min="8" max="8" width="18" customWidth="1"/>
    <col min="9" max="9" width="10" customWidth="1"/>
    <col min="10" max="10" width="22" customWidth="1"/>
    <col min="11" max="11" width="10" customWidth="1"/>
    <col min="12" max="12" width="18" customWidth="1"/>
    <col min="13" max="13" width="18" customWidth="1"/>
    <col min="14" max="14" width="68" customWidth="1"/>
  </cols>
  <sheetData>
    {''.join(sheet1_rows)}
  </sheetData>
  <mergeCells count="2">
    <mergeCell ref="A1:N1"/>
    <mergeCell ref="A2:N2"/>
  </mergeCells>
</worksheet>"""

    # --- CONSTRUIR HOJA 2: RESUMEN LIQUIDACIÓN MODELO 303 Y MODELO 130 ---
    s2 = []
    s2.append(
        '<row r="1" ht="28" customHeight="1">'
        + cell_str(1, 1, "RESUMEN EJECUTIVO DE AUTOLIQUIDACIÓN AEAT (MODELOS 303 Y 130)", style=1)
        + "".join(cell_str(1, c, "", style=1) for c in range(2, 6))
        + "</row>"
    )
    s2.append(
        '<row r="2" ht="20" customHeight="1">'
        + cell_str(2, 1, "Casillas listas para copiar en la Sede Electrónica de la Agencia Tributaria o enviar a gestoría", style=2)
        + "</row>"
    )
    s2.append('<row r="3" ht="10" customHeight="1"/>')
    s2.append(
        '<row r="4" ht="24" customHeight="1">'
        + cell_str(4, 1, "Modelo Oficial AEAT", style=3)
        + cell_str(4, 2, "Casilla / Concepto Tributario", style=3)
        + cell_str(4, 3, "Tramo / Tipo", style=3)
        + cell_str(4, 4, "Base Imponible (€)", style=3)
        + cell_str(4, 5, "Cuota Fiscal Deducible (€)", style=3)
        + "</row>"
    )

    summary_items = [
        ("Modelo 303 (IVA)", "Operaciones interiores corrientes al tipo general", "General (21%)", by_vat[21]["base"], by_vat[21]["cuota"]),
        ("Modelo 303 (IVA)", "Operaciones interiores corrientes al tipo reducido", "Reducido (10%)", by_vat[10]["base"], by_vat[10]["cuota"]),
        ("Modelo 303 (IVA)", "Operaciones interiores corrientes al tipo superreducido", "Superreducido (4%)", by_vat[4]["base"], by_vat[4]["cuota"]),
        ("Modelo 303 (IVA)", "Operaciones exentas o sujetas al 0%", "Exento (0%)", by_vat[0]["base"], 0.0),
    ]
    for idx, (mod, desc, tramo, b_val, c_val) in enumerate(summary_items):
        rr = 5 + idx
        s2.append(
            f'<row r="{rr}" ht="20" customHeight="1">'
            + cell_str(rr, 1, mod, style=5)
            + cell_str(rr, 2, desc, style=4)
            + cell_str(rr, 3, tramo, style=5)
            + cell_num(rr, 4, b_val, style=6)
            + cell_num(rr, 5, c_val, style=6)
            + "</row>"
        )

    # Casilla 28 y 29 Total Modelo 303
    s2.append(
        '<row r="9" ht="24" customHeight="1">'
        + cell_str(9, 1, "MODELO 303 TOTAL", style=8)
        + cell_str(9, 2, "Casilla [28] Base Deducible y Casilla [29] Cuota IVA Deducible", style=8)
        + cell_str(9, 3, "TOTAL IVA", style=8)
        + cell_num(9, 4, total_base, style=9)
        + cell_num(9, 5, total_iva, style=9)
        + "</row>"
    )
    # Modelo 130 IRPF
    ahorro_irpf = round(total_base * 0.20, 2)
    s2.append(
        '<row r="10" ht="24" customHeight="1">'
        + cell_str(10, 1, "MODELO 130 (IRPF)", style=8)
        + cell_str(10, 2, "Casilla [02] Gastos Fiscalmente Deducibles (Ahorro 20% IRPF)", style=8)
        + cell_str(10, 3, "Tramo 20% IRPF", style=8)
        + cell_num(10, 4, total_base, style=9)
        + cell_num(10, 5, ahorro_irpf, style=9)
        + "</row>"
    )
    s2.append(
        '<row r="11" ht="24" customHeight="1">'
        + cell_str(11, 1, "MODELO 111 / 130", style=8)
        + cell_str(11, 2, "Retenciones de IRPF practicadas en facturas profesionales", style=8)
        + cell_str(11, 3, "Retenciones", style=8)
        + cell_num(11, 4, total_base_irpf, style=9)
        + cell_num(11, 5, total_irpf, style=9)
        + "</row>"
    )

    sheet2_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="1" width="22" customWidth="1"/>
    <col min="2" max="2" width="64" customWidth="1"/>
    <col min="3" max="3" width="22" customWidth="1"/>
    <col min="4" max="4" width="22" customWidth="1"/>
    <col min="5" max="5" width="24" customWidth="1"/>
  </cols>
  <sheetData>
    {''.join(s2)}
  </sheetData>
  <mergeCells count="2">
    <mergeCell ref="A1:E1"/>
    <mergeCell ref="A2:E2"/>
  </mergeCells>
</worksheet>"""

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.00 &quot;€&quot;"/>
  </numFmts>
  <fonts count="5">
    <font><sz val="10"/><color rgb="FF0F172A"/><name val="Calibri"/></font>
    <font><b/><sz val="13"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><i/><sz val="10"/><color rgb="FF475569"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FF065F46"/><name val="Calibri"/></font>
  </fonts>
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF0F172A"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF059669"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD1FAE5"/></patternFill></fill>
  </fills>
  <borders count="3">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFCBD5E1"/></left>
      <right style="thin"><color rgb="FFCBD5E1"/></right>
      <top style="thin"><color rgb="FFCBD5E1"/></top>
      <bottom style="thin"><color rgb="FFCBD5E1"/></bottom>
    </border>
    <border>
      <left style="thin"><color rgb="FF059669"/></left>
      <right style="thin"><color rgb="FF059669"/></right>
      <top style="medium"><color rgb="FF059669"/></top>
      <bottom style="double"><color rgb="FF059669"/></bottom>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="10">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">
      <alignment vertical="center" indent="1"/>
    </xf>
    <xf numFmtId="0" fontId="2" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">
      <alignment vertical="center" indent="1"/>
    </xf>
    <xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="4" fillId="4" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="164" fontId="4" fillId="5" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
  </cellXfs>
</styleSheet>"""

    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="1. Libro Registro AEAT" sheetId="1" r:id="rId1"/>
    <sheet name="2. Resumen Mod 303 y 130" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"""

    workbook_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"""

    mem_zip = io.BytesIO()
    with zipfile.ZipFile(mem_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/styles.xml", styles_xml)
        zf.writestr("xl/worksheets/sheet1.xml", sheet1_xml)
        zf.writestr("xl/worksheets/sheet2.xml", sheet2_xml)

    return mem_zip.getvalue()
