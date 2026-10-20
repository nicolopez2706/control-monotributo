"""Genera el Excel de control (se regenera completo en cada ejecución)."""
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

F = "Arial"
NAVY, MID, GREY = "1F3864", "2F5597", "F2F2F2"
MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
MONEY = '"$" #,##0.00;[Red]-"$" #,##0.00;"-"'
MONEY0 = '"$" #,##0;[Red]-"$" #,##0;"-"'
CUIT = '00"-"00000000"-"0'
DATE = "dd/mm/yyyy"
ESTADOS = {  # estado: (fondo, texto)
    "OK": ("C6EFCE", "006100"), "ALERTA": ("FFEB9C", "9C5700"),
    "EXCEDIDO": ("FFC7CE", "9C0006"), "FALTA TOPE": ("E7E6E6", "595959"),
}
thin = Side(style="thin", color="D9D9D9")
box = Border(left=thin, right=thin, top=thin, bottom=thin)


def fnt(size=10, bold=False, color="000000", italic=False):
    return Font(name=F, size=size, bold=bold, color=color, italic=italic)


def fill(c):
    return PatternFill("solid", fgColor=c)


def etiqueta_mes(p):
    return f"{MESES[p.month - 1]}-{str(p.year)[2:]}"


def titulo(ws, texto, hasta_col, sub=None):
    ws["A1"] = texto
    ws["A1"].font = fnt(16, True, "FFFFFF")
    ws["A1"].alignment = Alignment(vertical="center", indent=1)
    for c in range(1, hasta_col + 1):
        ws.cell(row=1, column=c).fill = fill(NAVY)
    ws.row_dimensions[1].height = 32
    if sub:
        ws["A2"] = sub
        ws["A2"].font = fnt(9, italic=True, color="595959")
        ws["A2"].alignment = Alignment(indent=1)


def encabezado(ws, fila, etiquetas, col0=1):
    for i, e in enumerate(etiquetas):
        c = ws.cell(row=fila, column=col0 + i, value=e)
        c.font = fnt(10, True, "FFFFFF")
        c.fill = fill(MID)
        c.border = box
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[fila].height = 30


def anchos(ws, lista):
    for i, w in enumerate(lista, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w


def estilo_grafico(ch, texto, w=17.5, h=8.5):
    ch.title = texto
    ch.width, ch.height = w, h
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    if ch.legend is not None:
        ch.legend.position = "b"


def pintar(serie, color):
    serie.graphicalProperties.solidFill = color
    serie.graphicalProperties.line.solidFill = color


def construir(ruta, empresa, umbral, hoy, inicio, meses, estados, matriz, facturas,
              topes, vigencia_topes, sin_registrar, n_errores):
    """
    estados : lista de dicts (cuit, nombre, cat, tope, total, disp, uso, estado, ultima, ritmo, meses_tope)
    matriz  : {cuit: [12 importes mensuales]}
    facturas: lista de dicts del registro completo (fecha, cuit, trabajador, tipo, pto_vta, numero,
              importe, moneda, archivo, cargado, computa)
    """
    wb = Workbook()
    res = wb.active
    res.title = "Resumen"
    gra = wb.create_sheet("Gráficos")
    fic = wb.create_sheet("Ficha")
    men = wb.create_sheet("Mensual")
    fac = wb.create_sheet("Facturas")
    top = wb.create_sheet("Topes")
    ins = wb.create_sheet("Instrucciones")
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
    n = len(estados)
    periodo = f"{etiqueta_mes(meses[0])} a {etiqueta_mes(meses[-1])}"

    # ============================================================ RESUMEN
    titulo(res, f"Control de Monotributo — {empresa}", 11)
    res["A2"], res["A2"].font = "Actualizado:", fnt(10, True)
    res["A2"].alignment = Alignment(indent=1)
    res["B2"], res["B2"].number_format, res["B2"].font = hoy, DATE, fnt(10, True, MID)
    res["B2"].alignment = Alignment(horizontal="left")
    res["D2"], res["D2"].font = "Período evaluado:", fnt(10, True)
    res["E2"], res["E2"].font = periodo, fnt(10, True, MID)

    cuenta = lambda e: sum(1 for x in estados if x["estado"] == e)
    kpis = [("Trabajadores registrados", n, "0", NAVY),
            ("En regla", cuenta("OK"), "0", "548235"),
            ("En alerta", cuenta("ALERTA"), "0", "BF8F00"),
            ("Excedidos", cuenta("EXCEDIDO"), "0", "C00000"),
            ("Facturado total (período)", sum(x["total"] for x in estados), MONEY, NAVY)]
    for i, (lab, val, fmt, col) in enumerate(kpis):
        r = 4 + i
        res.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        a = res.cell(row=r, column=1, value=lab)
        a.font, a.alignment = fnt(10, True, "FFFFFF"), Alignment(vertical="center", indent=1)
        for c in (1, 2, 3):
            res.cell(row=r, column=c).fill = fill(col)
        v = res.cell(row=r, column=4, value=val)
        v.number_format, v.font, v.fill = fmt, fnt(12, True, col), fill(GREY)
        v.alignment = Alignment(horizontal="right", vertical="center")
        res.row_dimensions[r].height = 22

    res["F4"], res["F4"].font = "Cómo leer el estado", fnt(10, True)
    leyenda = [("OK", f"Menos del {umbral:.0%} del tope"), ("ALERTA", f"Desde el {umbral:.0%} del tope"),
               ("EXCEDIDO", "Superó el tope de su categoría"), ("FALTA TOPE", "Categoría sin tope cargado")]
    for i, (e, txt) in enumerate(leyenda):
        c = res.cell(row=5 + i, column=6, value=e)
        c.fill, c.font = fill(ESTADOS[e][0]), fnt(9, True, ESTADOS[e][1])
        c.alignment = Alignment(horizontal="center")
        res.cell(row=5 + i, column=7, value=txt).font = fnt(9, color="595959")

    avisos = []
    if sin_registrar:
        avisos.append("⚠ Facturas de CUIT que no están en trabajadores.csv: "
                      + ", ".join(str(c) for c in sorted(sin_registrar)))
    if n_errores:
        avisos.append(f"⚠ {n_errores} PDF no se pudieron leer (carpeta con_error): cargalos a mano en facturas.csv")
    for i, t in enumerate(avisos):
        c = res.cell(row=9 + i, column=1, value=t)
        c.font = fnt(10, True, "9C0006")

    encabezado(res, 11, ["CUIT", "Trabajador", "Categoría", "Tope anual", "Facturado en el período",
                         "Disponible", "Uso del tope", "Estado", "Ritmo mensual (últ. 90 días)",
                         "Meses hasta el tope", "Última factura"])
    for i, x in enumerate(estados):
        r = 12 + i
        vals = [x["cuit"], x["nombre"], x["cat"], x["tope"] if x["tope"] else "-", x["total"],
                x["disp"] if x["disp"] is not None else "-", x["uso"] if x["uso"] is not None else "-",
                x["estado"], x["ritmo"],
                "-" if x["meses_tope"] is None else ("Superado" if x["meses_tope"] == 0 else x["meses_tope"]),
                x["ultima"] if x["ultima"] else "-"]
        fmts = [CUIT, "@", "@", MONEY, MONEY, MONEY, "0.0%", "@", MONEY, '0.0 "meses"', DATE]
        for j, (v, fm) in enumerate(zip(vals, fmts), start=1):
            c = res.cell(row=r, column=j, value=v)
            c.number_format, c.font, c.border = fm, fnt(10), box
            if j in (1, 3, 7, 8, 10, 11):
                c.alignment = Alignment(horizontal="center")
            if i % 2 == 1 and j != 8:
                c.fill = fill(GREY)
        bg, fg = ESTADOS[x["estado"]]
        e = res.cell(row=r, column=8)
        e.fill, e.font = fill(bg), fnt(10, True, fg)
    if n:
        res.conditional_formatting.add(f"G12:G{11 + n}", DataBarRule(
            start_type="num", start_value=0, end_type="num", end_value=1, color="8FAADC", showValue=True))
        r = 12 + n
        res.cell(row=r, column=2, value="TOTAL").font = fnt(10, True)
        t = res.cell(row=r, column=5, value=round(sum(x["total"] for x in estados), 2))
        t.number_format, t.font = MONEY, fnt(10, True)
        for c in range(1, 12):
            res.cell(row=r, column=c).border = Border(top=Side(style="medium", color=NAVY))
    nota = res.cell(row=14 + n, column=1,
                    value="Ritmo = facturado en los últimos 90 días ÷ 3. Meses hasta el tope = disponible ÷ ritmo "
                          "(estimación lineal, orientativa). El período incluye el mes en curso.")
    nota.font = fnt(9, italic=True, color="595959")
    anchos(res, [17, 30, 11, 19, 21, 19, 13, 13, 22, 17, 14])
    res.freeze_panes = "A12"
    res.page_setup.orientation = "landscape"
    res.page_setup.fitToWidth, res.page_setup.fitToHeight = 1, 0
    res.sheet_properties.pageSetUpPr.fitToPage = True

    # ============================================================ MENSUAL
    titulo(men, "Facturación mensual por trabajador", 14,
           f"Período {periodo}. Las notas de crédito restan. Importes en pesos.")
    etiquetas = [etiqueta_mes(p) for p in meses]
    encabezado(men, 4, ["Trabajador"] + etiquetas + ["Total"])
    for i, x in enumerate(estados):
        r = 5 + i
        men.cell(row=r, column=1, value=x["nombre"]).font = fnt(10)
        for j, v in enumerate(matriz[x["cuit"]]):
            c = men.cell(row=r, column=2 + j, value=v)
            c.number_format, c.font = MONEY0, fnt(9)
        t = men.cell(row=r, column=14, value=round(sum(matriz[x["cuit"]]), 2))
        t.number_format, t.font = MONEY0, fnt(10, True)
        for c in range(1, 15):
            men.cell(row=r, column=c).border = box
    rt = 5 + n
    men.cell(row=rt, column=1, value="TOTAL").font = fnt(10, True)
    tot_mes = [round(sum(matriz[x["cuit"]][j] for x in estados), 2) for j in range(12)]
    for j, v in enumerate(tot_mes + [round(sum(tot_mes), 2)], start=2):
        c = men.cell(row=rt, column=j, value=v)
        c.number_format, c.font = MONEY0, fnt(10, True)
        c.border = Border(top=Side(style="medium", color=NAVY))
    if n:
        men.conditional_formatting.add(f"B5:M{rt - 1}", ColorScaleRule(
            start_type="min", start_color="FFFFFF", end_type="max", end_color="8FAADC"))
    # bloque auxiliar para el gráfico apilado: principales 8 + "Otros"
    orden = sorted(estados, key=lambda x: -x["total"])
    top8, resto = orden[:8], orden[8:]
    aux0 = rt + 3
    men.cell(row=aux0 - 1, column=1, value="Datos del gráfico (principales 8 + otros)").font = fnt(9, True, "595959")
    filas_aux = [(x["nombre"], matriz[x["cuit"]]) for x in top8]
    if resto:
        filas_aux.append(("Otros", [sum(matriz[x["cuit"]][j] for x in resto) for j in range(12)]))
    for i, (nom, vals) in enumerate(filas_aux):
        men.cell(row=aux0 + i, column=1, value=nom).font = fnt(9, color="595959")
        for j, v in enumerate(vals):
            c = men.cell(row=aux0 + i, column=2 + j, value=v)
            c.number_format, c.font = MONEY0, fnt(9, color="595959")
    anchos(men, [28] + [13] * 12 + [15])
    men.freeze_panes = "B5"

    # ============================================================ GRÁFICOS
    titulo(gra, "Gráficos", 20, "Se actualizan solos en cada ejecución.")
    if n:
        cats = Reference(res, min_col=2, min_row=12, max_row=11 + n)
        # 1) uso del tope
        c1 = BarChart()
        c1.type = "bar"
        c1.add_data(Reference(res, min_col=7, min_row=11, max_row=11 + n), titles_from_data=True)
        c1.set_categories(cats)
        c1.x_axis.scaling.orientation = "maxMin"
        c1.y_axis.scaling.min = 0
        pintar(c1.series[0], MID)
        estilo_grafico(c1, "Uso del tope por trabajador")
        c1.legend = None
        gra.add_chart(c1, "A4")
        # 2) facturado vs tope
        c2 = BarChart()
        c2.type = "col"
        c2.add_data(Reference(res, min_col=5, min_row=11, max_row=11 + n), titles_from_data=True)
        c2.add_data(Reference(res, min_col=4, min_row=11, max_row=11 + n), titles_from_data=True)
        c2.set_categories(cats)
        pintar(c2.series[0], MID)
        pintar(c2.series[1], "C00000")
        estilo_grafico(c2, "Facturado en el período vs. tope anual")
        gra.add_chart(c2, "L4")
        # 3) total por mes
        c3 = BarChart()
        c3.type = "col"
        c3.add_data(Reference(men, min_col=2, max_col=13, min_row=rt), from_rows=True, titles_from_data=False)
        c3.set_categories(Reference(men, min_col=2, max_col=13, min_row=4))
        pintar(c3.series[0], NAVY)
        estilo_grafico(c3, "Facturación total por mes")
        c3.legend = None
        gra.add_chart(c3, "A23")
        # 4) apilado por trabajador
        c4 = BarChart()
        c4.type = "col"
        c4.grouping, c4.overlap = "stacked", 100
        c4.add_data(Reference(men, min_col=1, max_col=13, min_row=aux0, max_row=aux0 + len(filas_aux) - 1),
                    from_rows=True, titles_from_data=True)
        c4.set_categories(Reference(men, min_col=2, max_col=13, min_row=4))
        estilo_grafico(c4, "Facturación mensual por trabajador")
        gra.add_chart(c4, "L23")
    anchos(gra, [9] * 20)

    # ============================================================ FICHA (gráfico individual)
    titulo(fic, "Ficha por trabajador", 8, "Elegí un trabajador en la celda amarilla.")
    fic["A3"], fic["A3"].font = "Trabajador:", fnt(11, True)
    fic["B3"] = estados[0]["nombre"] if n else ""
    fic["B3"].font, fic["B3"].fill, fic["B3"].border = fnt(11, True, "0000FF"), fill("FFF9E5"), box
    fic.merge_cells("B3:D3")
    if n:
        dv = DataValidation(type="list", formula1=f"=Resumen!$B$12:$B${11 + n}", allow_blank=False)
        fic.add_data_validation(dv)
        dv.add("B3")
        idx = f"MATCH($B$3,Resumen!$B$12:$B${11 + n},0)"
        campos = [("CUIT", "A", CUIT), ("Categoría", "C", "@"), ("Tope anual", "D", MONEY),
                  ("Facturado", "E", MONEY), ("Disponible", "F", MONEY), ("Uso del tope", "G", "0.0%"),
                  ("Estado", "H", "@"), ("Ritmo mensual", "I", MONEY)]
        for i, (lab, col, fm) in enumerate(campos):
            r = 5 + i // 2
            c0 = 1 + (i % 2) * 4
            fic.cell(row=r, column=c0, value=lab).font = fnt(10, True, "595959")
            v = fic.cell(row=r, column=c0 + 1, value=f"=INDEX(Resumen!${col}$12:${col}${11 + n},{idx})")
            v.number_format, v.font = fm, fnt(11, True, NAVY)
            v.alignment = Alignment(horizontal="left")
        encabezado(fic, 11, ["Mes", "Facturado en el mes", "Acumulado", "Tope anual"])
        for j in range(12):
            r = 12 + j
            L = men.cell(row=1, column=2 + j).column_letter
            fic.cell(row=r, column=1, value=etiquetas[j]).font = fnt(10)
            b = fic.cell(row=r, column=2, value=f"=INDEX(Mensual!{L}$5:{L}${4 + n},MATCH($B$3,Mensual!$A$5:$A${4 + n},0))")
            ac = fic.cell(row=r, column=3, value=f"=SUM($B$12:B{r})")
            tp = fic.cell(row=r, column=4, value="=IF(ISNUMBER($B$6),$B$6,NA())")
            for c in (b, ac, tp):
                c.number_format, c.font = MONEY0, fnt(10)
        for c in range(1, 5):
            for r in range(12, 24):
                fic.cell(row=r, column=c).border = box
        bar = BarChart()
        bar.type = "col"
        bar.add_data(Reference(fic, min_col=2, min_row=11, max_row=23), titles_from_data=True)
        bar.set_categories(Reference(fic, min_col=1, min_row=12, max_row=23))
        pintar(bar.series[0], "8FAADC")
        lin = LineChart()
        lin.add_data(Reference(fic, min_col=3, max_col=4, min_row=11, max_row=23), titles_from_data=True)
        lin.series[0].graphicalProperties.line.solidFill = NAVY
        lin.series[0].graphicalProperties.line.width = 28000
        lin.series[1].graphicalProperties.line.solidFill = "C00000"
        lin.series[1].graphicalProperties.line.dashStyle = "dash"
        lin.series[1].graphicalProperties.line.width = 22000
        estilo_grafico(bar, "Acumulado del período vs. tope", 19, 9.5)
        bar += lin
        fic.add_chart(bar, "F11")
    anchos(fic, [14, 24, 16, 16, 14, 18, 3, 3])
    fic.column_dimensions["G"].width = 3

    # ============================================================ FACTURAS
    titulo(fac, "Registro de facturas", 11, "Se completa solo al leer los PDF. Fuente: datos/facturas.csv")
    encabezado(fac, 4, ["Fecha", "CUIT emisor", "Trabajador", "Comprobante", "Pto. Vta.", "Número",
                        "Importe (ARS)", "Moneda orig.", "Archivo", "Cargado el", "Computa"])
    for i, f in enumerate(facturas):
        r = 5 + i
        vals = [f["fecha"], f["cuit"], f["trabajador"], f["tipo"], f["pto_vta"], f["numero"], f["importe"],
                f["moneda"], f["archivo"], f["cargado"], "Sí" if f["computa"] else "No (fuera del período)"]
        fmts = [DATE, CUIT, "@", "@", "00000", "00000000", MONEY, "@", "@", DATE, "@"]
        for j, (v, fm) in enumerate(zip(vals, fmts), start=1):
            c = fac.cell(row=r, column=j, value=v)
            c.number_format, c.border = fm, box
            c.font = fnt(10, color="808080") if (j >= 9 or not f["computa"]) else fnt(10)
            if i % 2 == 1:
                c.fill = fill(GREY)
        if f["trabajador"] == "No registrado":
            fac.cell(row=r, column=3).fill = fill("FFC7CE")
            fac.cell(row=r, column=3).font = fnt(10, True, "9C0006")
    fac.auto_filter.ref = f"A4:K{4 + max(len(facturas), 1)}"
    fac.freeze_panes = "A5"
    anchos(fac, [12, 16, 28, 26, 9, 12, 18, 11, 26, 12, 22])

    # ============================================================ TOPES
    titulo(top, "Topes anuales por categoría", 2, vigencia_topes)
    encabezado(top, 4, ["Categoría", "Tope anual (ingresos brutos)"])
    for i, (cat, v) in enumerate(topes):
        a = top.cell(row=5 + i, column=1, value=cat)
        a.font, a.alignment, a.border = fnt(10, True), Alignment(horizontal="center"), box
        b = top.cell(row=5 + i, column=2, value=v)
        b.number_format, b.font, b.border = MONEY, fnt(10), box
    top.cell(row=6 + len(topes), column=1, value="Se editan en datos/topes.csv").font = fnt(9, italic=True, color="595959")
    anchos(top, [14, 34])

    # ============================================================ INSTRUCCIONES
    titulo(ins, "Cómo se usa esta planilla", 2)
    lineas = [
        ("Qué es", "Un informe que se genera solo: no se edita a mano (los cambios se pisan en la próxima ejecución)."),
        ("Dónde se cargan los datos", "En el repositorio: datos/trabajadores.csv, datos/topes.csv y config.ini."),
        ("Facturas", "Subí los PDF a la carpeta facturas del repositorio; el script lee el QR y las agrega."),
        ("Resumen", "Estado de cada trabajador: facturado en el período vs. tope, disponible y ritmo."),
        ("Gráficos / Ficha", "Gráficos generales y un gráfico individual (elegí el trabajador en la hoja Ficha)."),
        ("Mensual", "Cuánto facturó cada uno mes a mes (mapa de calor)."),
        ("Período", "Los 12 meses calendario que terminan en el mes actual (incluido). ARCA usa meses cerrados: "
                    "este criterio es algo más conservador."),
        ("Limitación", "Solo suma lo facturado por este circuito. Si el trabajador tiene otros clientes, su total "
                       "real puede ser mayor."),
        ("Aviso", "Herramienta de control interno. No reemplaza el asesoramiento de un contador."),
    ]
    for i, (k, v) in enumerate(lineas):
        a = ins.cell(row=3 + i, column=1, value=k)
        a.font, a.alignment = fnt(10, True, MID), Alignment(vertical="top", indent=1)
        b = ins.cell(row=3 + i, column=2, value=v)
        b.font, b.alignment = fnt(10), Alignment(wrap_text=True, vertical="top")
        ins.row_dimensions[3 + i].height = 32
    anchos(ins, [28, 100])

    for ws, col in [(res, NAVY), (gra, MID), (fic, MID), (men, "BF8F00"), (fac, "BF8F00"), (top, "7F7F7F"), (ins, "7F7F7F")]:
        ws.sheet_properties.tabColor = col
    wb.calculation.fullCalcOnLoad = True
    wb.save(ruta)
