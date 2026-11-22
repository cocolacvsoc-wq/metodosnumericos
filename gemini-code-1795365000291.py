#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Software de Cálculo Hidráulico - Ecuación de Manning (Canal Trapezoidal)
Método Numérico de Newton-Raphson con Interfaz Gráfica (Tkinter),
Gráfico Dinámico (Matplotlib) y Exportación a Excel (OpenPyXL).
"""

import math
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.patches as patches

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.utils import get_column_letter


def calcular_propiedades_hidraulicas(y: float, b: float, z: float, S: float, n: float) -> dict:
    """Calcula área, perímetro, radio hidráulico, ancho superficial, caudal y velocidad."""
    A = b * y + z * (y ** 2)
    P = b + 2.0 * y * math.sqrt(1.0 + z ** 2)
    R = A / P if P > 0 else 0.0
    T = b + 2.0 * z * y
    Q_calc = (1.0 / n) * A * (R ** (2.0 / 3.0)) * math.sqrt(S)
    V = Q_calc / A if A > 0 else 0.0
    return {"y": y, "A": A, "P": P, "R": R, "T": T, "Q_calc": Q_calc, "V": V}


def calcular_derivadas_analiticas(y: float, b: float, z: float, A: float, P: float, Q_calc: float) -> dict:
    """Calcula las derivadas A'(y), P'(y) y f'(y)."""
    dA = b + 2.0 * z * y
    dP = 2.0 * math.sqrt(1.0 + z ** 2)
    termino_A = (5.0 / 3.0) * (dA / A)
    termino_P = (2.0 / 3.0) * (dP / P)
    df = Q_calc * (termino_A - termino_P)
    return {"dA": dA, "dP": dP, "termino_A": termino_A, "termino_P": termino_P, "df": df}


def resolver_newton_raphson(
    Q_diseno: float, b: float, z: float, S: float, n: float, y0: float,
    tolerancia: float = 1e-10, max_iter: int = 50
) -> dict:
    """Ejecuta el algoritmo iterativo de Newton-Raphson sin omitir pasos."""
    iteraciones = []
    y_actual = y0
    convergio = False
    mensaje_estado = ""

    for i in range(1, max_iter + 1):
        if y_actual <= 0:
            mensaje_estado = f"Error numérico en iteración {i}: profundidad no positiva (y = {y_actual:.6f} m)."
            break

        hid = calcular_propiedades_hidraulicas(y_actual, b, z, S, n)
        fy = hid["Q_calc"] - Q_diseno
        der = calcular_derivadas_analiticas(y_actual, b, z, hid["A"], hid["P"], hid["Q_calc"])

        if abs(der["df"]) < 1e-15:
            mensaje_estado = f"Error numérico en iteración {i}: derivada f'(y) cercana a cero."
            break

        y_siguiente = y_actual - (fy / der["df"])
        error = abs(y_siguiente - y_actual)

        iteraciones.append({
            "iter": i,
            "y_actual": y_actual,
            "A": hid["A"],
            "P": hid["P"],
            "R": hid["R"],
            "T": hid["T"],
            "Q_calc": hid["Q_calc"],
            "V": hid["V"],
            "fy": fy,
            "dA": der["dA"],
            "dP": der["dP"],
            "dfy": der["df"],
            "y_siguiente": y_siguiente,
            "error": error,
        })

        y_actual = y_siguiente
        if error < tolerancia:
            convergio = True
            mensaje_estado = f"Convergencia alcanzada exitosamente en {i} iteraciones (error < {tolerancia:.1e} m)."
            break

    if not convergio and not mensaje_estado:
        mensaje_estado = f"Advertencia: No se alcanzó la tolerancia en {max_iter} iteraciones."

    final_hid = calcular_propiedades_hidraulicas(y_actual, b, z, S, n) if y_actual > 0 else {}

    return {
        "convergio": convergio,
        "mensaje": mensaje_estado,
        "iteraciones": iteraciones,
        "resultado_final": final_hid,
        "parametros": {
            "Q": Q_diseno, "b": b, "z": z, "S": S, "n": n,
            "y0": y0, "tolerancia": tolerancia, "max_iter": max_iter,
        },
    }


def exportar_reporte_excel(ruta_archivo: str, solucion: dict) -> None:
    """Genera el libro Excel con Resumen, Iteraciones Newton-Raphson y Sección Transversal."""
    params = solucion["parametros"]
    iters = solucion["iteraciones"]

    wb = openpyxl.Workbook()
    navy_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    sub_fill = PatternFill(start_color="2C5282", end_color="2C5282", fill_type="solid")
    zebra_fill = PatternFill(start_color="F0F4F8", end_color="F0F4F8", fill_type="solid")
    accent_fill = PatternFill(start_color="E6FFFA", end_color="E6FFFA", fill_type="solid")

    title_font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    white_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=11, bold=True, color="1A202C")
    reg_font = Font(name="Calibri", size=11, color="1A202C")
    thin_side = Side(border_style="thin", color="CBD5E0")
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # HOJA 1: RESUMEN DE RESULTADOS
    ws1 = wb.active
    ws1.title = "RESUMEN DE RESULTADOS"
    ws1.views.sheetView[0].showGridLines = True
    ws1.merge_cells("A1:D1")
    ws1["A1"] = "INFORME HIDRÁULICO — CÁLCULO DE PROFUNDIDAD NORMAL (MANNING)"
    ws1["A1"].font = title_font
    ws1["A1"].fill = navy_fill
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 28

    ws1["A3"] = "1. DATOS INICIALES DEL CANAL"
    ws1["A3"].font = bold_font
    for col_i, h in enumerate(["Parámetro", "Símbolo", "Valor", "Unidad"], 1):
        c = ws1.cell(row=4, column=col_i, value=h)
        c.font, c.fill, c.border = white_font, sub_fill, border_all
        c.alignment = Alignment(horizontal="center")

    datos_rows = [
        ("Caudal de diseño", "Q", params["Q"], "m³/s", "0.0000"),
        ("Ancho del fondo del canal", "b", params["b"], "m", "0.0000"),
        ("Talud lateral (H:V)", "z", params["z"], "adim.", "0.0000"),
        ("Pendiente longitudinal", "S", params["S"], "m/m", "0.000000"),
        ("Coeficiente de rugosidad de Manning", "n", params["n"], "s/m^(1/3)", "0.0000"),
        ("Profundidad inicial (Newton-Raphson)", "y₀", params["y0"], "m", "0.0000"),
        ("Tolerancia numérica", "ε", params["tolerancia"], "m", "0.00E+00"),
    ]
    for idx, (param, simb, val, unid, fmt) in enumerate(datos_rows, start=5):
        ws1.cell(row=idx, column=1, value=param).font = reg_font
        ws1.cell(row=idx, column=2, value=simb).font = bold_font
        c_v = ws1.cell(row=idx, column=3, value=val)
        c_v.font, c_v.number_format = reg_font, fmt
        ws1.cell(row=idx, column=4, value=unid).font = reg_font
        for col_i in range(1, 5):
            cell = ws1.cell(row=idx, column=col_i)
            cell.border = border_all
            if col_i in [2, 4]:
                cell.alignment = Alignment(horizontal="center")
            if idx % 2 == 0:
                cell.fill = zebra_fill

    ws1["A13"] = "2. RESULTADOS HIDRÁULICOS FINALES"
    ws1["A13"].font = bold_font
    for col_i, h in enumerate(["Variable Hidráulica", "Símbolo", "Resultado", "Unidad"], 1):
        c = ws1.cell(row=14, column=col_i, value=h)
        c.font, c.fill, c.border = white_font, sub_fill, border_all
        c.alignment = Alignment(horizontal="center")

    last_nr_row = 4 + len(iters)
    res_rows = [
        ("Profundidad normal", "yn", f"='NEWTON-RAPHSON'!F{last_nr_row}", "m", "0.000000"),
        ("Área hidráulica", "A", "=C6*C15+C7*(C15^2)", "m²", "0.000000"),
        ("Perímetro mojado", "P", "=C6+2*C15*SQRT(1+C7^2)", "m", "0.000000"),
        ("Radio hidráulico", "R", "=C16/C17", "m", "0.000000"),
        ("Ancho superficial", "T", "=C6+2*C7*C15", "m", "0.000000"),
        ("Velocidad media", "V", "=C21/C16", "m/s", "0.000000"),
        ("Caudal comprobado", "Q_calc", "=(1/C9)*C16*(C18^(2/3))*SQRT(C8)", "m³/s", "0.000000"),
        ("Número de iteraciones", "N", len(iters), "iter.", "0"),
    ]
    for idx, (var, simb, form, unid, fmt) in enumerate(res_rows, start=15):
        ws1.cell(row=idx, column=1, value=var).font = bold_font if idx == 15 else reg_font
        ws1.cell(row=idx, column=2, value=simb).font = bold_font
        c_v = ws1.cell(row=idx, column=3, value=form)
        c_v.font = bold_font if idx in [15, 21] else reg_font
        c_v.number_format = fmt
        ws1.cell(row=idx, column=4, value=unid).font = reg_font
        for col_i in range(1, 5):
            cell = ws1.cell(row=idx, column=col_i)
            cell.border = border_all
            if col_i in [2, 4]:
                cell.alignment = Alignment(horizontal="center")
            if idx == 15:
                cell.fill = accent_fill
            elif idx % 2 == 0:
                cell.fill = zebra_fill

    ws1["A24"] = "3. ECUACIONES UTILIZADAS"
    ws1["A24"].font = bold_font
    for col_i, h in enumerate(["Concepto", "Ecuación Matemática", "Descripción", "Unidad"], 1):
        c = ws1.cell(row=25, column=col_i, value=h)
        c.font, c.fill, c.border = white_font, sub_fill, border_all
        c.alignment = Alignment(horizontal="center")

    eq_rows = [
        ("Ecuación de Manning", "Q = (1/n) · A · R^(2/3) · S^(1/2)", "Flujo uniforme en canales abiertos", "m³/s"),
        ("Área hidráulica", "A(y) = b·y + z·y²", "Sección transversal mojada", "m²"),
        ("Perímetro mojado", "P(y) = b + 2y·√(1 + z²)", "Contorno mojado del canal", "m"),
        ("Radio hidráulico", "R(y) = A(y) / P(y)", "Relación área / perímetro", "m"),
        ("Ancho superficial", "T(y) = b + 2·z·y", "Espejo de agua", "m"),
        ("Derivada del área", "A'(y) = b + 2·z·y", "Derivada de A respecto a y", "m"),
        ("Derivada del perímetro", "P'(y) = 2·√(1 + z²)", "Derivada de P respecto a y", "adim."),
        ("Derivada de f(y)", "f'(y) = Q_calc·[(5/3)(A'/A) - (2/3)(P'/P)]", "Derivada analítica de f(y)", "m²/s"),
        ("Iteración Newton-Raphson", "y_{i+1} = y_i - f(y_i) / f'(y_i)", "Actualización iterativa", "m"),
    ]
    for idx, (conc, expr, desc, unid) in enumerate(eq_rows, start=26):
        ws1.cell(row=idx, column=1, value=conc).font = bold_font
        ws1.cell(row=idx, column=2, value=expr).font = reg_font
        ws1.cell(row=idx, column=3, value=desc).font = reg_font
        ws1.cell(row=idx, column=4, value=unid).font = reg_font
        for col_i in range(1, 5):
            cell = ws1.cell(row=idx, column=col_i)
            cell.border = border_all
            if col_i == 4:
                cell.alignment = Alignment(horizontal="center")
            if idx % 2 == 0:
                cell.fill = zebra_fill

    # HOJA 2: NEWTON-RAPHSON
    ws2 = wb.create_sheet(title="NEWTON-RAPHSON")
    ws2.views.sheetView[0].showGridLines = True
    ws2.merge_cells("A1:G1")
    ws2["A1"] = "DESARROLLO NUMÉRICO — ITERACIONES DE NEWTON-RAPHSON"
    ws2["A1"].font, ws2["A1"].fill = title_font, navy_fill
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 28

    headers_nr = [
        "Número de iteración", "Profundidad actual y_i (m)", "Caudal calculado Q_calc (m³/s)",
        "Valor de f(y_i) (m³/s)", "Valor de f'(y_i) (m²/s)", "Nueva profundidad y_{i+1} (m)",
        "Error entre iteraciones (m)",
    ]
    for col_i, h in enumerate(headers_nr, 1):
        c = ws2.cell(row=4, column=col_i, value=h)
        c.font, c.fill, c.border = white_font, sub_fill, border_all
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    b_ref, z_ref, s_ref, n_ref, q_ref = (
        "'RESUMEN DE RESULTADOS'!$C$6", "'RESUMEN DE RESULTADOS'!$C$7",
        "'RESUMEN DE RESULTADOS'!$C$8", "'RESUMEN DE RESULTADOS'!$C$9", "'RESUMEN DE RESULTADOS'!$C$5"
    )
    for idx, it in enumerate(iters, start=5):
        ws2.cell(row=idx, column=1, value=it["iter"]).alignment = Alignment(horizontal="center")
        c_y = ws2.cell(row=idx, column=2, value=it["y_actual"] if idx == 5 else f"=F{idx-1}")
        c_y.number_format = "0.0000000000"

        A_expr = f"({b_ref}*B{idx}+{z_ref}*(B{idx}^2))"
        P_expr = f"({b_ref}+2*B{idx}*SQRT(1+{z_ref}^2))"
        dA_expr = f"({b_ref}+2*{z_ref}*B{idx})"
        dP_expr = f"(2*SQRT(1+{z_ref}^2))"

        ws2.cell(row=idx, column=3, value=f"=(1/{n_ref})*{A_expr}*(({A_expr}/{P_expr})^(2/3))*SQRT({s_ref})").number_format = "0.0000000000"
        ws2.cell(row=idx, column=4, value=f"=C{idx}-{q_ref}").number_format = "0.0000000000"
        ws2.cell(row=idx, column=5, value=f"=C{idx}*((5/3)*({dA_expr}/{A_expr})-(2/3)*({dP_expr}/{P_expr}))").number_format = "0.00000000"
        ws2.cell(row=idx, column=6, value=f"=B{idx}-(D{idx}/E{idx})").number_format = "0.0000000000"
        ws2.cell(row=idx, column=7, value=f"=ABS(F{idx}-B{idx})").number_format = "0.00E+00"

        for col_i in range(1, 8):
            cell = ws2.cell(row=idx, column=col_i)
            cell.font, cell.border = reg_font, border_all
            if idx % 2 == 0:
                cell.fill = zebra_fill

    # HOJA 3: SECCIÓN TRANSVERSAL
    ws3 = wb.create_sheet(title="SECCIÓN TRANSVERSAL")
    ws3.views.sheetView[0].showGridLines = True
    ws3.merge_cells("A1:F1")
    ws3["A1"] = "SECCIÓN TRANSVERSAL TRAPEZOIDAL Y PERFIL DE FLUJO"
    ws3["A1"].font, ws3["A1"].fill = title_font, navy_fill
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 28

    for col_i, h in enumerate(["Dimensión", "Valor", "Unidad"], 1):
        c = ws3.cell(row=4, column=col_i, value=h)
        c.font, c.fill, c.border = white_font, sub_fill, border_all
        c.alignment = Alignment(horizontal="center")

    geom_rows = [
        ("Ancho del fondo (b)", "='RESUMEN DE RESULTADOS'!C6", "m"),
        ("Talud lateral (z)", "='RESUMEN DE RESULTADOS'!C7", "H:V"),
        ("Profundidad calculada (yn)", "='RESUMEN DE RESULTADOS'!C15", "m"),
        ("Ancho superficial (T)", "='RESUMEN DE RESULTADOS'!C19", "m"),
        ("Altura total dibujada (H)", "=B7*1.35", "m"),
    ]
    for idx, (dim, form, unid) in enumerate(geom_rows, start=5):
        ws3.cell(row=idx, column=1, value=dim).font = reg_font
        c_v = ws3.cell(row=idx, column=2, value=form)
        c_v.font, c_v.number_format = bold_font, "0.0000"
        ws3.cell(row=idx, column=3, value=unid).font = reg_font
        for col_i in range(1, 4):
            ws3.cell(row=idx, column=col_i).border = border_all

    for col_i, h in enumerate(["Vértice Canal", "X Canal (m)", "Y Canal (m)", "Vértice Agua", "X Agua (m)", "Y Agua (m)"], 1):
        c = ws3.cell(row=12, column=col_i, value=h)
        c.font, c.fill, c.border = white_font, sub_fill, border_all
        c.alignment = Alignment(horizontal="center")

    verts = [
        ("Borde superior izquierdo", "=-(B5/2 + B6*B9)", "=B9", "Nivel agua izquierdo", "=-B8/2", "=B7"),
        ("Fondo izquierdo", "=-B5/2", 0.0, "Nivel agua derecho", "=B8/2", "=B7"),
        ("Fondo derecho", "=B5/2", 0.0, "", None, None),
        ("Borde superior derecho", "=(B5/2 + B6*B9)", "=B9", "", None, None),
    ]
    for idx, (vc, xc, yc, va, xa, ya) in enumerate(verts, start=13):
        ws3.cell(row=idx, column=1, value=vc).font = reg_font
        ws3.cell(row=idx, column=2, value=xc).number_format = "0.0000"
        ws3.cell(row=idx, column=3, value=yc).number_format = "0.0000"
        ws3.cell(row=idx, column=4, value=va).font = reg_font
        if xa is not None:
            ws3.cell(row=idx, column=5, value=xa).number_format = "0.0000"
            ws3.cell(row=idx, column=6, value=ya).number_format = "0.0000"
        for col_i in range(1, 7):
            ws3.cell(row=idx, column=col_i).border = border_all

    chart = ScatterChart()
    chart.title = "Sección Transversal del Canal Trapezoidal"
    chart.style = 2
    chart.x_axis.title = "Distancia horizontal X (m)"
    chart.y_axis.title = "Elevación Y (m)"
    chart.width, chart.height = 18, 11

    s_canal = Series(Reference(ws3, min_col=3, min_row=13, max_row=16),
                     Reference(ws3, min_col=2, min_row=13, max_row=16), title="Sección del Canal")
    s_canal.graphicalProperties.line.solidFill = "1B365D"
    s_canal.graphicalProperties.line.width = 28000
    chart.series.append(s_canal)

    s_agua = Series(Reference(ws3, min_col=6, min_row=13, max_row=14),
                    Reference(ws3, min_col=5, min_row=13, max_row=14), title="Nivel de Agua (yn)")
    s_agua.graphicalProperties.line.solidFill = "0284C7"
    s_agua.graphicalProperties.line.width = 25000
    s_agua.graphicalProperties.line.dashStyle = "dash"
    chart.series.append(s_agua)

    ws3.add_chart(chart, "A18")

    for ws in [ws1, ws2, ws3]:
        for col in ws.columns:
            max_len = max(
                (len(str(cell.value or "")) for cell in col if cell.row != 1 and not str(cell.value or "").startswith("=")),
                default=10
            )
            ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 5, 16)

    wb.save(ruta_archivo)


class AplicacionManning(tk.Tk):
    def __init__(self):
        super