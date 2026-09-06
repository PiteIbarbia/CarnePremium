from __future__ import annotations


def preparar_dashboard(workbook, fila_inicio=2, fila_fin=501):
    """
    Prepara un Dashboard compatible con Excel 2019 y Google Sheets.

    No depende de tablas dinámicas ni funciones modernas.
    """

    if "10_DASHBOARD" in workbook.sheetnames:
        ws = workbook["10_DASHBOARD"]
    else:
        ws = workbook.create_sheet("10_DASHBOARD")

    ws["A1"] = "CARNE PREMIUM ERP — DASHBOARD OPERATIVO"

    metricas = [
        ("A3", "Ventas registradas", "B3",
         f'=COUNTIF(\'05_VENTAS\'!$A${fila_inicio}:$A${fila_fin},"<>")'),

        ("A4", "Facturación", "B4",
         f'=SUM(\'05_VENTAS\'!$I${fila_inicio}:$I${fila_fin})'),

        ("A5", "Ganancia", "B5",
         f'=SUM(\'05_VENTAS\'!$J${fila_inicio}:$J${fila_fin})'),

        ("A6", "Margen", "B6",
         '=IFERROR(B5/B4,"")'),

        ("A7", "Kg vendidos", "B7",
         f'=SUM(\'05_VENTAS\'!$G${fila_inicio}:$G${fila_fin})'),

        ("A8", "Ticket promedio", "B8",
         '=IFERROR(B4/B3,"")'),

        ("A9", "Ventas con error", "B9",
         f'=COUNTIF(\'05_VENTAS\'!$Q${fila_inicio}:$Q${fila_fin},"<>OK")-COUNTIF(\'05_VENTAS\'!$Q${fila_inicio}:$Q${fila_fin},"")'),
    ]

    for label_cell, label, value_cell, formula in metricas:
        ws[label_cell] = label
        ws[value_cell] = formula

    ws["B4"].number_format = '$ #,##0.00'
    ws["B5"].number_format = '$ #,##0.00'
    ws["B6"].number_format = '0.0%'
    ws["B7"].number_format = '0.000'
    ws["B8"].number_format = '$ #,##0.00'

    return ws
