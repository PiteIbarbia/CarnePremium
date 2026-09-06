# ============================================================
# PATCH MVP 0.1.1 — VENTAS POR NOMBRE
# ============================================================
#
# Reemplaza la función setup_sales() de builder.py por esta.
#
# Cambio:
# - El usuario escribe Cliente y Producto por nombre.
# - ClienteID y ProductoID quedan como columnas técnicas.
# - Los IDs técnicos pueden ocultarse después.
# - Se usa INDEX/MATCH, compatible con Excel 2019.
# ============================================================

def setup_sales(workbook):

    ws = ensure_sheet(
        workbook,
        "05_VENTAS"
    )

    clear_validations(ws)

    headers = [
        "VentaID",
        "Fecha",
        "Hora",
        "Cliente",
        "TipoPrecio",
        "Producto",
        "Kg",
        "PrecioKg",
        "Importe",
        "Ganancia$",
        "Ganancia%",
        "MedioPago",
        "Vendedor",
        "Observaciones",
        "_ClienteID",
        "_ProductoID",
    ]

    set_headers(ws, headers)

    # --------------------------------------------------------
    # Campos calculados / internos
    # --------------------------------------------------------

    for row in range(2, MAX_ROWS + 1):

        # Buscar ID interno por nombre de cliente.
        # MATCH exacto; luego recupera ClienteID.
        ws.cell(
            row=row,
            column=15
        ).value = (
            f'=IF(D{row}="","",'
            f'IFERROR(INDEX(\'03_CLIENTES\'!$A$2:$A$501,'
            f'MATCH(D{row},\'03_CLIENTES\'!$B$2:$B$501,0)),""))'
        )

        # Buscar ID interno por nombre de producto.
        ws.cell(
            row=row,
            column=16
        ).value = (
            f'=IF(F{row}="","",'
            f'IFERROR(INDEX(\'04_PRODUCTOS\'!$A$2:$A$501,'
            f'MATCH(F{row},\'04_PRODUCTOS\'!$B$2:$B$501,0)),""))'
        )

        # VentaID
        ws.cell(
            row=row,
            column=1
        ).value = (
            f'=IF(OR(D{row}="",F{row}=""),"",'
            f'"VEN"&TEXT(ROW()-1,"000000"))'
        )

        # Fecha
        ws.cell(
            row=row,
            column=2
        ).value = (
            f'=IF(A{row}="","",TODAY())'
        )

        # Hora
        ws.cell(
            row=row,
            column=3
        ).value = (
            f'=IF(A{row}="","",NOW())'
        )

    # --------------------------------------------------------
    # Listas que siguen siendo apropiadas como listas
    # --------------------------------------------------------

    add_list_validation(
        ws,
        "E2:E501",
        formula1=(
            "'02_CATALOGOS'!"
            "$A$2:$A$4"
        )
    )

    add_list_validation(
        ws,
        "L2:L501",
        formula1=(
            "'02_CATALOGOS'!"
            "$C$2:$C$6"
        )
    )

    # --------------------------------------------------------
    # Precio por tipo
    # --------------------------------------------------------
    #
    # Busca el producto por nombre directamente.
    # Columnas de 04_PRODUCTOS:
    # B=Producto
    # E=PrecioMinorista
    # F=PrecioMayorista
    # G=PrecioEspecial
    # --------------------------------------------------------

    for row in range(2, MAX_ROWS + 1):

        retail = (
            f'INDEX(\'04_PRODUCTOS\'!$E$2:$E$501,'
            f'MATCH(F{row},\'04_PRODUCTOS\'!$B$2:$B$501,0))'
        )

        wholesale = (
            f'INDEX(\'04_PRODUCTOS\'!$F$2:$F$501,'
            f'MATCH(F{row},\'04_PRODUCTOS\'!$B$2:$B$501,0))'
        )

        special = (
            f'INDEX(\'04_PRODUCTOS\'!$G$2:$G$501,'
            f'MATCH(F{row},\'04_PRODUCTOS\'!$B$2:$B$501,0))'
        )

        cost = (
            f'INDEX(\'04_PRODUCTOS\'!$D$2:$D$501,'
            f'MATCH(F{row},\'04_PRODUCTOS\'!$B$2:$B$501,0))'
        )

        ws.cell(
            row=row,
            column=8
        ).value = (
            f'=IFERROR('
            f'IF(E{row}="Mayorista",'
                f'IF({wholesale}="",{retail},{wholesale}),'
                f'IF(E{row}="Especial",'
                    f'IF({special}="",{retail},{special}),'
                    f'{retail}'
                f')'
            f'),"")'
        )

        # Importe
        ws.cell(
            row=row,
            column=9
        ).value = (
            f'=IF(AND(G{row}<>"",H{row}<>""),'
            f'G{row}*H{row},"")'
        )

        # Ganancia
        ws.cell(
            row=row,
            column=10
        ).value = (
            f'=IFERROR('
            f'IF({cost}="","",'
            f'I{row}-(G{row}*{cost})),'
            f'"")'
        )

        # Ganancia %
        ws.cell(
            row=row,
            column=11
        ).value = (
            f'=IF(OR(I{row}="",J{row}=""),"",'
            f'IFERROR(J{row}/I{row},""))'
        )

        # Formatos
        ws.cell(
            row=row,
            column=7
        ).number_format = "0.000"

        for column in range(8, 11):
            ws.cell(
                row=row,
                column=column
            ).number_format = "$ #,##0.00"

        ws.cell(
            row=row,
            column=11
        ).number_format = "0.0%"

        ws.cell(
            row=row,
            column=2
        ).number_format = "dd/mm/yyyy"

        ws.cell(
            row=row,
            column=3
        ).number_format = "hh:mm:ss"

    # --------------------------------------------------------
    # Tablas y presentación
    # --------------------------------------------------------

    ensure_table(
        ws,
        "VentasTable",
        "A1:P501"
    )

    widths = {
        "A": 15, "B": 12, "C": 10, "D": 28, "E": 15,
        "F": 30, "G": 10, "H": 16, "I": 16, "J": 16,
        "K": 12, "L": 20, "M": 20, "N": 32,
        "O": 16, "P": 16
    }

    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    # Columnas técnicas: quedan ocultas.
    ws.column_dimensions["O"].hidden = True
    ws.column_dimensions["P"].hidden = True

    ws.freeze_panes = "A2"
