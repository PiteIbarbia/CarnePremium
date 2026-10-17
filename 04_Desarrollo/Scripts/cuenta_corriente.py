from __future__ import annotations

from datetime import date, datetime

from openpyxl.styles import Alignment, Font, PatternFill

MAX_SALES_ROW = 501
ACCOUNT_HEADERS = [
    "Fecha",
    "ClienteID",
    "Cliente",
    "KgVendidos",
    "VentasDelDia",
    "CobrosDeVentas",
    "CobrosDeudaAnterior",
    "SaldoDelDia",
    "SaldoAcumulado",
    "EstadoCuenta",
    "TicketID",
    "CantidadRegistros",
]


def _text(value):
    return str(value or "").strip()


def _number(value):
    if isinstance(value, bool) or value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        raw = value.strip().replace("$", "").replace(" ", "")
        if "," in raw and "." in raw:
            raw = raw.replace(".", "").replace(",", ".")
        elif "," in raw:
            raw = raw.replace(",", ".")
        try:
            return float(raw)
        except ValueError:
            return None
    return None


def _as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        raw = value.strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(raw[:10], fmt).date()
            except ValueError:
                continue
    return None


def _client_lookup(workbook):
    clients = {}
    if "03_CLIENTES" not in workbook.sheetnames:
        return clients
    ws = workbook["03_CLIENTES"]
    for row in range(2, min(MAX_SALES_ROW, ws.max_row) + 1):
        name = _text(ws.cell(row, 2).value)
        if not name:
            continue
        raw_id = ws.cell(row, 1).value
        # IDs de cliente se calculan con una fórmula determinista CLI + número de fila.
        client_id = (
            f"CLI{row - 1:06d}"
            if raw_id in (None, "") or (isinstance(raw_id, str) and raw_id.startswith("="))
            else str(raw_id).strip()
        )
        clients[name.casefold()] = (client_id, name)
    return clients


def _state(balance):
    if balance > 0.005:
        return "DEBE"
    if balance < -0.005:
        return "A FAVOR"
    return "PAGADO / OK"


def sincronizar_cuenta_corriente(workbook):
    """
    Reconstruye 09_CTACTE como vista automática desde 05_VENTAS.
    Las ventas, cobros aplicados a ventas y cobros de deuda anterior
    se consolidan por cliente y día. Saldo positivo = deuda; negativo = crédito.
    """
    if "05_VENTAS" not in workbook.sheetnames:
        raise RuntimeError("Falta 05_VENTAS.")
    sales = workbook["05_VENTAS"]
    columns = {
        _text(sales.cell(1, col).value): col
        for col in range(1, sales.max_column + 1)
        if sales.cell(1, col).value not in (None, "")
    }
    required = [
        "Fecha", "Cliente", "Producto", "Kg", "Importe",
        "PagoRecibido", "MedioPago", "EstadoVenta", "TipoRegistro",
    ]
    missing = [name for name in required if name not in columns]
    if missing:
        raise RuntimeError("05_VENTAS: faltan columnas para Cuenta Corriente: " + ", ".join(missing))

    clients = _client_lookup(workbook)
    groups = {}

    def get_group(day, client_id, client_name):
        key = (day, client_id)
        if key not in groups:
            groups[key] = {
                "fecha": day,
                "cliente_id": client_id,
                "cliente": client_name,
                "kg": 0.0,
                "ventas": 0.0,
                "cobros_ventas": 0.0,
                "cobros_deuda": 0.0,
                "lineas_venta": 0,
                "registros": 0,
                "ticket_id": "",
            }
        return groups[key]

    for row in range(2, min(MAX_SALES_ROW, sales.max_row) + 1):
        status = _text(sales.cell(row, columns["EstadoVenta"]).value).upper()
        kind = _text(sales.cell(row, columns["TipoRegistro"]).value).upper() or "VENTA"
        name = _text(sales.cell(row, columns["Cliente"]).value)
        client_match = clients.get(name.casefold())
        if not name or client_match is None:
            continue
        client_id, canonical_name = client_match
        sale_day = _as_date(sales.cell(row, columns["Fecha"]).value)
        if sale_day is None:
            continue

        payment_method = _text(sales.cell(row, columns["MedioPago"]).value)
        paid = _number(sales.cell(row, columns["PagoRecibido"]).value) or 0.0
        if (
            not payment_method
            or payment_method.casefold().replace(" ", "") == "cuentacorriente"
            or paid <= 0
        ):
            paid = 0.0

        # Un cobro de deuda antigua es un movimiento, no una venta nueva.
        if kind == "COBRO":
            if status != "COBRO OK":
                continue
            payment_day = (
                _as_date(sales.cell(row, columns["FechaCobro"]).value)
                if "FechaCobro" in columns
                else None
            ) or sale_day
            group = get_group(payment_day, client_id, canonical_name)
            group["registros"] += 1
            if paid > 0:
                group["cobros_deuda"] += paid
            continue

        if status != "OK":
            continue
        product = _text(sales.cell(row, columns["Producto"]).value)
        kg = _number(sales.cell(row, columns["Kg"]).value)
        amount = _number(sales.cell(row, columns["Importe"]).value)
        if not product or kg is None or kg <= 0 or amount is None or amount < 0:
            continue

        # La venta permanece registrada en su fecha original, aunque todavía
        # no se haya cobrado o se cobre varios días después.
        sale_group = get_group(sale_day, client_id, canonical_name)
        sale_group["kg"] += kg
        sale_group["ventas"] += amount
        sale_group["lineas_venta"] += 1
        sale_group["registros"] += 1
        ticket_id = f"T-{sale_day:%Y%m%d}-{client_id}"
        sale_group["ticket_id"] = ticket_id

        # El cobro es independiente de la venta. Solo el dinero efectivamente
        # recibido se registra en el día de cobro, sin duplicar la venta.
        if paid > 0:
            payment_day = (
                _as_date(sales.cell(row, columns["FechaCobro"]).value)
                if "FechaCobro" in columns
                else None
            ) or sale_day
            payment_group = get_group(payment_day, client_id, canonical_name)
            payment_group["cobros_ventas"] += paid
            payment_group["registros"] += 1
            if not payment_group["ticket_id"]:
                payment_group["ticket_id"] = ticket_id

    # Preparar saldos y resumen por cliente antes de escribir la vista.
    running = {}
    customer_totals = {}
    ordered = sorted(
        groups.values(),
        key=lambda item: (
            item["fecha"],
            item["cliente"].casefold(),
            item["cliente_id"],
        ),
    )
    ledger_rows = []
    for item in ordered:
        day_balance = round(
            item["ventas"] - item["cobros_ventas"] - item["cobros_deuda"], 2
        )
        prior_balance = running.get(item["cliente_id"], 0.0)
        ending_balance = round(prior_balance + day_balance, 2)
        running[item["cliente_id"]] = ending_balance
        ledger_rows.append((item, day_balance, ending_balance, _state(ending_balance)))

        total = customer_totals.setdefault(
            item["cliente_id"],
            {
                "cliente_id": item["cliente_id"],
                "cliente": item["cliente"],
                "kg": 0.0,
                "ventas": 0.0,
                "cobros": 0.0,
                "saldo": 0.0,
                "ultima_fecha": item["fecha"],
                "registros": 0,
            },
        )
        total["kg"] += item["kg"]
        total["ventas"] += item["ventas"]
        total["cobros"] += item["cobros_ventas"] + item["cobros_deuda"]
        total["saldo"] += day_balance
        total["ultima_fecha"] = max(total["ultima_fecha"], item["fecha"])
        total["registros"] += item["registros"]

    customer_rows = sorted(
        customer_totals.values(),
        key=lambda item: (item["cliente"].casefold(), item["cliente_id"]),
    )
    due_from_customers = round(
        sum(max(round(item["saldo"], 2), 0) for item in customer_rows), 2
    )
    customer_credits = round(
        sum(abs(min(round(item["saldo"], 2), 0)) for item in customer_rows), 2
    )
    customers_owing = sum(1 for item in customer_rows if round(item["saldo"], 2) > 0.005)
    customers_in_credit = sum(1 for item in customer_rows if round(item["saldo"], 2) < -0.005)
    customers_settled = sum(
        1 for item in customer_rows if abs(round(item["saldo"], 2)) <= 0.005
    )

    ws = (
        workbook["09_CTACTE"]
        if "09_CTACTE" in workbook.sheetnames
        else workbook.create_sheet("09_CTACTE")
    )
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    ws.sheet_properties.tabColor = "7A3E13"
    ws.freeze_panes = "A10"

    # Esta hoja es un reporte generado. Se reconstruye desde 05_VENTAS.
    clear_to = max(ws.max_row, 501, len(customer_rows) + len(ledger_rows) + 20)
    for row in range(1, clear_to + 1):
        for col in range(1, 13):
            ws.cell(row, col).value = None
            ws.cell(row, col).fill = PatternFill(fill_type=None)
            ws.cell(row, col).font = Font()
            ws.cell(row, col).alignment = Alignment()

    title_fill = PatternFill("solid", fgColor="7A3E13")
    section_fill = PatternFill("solid", fgColor="E9D8C9")
    header_fill = PatternFill("solid", fgColor="7A3E13")
    card_fill = PatternFill("solid", fgColor="F7F1EC")
    title_font = Font(bold=True, color="FFFFFF", size=14)
    section_font = Font(bold=True, color="5B341B", size=11)
    header_font = Font(bold=True, color="FFFFFF")
    label_font = Font(bold=True, color="5B341B")
    money_format = '$ #,##0.00;[Red]-$ #,##0.00'

    ws["A1"] = "CUENTA CORRIENTE — RESUMEN Y MOVIMIENTOS"
    for col in range(1, 13):
        cell = ws.cell(1, col)
        cell.fill = title_fill
        cell.font = title_font
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32

    ws["A2"] = "Lectura rápida: saldo positivo = el cliente te debe; saldo negativo = saldo a favor del cliente."
    ws["A3"] = (
        "Cuentas a pagar a proveedores: pendiente de integrar con Compras/Proveedores. "
        "No se muestran deudas estimadas ni inventadas."
    )
    for row in (2, 3):
        ws.cell(row, 1).font = Font(italic=True, color="5B341B")
        ws.cell(row, 1).alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 24
    ws.row_dimensions[3].height = 30

    cards = [
        (1, "A COBRAR DE CLIENTES", due_from_customers, "money"),
        (3, "SALDOS A FAVOR DE CLIENTES", customer_credits, "money"),
        (5, "CLIENTES QUE DEBEN", customers_owing, "count"),
        (7, "CLIENTES CON SALDO A FAVOR", customers_in_credit, "count"),
        (9, "CLIENTES AL DÍA", customers_settled, "count"),
    ]
    for start_col, label, value, value_type in cards:
        label_cell = ws.cell(5, start_col, label)
        value_cell = ws.cell(6, start_col, value)
        label_cell.font = label_font
        label_cell.fill = section_fill
        label_cell.alignment = Alignment(vertical="center", wrap_text=True)
        value_cell.font = Font(bold=True, size=12, color="7A3E13")
        value_cell.fill = card_fill
        value_cell.alignment = Alignment(vertical="center")
        if value_type == "money":
            value_cell.number_format = money_format
    ws.row_dimensions[5].height = 34
    ws.row_dimensions[6].height = 26

    customer_headers = [
        "ClienteID", "Cliente", "Kg vendidos", "Ventas históricas",
        "Cobros registrados", "Saldo actual", "Estado", "Último movimiento",
    ]
    summary_title_row = 8
    summary_header_row = 9
    ws.cell(summary_title_row, 1, "RESUMEN POR CLIENTE — ¿QUIÉN ME DEBE?")
    for col in range(1, len(customer_headers) + 1):
        ws.cell(summary_title_row, col).fill = section_fill
        ws.cell(summary_title_row, col).font = section_font
    for col, header in enumerate(customer_headers, start=1):
        cell = ws.cell(summary_header_row, col, header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row, item in enumerate(customer_rows, start=summary_header_row + 1):
        balance = round(item["saldo"], 2)
        state = _state(balance)
        if balance > 0.005:
            state = "TE DEBE"
        elif balance < -0.005:
            state = "SALDO A FAVOR"
        else:
            state = "AL DÍA"
        values = [
            item["cliente_id"],
            item["cliente"],
            round(item["kg"], 3),
            round(item["ventas"], 2),
            round(item["cobros"], 2),
            balance,
            state,
            item["ultima_fecha"],
        ]
        for col, value in enumerate(values, start=1):
            ws.cell(row, col, value)
        ws.cell(row, 3).number_format = "0.000"
        for col in (4, 5, 6):
            ws.cell(row, col).number_format = money_format
        ws.cell(row, 8).number_format = "dd/mm/yyyy"
        ws.cell(row, 7).alignment = Alignment(horizontal="center")

    detail_title_row = summary_header_row + len(customer_rows) + 2
    detail_header_row = detail_title_row + 1
    ws.cell(detail_title_row, 1, "MOVIMIENTOS POR DÍA — HISTORIAL DE CADA CLIENTE")
    for col in range(1, len(ACCOUNT_HEADERS) + 1):
        ws.cell(detail_title_row, col).fill = section_fill
        ws.cell(detail_title_row, col).font = section_font
    for col, header in enumerate(ACCOUNT_HEADERS, start=1):
        cell = ws.cell(detail_header_row, col, header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row, (item, day_balance, ending_balance, state) in enumerate(
        ledger_rows, start=detail_header_row + 1
    ):
        values = [
            item["fecha"],
            item["cliente_id"],
            item["cliente"],
            round(item["kg"], 3),
            round(item["ventas"], 2),
            round(item["cobros_ventas"], 2),
            round(item["cobros_deuda"], 2),
            day_balance,
            ending_balance,
            state,
            item["ticket_id"],
            item["registros"],
        ]
        for col, value in enumerate(values, start=1):
            ws.cell(row, col, value)
        ws.cell(row, 1).number_format = "dd/mm/yyyy"
        ws.cell(row, 4).number_format = "0.000"
        for col in (5, 6, 7, 8, 9):
            ws.cell(row, col).number_format = money_format
        ws.cell(row, 10).alignment = Alignment(horizontal="center")

    widths = {
        "A": 16, "B": 18, "C": 30, "D": 14, "E": 19, "F": 19,
        "G": 23, "H": 19, "I": 19, "J": 19, "K": 30, "L": 19,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.row_dimensions[summary_header_row].height = 34
    ws.row_dimensions[detail_header_row].height = 34
    ws.auto_filter.ref = f"A{summary_header_row}:H{summary_header_row + len(customer_rows)}"
    ws.freeze_panes = f"A{summary_header_row + 1}"

    return {
        "grupos_cliente_dia": len(ordered),
        "clientes": len(running),
        "a_cobrar_clientes": due_from_customers,
        "saldos_a_favor_clientes": customer_credits,
        "clientes_que_deben": customers_owing,
        "clientes_con_saldo_a_favor": customers_in_credit,
        "clientes_al_dia": customers_settled,
    }
