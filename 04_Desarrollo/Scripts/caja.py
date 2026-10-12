from __future__ import annotations

from datetime import date, datetime
from openpyxl.styles import Alignment, Font, PatternFill

MAX_SALES_ROW = 501
CAJA_HEADERS = [
    "MovimientoID", "Fecha", "Tipo", "Concepto", "ReferenciaVenta",
    "MedioPago", "Ingreso", "Egreso", "Saldo", "Origen",
]


def _column_map(ws):
    result = {}
    for column in range(1, ws.max_column + 1):
        value = ws.cell(1, column).value
        if value not in (None, ""):
            result[str(value).strip()] = column
    return result


def _text(value):
    return str(value or "").strip()


def _number(value):
    if isinstance(value, bool) or value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace("$", "").replace(" ", "")
        if "," in cleaned and "." in cleaned:
            cleaned = cleaned.replace(".", "").replace(",", ".")
        elif "," in cleaned:
            cleaned = cleaned.replace(",", ".")
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def _sale_reference(ws, row, columns):
    value = ws.cell(row, columns["VentaID"]).value
    if isinstance(value, str) and value.startswith("="):
        return f"VEN{row - 1:06d}"
    if value in (None, ""):
        return f"VEN{row - 1:06d}"
    return str(value).strip()


def preparar_caja(workbook):
    """Crea encabezados de Caja sin borrar movimientos previos."""
    ws = workbook["07_CAJA"] if "07_CAJA" in workbook.sheetnames else workbook.create_sheet("07_CAJA")
    current = [_text(ws.cell(1, c).value) for c in range(1, len(CAJA_HEADERS) + 1)]
    has_rows = any(ws.cell(row, 1).value not in (None, "") for row in range(2, ws.max_row + 1))
    if all(value == "" for value in current) and not has_rows:
        for column, header in enumerate(CAJA_HEADERS, start=1):
            ws.cell(1, column).value = header
    elif current != CAJA_HEADERS:
        raise RuntimeError(
            "La estructura de 07_CAJA no coincide con el esquema de este Build. "
            "No se modificó para proteger los movimientos existentes."
        )

    ws.sheet_view.showGridLines = False
    for column, header in enumerate(CAJA_HEADERS, start=1):
        cell = ws.cell(1, column)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="7A3E13")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col, width in {"A": 16, "B": 18, "C": 16, "D": 34, "E": 20, "F": 22, "G": 16, "H": 16, "I": 17, "J": 18}.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    return ws


def _next_movement_id(ws, movement_col):
    largest = 0
    for row in range(2, ws.max_row + 1):
        value = _text(ws.cell(row, movement_col).value)
        if value.startswith("MOV"):
            try:
                largest = max(largest, int(value[3:]))
            except ValueError:
                pass
    return f"MOV{largest + 1:06d}"


def _client_id_for_name(workbook, client_name):
    target = _text(client_name).casefold()
    if not target or "03_CLIENTES" not in workbook.sheetnames:
        return ""
    clients = workbook["03_CLIENTES"]
    for row in range(2, min(MAX_SALES_ROW, clients.max_row) + 1):
        name = _text(clients.cell(row, 2).value)
        if not name or name.casefold() != target:
            continue
        raw_id = clients.cell(row, 1).value
        if raw_id in (None, "") or (isinstance(raw_id, str) and raw_id.startswith("=")):
            return f"CLI{row - 1:06d}"
        return str(raw_id).strip()
    return ""


def _cash_date(value):
    if isinstance(value, (date, datetime)):
        return value
    return datetime.now()


def sincronizar_ventas_en_caja(workbook):
    """
    Registra en Caja solo el dinero efectivamente cobrado en PagoRecibido.
    Las ventas a cuenta corriente sin cobro no son ingresos.
    Los cobros de deuda ingresados con TipoRegistro=COBRO sí crean ingresos.
    """
    cash = preparar_caja(workbook)
    sales = workbook["05_VENTAS"]
    sc = _column_map(sales)
    cc = _column_map(cash)
    required = [
        "VentaID", "Fecha", "Cliente", "Producto", "Importe",
        "PagoRecibido", "MedioPago", "EstadoVenta", "TipoRegistro",
    ]
    missing = [name for name in required if name not in sc]
    if missing:
        raise RuntimeError("05_VENTAS: faltan columnas para Caja: " + ", ".join(missing))

    expected = {}
    known_refs = set()
    for row in range(2, min(MAX_SALES_ROW, sales.max_row) + 1):
        status = _text(sales.cell(row, sc["EstadoVenta"]).value).upper()
        kind = _text(sales.cell(row, sc["TipoRegistro"]).value).upper() or "VENTA"
        if status == "OK" and kind != "COBRO":
            ref = _sale_reference(sales, row, sc)
            known_refs.add(ref)
        elif status == "COBRO OK" and kind == "COBRO":
            ref = f"COB{row - 1:06d}"
            known_refs.add(ref)
        else:
            continue

        client = _text(sales.cell(row, sc["Cliente"]).value)
        client_id = _client_id_for_name(workbook, client)
        paid = _number(sales.cell(row, sc["PagoRecibido"]).value)
        payment = _text(sales.cell(row, sc["MedioPago"]).value)
        sale_date = _cash_date(sales.cell(row, sc["Fecha"]).value)
        if (
            not client or not client_id or paid is None or paid <= 0
            or not payment
            or payment.casefold().replace(" ", "") == "cuentacorriente"
        ):
            continue

        if kind == "COBRO" and status == "COBRO OK":
            concept = f"Cobro de saldo anterior - {client}"
            ref = f"COB{row - 1:06d}"
        else:
            product = _text(sales.cell(row, sc["Producto"]).value)
            if not product:
                continue
            ticket_id = f"T-{sale_date:%Y%m%d}-{client_id}"
            concept = f"Cobro {ticket_id} - {client} / {product}"
            ref = _sale_reference(sales, row, sc)

        expected[ref] = {
            "fecha": sale_date,
            "concepto": concept,
            "medio_pago": payment,
            "importe": round(paid, 2),
        }

    # Evitar duplicados históricos para una misma referencia de venta/cobro.
    duplicates = []
    seen = set()
    for row in range(2, cash.max_row + 1):
        if _text(cash.cell(row, cc["Origen"]).value) != "VENTA_AUTO":
            continue
        ref = _text(cash.cell(row, cc["ReferenciaVenta"]).value)
        if ref and ref in seen:
            duplicates.append(row)
        elif ref:
            seen.add(ref)
    for row in reversed(duplicates):
        cash.delete_rows(row, 1)

    existing = {}
    for row in range(2, cash.max_row + 1):
        if _text(cash.cell(row, cc["Origen"]).value) == "VENTA_AUTO":
            ref = _text(cash.cell(row, cc["ReferenciaVenta"]).value)
            if ref:
                existing[ref] = row

    for ref, row in list(existing.items()):
        if ref in expected:
            continue
        if ref in known_refs:
            # No invertir el importe de una venta que sigue aprobada; neutralizar
            # un cobro automático heredado si ahora el usuario no declara cobro.
            cash.cell(row, cc["Tipo"]).value = "REVISION COBRO"
            cash.cell(row, cc["Concepto"]).value = f"Sin cobro confirmado: {ref}"
            cash.cell(row, cc["Ingreso"]).value = 0
            cash.cell(row, cc["Egreso"]).value = 0
            cash.cell(row, cc["Origen"]).value = "VENTA_AUTO_MIGRADA"
        else:
            income = _number(cash.cell(row, cc["Ingreso"]).value) or 0
            cash.cell(row, cc["Tipo"]).value = "ANULACION"
            cash.cell(row, cc["Concepto"]).value = f"Anulación automática: {ref}"
            cash.cell(row, cc["Ingreso"]).value = 0
            cash.cell(row, cc["Egreso"]).value = income
            cash.cell(row, cc["Origen"]).value = "VENTA_AUTO_ANULADA"

    active = {}
    for row in range(2, cash.max_row + 1):
        if _text(cash.cell(row, cc["Origen"]).value) == "VENTA_AUTO":
            ref = _text(cash.cell(row, cc["ReferenciaVenta"]).value)
            if ref:
                active[ref] = row

    added = updated = 0
    for ref, receipt in expected.items():
        row = active.get(ref)
        if row is None:
            row = cash.max_row + 1
            cash.cell(row, cc["MovimientoID"]).value = _next_movement_id(cash, cc["MovimientoID"])
            added += 1
        else:
            updated += 1
        values = {
            "Fecha": receipt["fecha"],
            "Tipo": "INGRESO",
            "Concepto": receipt["concepto"],
            "ReferenciaVenta": ref,
            "MedioPago": receipt["medio_pago"],
            "Ingreso": receipt["importe"],
            "Egreso": 0,
            "Origen": "VENTA_AUTO",
        }
        for header, value in values.items():
            cash.cell(row, cc[header]).value = value

    for row in range(2, cash.max_row + 1):
        cash.cell(row, cc["Saldo"]).value = f"=SUM($G$2:G{row})-SUM($H$2:H{row})"
        cash.cell(row, cc["Fecha"]).number_format = "dd/mm/yyyy hh:mm"
        for header in ("Ingreso", "Egreso", "Saldo"):
            cash.cell(row, cc[header]).number_format = '$ #,##0.00'

    return {
        "agregados": added,
        "actualizados": updated,
        "movimientos": max(0, cash.max_row - 1),
        "ingresos_cobrados": round(sum(item["importe"] for item in expected.values()), 2),
    }
