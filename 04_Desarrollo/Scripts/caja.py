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


def sincronizar_ventas_en_caja(workbook):
    """
    Sincroniza ventas aprobadas con Caja de forma idempotente.
    Cuenta Corriente queda fuera de ingresos cobrados.
    """
    cash = preparar_caja(workbook)
    sales = workbook["05_VENTAS"]
    sc = _column_map(sales)
    cc = _column_map(cash)
    required = ["VentaID", "Fecha", "Cliente", "Producto", "Importe", "MedioPago", "EstadoVenta"]
    missing = [x for x in required if x not in sc]
    if missing:
        raise RuntimeError("05_VENTAS: faltan columnas para Caja: " + ", ".join(missing))

    expected = {}
    for row in range(2, min(MAX_SALES_ROW, sales.max_row) + 1):
        if _text(sales.cell(row, sc["EstadoVenta"]).value).upper() != "OK":
            continue
        client = _text(sales.cell(row, sc["Cliente"]).value)
        product = _text(sales.cell(row, sc["Producto"]).value)
        amount = _number(sales.cell(row, sc["Importe"]).value)
        payment = _text(sales.cell(row, sc["MedioPago"]).value)
        if not client or not product or amount is None or amount <= 0 or not payment:
            continue
        if payment.casefold().replace(" ", "") == "cuentacorriente":
            continue
        sale_date = sales.cell(row, sc["Fecha"]).value
        if isinstance(sale_date, str) and sale_date.startswith("="):
            sale_date = None
        if not isinstance(sale_date, (date, datetime)):
            sale_date = datetime.now()
        ref = _sale_reference(sales, row, sc)
        expected[ref] = {
            "fecha": sale_date, "concepto": f"Venta {client} - {product}",
            "medio_pago": payment, "importe": amount,
        }

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
        if ref not in expected:
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
    for ref, sale in expected.items():
        row = active.get(ref)
        if row is None:
            row = cash.max_row + 1
            cash.cell(row, cc["MovimientoID"]).value = _next_movement_id(cash, cc["MovimientoID"])
            added += 1
        else:
            updated += 1
        values = {
            "Fecha": sale["fecha"], "Tipo": "INGRESO",
            "Concepto": sale["concepto"], "ReferenciaVenta": ref,
            "MedioPago": sale["medio_pago"], "Ingreso": sale["importe"],
            "Egreso": 0, "Origen": "VENTA_AUTO",
        }
        for header, value in values.items():
            cash.cell(row, cc[header]).value = value

    for row in range(2, cash.max_row + 1):
        cash.cell(row, cc["Saldo"]).value = f"=SUM($G$2:G{row})-SUM($H$2:H{row})"
        cash.cell(row, cc["Fecha"]).number_format = "dd/mm/yyyy hh:mm"
        for header in ("Ingreso", "Egreso", "Saldo"):
            cash.cell(row, cc[header]).number_format = '$ #,##0.00'

    return {"agregados": added, "actualizados": updated, "movimientos": max(0, cash.max_row - 1)}
