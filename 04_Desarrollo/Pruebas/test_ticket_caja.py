from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

SCRIPT_DIR = Path(__file__).resolve().parents[1] / "Scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from caja import sincronizar_ventas_en_caja
from cuenta_corriente import sincronizar_cuenta_corriente
from ticket import preparar_ticket


def make_workbook():
    wb = Workbook()
    sales = wb.active
    sales.title = "05_VENTAS"
    sales.append([
        "VentaID", "Fecha", "Hora", "Cliente", "TipoPrecio", "Producto",
        "Kg", "PrecioKg", "Importe", "Ganancia$", "Ganancia%", "MedioPago",
        "Vendedor", "Observaciones", "_ClienteID", "_ProductoID", "EstadoVenta",
        "TipoAplicado", "MotivoPrecio", "OfertaID", "OfertaNombre",
        "PrecioEspecialID", "PagoRecibido", "_TicketLineaKey", "EstadoCobro",
        "SaldoTicket", "TipoRegistro",
    ])

    clients = wb.create_sheet("03_CLIENTES")
    clients.append(["ClienteID", "Nombre"])
    clients.append(['=IF(B2="","","CLI"&TEXT(ROW()-1,"000000"))', "Cliente Demo"])
    clients.append(['=IF(B3="","","CLI"&TEXT(ROW()-1,"000000"))', "Cliente Cuenta"])

    wb.create_sheet("06_TICKET")
    wb.create_sheet("07_CAJA")
    wb.create_sheet("09_CTACTE")

    rows = [
        ["VEN000001", datetime(2026, 10, 11, 10), datetime(2026, 10, 11, 10), "Cliente Demo", "Minorista", "Asado", 2, 10000, 20000, 3000, .15, "Efectivo", "Vendedor Demo", None, "CLI000001", "PRO000001", "OK", "Minorista", "Precio base", None, None, None, 15000, None, "SALDO PENDIENTE", 5000, "VENTA"],
        ["VEN000002", datetime(2026, 10, 11, 10, 5), datetime(2026, 10, 11, 10, 5), "Cliente Demo", "Minorista", "Vacío", 1.5, 10000, 15000, 2500, .1667, "Efectivo", "Vendedor Demo", None, "CLI000001", "PRO000002", "OK", "Minorista", "Precio base", None, None, None, 15000, None, "SALDO PENDIENTE", 5000, "VENTA"],
        ["VEN000003", datetime(2026, 10, 12, 10), datetime(2026, 10, 12, 10), "Cliente Demo", "Minorista", "Bife", .5, 10000, 5000, 1000, .2, "Efectivo", "Vendedor Demo", None, "CLI000001", "PRO000003", "OK", "Minorista", "Precio base", None, None, None, None, None, "PENDIENTE", 5000, "VENTA"],
        ["VEN000004", datetime(2026, 10, 11, 10, 15), datetime(2026, 10, 11, 10, 15), "Cliente Desconocido", "Minorista", "Colita", .5, None, None, None, None, "Efectivo", "Vendedor Demo", None, None, None, "CLIENTE NO ENCONTRADO", None, None, None, None, None, None, None, None, None, "VENTA"],
        ["VEN000005", datetime(2026, 10, 11, 10, 20), datetime(2026, 10, 11, 10, 20), "Cliente Cuenta", "Minorista", "Matambre", 1, 10000, 10000, 2000, .2, "Cuenta Corriente", "Vendedor Demo", None, "CLI000002", "PRO000004", "OK", "Minorista", "Precio base", None, None, None, None, None, "PENDIENTE", 10000, "VENTA"],
        ["", datetime(2026, 10, 12, 12), datetime(2026, 10, 12, 12), "Cliente Cuenta", None, None, None, None, None, None, None, "Transferencia", "Vendedor Demo", "Cobro parcial de saldo", "CLI000002", None, "COBRO OK", None, None, None, None, None, 3000, None, None, None, "COBRO"],
    ]
    for row in rows:
        sales.append(row)
    return wb


def main():
    wb = make_workbook()
    try:
        first_cash = sincronizar_ventas_en_caja(wb)
        preparar_ticket(wb)
        first_account = sincronizar_cuenta_corriente(wb)

        assert first_cash["agregados"] == 3, first_cash
        assert first_cash["actualizados"] == 0, first_cash
        assert first_cash["ingresos_cobrados"] == 33000, first_cash

        cash = wb["07_CAJA"]
        by_ref = {
            str(cash.cell(row, 5).value): row
            for row in range(2, cash.max_row + 1)
            if cash.cell(row, 10).value == "VENTA_AUTO"
        }
        assert set(by_ref) == {"VEN000001", "VEN000002", "COB000006"}, by_ref
        cash_total = sum(float(cash.cell(row, 7).value or 0) for row in range(2, cash.max_row + 1))
        assert cash_total == 33000, cash_total

        ticket = wb["06_TICKET"]
        listed = [
            ticket.cell(row, 7).value
            for row in range(2, 502)
            if ticket.cell(row, 7).value not in (None, "")
        ]
        assert listed == [
            "T-20261011-CLI000001",
            "T-20261012-CLI000001",
            "T-20261011-CLI000002",
        ], listed
        assert ticket["B3"].value == "T-20261011-CLI000001"
        assert ticket["A9"].value.startswith("=IFERROR(INDEX(")
        assert ticket["B39"].value == "=SUM(B9:B38)"
        assert ticket["B39"].number_format == "0.000"
        assert ticket["D39"].value == "=SUM(D9:D38)"

        second_cash = sincronizar_ventas_en_caja(wb)
        preparar_ticket(wb)
        second_account = sincronizar_cuenta_corriente(wb)
        assert second_cash["agregados"] == 0, second_cash
        assert second_cash["actualizados"] == 3, second_cash
        assert cash.max_row == 4, "Se duplicaron movimientos al repetir el proceso."

        # Editar un cobro existente actualiza el movimiento correspondiente, no crea otro.
        wb["05_VENTAS"]["W2"] = 16000
        third_cash = sincronizar_ventas_en_caja(wb)
        sincronizar_cuenta_corriente(wb)
        assert third_cash["agregados"] == 0, third_cash
        assert cash.max_row == 4
        assert cash.cell(by_ref["VEN000001"], 7).value == 16000

        # Cuenta Corriente suma ventas, cobros de ventas y cobros posteriores.
        account = wb["09_CTACTE"]
        account_rows = {
            (account.cell(row, 3).value, account.cell(row, 1).value): row
            for row in range(2, account.max_row + 1)
        }
        demo_day1 = account_rows[("Cliente Demo", datetime(2026, 10, 11).date())]
        account_day1 = account_rows[("Cliente Cuenta", datetime(2026, 10, 11).date())]
        account_day2 = account_rows[("Cliente Cuenta", datetime(2026, 10, 12).date())]
        assert account.cell(demo_day1, 9).value == 4000
        assert account.cell(account_day1, 9).value == 10000
        assert account.cell(account_day2, 9).value == 7000
        assert account.cell(account_day2, 10).value == "DEBE"

        print("OK: Ticket automático cliente/fecha, kilos totales, cobros, Caja idempotente y Cuenta Corriente.")
    finally:
        wb.close()


if __name__ == "__main__":
    main()
