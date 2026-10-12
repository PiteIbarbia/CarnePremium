from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

SCRIPT_DIR = Path(__file__).resolve().parents[1] / "Scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from caja import sincronizar_ventas_en_caja
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
        "PrecioEspecialID", "TicketID", "_TicketLineaKey",
    ])
    wb.create_sheet("06_TICKET")
    wb.create_sheet("07_CAJA")

    rows = [
        ["VEN000001", datetime(2026, 10, 11, 10), datetime(2026, 10, 11, 10), "Cliente Demo", "Minorista", "Asado", 2, 10000, 20000, 3000, .15, "Efectivo", "Vendedor Demo", None, "CLI000001", "PRO000001", "OK", "Minorista", "Precio base", None, None, None, "T-20261011-001", None],
        ["VEN000002", datetime(2026, 10, 11, 10, 5), datetime(2026, 10, 11, 10, 5), "Cliente Demo", "Minorista", "Vacío", 1.5, 10000, 15000, 2500, .1667, "Transferencia", "Vendedor Demo", None, "CLI000001", "PRO000002", "OK", "Minorista", "Precio base", None, None, None, "T-20261011-001", None],
        ["VEN000003", datetime(2026, 10, 11, 10, 10), datetime(2026, 10, 11, 10, 10), "Cliente Demo", "Minorista", "Bife", .5, 10000, 5000, 1000, .2, "Efectivo", "Vendedor Demo", None, "CLI000001", "PRO000003", "OK", "Minorista", "Precio base", None, None, None, None, None],
        ["VEN000004", datetime(2026, 10, 11, 10, 15), datetime(2026, 10, 11, 10, 15), "Cliente Desconocido", "Minorista", "Colita", .5, None, None, None, None, "Efectivo", "Vendedor Demo", None, None, None, "CLIENTE NO ENCONTRADO", None, None, None, None, None, "T-INVALIDO", None],
        ["VEN000005", datetime(2026, 10, 11, 10, 20), datetime(2026, 10, 11, 10, 20), "Cliente Cuenta", "Minorista", "Matambre", 1, 10000, 10000, 2000, .2, "Cuenta Corriente", "Vendedor Demo", None, "CLI000002", "PRO000004", "OK", "Minorista", "Precio base", None, None, None, "T-CTA", None],
    ]
    for row in rows:
        sales.append(row)
    return wb


def main():
    wb = make_workbook()
    try:
        first = sincronizar_ventas_en_caja(wb)
        preparar_ticket(wb)

        assert first["agregados"] == 3, first
        assert first["actualizados"] == 0, first

        cash = wb["07_CAJA"]
        by_ref = {
            cash.cell(row, 5).value: row
            for row in range(2, cash.max_row + 1)
            if cash.cell(row, 10).value == "VENTA_AUTO"
        }
        assert set(by_ref) == {"VEN000001", "VEN000002", "VEN000003"}, by_ref
        total = sum(float(cash.cell(row, 7).value or 0) for row in range(2, cash.max_row + 1))
        assert total == 40000, total

        ticket = wb["06_TICKET"]
        listed = [
            ticket.cell(row, 7).value
            for row in range(2, 502)
            if ticket.cell(row, 7).value not in (None, "")
        ]
        assert listed == ["T-20261011-001", "VEN000003", "T-CTA"], listed
        assert ticket["B3"].value == "T-20261011-001"
        assert ticket["A9"].value.startswith("=IFERROR(INDEX(")
        assert ticket["B39"].value == "=SUM(B9:B38)"
        assert ticket["B39"].number_format == "0.000"
        assert ticket["D39"].value == "=SUM(D9:D38)"

        second = sincronizar_ventas_en_caja(wb)
        preparar_ticket(wb)
        assert second["agregados"] == 0, second
        assert cash.max_row == 4, "Se duplicaron movimientos al repetir el proceso."

        wb["05_VENTAS"]["I2"] = 22000
        third = sincronizar_ventas_en_caja(wb)
        assert third["agregados"] == 0, third
        assert cash.max_row == 4
        updated_row = by_ref["VEN000001"]
        assert cash.cell(updated_row, 7).value == 22000

        print("OK: Ticket + Caja, múltiples cortes, exclusiones y no duplicación.")
    finally:
        wb.close()


if __name__ == "__main__":
    main()
