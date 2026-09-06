from __future__ import annotations

from datetime import datetime

import sys
from pathlib import Path

from openpyxl import Workbook


TEST_DIR = Path(
    __file__
).resolve().parent

SCRIPTS_DIR = (
    TEST_DIR.parent
    / "Scripts"
)

sys.path.insert(
    0,
    str(SCRIPTS_DIR)
)

from motor_comercial_v02 import (
    calculate_sale
)


def create_workbook():

    wb = Workbook()

    wb.remove(
        wb.active
    )

    # ========================================================
    # CLIENTES
    # ========================================================

    ws = wb.create_sheet(
        "03_CLIENTES"
    )

    ws.append([
        "ClienteID",
        "Nombre",
        "Telefono",
        "Direccion",
        "TipoCliente",
        "Estado",
        "Observaciones"
    ])

    ws.append([
        "CLI001",
        "Cliente Minorista",
        "",
        "",
        "Minorista",
        "Activo",
        ""
    ])

    ws.append([
        "CLI002",
        "Cliente Mayorista",
        "",
        "",
        "Mayorista",
        "Activo",
        ""
    ])

    ws.append([
        "CLI003",
        "Cliente Especial",
        "",
        "",
        "Minorista",
        "Activo",
        ""
    ])

    # ========================================================
    # PRODUCTOS
    # ========================================================

    ws = wb.create_sheet(
        "04_PRODUCTOS"
    )

    ws.append([
        "ProductoID",
        "Producto",
        "Categoria",
        "Costo",
        "PrecioMinorista",
        "PrecioMayorista",
        "PrecioEspecial",
        "Estado",
        "Observaciones"
    ])

    ws.append([
        "PRO001",
        "Asado",
        "Vacuno",
        10000,
        16000,
        14500,
        None,
        "Activo",
        ""
    ])

    # ========================================================
    # 22_PRECIOS
    # ========================================================

    ws = wb.create_sheet(
        "22_PRECIOS"
    )

    ws.append([
        "PrecioID",
        "ClienteID",
        "Cliente",
        "ProductoID",
        "Producto",
        "TipoPrecio",
        "Precio",
        "Descuento%",
        "FechaInicio",
        "FechaFin",
        "Estado",
        "Prioridad",
        "Motivo"
    ])

    # Precio especial Cliente Especial
    ws.append([
        "PRE001",
        "CLI003",
        "Cliente Especial",
        "PRO001",
        "Asado",
        "Especial",
        14000,
        None,
        datetime(2026, 9, 1),
        datetime(2026, 9, 30),
        "Activo",
        10,
        "Precio cliente especial"
    ])

    # ========================================================
    # 23_OFERTAS
    # ========================================================

    ws = wb.create_sheet(
        "23_OFERTAS"
    )

    ws.append([
        "OfertaID",
        "ProductoID",
        "Producto",
        "NombreOferta",
        "PrecioOferta",
        "Descuento%",
        "FechaInicio",
        "FechaFin",
        "MinKg",
        "Estado",
        "Prioridad"
    ])

    return wb


def test_sale(
    wb,
    title,
    client,
    product,
    date,
    kg
):

    print()
    print(
        title
    )

    result = calculate_sale(
        wb,
        client,
        product,
        date,
        kg
    )

    print(
        result
    )

    return result


def main():

    print()
    print("=" * 60)
    print(
        "BLOQUE COMERCIAL v0.2"
    )
    print("=" * 60)

    wb = create_workbook()

    try:

        # ====================================================
        # CASO 1 — MINORISTA
        # ====================================================

        result = test_sale(
            wb,
            "CASO 1 — MINORISTA",
            "Cliente Minorista",
            "Asado",
            datetime(2026, 9, 10),
            2
        )

        assert result["price"] == 16000
        assert result["type"] == "Minorista"
        assert result["importe"] == 32000
        assert result["ganancia"] == 12000

        # ====================================================
        # CASO 2 — MAYORISTA
        # ====================================================

        result = test_sale(
            wb,
            "CASO 2 — MAYORISTA",
            "Cliente Mayorista",
            "Asado",
            datetime(2026, 9, 10),
            2
        )

        assert result["price"] == 14500
        assert result["type"] == "Mayorista"
        assert result["importe"] == 29000
        assert result["ganancia"] == 9000

        # ====================================================
        # CASO 3 — ESPECIAL CLIENTE
        # ====================================================

        result = test_sale(
            wb,
            "CASO 3 — ESPECIAL CLIENTE",
            "Cliente Especial",
            "Asado",
            datetime(2026, 9, 10),
            2
        )

        assert result["price"] == 14000
        assert result["type"] == "Especial"
        assert result["special_price_id"] == "PRE001"

        # ====================================================
        # CASO 4 — OFERTA GANA AL ESPECIAL
        # ====================================================

        offers = wb[
            "23_OFERTAS"
        ]

        offers.append([
            "OF001",
            "PRO001",
            "Asado",
            "Oferta Septiembre",
            12000,
            None,
            datetime(2026, 9, 1),
            datetime(2026, 9, 30),
            2,
            "Activo",
            100
        ])

        result = test_sale(
            wb,
            "CASO 4 — OFERTA SOBRE ESPECIAL",
            "Cliente Especial",
            "Asado",
            datetime(2026, 9, 15),
            2
        )

        assert result["price"] == 12000
        assert result["type"] == "Oferta"
        assert result["offer_id"] == "OF001"
        assert result["importe"] == 24000
        assert result["ganancia"] == 4000

        # ====================================================
        # CASO 5 — OFERTA VENCIDA
        # ====================================================

        result = test_sale(
            wb,
            "CASO 5 — OFERTA VENCIDA",
            "Cliente Especial",
            "Asado",
            datetime(2026, 10, 1),
            2
        )

        assert result["price"] == 14000
        assert result["type"] == "Especial"
        assert result["offer_id"] is None
        assert result["special_price_id"] == "PRE001"

        # ====================================================
        # CASO 6 — OFERTA NO CUMPLE KG
        # ====================================================

        result = test_sale(
            wb,
            "CASO 6 — KG MINIMO NO CUMPLIDO",
            "Cliente Especial",
            "Asado",
            datetime(2026, 9, 15),
            1
        )

        assert result["price"] == 14000
        assert result["type"] == "Especial"

        # ====================================================
        # FINAL
        # ====================================================

        print()
        print("=" * 60)

        print(
            "OK: BLOQUE COMERCIAL v0.2 SUPERADO"
        )

        print("=" * 60)

        print()
        print(
            "Minorista: OK"
        )

        print(
            "Mayorista: OK"
        )

        print(
            "Precio especial cliente: OK"
        )

        print(
            "Oferta prioritaria: OK"
        )

        print(
            "Vigencia por día: OK"
        )

        print(
            "Fallback: OK"
        )

        print(
            "Ganancia: OK"
        )

        print()

    finally:

        wb.close()


if __name__ == "__main__":
    main()