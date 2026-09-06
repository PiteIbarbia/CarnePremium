from __future__ import annotations

from datetime import datetime

import sys
from pathlib import Path

from openpyxl import Workbook


TEST_DIR = (
    Path(__file__).resolve().parent
)

SCRIPTS_DIR = (
    TEST_DIR.parent
    / "Scripts"
)

sys.path.insert(
    0,
    str(SCRIPTS_DIR)
)

from motor_comercial import (
    determinar_precio_venta
)


# ============================================================
# WORKBOOK DE PRUEBA
# ============================================================

def crear_workbook():

    workbook = Workbook()

    workbook.remove(
        workbook.active
    )

    # ========================================================
    # CLIENTES
    # ========================================================

    clientes = workbook.create_sheet(
        "03_CLIENTES"
    )

    clientes.append([
        "ClienteID",
        "Nombre",
        "Telefono",
        "Direccion",
        "TipoCliente",
        "Estado",
        "Observaciones"
    ])

    clientes.append([
        "CLI001",
        "Cliente Prueba",
        "",
        "",
        "Minorista",
        "Activo",
        ""
    ])

    clientes.append([
        "CLI002",
        "Otro Cliente",
        "",
        "",
        "Minorista",
        "Activo",
        ""
    ])

    # ========================================================
    # PRODUCTOS
    # ========================================================

    productos = workbook.create_sheet(
        "04_PRODUCTOS"
    )

    productos.append([
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

    productos.append([
        "PRO001",
        "Asado Prueba",
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

    precios = workbook.create_sheet(
        "22_PRECIOS"
    )

    precios.append([
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

    # Especial vigente
    precios.append([
        "PRE001",
        "CLI001",
        "Cliente Prueba",
        "PRO001",
        "Asado Prueba",
        "Especial",
        14000,
        None,
        datetime(
            2026,
            9,
            1
        ),
        datetime(
            2026,
            9,
            30
        ),
        "Activo",
        10,
        "Precio especial Septiembre"
    ])

    # Especial de mayor prioridad
    precios.append([
        "PRE002",
        "CLI001",
        "Cliente Prueba",
        "PRO001",
        "Asado Prueba",
        "Especial",
        13500,
        None,
        datetime(
            2026,
            9,
            10
        ),
        datetime(
            2026,
            9,
            30
        ),
        "Activo",
        20,
        "Precio especial Premium"
    ])

    # Precio de otro cliente
    precios.append([
        "PRE003",
        "CLI002",
        "Otro Cliente",
        "PRO001",
        "Asado Prueba",
        "Especial",
        13000,
        None,
        datetime(
            2026,
            9,
            1
        ),
        datetime(
            2026,
            9,
            30
        ),
        "Activo",
        100,
        "Otro cliente"
    ])

    # ========================================================
    # OFERTAS
    # ========================================================

    ofertas = workbook.create_sheet(
        "23_OFERTAS"
    )

    ofertas.append([
        "OfertaID",
        "ProductoID",
        "NombreOferta",
        "PrecioOferta",
        "Descuento%",
        "FechaInicio",
        "FechaFin",
        "MinKg",
        "Estado",
        "Prioridad"
    ])

    return workbook


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("=" * 60)
    print(
        "MVP 0.1.6 - PRECIOS ESPECIALES POR CLIENTE"
    )
    print("=" * 60)
    print()

    workbook = crear_workbook()

    try:

        # ====================================================
        # CASO 1
        # ====================================================

        print(
            "CASO 1 - ESPECIAL VIGENTE"
        )

        resultado = determinar_precio_venta(

            workbook,

            "Cliente Prueba",

            "Asado Prueba",

            datetime(
                2026,
                9,
                5
            ),

            2
        )

        print(
            resultado
        )

        assert (
            resultado["precio_kg"]
            == 14000
        )

        assert (
            resultado["tipo_aplicado"]
            == "Especial"
        )

        assert (
            resultado[
                "precio_especial_id"
            ]
            == "PRE001"
        )

        # ====================================================
        # CASO 2
        # ====================================================

        print()
        print(
            "CASO 2 - PRIORIDAD MAYOR"
        )

        resultado = determinar_precio_venta(

            workbook,

            "Cliente Prueba",

            "Asado Prueba",

            datetime(
                2026,
                9,
                15
            ),

            2
        )

        print(
            resultado
        )

        assert (
            resultado["precio_kg"]
            == 13500
        )

        assert (
            resultado[
                "precio_especial_id"
            ]
            == "PRE002"
        )

        # ====================================================
        # CASO 3
        # ====================================================

        print()
        print(
            "CASO 3 - FECHA ANTERIOR"
        )

        resultado = determinar_precio_venta(

            workbook,

            "Cliente Prueba",

            "Asado Prueba",

            datetime(
                2026,
                8,
                31
            ),

            2
        )

        print(
            resultado
        )

        assert (
            resultado["precio_kg"]
            == 16000
        )

        assert (
            resultado[
                "tipo_aplicado"
            ]
            == "Minorista"
        )

        # ====================================================
        # CASO 4
        # ====================================================

        print()
        print(
            "CASO 4 - CLIENTE DIFERENTE"
        )

        resultado = determinar_precio_venta(

            workbook,

            "Otro Cliente",

            "Asado Prueba",

            datetime(
                2026,
                9,
                15
            ),

            2
        )

        print(
    resultado
)

    assert (
        resultado["precio_kg"]
        == 13000
    )

    assert (
        resultado["tipo_aplicado"]
        == "Especial"
    )

    assert (
        resultado["precio_especial_id"]
        == "PRE003"
    )

        # ====================================================
        # CASO 5
        # ====================================================

        print()
        print(
            "CASO 5 - OFERTA DEBE GANAR"
        )

        ofertas = workbook[
            "23_OFERTAS"
        ]

        ofertas.append([
            "OF016",
            "PRO001",
            "Oferta que gana",
            12000,
            None,
            datetime(
                2026,
                9,
                1
            ),
            datetime(
                2026,
                9,
                30
            ),
            1,
            "Activo",
            100
        ])

        resultado = determinar_precio_venta(

            workbook,

            "Cliente Prueba",

            "Asado Prueba",

            datetime(
                2026,
                9,
                15
            ),

            2
        )

        print(
            resultado
        )

        assert (
            resultado["precio_kg"]
            == 12000
        )

        assert (
            resultado[
                "tipo_aplicado"
            ]
            == "Oferta"
        )

        # ====================================================
        # RESULTADO
        # ====================================================

        print()
        print("=" * 60)
        print(
            "OK: MVP 0.1.6 SUPERADO"
        )
        print("=" * 60)
        print()

        print(
            "Precio especial por cliente: OK"
        )

        print(
            "Vigencia por día: OK"
        )

        print(
            "Prioridad: OK"
        )

        print(
            "Cliente distinto no recibe precio especial: OK"
        )

        print(
            "Oferta mantiene prioridad superior: OK"
        )

        print()

        print(
            "La prueba no modifica el Excel maestro."
        )

    finally:

        workbook.close()


if __name__ == "__main__":
    main()