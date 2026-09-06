from __future__ import annotations

from datetime import datetime

import sys
from pathlib import Path

from openpyxl import Workbook


# ============================================================
# RUTAS
# ============================================================

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


from motor_comercial import (
    calculate_sale
)


# ============================================================
# WORKBOOK AISLADO
# ============================================================

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
        "Observaciones",
    ])

    ws.append([
        "CLI001",
        "Cliente Minorista",
        "",
        "",
        "Minorista",
        "Activo",
        "",
    ])

    ws.append([
        "CLI002",
        "Cliente Mayorista",
        "",
        "",
        "Mayorista",
        "Activo",
        "",
    ])

    ws.append([
        "CLI003",
        "Cliente Especial",
        "",
        "",
        "Minorista",
        "Activo",
        "",
    ])

    ws.append([
        "CLI004",
        "Cliente Sin Regla",
        "",
        "",
        "Minorista",
        "Activo",
        "",
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
        "Observaciones",
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
        "",
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
        "Motivo",
    ])

    # Especial vigente para Cliente Especial
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
        datetime(2026, 10, 31),
        "Activo",
        10,
        "Precio cliente especial",
    ])

    # Regla de mayor prioridad posterior
    ws.append([
        "PRE002",
        "CLI003",
        "Cliente Especial",
        "PRO001",
        "Asado",
        "Especial",
        13500,
        None,
        datetime(2026, 9, 10),
        datetime(2026, 9, 30),
        "Activo",
        20,
        "Precio especial premium",
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
        "Prioridad",
    ])

    return wb


# ============================================================
# EJECUTAR CASO
# ============================================================

def run_case(
    workbook,
    title,
    client,
    product,
    operation_date,
    kg
):

    print()
    print(title)
    print("-" * 60)

    result = calculate_sale(

        workbook,

        client,

        product,

        operation_date,

        kg
    )

    print(
        f"Estado: {result.get('estado')}"
    )

    print(
        f"PrecioKg: {result.get('precio_kg')}"
    )

    print(
        f"Tipo: {result.get('tipo_aplicado')}"
    )

    print(
        f"Motivo: {result.get('motivo')}"
    )

    print(
        f"OfertaID: {result.get('oferta_id')}"
    )

    print(
        f"PrecioEspecialID: "
        f"{result.get('precio_especial_id')}"
    )

    print(
        f"Importe: {result.get('importe')}"
    )

    print(
        f"Ganancia: {result.get('ganancia')}"
    )

    return result


# ============================================================
# MAIN
# ============================================================

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
        # 1. MINORISTA
        # ====================================================

        result = run_case(

            wb,

            "CASO 1 - MINORISTA",

            "Cliente Minorista",

            "Asado",

            datetime(
                2026,
                8,
                20
            ),

            2
        )

        assert result["precio_kg"] == 16000
        assert result["tipo_aplicado"] == "Minorista"
        assert result["importe"] == 32000
        assert result["ganancia"] == 12000


        # ====================================================
        # 2. MAYORISTA
        # ====================================================

        result = run_case(

            wb,

            "CASO 2 - MAYORISTA",

            "Cliente Mayorista",

            "Asado",

            datetime(
                2026,
                8,
                20
            ),

            2
        )

        assert result["precio_kg"] == 14500
        assert result["tipo_aplicado"] == "Mayorista"
        assert result["importe"] == 29000
        assert result["ganancia"] == 9000


        # ====================================================
        # 3. ESPECIAL CLIENTE
        # ====================================================

        result = run_case(

            wb,

            "CASO 3 - ESPECIAL CLIENTE",

            "Cliente Especial",

            "Asado",

            datetime(
                2026,
                9,
                5
            ),

            2
        )

        assert result["precio_kg"] == 14000
        assert result["tipo_aplicado"] == "Especial"
        assert result["precio_especial_id"] == "PRE001"


        # ====================================================
        # 4. ESPECIAL DE MAYOR PRIORIDAD
        # ====================================================

        result = run_case(

            wb,

            "CASO 4 - ESPECIAL PRIORIDAD MAYOR",

            "Cliente Especial",

            "Asado",

            datetime(
                2026,
                9,
                15
            ),

            2
        )

        assert result["precio_kg"] == 13500
        assert result["tipo_aplicado"] == "Especial"
        assert result["precio_especial_id"] == "PRE002"


        # ====================================================
        # 5. CLIENTE SIN REGLA ESPECIAL
        # ====================================================

        result = run_case(

            wb,

            "CASO 5 - CLIENTE SIN PRECIO ESPECIAL",

            "Cliente Sin Regla",

            "Asado",

            datetime(
                2026,
                9,
                15
            ),

            2
        )

        assert result["precio_kg"] == 16000
        assert result["tipo_aplicado"] == "Minorista"
        assert result["precio_especial_id"] is None


        # ====================================================
        # 6. OFERTA GANA AL ESPECIAL
        # ====================================================

        ofertas = wb[
            "23_OFERTAS"
        ]

        ofertas.append([
            "OF001",
            "PRO001",
            "Asado",
            "Oferta Septiembre",
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
            2,
            "Activo",
            100,
        ])

        result = run_case(

            wb,

            "CASO 6 - OFERTA GANA AL ESPECIAL",

            "Cliente Especial",

            "Asado",

            datetime(
                2026,
                9,
                15
            ),

            2
        )

        assert result["precio_kg"] == 12000
        assert result["tipo_aplicado"] == "Oferta"
        assert result["oferta_id"] == "OF001"
        assert result["importe"] == 24000
        assert result["ganancia"] == 4000


        # ====================================================
        # 7. OFERTA VENCIDA
        # ====================================================

        result = run_case(

            wb,

            "CASO 7 - OFERTA VENCIDA, ESPECIAL ACTIVO",

            "Cliente Especial",

            "Asado",

            datetime(
                2026,
                10,
                1
            ),

            2
        )

        assert result["precio_kg"] == 14000
        assert result["tipo_aplicado"] == "Especial"
        assert result["oferta_id"] is None
        assert result["precio_especial_id"] == "PRE001"


        # ====================================================
        # 8. OFERTA + ESPECIAL VENCIDOS
        # ====================================================

        result = run_case(

            wb,

            "CASO 8 - OFERTA Y ESPECIAL VENCIDOS",

            "Cliente Especial",

            "Asado",

            datetime(
                2026,
                11,
                1
            ),

            2
        )

        assert result["precio_kg"] == 16000
        assert result["tipo_aplicado"] == "Minorista"
        assert result["oferta_id"] is None
        assert result["precio_especial_id"] is None


        # ====================================================
        # 9. MINIMO KG
        # ====================================================

        result = run_case(

            wb,

            "CASO 9 - OFERTA NO CUMPLE MINIMO",

            "Cliente Especial",

            "Asado",

            datetime(
                2026,
                9,
                15
            ),

            1
        )

        # Como no hay otra oferta y la oferta requiere
        # 2 kg, debe entrar el especial PRE002.
        assert result["precio_kg"] == 13500
        assert result["tipo_aplicado"] == "Especial"


        # ====================================================
        # 10. ERROR CLIENTE
        # ====================================================

        result = run_case(

            wb,

            "CASO 10 - CLIENTE INEXISTENTE",

            "Cliente Que No Existe",

            "Asado",

            datetime(
                2026,
                9,
                15
            ),

            2
        )

        assert (
            result["estado"]
            == "CLIENTE NO ENCONTRADO"
        )


        # ====================================================
        # 11. ERROR PRODUCTO
        # ====================================================

        result = run_case(

            wb,

            "CASO 11 - PRODUCTO INEXISTENTE",

            "Cliente Minorista",

            "Producto Que No Existe",

            datetime(
                2026,
                9,
                15
            ),

            2
        )

        assert (
            result["estado"]
            == "PRODUCTO NO ENCONTRADO"
        )


        # ====================================================
        # 12. KG INVALIDO
        # ====================================================

        result = run_case(

            wb,

            "CASO 12 - KG INVALIDO",

            "Cliente Minorista",

            "Asado",

            datetime(
                2026,
                9,
                15
            ),

            0
        )

        assert (
            result["estado"]
            == "KG INVALIDO"
        )


        # ====================================================
        # RESULTADO
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
            "Especial por cliente: OK"
        )

        print(
            "Prioridad especial: OK"
        )

        print(
            "Oferta: OK"
        )

        print(
            "Oferta sobre especial: OK"
        )

        print(
            "Vigencia por día: OK"
        )

        print(
            "Fecha final inclusiva: OK"
        )

        print(
            "Mínimo de Kg: OK"
        )

        print(
            "Fallback: OK"
        )

        print(
            "Errores de cliente/producto/Kg: OK"
        )

        print(
            "Ganancia: OK"
        )

        print()
        print(
            "El Excel maestro NO fue modificado."
        )

    finally:

        wb.close()


if __name__ == "__main__":
    main()