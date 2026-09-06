from __future__ import annotations

from datetime import datetime


import sys
from pathlib import Path


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


from precios import (
    determinar_precio_final
)


# ============================================================
# PRODUCTO DE PRUEBA
# ============================================================

PRODUCTO = {

    "id":
        "PRO_TEST_001",

    "nombre":
        "Asado Prueba",

    "categoria":
        "Vacuno",

    "costo":
        10000,

    "precio_minorista":
        16000,

    "precio_mayorista":
        None,

    "precio_especial":
        None,

    "estado":
        "Activo",
}


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print(
        "=" * 60
    )

    print(
        "MVP 0.1.3 - TEST MOTOR DE PRECIOS"
    )

    print(
        "=" * 60
    )

    # --------------------------------------------------------
    # Caso 1
    # --------------------------------------------------------
    # Sin oferta, sin especial,
    # minorista.
    # --------------------------------------------------------

    resultado = determinar_precio_final(

        PRODUCTO,

        tipo_precio="Minorista",

        fecha_operacion=datetime.now(),

        kg=2
    )

    print()
    print(
        "CASO 1 — MINORISTA"
    )

    print(
        resultado
    )

    assert (
        resultado[
            "precio"
        ]
        == 16000
    )

    # --------------------------------------------------------
    # Caso 2
    # --------------------------------------------------------
    # Mayorista inexistente:
    # debe caer a minorista.
    # --------------------------------------------------------

    resultado = determinar_precio_final(

        PRODUCTO,

        tipo_precio="Mayorista",

        fecha_operacion=datetime.now(),

        kg=2
    )

    print()
    print(
        "CASO 2 — MAYORISTA SIN PRECIO"
    )

    print(
        resultado
    )

    assert (
        resultado[
            "precio"
        ]
        == 16000
    )

    # --------------------------------------------------------
    # Caso 3
    # --------------------------------------------------------
    # Especial inexistente:
    # debe caer a minorista.
    # --------------------------------------------------------

    resultado = determinar_precio_final(

        PRODUCTO,

        tipo_precio="Especial",

        fecha_operacion=datetime.now(),

        kg=2
    )

    print()
    print(
        "CASO 3 — ESPECIAL SIN PRECIO"
    )

    print(
        resultado
    )

    assert (
        resultado[
            "precio"
        ]
        == 16000
    )

    # --------------------------------------------------------
    # Caso 4
    # --------------------------------------------------------
    # Precio mayorista real.
    # --------------------------------------------------------

    producto_mayorista = dict(
        PRODUCTO
    )

    producto_mayorista[
        "precio_mayorista"
    ] = 14500

    resultado = determinar_precio_final(

        producto_mayorista,

        tipo_precio="Mayorista",

        fecha_operacion=datetime.now(),

        kg=2
    )

    print()
    print(
        "CASO 4 — MAYORISTA"
    )

    print(
        resultado
    )

    assert (
        resultado[
            "precio"
        ]
        == 14500
    )

    # --------------------------------------------------------
    # Caso 5
    # --------------------------------------------------------
    # Precio especial cliente.
    # --------------------------------------------------------

    resultado = determinar_precio_final(

        PRODUCTO,

        tipo_precio="Minorista",

        fecha_operacion=datetime.now(),

        kg=2,

        precio_especial_cliente=14000
    )

    print()
    print(
        "CASO 5 — ESPECIAL CLIENTE"
    )

    print(
        resultado
    )

    assert (
        resultado[
            "precio"
        ]
        == 14000
    )

    # --------------------------------------------------------
    # Caso 6
    # --------------------------------------------------------
    # Oferta fija válida.
    # --------------------------------------------------------

    oferta = {

        "nombre_oferta":
            "Oferta Prueba",

        "precio_oferta":
            13500,

        "descuento_pct":
            None,

        "fecha_inicio":
            datetime.now(),

        "fecha_fin":
            datetime.now(),

        "min_kg":
            1,

        "estado":
            "Activo",
    }

    resultado = determinar_precio_final(

        PRODUCTO,

        tipo_precio="Minorista",

        fecha_operacion=datetime.now(),

        kg=2,

        oferta=oferta
    )

    print()
    print(
        "CASO 6 — OFERTA"
    )

    print(
        resultado
    )

    assert (
        resultado[
            "precio"
        ]
        == 13500
    )

    print()
    print(
        "=" * 60
    )

    print(
        "OK: MVP 0.1.3 — MOTOR DE PRECIOS SUPERADO"
    )

    print(
        "=" * 60
    )

    print()
    print(
        "Prioridad Oferta → Especial → TipoPrecio: OK"
    )

    print(
        "Fallback a Minorista: OK"
    )

    print(
        "Mayorista: OK"
    )

    print(
        "Especial: OK"
    )

    print(
        "Oferta fija: OK"
    )

    print()


if __name__ == "__main__":
    main()