from __future__ import annotations

from datetime import datetime


def buscar_cliente_por_nombre(
    workbook,
    nombre_cliente: str
):
    """
    Busca un cliente por nombre.

    Devuelve:
        dict con los datos del cliente
        None si no existe.
    """

    if not nombre_cliente:
        return None

    hoja = workbook["03_CLIENTES"]

    nombre_buscado = (
        str(nombre_cliente)
        .strip()
        .lower()
    )

    for fila in range(
        2,
        hoja.max_row + 1
    ):

        nombre = hoja.cell(
            fila,
            2
        ).value

        if not nombre:
            continue

        if (
            str(nombre)
            .strip()
            .lower()
            == nombre_buscado
        ):

            return {
                "id": hoja.cell(
                    fila,
                    1
                ).value,

                "nombre": nombre,

                "tipo_cliente": hoja.cell(
                    fila,
                    5
                ).value,

                "estado": hoja.cell(
                    fila,
                    6
                ).value,
            }

    return None


def buscar_producto_por_nombre(
    workbook,
    nombre_producto: str
):
    """
    Busca un producto por nombre.

    Devuelve:
        dict con los datos del producto
        None si no existe.
    """

    if not nombre_producto:
        return None

    hoja = workbook["04_PRODUCTOS"]

    nombre_buscado = (
        str(nombre_producto)
        .strip()
        .lower()
    )

    for fila in range(
        2,
        hoja.max_row + 1
    ):

        nombre = hoja.cell(
            fila,
            2
        ).value

        if not nombre:
            continue

        if (
            str(nombre)
            .strip()
            .lower()
            == nombre_buscado
        ):

            return {
                "id": hoja.cell(
                    fila,
                    1
                ).value,

                "nombre": nombre,

                "categoria": hoja.cell(
                    fila,
                    3
                ).value,

                "costo": hoja.cell(
                    fila,
                    4
                ).value,

                "precio_minorista": hoja.cell(
                    fila,
                    5
                ).value,

                "precio_mayorista": hoja.cell(
                    fila,
                    6
                ).value,

                "precio_especial": hoja.cell(
                    fila,
                    7
                ).value,

                "estado": hoja.cell(
                    fila,
                    8
                ).value,
            }

    return None


def determinar_precio(
    producto,
    tipo_precio: str
):
    """
    Determina el precio base según el tipo de precio.

    Reglas actuales:

    Minorista:
        PrecioMinorista

    Mayorista:
        PrecioMayorista
        si no existe -> PrecioMinorista

    Especial:
        PrecioEspecial
        si no existe -> PrecioMinorista
    """

    if not producto:
        return None

    tipo = (
        str(tipo_precio or "")
        .strip()
        .lower()
    )

    minorista = producto.get(
        "precio_minorista"
    )

    mayorista = producto.get(
        "precio_mayorista"
    )

    especial = producto.get(
        "precio_especial"
    )

    if tipo == "mayorista":

        if mayorista not in (
            None,
            ""
        ):

            return float(
                mayorista
            )

        if minorista not in (
            None,
            ""
        ):

            return float(
                minorista
            )

        return None

    if tipo == "especial":

        if especial not in (
            None,
            ""
        ):

            return float(
                especial
            )

        if minorista not in (
            None,
            ""
        ):

            return float(
                minorista
            )

        return None

    if minorista not in (
        None,
        ""
    ):

        return float(
            minorista
        )

    return None


def calcular_venta(
    workbook,
    nombre_cliente: str,
    nombre_producto: str,
    tipo_precio: str,
    kg
):
    """
    Procesa una venta completa.

    Devuelve un diccionario con:
        cliente
        producto
        precio
        kg
        importe
        costo
        ganancia
        ganancia_porcentaje
        estado
    """

    # --------------------------------------------------------
    # Cliente
    # --------------------------------------------------------

    cliente = buscar_cliente_por_nombre(
        workbook,
        nombre_cliente
    )

    if cliente is None:

        return {
            "estado": "CLIENTE NO ENCONTRADO"
        }

    # --------------------------------------------------------
    # Producto
    # --------------------------------------------------------

    producto = buscar_producto_por_nombre(
        workbook,
        nombre_producto
    )

    if producto is None:

        return {
            "estado": "PRODUCTO NO ENCONTRADO"
        }

    # --------------------------------------------------------
    # Kg
    # --------------------------------------------------------

    try:

        cantidad_kg = float(
            kg
        )

    except (
        TypeError,
        ValueError
    ):

        return {
            "estado": "KG INVALIDO"
        }

    if cantidad_kg <= 0:

        return {
            "estado": "KG INVALIDO"
        }

    # --------------------------------------------------------
    # Precio
    # --------------------------------------------------------

    precio = determinar_precio(
        producto,
        tipo_precio
    )

    if precio is None:

        return {
            "estado": "PRECIO NO DISPONIBLE"
        }

    # --------------------------------------------------------
    # Importe
    # --------------------------------------------------------

    importe = (
        cantidad_kg
        * precio
    )

    # --------------------------------------------------------
    # Costo
    # --------------------------------------------------------

    costo_unitario = producto.get(
        "costo"
    )

    ganancia = None
    ganancia_porcentaje = None

    if costo_unitario not in (
        None,
        ""
    ):

        try:

            costo_unitario = float(
                costo_unitario
            )

            costo_total = (
                cantidad_kg
                * costo_unitario
            )

            ganancia = (
                importe
                - costo_total
            )

            if importe != 0:

                ganancia_porcentaje = (
                    ganancia
                    / importe
                )

        except (
            TypeError,
            ValueError
        ):

            ganancia = None
            ganancia_porcentaje = None

    # --------------------------------------------------------
    # Resultado
    # --------------------------------------------------------

    return {

        "estado": "OK",

        "fecha": datetime.now(),

        "cliente_id": cliente["id"],

        "cliente":
            cliente["nombre"],

        "tipo_cliente":
            cliente["tipo_cliente"],

        "producto_id":
            producto["id"],

        "producto":
            producto["nombre"],

        "tipo_precio":
            tipo_precio,

        "kg":
            cantidad_kg,

        "precio_kg":
            precio,

        "importe":
            importe,

        "costo_unitario":
            costo_unitario,

        "ganancia":
            ganancia,

        "ganancia_porcentaje":
            ganancia_porcentaje,
    }