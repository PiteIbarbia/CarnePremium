from __future__ import annotations

from datetime import datetime

from motor_comercial import determinar_precio_venta

from motor_comercial import (
    determinar_precio_venta
)


STATUS_OK = "OK"


def procesar_venta_excel(
    workbook,
    fila: int
):
    """
    Procesa una fila de 05_VENTAS usando
    el motor comercial central.

    INTERFAZ VISIBLE:

        D = Cliente
        E = TipoPrecio
        F = Producto
        G = Kg

    CAMPOS INTERNOS:

        O = _ClienteID
        P = _ProductoID
        Q = EstadoVenta

    CAMPOS COMERCIALES:

        R = TipoAplicado
        S = MotivoPrecio
        T = OfertaID
        U = OfertaNombre
    """

    ws = workbook["05_VENTAS"]

    # ========================================================
    # DATOS DE ENTRADA
    # ========================================================

    nombre_cliente = ws.cell(
        fila,
        4
    ).value

    tipo_precio = ws.cell(
        fila,
        5
    ).value

    nombre_producto = ws.cell(
        fila,
        6
    ).value

    kg = ws.cell(
        fila,
        7
    ).value

    # ========================================================
    # FECHA
    # ========================================================

    fecha_venta = ws.cell(
        fila,
        2
    ).value

    if not fecha_venta:

        fecha_venta = datetime.now()

    # ========================================================
    # MOTOR COMERCIAL
    # ========================================================

    resultado = determinar_precio_venta(

        workbook,

        nombre_cliente,

        nombre_producto,

        fecha_venta,

        kg
    )

    estado = resultado.get(
        "estado"
    )

    # ========================================================
    # ERROR
    # ========================================================

    if estado != STATUS_OK:

        ws.cell(
            fila,
            17
        ).value = estado

        # Limpiar campos comerciales
        for columna in range(
            18,
            22
        ):

            ws.cell(
                fila,
                columna
            ).value = None

        return resultado

    # ========================================================
    # IDENTIFICADORES INTERNOS
    # ========================================================

    ws.cell(
        fila,
        15
    ).value = resultado.get(
        "cliente_id"
    )

    ws.cell(
        fila,
        16
    ).value = resultado.get(
        "producto_id"
    )

    # ========================================================
    # PRECIO
    # ========================================================

    precio_kg = resultado.get(
        "precio_kg"
    )

    ws.cell(
        fila,
        8
    ).value = precio_kg

    # ========================================================
    # IMPORTE
    # ========================================================

    try:

        cantidad_kg = float(
            kg
        )

        importe = (
            cantidad_kg
            * float(precio_kg)
        )

    except (
        TypeError,
        ValueError
    ):

        importe = None

    ws.cell(
        fila,
        9
    ).value = importe

    # ========================================================
    # COSTO / GANANCIA
    # ========================================================

    costo = resultado.get(
        "costo_unitario"
    )

    ganancia = None
    ganancia_porcentaje = None

    if (
        importe is not None
        and costo not in (
            None,
            ""
        )
    ):

        try:

            costo_total = (
                cantidad_kg
                * float(costo)
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

            pass

    ws.cell(
        fila,
        10
    ).value = ganancia

    ws.cell(
        fila,
        11
    ).value = ganancia_porcentaje

    # ========================================================
    # INFORMACIÓN COMERCIAL
    # ========================================================

    ws.cell(
        fila,
        18
    ).value = resultado.get(
        "tipo_aplicado"
    )

    ws.cell(
        fila,
        19
    ).value = resultado.get(
        "motivo"
    )

    ws.cell(
        fila,
        20
    ).value = resultado.get(
        "oferta_id"
    )

    ws.cell(
        fila,
        21
    ).value = resultado.get(
        "oferta_nombre"
    )

    # ========================================================
    # ESTADO
    # ========================================================

    ws.cell(
        fila,
        17
    ).value = STATUS_OK

    # ========================================================
    # FORMATOS
    # ========================================================

    ws.cell(
        fila,
        7
    ).number_format = "0.000"

    for columna in (
        8,
        9,
        10
    ):

        ws.cell(
            fila,
            columna
        ).number_format = (
            '$ #,##0.00'
        )

    ws.cell(
        fila,
        11
    ).number_format = "0.0%"

    return resultado


def procesar_todas_las_ventas(
    workbook,
    fila_inicio=2,
    fila_fin=501
):
    """
    Procesa todas las ventas no vacías.
    """

    resultados = []

    ws = workbook[
        "05_VENTAS"
    ]

    for fila in range(
        fila_inicio,
        fila_fin + 1
    ):

        cliente = ws.cell(
            fila,
            4
        ).value

        producto = ws.cell(
            fila,
            6
        ).value

        if not cliente and not producto:

            continue

        resultado = (
            procesar_venta_excel(
                workbook,
                fila
            )
        )

        resultados.append(
            resultado
        )

    return resultados