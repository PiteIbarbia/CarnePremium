from __future__ import annotations

from datetime import datetime

from motor_comercial import determinar_precio_venta, find_client


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
    hora_venta = ws.cell(
        fila,
        3
    ).value
    ahora = datetime.now()

    # Fijar fecha y hora al procesar para que no cambien con TODAY/NOW.
    if (
        fecha_venta in (None, "")
        or (isinstance(fecha_venta, str) and fecha_venta.startswith("="))
    ):
        fecha_venta = ahora
        ws.cell(fila, 2).value = fecha_venta
    if (
        hora_venta in (None, "")
        or (isinstance(hora_venta, str) and hora_venta.startswith("="))
    ):
        ws.cell(fila, 3).value = ahora

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

    medio_pago = ws.cell(fila, 12).value
    if estado == STATUS_OK and not str(medio_pago or "").strip():
        estado = "MEDIO DE PAGO REQUERIDO"
        resultado["estado"] = estado
        ws.cell(fila, 17).value = estado
        for columna in (8, 9, 10, 11, 15, 16, 18, 19, 20, 21, 22):
            ws.cell(fila, columna).value = None
        return resultado

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
            23
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

    # ID del precio especial aplicado (columna V).
    ws.cell(
        fila,
        22
    ).value = resultado.get(
        "precio_especial_id"
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


def procesar_cobro_excel(workbook, fila: int):
    """Valida y persiste un cobro de saldo anterior ingresado desde 05_VENTAS."""
    ws = workbook["05_VENTAS"]
    cliente_nombre = ws.cell(fila, 4).value
    medio_pago = str(ws.cell(fila, 12).value or "").strip()
    importe_cobrado = ws.cell(fila, 23).value
    fecha = ws.cell(fila, 2).value
    hora = ws.cell(fila, 3).value
    ahora = datetime.now()

    if fecha in (None, "") or (isinstance(fecha, str) and fecha.startswith("=")):
        fecha = ahora
        ws.cell(fila, 2).value = fecha
    if hora in (None, "") or (isinstance(hora, str) and hora.startswith("=")):
        ws.cell(fila, 3).value = ahora

    try:
        importe_cobrado = float(importe_cobrado)
    except (TypeError, ValueError):
        importe_cobrado = None

    client = find_client(workbook, cliente_nombre)
    if client is None:
        estado = "CLIENTE NO ENCONTRADO"
    elif str(client.get("estado") or "").strip().casefold() != "activo":
        estado = "CLIENTE INACTIVO"
    elif importe_cobrado is None or importe_cobrado <= 0:
        estado = "IMPORTE COBRO INVALIDO"
    elif not medio_pago:
        estado = "MEDIO DE PAGO REQUERIDO"
    elif medio_pago.casefold().replace(" ", "") == "cuentacorriente":
        estado = "COBRO NO VÁLIDO PARA CUENTA CORRIENTE"
    else:
        estado = "COBRO OK"

    # Las filas COBRO no son ventas de carne; limpiar resultados comerciales
    # para evitar que queden valores de una fila reutilizada.
    for columna in (8, 9, 10, 11, 16, 18, 19, 20, 21, 22):
        ws.cell(fila, columna).value = None
    if client is not None:
        ws.cell(fila, 15).value = client.get("id")
    else:
        ws.cell(fila, 15).value = None
    ws.cell(fila, 17).value = estado
    ws.cell(fila, 23).number_format = '$ #,##0.00'
    return {
        "estado": estado,
        "tipo_registro": "COBRO",
        "importe_cobrado": importe_cobrado,
        "cliente": cliente_nombre,
    }


def procesar_todas_las_ventas(
    workbook,
    fila_inicio=2,
    fila_fin=501
):
    """Procesa ventas de carne y cobros de saldo ingresados en 05_VENTAS."""
    resultados = []
    ws = workbook["05_VENTAS"]

    for fila in range(fila_inicio, fila_fin + 1):
        cliente = ws.cell(fila, 4).value
        producto = ws.cell(fila, 6).value
        tipo_registro = str(ws.cell(fila, 27).value or "").strip().upper()

        if tipo_registro == "COBRO":
            if cliente or ws.cell(fila, 23).value not in (None, ""):
                resultados.append(procesar_cobro_excel(workbook, fila))
            continue

        if not cliente and not producto:
            continue

        resultado = procesar_venta_excel(workbook, fila)
        resultados.append(resultado)

    return resultados