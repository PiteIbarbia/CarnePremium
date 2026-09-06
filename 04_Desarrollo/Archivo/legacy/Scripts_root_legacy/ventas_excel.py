from __future__ import annotations

from datetime import datetime

from ventas import calcular_venta


STATUS_OK = "OK"


def procesar_venta_excel(workbook, fila: int):
    """
    Procesa una fila de 05_VENTAS usando el motor de negocio.

    Interfaz visible:
        D = Cliente (nombre)
        E = TipoPrecio
        F = Producto (nombre)
        G = Kg

    Campos internos:
        O = _ClienteID
        P = _ProductoID
        Q = EstadoVenta
    """

    ws = workbook["05_VENTAS"]

    cliente = ws.cell(fila, 4).value
    tipo_precio = ws.cell(fila, 5).value
    producto = ws.cell(fila, 6).value
    kg = ws.cell(fila, 7).value

    resultado = calcular_venta(
        workbook,
        nombre_cliente=cliente,
        nombre_producto=producto,
        tipo_precio=tipo_precio,
        kg=kg,
    )

    estado = resultado.get("estado")

    if estado != STATUS_OK:
        ws.cell(fila, 17).value = estado
        return resultado

    ws.cell(fila, 1).value = f"VEN{fila - 1:06d}"
    ws.cell(fila, 2).value = datetime.now().date()
    ws.cell(fila, 3).value = datetime.now().time()

    ws.cell(fila, 15).value = resultado["cliente_id"]
    ws.cell(fila, 16).value = resultado["producto_id"]

    ws.cell(fila, 8).value = resultado["precio_kg"]
    ws.cell(fila, 9).value = resultado["importe"]
    ws.cell(fila, 10).value = resultado["ganancia"]
    ws.cell(fila, 11).value = resultado["ganancia_porcentaje"]
    ws.cell(fila, 17).value = STATUS_OK

    return resultado


def procesar_todas_las_ventas(workbook, fila_inicio=2, fila_fin=501):
    resultados = []

    ws = workbook["05_VENTAS"]

    for fila in range(fila_inicio, fila_fin + 1):
        cliente = ws.cell(fila, 4).value
        producto = ws.cell(fila, 6).value

        if not cliente and not producto:
            continue

        resultados.append(
            procesar_venta_excel(workbook, fila)
        )

    return resultados
