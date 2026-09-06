from __future__ import annotations

from datetime import datetime
from typing import Any

from motor_comercial import calculate_sale


# ============================================================
# CONFIGURACIÓN
# ============================================================

DEFAULT_MAX_ROWS = 501


# ============================================================
# UTILIDADES
# ============================================================

def normalizar_texto(
    value: Any
) -> str:

    if value is None:
        return ""

    return (
        str(value)
        .strip()
        .casefold()
    )


def obtener_mapa_columnas(
    worksheet
) -> dict[str, int]:
    """
    Devuelve:

        {
            "Cliente": 4,
            "Producto": 6,
            ...
        }

    según los encabezados reales de la hoja.

    Esto evita depender de posiciones fijas.
    """

    result: dict[str, int] = {}

    for column in range(
        1,
        worksheet.max_column + 1
    ):

        header = worksheet.cell(
            1,
            column
        ).value

        if header in (
            None,
            ""
        ):
            continue

        result[
            str(header).strip()
        ] = column

    return result


def exigir_columnas(
    column_map: dict[str, int],
    required: list[str]
):
    """
    Verifica que existan todas las columnas necesarias.
    """

    missing = [
        name
        for name in required
        if name not in column_map
    ]

    if missing:

        raise RuntimeError(
            "05_VENTAS: faltan columnas: "
            + ", ".join(missing)
        )


def escribir_si_existe(
    worksheet,
    column_map: dict[str, int],
    row: int,
    header: str,
    value: Any
):
    """
    Escribe solo si la columna existe.

    Esto permite compatibilidad entre builds.
    """

    column = column_map.get(
        header
    )

    if column is None:
        return

    worksheet.cell(
        row,
        column
    ).value = value


def limpiar_si_existe(
    worksheet,
    column_map: dict[str, int],
    row: int,
    headers: list[str]
):
    """
    Limpia campos calculados/comerciales
    cuando la venta contiene un error.
    """

    for header in headers:

        escribir_si_existe(
            worksheet,
            column_map,
            row,
            header,
            None
        )


# ============================================================
# CAMPOS DE VENTA
# ============================================================

REQUIRED_SALES_COLUMNS = [

    "Cliente",
    "TipoPrecio",
    "Producto",
    "Kg",
]


OPTIONAL_OUTPUT_COLUMNS = [

    "PrecioKg",
    "Importe",
    "Ganancia$",
    "Ganancia%",

    "EstadoVenta",

    "TipoAplicado",
    "MotivoPrecio",

    "OfertaID",
    "OfertaNombre",

    "PrecioEspecialID",

]


# ============================================================
# PROCESAR UNA VENTA
# ============================================================

def procesar_venta(
    workbook,
    fila: int
):
    """
    Procesa una venta individual de 05_VENTAS.

    Entrada mínima:

        Cliente
        TipoPrecio
        Producto
        Kg

    La fecha se toma de la columna Fecha
    si existe; de lo contrario se utiliza
    la fecha actual.

    El resultado del motor comercial se
    escribe nuevamente en la misma fila.
    """

    worksheet = workbook[
        "05_VENTAS"
    ]

    column_map = (
        obtener_mapa_columnas(
            worksheet
        )
    )

    exigir_columnas(
        column_map,
        REQUIRED_SALES_COLUMNS
    )

    # --------------------------------------------------------
    # ENTRADA
    # --------------------------------------------------------

    client = worksheet.cell(
        fila,
        column_map["Cliente"]
    ).value

    client_type = worksheet.cell(
        fila,
        column_map["TipoPrecio"]
    ).value

    product = worksheet.cell(
        fila,
        column_map["Producto"]
    ).value

    kg = worksheet.cell(
        fila,
        column_map["Kg"]
    ).value

    # --------------------------------------------------------
    # FILA VACÍA
    # --------------------------------------------------------

    if (
        normalizar_texto(client)
        == ""
        and
        normalizar_texto(product)
        == ""
        and
        kg in (
            None,
            ""
        )
    ):

        return {
            "estado": "VACIA"
        }

    # --------------------------------------------------------
    # FECHA
    # --------------------------------------------------------

    if "Fecha" in column_map:

        operation_date = worksheet.cell(
            fila,
            column_map["Fecha"]
        ).value

    else:

        operation_date = None

    if not operation_date:

        operation_date = datetime.now()

    # --------------------------------------------------------
    # MOTOR COMERCIAL
    # --------------------------------------------------------

    resultado = calculate_sale(

        workbook,

        client,

        product,

        operation_date,

        kg
    )

    # --------------------------------------------------------
    # ERROR
    # --------------------------------------------------------

    if (
        resultado.get("status")
        != "OK"
    ):

        escribir_si_existe(
            worksheet,
            column_map,
            fila,
            "EstadoVenta",
            resultado.get(
                "status"
            )
        )

        limpiar_si_existe(

            worksheet,

            column_map,

            fila,

            [
                "PrecioKg",
                "Importe",
                "Ganancia$",
                "Ganancia%",
                "TipoAplicado",
                "MotivoPrecio",
                "OfertaID",
                "OfertaNombre",
                "PrecioEspecialID",
            ]
        )

        return resultado

    # --------------------------------------------------------
    # ID CLIENTE
    # --------------------------------------------------------

    escribir_si_existe(
        worksheet,
        column_map,
        fila,
        "_ClienteID",
        resultado.get(
            "cliente_id"
        )
    )

    # --------------------------------------------------------
    # ID PRODUCTO
    # --------------------------------------------------------

    escribir_si_existe(
        worksheet,
        column_map,
        fila,
        "_ProductoID",
        resultado.get(
            "producto_id"
        )
    )

    # --------------------------------------------------------
    # PRECIO
    # --------------------------------------------------------

    escribir_si_existe(
        worksheet,
        column_map,
        fila,
        "PrecioKg",
        resultado.get(
            "precio_kg"
        )
    )

    # --------------------------------------------------------
    # IMPORTE
    # --------------------------------------------------------

    escribir_si_existe(
        worksheet,
        column_map,
        fila,
        "Importe",
        resultado.get(
            "importe"
        )
    )

    # --------------------------------------------------------
    # GANANCIA
    # --------------------------------------------------------

    escribir_si_existe(
        worksheet,
        column_map,
        fila,
        "Ganancia$",
        resultado.get(
            "ganancia"
        )
    )

    # --------------------------------------------------------
    # GANANCIA %
    # --------------------------------------------------------

    escribir_si_existe(
        worksheet,
        column_map,
        fila,
        "Ganancia%",
        resultado.get(
            "ganancia_pct"
        )
    )

    # --------------------------------------------------------
    # TIPO APLICADO
    # --------------------------------------------------------

    escribir_si_existe(
        worksheet,
        column_map,
        fila,
        "TipoAplicado",
        resultado.get(
            "tipo_aplicado"
        )
    )

    # --------------------------------------------------------
    # MOTIVO
    # --------------------------------------------------------

    escribir_si_existe(
        worksheet,
        column_map,
        fila,
        "MotivoPrecio",
        resultado.get(
            "motivo"
        )
    )

    # --------------------------------------------------------
    # OFERTA
    # --------------------------------------------------------

    escribir_si_existe(
        worksheet,
        column_map,
        fila,
        "OfertaID",
        resultado.get(
            "oferta_id"
        )
    )

    escribir_si_existe(
        worksheet,
        column_map,
        fila,
        "OfertaNombre",
        resultado.get(
            "oferta_nombre"
        )
    )

    # --------------------------------------------------------
    # PRECIO ESPECIAL
    # --------------------------------------------------------

    escribir_si_existe(
        worksheet,
        column_map,
        fila,
        "PrecioEspecialID",
        resultado.get(
            "precio_especial_id"
        )
    )

    # --------------------------------------------------------
    # ESTADO
    # --------------------------------------------------------

    escribir_si_existe(
        worksheet,
        column_map,
        fila,
        "EstadoVenta",
        "OK"
    )

    # --------------------------------------------------------
    # FORMATOS
    # --------------------------------------------------------

    if "Kg" in column_map:

        worksheet.cell(
            fila,
            column_map["Kg"]
        ).number_format = (
            "0.000"
        )

    for header in (
        "PrecioKg",
        "Importe",
        "Ganancia$",
    ):

        column = column_map.get(
            header
        )

        if column:

            worksheet.cell(
                fila,
                column
            ).number_format = (
                '$ #,##0.00'
            )

    if "Ganancia%" in column_map:

        worksheet.cell(
            fila,
            column_map["Ganancia%"]
        ).number_format = (
            "0.0%"
        )

    return resultado


# ============================================================
# PROCESAR TODAS LAS VENTAS
# ============================================================

def procesar_todas_las_ventas(
    workbook,
    fila_inicio: int = 2,
    fila_fin: int = DEFAULT_MAX_ROWS
):
    """
    Procesa todas las filas de 05_VENTAS
    que contengan datos.
    """

    worksheet = workbook[
        "05_VENTAS"
    ]

    results = []

    column_map = (
        obtener_mapa_columnas(
            worksheet
        )
    )

    exigir_columnas(
        column_map,
        REQUIRED_SALES_COLUMNS
    )

    for row in range(
        fila_inicio,
        fila_fin + 1
    ):

        client = worksheet.cell(
            row,
            column_map["Cliente"]
        ).value

        product = worksheet.cell(
            row,
            column_map["Producto"]
        ).value

        kg = worksheet.cell(
            row,
            column_map["Kg"]
        ).value

        if (
            normalizar_texto(client)
            == ""
            and
            normalizar_texto(product)
            == ""
            and
            kg in (
                None,
                ""
            )
        ):
            continue

        result = procesar_venta(
            workbook,
            row
        )

        results.append(
            result
        )

    return results


# ============================================================
# ALIAS DE COMPATIBILIDAD
# ============================================================

def procesar_venta_excel(
    workbook,
    fila: int
):
    """
    Alias compatible con el módulo anterior.
    """

    return procesar_venta(
        workbook,
        fila
    )


def procesar_todas_las_ventas_excel(
    workbook,
    fila_inicio: int = 2,
    fila_fin: int = DEFAULT_MAX_ROWS
):
    """
    Alias compatible con el módulo anterior.
    """

    return procesar_todas_las_ventas(
        workbook,
        fila_inicio,
        fila_fin
    )