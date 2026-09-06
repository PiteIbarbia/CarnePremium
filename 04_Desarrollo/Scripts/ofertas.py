from __future__ import annotations

from datetime import date, datetime


# ============================================================
# UTILIDADES
# ============================================================

def convertir_fecha(valor):

    if valor in (
        None,
        ""
    ):
        return None

    if isinstance(
        valor,
        datetime
    ):
        return valor.date()

    if isinstance(
        valor,
        date
    ):
        return valor

    if isinstance(
        valor,
        str
    ):

        formatos = [
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%d-%m-%Y"
        ]

        for formato in formatos:

            try:

                return datetime.strptime(
                    valor.strip(),
                    formato
                ).date()

            except ValueError:

                continue

    return None


# ============================================================
# LEER UNA OFERTA
# ============================================================

def leer_oferta(
    hoja,
    fila
):

    return {

        "id":
            hoja.cell(
                fila,
                1
            ).value,

        "producto_id":
            hoja.cell(
                fila,
                2
            ).value,

        "nombre_oferta":
            hoja.cell(
                fila,
                3
            ).value,

        "precio_oferta":
            hoja.cell(
                fila,
                4
            ).value,

        "descuento_pct":
            hoja.cell(
                fila,
                5
            ).value,

        "fecha_inicio":
            hoja.cell(
                fila,
                6
            ).value,

        "fecha_fin":
            hoja.cell(
                fila,
                7
            ).value,

        "min_kg":
            hoja.cell(
                fila,
                8
            ).value,

        "estado":
            hoja.cell(
                fila,
                9
            ).value,

        "prioridad":
            hoja.cell(
                fila,
                10
            ).value,
    }


# ============================================================
# OFERTAS DEL PRODUCTO
# ============================================================

def obtener_ofertas_producto(
    workbook,
    producto_id
):

    if (
        "23_OFERTAS"
        not in workbook.sheetnames
    ):

        return []

    hoja = workbook[
        "23_OFERTAS"
    ]

    resultados = []

    for fila in range(
        2,
        hoja.max_row + 1
    ):

        oferta_producto_id = (
            hoja.cell(
                fila,
                2
            ).value
        )

        if (
            oferta_producto_id
            == producto_id
        ):

            resultados.append(
                leer_oferta(
                    hoja,
                    fila
                )
            )

    return resultados


# ============================================================
# VIGENCIA
# ============================================================

def oferta_vigente(
    oferta,
    fecha_operacion
):

    estado = str(
        oferta.get(
            "estado",
            "Activo"
        )
    ).strip().lower()

    if estado != "activo":

        return False

    fecha = convertir_fecha(
        fecha_operacion
    )

    inicio = convertir_fecha(
        oferta.get(
            "fecha_inicio"
        )
    )

    fin = convertir_fecha(
        oferta.get(
            "fecha_fin"
        )
    )

    if fecha is None:

        return False

    if (
        inicio is not None
        and fecha < inicio
    ):

        return False

    # La fecha final es INCLUSIVA.
    if (
        fin is not None
        and fecha > fin
    ):

        return False

    return True


# ============================================================
# KG MINIMOS
# ============================================================

def oferta_cumple_minimo(
    oferta,
    kg
):

    try:

        cantidad = float(
            kg
            or 0
        )

    except (
        TypeError,
        ValueError
    ):

        return False

    try:

        minimo = float(
            oferta.get(
                "min_kg",
                0
            )
            or 0
        )

    except (
        TypeError,
        ValueError
    ):

        minimo = 0

    return (
        cantidad >= minimo
    )


# ============================================================
# PRIORIDAD
# ============================================================

def prioridad_oferta(
    oferta
):

    try:

        return int(
            oferta.get(
                "prioridad",
                0
            )
            or 0
        )

    except (
        TypeError,
        ValueError
    ):

        return 0


# ============================================================
# MEJOR OFERTA VALIDA
# ============================================================

def obtener_mejor_oferta(
    workbook,
    producto_id,
    fecha_operacion,
    kg
):

    ofertas = obtener_ofertas_producto(
        workbook,
        producto_id
    )

    candidatas = []

    for oferta in ofertas:

        if not oferta_vigente(
            oferta,
            fecha_operacion
        ):

            continue

        if not oferta_cumple_minimo(
            oferta,
            kg
        ):

            continue

        precio = oferta.get(
            "precio_oferta"
        )

        descuento = oferta.get(
            "descuento_pct"
        )

        tiene_precio = (
            precio
            not in (
                None,
                ""
            )
        )

        tiene_descuento = (
            descuento
            not in (
                None,
                ""
            )
        )

        if not (
            tiene_precio
            or
            tiene_descuento
        ):

            continue

        candidatas.append(
            oferta
        )

    if not candidatas:

        return None

    candidatas.sort(
        key=prioridad_oferta,
        reverse=True
    )

    return candidatas[0]


# ============================================================
# COMPATIBILIDAD
# ============================================================

def obtener_oferta_producto(
    workbook,
    producto_id=None,
    producto_nombre=None
):

    """
    Función de compatibilidad con la versión anterior.

    Devuelve la oferta válida de mayor prioridad
    cuando sea posible identificar el producto.
    """

    if producto_id is None:

        return None

    return obtener_mejor_oferta(
        workbook,
        producto_id,
        date.today(),
        0
    )