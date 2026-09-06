from __future__ import annotations

from datetime import date, datetime


# ============================================================
# FECHAS
# ============================================================

def convertir_fecha(valor):

    if valor in (None, ""):
        return None

    if isinstance(valor, datetime):
        return valor.date()

    if isinstance(valor, date):
        return valor

    if isinstance(valor, str):

        formatos = [
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%d-%m-%Y",
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


def fecha_vigente(
    fecha_operacion,
    fecha_inicio,
    fecha_fin
):
    """
    Vigencia inclusiva por día.
    """

    fecha = convertir_fecha(
        fecha_operacion
    )

    inicio = convertir_fecha(
        fecha_inicio
    )

    fin = convertir_fecha(
        fecha_fin
    )

    if fecha is None:
        return False

    if inicio is not None and fecha < inicio:
        return False

    if fin is not None and fecha > fin:
        return False

    return True


# ============================================================
# LEER REGLA
# ============================================================

def leer_regla(
    hoja,
    fila
):

    return {
        "precio_id": hoja.cell(
            fila,
            1
        ).value,

        "cliente_id": hoja.cell(
            fila,
            2
        ).value,

        "cliente": hoja.cell(
            fila,
            3
        ).value,

        "producto_id": hoja.cell(
            fila,
            4
        ).value,

        "producto": hoja.cell(
            fila,
            5
        ).value,

        "tipo_precio": hoja.cell(
            fila,
            6
        ).value,

        "precio": hoja.cell(
            fila,
            7
        ).value,

        "descuento_pct": hoja.cell(
            fila,
            8
        ).value,

        "fecha_inicio": hoja.cell(
            fila,
            9
        ).value,

        "fecha_fin": hoja.cell(
            fila,
            10
        ).value,

        "estado": hoja.cell(
            fila,
            11
        ).value,

        "prioridad": hoja.cell(
            fila,
            12
        ).value,
    }


# ============================================================
# REGLAS DEL CLIENTE + PRODUCTO
# ============================================================

def obtener_precios_especiales(
    workbook,
    cliente_id,
    producto_id
):

    if "22_PRECIOS" not in workbook.sheetnames:
        return []

    hoja = workbook[
        "22_PRECIOS"
    ]

    reglas = []

    for fila in range(
        2,
        hoja.max_row + 1
    ):

        regla_cliente = hoja.cell(
            fila,
            2
        ).value

        regla_producto = hoja.cell(
            fila,
            4
        ).value

        if (
            regla_cliente == cliente_id
            and
            regla_producto == producto_id
        ):

            reglas.append(
                leer_regla(
                    hoja,
                    fila
                )
            )

    return reglas


# ============================================================
# PRIORIDAD
# ============================================================

def obtener_prioridad(
    regla
):

    try:
        return int(
            regla.get(
                "prioridad",
                0
            ) or 0
        )

    except (
        TypeError,
        ValueError
    ):

        return 0


# ============================================================
# REGLA ESPECIAL VIGENTE
# ============================================================

def es_regla_especial_valida(
    regla,
    fecha_operacion
):

    estado = str(
        regla.get(
            "estado",
            ""
        )
    ).strip().lower()

    if estado != "activo":
        return False

    tipo = str(
        regla.get(
            "tipo_precio",
            ""
        )
    ).strip().lower()

    if tipo != "especial":
        return False

    if not fecha_vigente(
        fecha_operacion,
        regla.get("fecha_inicio"),
        regla.get("fecha_fin")
    ):
        return False

    precio = regla.get(
        "precio"
    )

    descuento = regla.get(
        "descuento_pct"
    )

    if precio in (
        None,
        ""
    ) and descuento in (
        None,
        ""
    ):
        return False

    return True


# ============================================================
# OBTENER MEJOR PRECIO ESPECIAL
# ============================================================

def obtener_precio_especial_cliente(
    workbook,
    cliente_id,
    producto_id,
    fecha_operacion
):

    reglas = obtener_precios_especiales(
        workbook,
        cliente_id,
        producto_id
    )

    candidatas = []

    for regla in reglas:

        if es_regla_especial_valida(
            regla,
            fecha_operacion
        ):

            candidatas.append(
                regla
            )

    if not candidatas:
        return None

    candidatas.sort(
        key=obtener_prioridad,
        reverse=True
    )

    regla = candidatas[0]

    precio = regla.get(
        "precio"
    )

    if precio not in (
        None,
        ""
    ):

        try:
            precio = float(
                precio
            )

            return {
                "precio": precio,
                "tipo_aplicado": "Especial",
                "motivo": regla.get(
                    "nombre_oferta"
                )
                or
                "Precio especial cliente",
                "precio_id": regla.get(
                    "precio_id"
                ),
            }

        except (
            TypeError,
            ValueError
        ):
            return None

    descuento = regla.get(
        "descuento_pct"
    )

    if descuento not in (
        None,
        ""
    ):

        try:

            descuento = float(
                descuento
            )

            return {
                "descuento_pct":
                    descuento,

                "tipo_aplicado":
                    "Especial",

                "motivo":
                    "Descuento especial cliente",

                "precio_id":
                    regla.get(
                        "precio_id"
                    ),
            }

        except (
            TypeError,
            ValueError
        ):

            return None

    return None