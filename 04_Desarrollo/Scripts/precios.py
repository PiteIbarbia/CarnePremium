from __future__ import annotations

from datetime import date, datetime


# ============================================================
# UTILIDADES
# ============================================================

def convertir_fecha(valor):
    """
    Convierte distintos formatos de fecha
    a datetime/date cuando es posible.
    """

    if valor in (None, ""):
        return None

    if isinstance(valor, datetime):
        return valor

    if isinstance(valor, date):
        return datetime.combine(
            valor,
            datetime.min.time()
        )

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
                )

            except ValueError:
                continue

    return None


def fecha_dentro_de_vigencia(
    fecha_operacion,
    fecha_inicio,
    fecha_fin
):
    """
    Determina si una fecha está dentro de una vigencia.

    Inicio vacío = sin límite inferior.
    Fin vacío = sin límite superior.
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
# PRECIO BASE DEL PRODUCTO
# ============================================================

def determinar_precio_base(
    producto,
    tipo_precio
):
    """
    Obtiene el precio base del producto.

    Reglas:

    Minorista:
        PrecioMinorista

    Mayorista:
        PrecioMayorista
        si está vacío:
            PrecioMinorista

    Especial:
        PrecioEspecial
        si está vacío:
            PrecioMinorista
    """

    if not producto:
        return None

    tipo = str(
        tipo_precio or ""
    ).strip().lower()

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


# ============================================================
# OFERTA
# ============================================================

def evaluar_oferta(
    oferta,
    fecha_operacion,
    kg
):
    """
    Evalúa una oferta individual.

    Una oferta es válida si:

    - está activa;
    - la fecha está vigente;
    - se cumple MinKg;
    - tiene PrecioOferta o Descuento%.
    """

    if not oferta:
        return None

    estado = str(
        oferta.get(
            "estado",
            "Activo"
        )
    ).strip().lower()

    if estado != "activo":
        return None

    if not fecha_dentro_de_vigencia(
        fecha_operacion,
        oferta.get("fecha_inicio"),
        oferta.get("fecha_fin")
    ):
        return None

    try:
        cantidad_kg = float(
            kg or 0
        )
    except (
        TypeError,
        ValueError
    ):
        return None

    try:
        minimo = float(
            oferta.get(
                "min_kg",
                0
            ) or 0
        )
    except (
        TypeError,
        ValueError
    ):
        minimo = 0

    if cantidad_kg < minimo:
        return None

    precio_oferta = oferta.get(
        "precio_oferta"
    )

    descuento = oferta.get(
        "descuento_pct"
    )

    # Precio fijo de oferta
    if precio_oferta not in (
        None,
        ""
    ):
        try:
            return {
                "precio": float(
                    precio_oferta
                ),
                "tipo": "Oferta",
                "motivo": oferta.get(
                    "nombre_oferta"
                )
            }

        except (
            TypeError,
            ValueError
        ):
            pass

    # Descuento porcentual
    if descuento not in (
        None,
        ""
    ):
        try:
            return {
                "descuento_pct": float(
                    descuento
                ),
                "tipo": "OfertaDescuento",
                "motivo": oferta.get(
                    "nombre_oferta"
                )
            }

        except (
            TypeError,
            ValueError
        ):
            pass

    return None


# ============================================================
# PRECIO FINAL
# ============================================================

def determinar_precio_final(
    producto,
    tipo_precio,
    fecha_operacion=None,
    kg=0,
    oferta=None,
    precio_especial_cliente=None
):
    """
    Determina el precio final.

    PRIORIDAD:

    1. Oferta válida
    2. Precio especial por cliente
    3. Precio según tipo
    4. Precio minorista como fallback

    Devuelve:

        {
            "precio": ...,
            "tipo_aplicado": ...,
            "motivo": ...
        }
    """

    if not producto:
        return {
            "precio": None,
            "tipo_aplicado": None,
            "motivo": "Producto inexistente"
        }

    # --------------------------------------------------------
    # PRECIO BASE
    # --------------------------------------------------------

    precio_base = determinar_precio_base(
        producto,
        tipo_precio
    )

    if precio_base is None:
        return {
            "precio": None,
            "tipo_aplicado": None,
            "motivo": "Precio no disponible"
        }

    # --------------------------------------------------------
    # 1 — OFERTA
    # --------------------------------------------------------

    resultado_oferta = evaluar_oferta(
        oferta,
        fecha_operacion,
        kg
    )

    if resultado_oferta:

        if (
            resultado_oferta[
                "tipo"
            ]
            == "Oferta"
        ):

            return {
                "precio":
                    resultado_oferta[
                        "precio"
                    ],

                "tipo_aplicado":
                    "Oferta",

                "motivo":
                    resultado_oferta[
                        "motivo"
                    ]
            }

        if (
            resultado_oferta[
                "tipo"
            ]
            == "OfertaDescuento"
        ):

            descuento = (
                resultado_oferta[
                    "descuento_pct"
                ]
            )

            precio_final = (
                precio_base
                * (
                    1
                    - (
                        descuento
                        / 100
                    )
                )
            )

            return {
                "precio":
                    precio_final,

                "tipo_aplicado":
                    "Oferta",

                "motivo":
                    resultado_oferta[
                        "motivo"
                    ]
            }

    # --------------------------------------------------------
    # 2 — PRECIO ESPECIAL POR CLIENTE
    # --------------------------------------------------------

    if precio_especial_cliente not in (
        None,
        ""
    ):

        try:

            return {
                "precio":
                    float(
                        precio_especial_cliente
                    ),

                "tipo_aplicado":
                    "Especial",

                "motivo":
                    "Precio especial cliente"
            }

        except (
            TypeError,
            ValueError
        ):
            pass

    # --------------------------------------------------------
    # 3 — PRECIO SEGÚN TIPO
    # --------------------------------------------------------

    return {
        "precio":
            precio_base,

        "tipo_aplicado":
            tipo_precio
            if tipo_precio
            else "Minorista",

        "motivo":
            "Precio base"
    }