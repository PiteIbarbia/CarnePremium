from __future__ import annotations

from datetime import date, datetime


# ============================================================
# MOTOR COMERCIAL CARNE PREMIUM
# BLOQUE COMERCIAL v0.2
# ============================================================
#
# PRIORIDAD:
#
# 1. OFERTA
# 2. PRECIO ESPECIAL CLIENTE
# 3. MAYORISTA
# 4. MINORISTA
#
# Las vigencias se manejan POR DÍA e INCLUSIVAS.
#
# Este archivo es autosuficiente.
# No depende de precios.py, ofertas.py ni
# precios_cliente.py.
#
# Devuelve nombres de claves compatibles
# con ventas_excel.py:
#
#   estado
#   cliente_id
#   cliente
#   tipo_cliente
#   producto_id
#   producto
#   kg
#   precio_kg
#   tipo_aplicado
#   motivo
#   oferta_id
#   oferta_nombre
#   precio_especial_id
#   costo_unitario
#
# ============================================================


# ============================================================
# FECHAS
# ============================================================

def to_date(value):
    """
    Convierte datetime, date o string a date.

    Formatos admitidos:
        dd/mm/yyyy
        yyyy-mm-dd
        dd-mm-yyyy
    """

    if value in (None, ""):
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    if isinstance(value, str):

        value = value.strip()

        for fmt in (
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%d-%m-%Y",
        ):

            try:

                return datetime.strptime(
                    value,
                    fmt
                ).date()

            except ValueError:
                continue

    return None


def date_is_valid(
    operation_date,
    start_date=None,
    end_date=None
):
    """
    Vigencia inclusiva por día.

    Ejemplo:
        Inicio 01/09/2026
        Fin    07/09/2026

    Es válida:
        01, 02, 03, 04, 05, 06 y 07.
    """

    current = to_date(
        operation_date
    )

    start = to_date(
        start_date
    )

    end = to_date(
        end_date
    )

    if current is None:
        return False

    if (
        start is not None
        and current < start
    ):
        return False

    if (
        end is not None
        and current > end
    ):
        return False

    return True


# ============================================================
# NUMEROS
# ============================================================

def safe_float(value):

    if value in (None, ""):
        return None

    try:
        return float(value)

    except (
        TypeError,
        ValueError
    ):
        return None


# ============================================================
# CLIENTE
# ============================================================

def find_client(
    workbook,
    client_name
):
    """
    Busca cliente por nombre.

    Comparación:
        sin espacios laterales
        sin distinguir mayúsculas/minúsculas
    """

    if not client_name:
        return None

    ws = workbook[
        "03_CLIENTES"
    ]

    target = (
        str(client_name)
        .strip()
        .casefold()
    )

    for row in range(
        2,
        ws.max_row + 1
    ):

        name = ws.cell(
            row,
            2
        ).value

        if not name:
            continue

        current = (
            str(name)
            .strip()
            .casefold()
        )

        if current == target:

            # Los ID de la plantilla son fórmulas de Excel. openpyxl lee el
            # texto de la fórmula, no su resultado calculado. No copiar esa
            # fórmula a 05_VENTAS porque sus referencias relativas cambiarían.
            raw_id = ws.cell(row, 1).value
            client_id = raw_id
            if (
                raw_id in (None, "")
                or (
                    isinstance(raw_id, str)
                    and raw_id.startswith("=")
                )
            ):
                client_id = f"CLI{row - 1:06d}"

            return {

                "id":
                    client_id,

                "nombre":
                    name,

                "tipo_cliente":
                    ws.cell(
                        row,
                        5
                    ).value,

                "estado":
                    ws.cell(
                        row,
                        6
                    ).value,
            }

    return None


# ============================================================
# PRODUCTO
# ============================================================

def find_product(
    workbook,
    product_name
):
    """
    Busca producto por nombre.
    """

    if not product_name:
        return None

    ws = workbook[
        "04_PRODUCTOS"
    ]

    target = (
        str(product_name)
        .strip()
        .casefold()
    )

    for row in range(
        2,
        ws.max_row + 1
    ):

        name = ws.cell(
            row,
            2
        ).value

        if not name:
            continue

        current = (
            str(name)
            .strip()
            .casefold()
        )

        if current == target:

            # Resolver el ID de producto si la celda contiene fórmula Excel.
            # Nunca devolver el texto de la fórmula como si fuera un ID.
            raw_id = ws.cell(row, 1).value
            product_id = raw_id
            if (
                raw_id in (None, "")
                or (
                    isinstance(raw_id, str)
                    and raw_id.startswith("=")
                )
            ):
                product_id = f"PRO{row - 1:06d}"

            return {

                "id":
                    product_id,

                "nombre":
                    name,

                "costo":
                    ws.cell(
                        row,
                        4
                    ).value,

                "precio_minorista":
                    ws.cell(
                        row,
                        5
                    ).value,

                "precio_mayorista":
                    ws.cell(
                        row,
                        6
                    ).value,

                "precio_especial":
                    ws.cell(
                        row,
                        7
                    ).value,

                "estado":
                    ws.cell(
                        row,
                        8
                    ).value,
            }

    return None


# ============================================================
# OFERTAS
# ============================================================

def read_offers(
    workbook,
    product_id
):
    """
    Lee todas las ofertas del producto.

    Estructura actual 23_OFERTAS:

        A OfertaID
        B ProductoID
        C Producto
        D NombreOferta
        E PrecioOferta
        F Descuento%
        G FechaInicio
        H FechaFin
        I MinKg
        J Estado
        K Prioridad
    """

    if "23_OFERTAS" not in workbook.sheetnames:
        return []

    ws = workbook[
        "23_OFERTAS"
    ]

    results = []

    for row in range(
        2,
        ws.max_row + 1
    ):

        row_product_id = (
            ws.cell(
                row,
                2
            ).value
        )

        if row_product_id != product_id:
            continue

        results.append({

            "id":
                ws.cell(
                    row,
                    1
                ).value,

            "product_id":
                row_product_id,

            "product":
                ws.cell(
                    row,
                    3
                ).value,

            "name":
                ws.cell(
                    row,
                    4
                ).value,

            "price":
                ws.cell(
                    row,
                    5
                ).value,

            "discount":
                ws.cell(
                    row,
                    6
                ).value,

            "start":
                ws.cell(
                    row,
                    7
                ).value,

            "end":
                ws.cell(
                    row,
                    8
                ).value,

            "min_kg":
                ws.cell(
                    row,
                    9
                ).value,

            "status":
                ws.cell(
                    row,
                    10
                ).value,

            "priority":
                ws.cell(
                    row,
                    11
                ).value,
        })

    return results


def get_best_offer(
    workbook,
    product_id,
    operation_date,
    kg
):
    """
    Busca la mejor oferta válida.

    Condiciones:

        Estado = Activo
        Fecha vigente
        Kg >= MinKg
        Precio o descuento informado

    Selección:

        Mayor prioridad.
    """

    quantity = safe_float(
        kg
    )

    if (
        quantity is None
        or quantity <= 0
    ):
        return None

    candidates = []

    for offer in read_offers(
        workbook,
        product_id
    ):

        status = str(
            offer.get(
                "status",
                ""
            )
        ).strip().casefold()

        if status != "activo":
            continue

        if not date_is_valid(
            operation_date,
            offer.get("start"),
            offer.get("end")
        ):
            continue

        minimum = (
            safe_float(
                offer.get(
                    "min_kg"
                )
            )
            or 0
        )

        if quantity < minimum:
            continue

        fixed_price = safe_float(
            offer.get(
                "price"
            )
        )

        discount = safe_float(
            offer.get(
                "discount"
            )
        )

        if (
            fixed_price is None
            and discount is None
        ):
            continue

        priority = (
            safe_float(
                offer.get(
                    "priority"
                )
            )
            or 0
        )

        candidates.append(
            (
                priority,
                offer
            )
        )

    if not candidates:
        return None

    candidates.sort(
        key=lambda item: item[0],
        reverse=True
    )

    return candidates[0][1]


# ============================================================
# PRECIOS ESPECIALES
# ============================================================

def read_special_prices(
    workbook,
    client_id,
    product_id
):
    """
    Lee las reglas especiales cliente/producto.

    Estructura actual 22_PRECIOS:

        A PrecioID
        B ClienteID
        C Cliente
        D ProductoID
        E Producto
        F TipoPrecio
        G Precio
        H Descuento%
        I FechaInicio
        J FechaFin
        K Estado
        L Prioridad
        M Motivo
    """

    if "22_PRECIOS" not in workbook.sheetnames:
        return []

    ws = workbook[
        "22_PRECIOS"
    ]

    results = []

    for row in range(
        2,
        ws.max_row + 1
    ):

        row_client_id = (
            ws.cell(
                row,
                2
            ).value
        )

        row_product_id = (
            ws.cell(
                row,
                4
            ).value
        )

        if (
            row_client_id != client_id
            or
            row_product_id != product_id
        ):
            continue

        results.append({

            "id":
                ws.cell(
                    row,
                    1
                ).value,

            "client_id":
                row_client_id,

            "client":
                ws.cell(
                    row,
                    3
                ).value,

            "product_id":
                row_product_id,

            "product":
                ws.cell(
                    row,
                    5
                ).value,

            "type":
                ws.cell(
                    row,
                    6
                ).value,

            "price":
                ws.cell(
                    row,
                    7
                ).value,

            "discount":
                ws.cell(
                    row,
                    8
                ).value,

            "start":
                ws.cell(
                    row,
                    9
                ).value,

            "end":
                ws.cell(
                    row,
                    10
                ).value,

            "status":
                ws.cell(
                    row,
                    11
                ).value,

            "priority":
                ws.cell(
                    row,
                    12
                ).value,

            "reason":
                ws.cell(
                    row,
                    13
                ).value,
        })

    return results


def get_best_special_price(
    workbook,
    client_id,
    product_id,
    operation_date
):
    """
    Obtiene el precio especial vigente
    de mayor prioridad.
    """

    candidates = []

    for rule in read_special_prices(
        workbook,
        client_id,
        product_id
    ):

        status = str(
            rule.get(
                "status",
                ""
            )
        ).strip().casefold()

        if status != "activo":
            continue

        price_type = str(
            rule.get(
                "type",
                ""
            )
        ).strip().casefold()

        if price_type != "especial":
            continue

        if not date_is_valid(
            operation_date,
            rule.get("start"),
            rule.get("end")
        ):
            continue

        fixed_price = safe_float(
            rule.get(
                "price"
            )
        )

        discount = safe_float(
            rule.get(
                "discount"
            )
        )

        if (
            fixed_price is None
            and discount is None
        ):
            continue

        priority = (
            safe_float(
                rule.get(
                    "priority"
                )
            )
            or 0
        )

        candidates.append(
            (
                priority,
                rule
            )
        )

    if not candidates:
        return None

    candidates.sort(
        key=lambda item: item[0],
        reverse=True
    )

    return candidates[0][1]


# ============================================================
# PRECIO NORMAL
# ============================================================

def get_normal_price(
    product,
    client_type
):
    """
    Determina precio mayorista/minorista.

    Fallback:

        Mayorista sin precio
            -> Minorista

        Especial sin precio
            -> Minorista
    """

    client_type = str(
        client_type
        or "Minorista"
    ).strip().casefold()

    retail = safe_float(
        product.get(
            "precio_minorista"
        )
    )

    wholesale = safe_float(
        product.get(
            "precio_mayorista"
        )
    )

    special = safe_float(
        product.get(
            "precio_especial"
        )
    )

    if client_type == "mayorista":

        if wholesale is not None:
            return wholesale

        return retail

    if client_type == "especial":

        if special is not None:
            return special

        return retail

    return retail


# ============================================================
# MOTOR DE PRECIO
# ============================================================

def calculate_price(
    workbook,
    client_name,
    product_name,
    operation_date,
    kg
):
    """
    Determina exclusivamente el precio comercial.

    Prioridad:

        Oferta
        Especial cliente
        Mayorista
        Minorista
    """

    # --------------------------------------------------------
    # CLIENTE
    # --------------------------------------------------------

    client = find_client(
        workbook,
        client_name
    )

    if client is None:

        return {
            "estado":
                "CLIENTE NO ENCONTRADO"
        }

    if str(
        client.get(
            "estado",
            ""
        )
    ).strip().casefold() != "activo":

        return {
            "estado":
                "CLIENTE INACTIVO"
        }

    # --------------------------------------------------------
    # PRODUCTO
    # --------------------------------------------------------

    product = find_product(
        workbook,
        product_name
    )

    if product is None:

        return {
            "estado":
                "PRODUCTO NO ENCONTRADO"
        }

    if str(
        product.get(
            "estado",
            ""
        )
    ).strip().casefold() != "activo":

        return {
            "estado":
                "PRODUCTO INACTIVO"
        }

    # --------------------------------------------------------
    # KG
    # --------------------------------------------------------

    quantity = safe_float(
        kg
    )

    if (
        quantity is None
        or quantity <= 0
    ):

        return {
            "estado":
                "KG INVALIDO"
        }

    # --------------------------------------------------------
    # FECHA
    # --------------------------------------------------------

    if operation_date is None:

        operation_date = datetime.now()

    # --------------------------------------------------------
    # IDs
    # --------------------------------------------------------

    client_id = client[
        "id"
    ]

    product_id = product[
        "id"
    ]

    # ========================================================
    # 1. OFERTA
    # ========================================================

    offer = get_best_offer(

        workbook,

        product_id,

        operation_date,

        quantity
    )

    if offer is not None:

        fixed_price = safe_float(
            offer.get(
                "price"
            )
        )

        if fixed_price is not None:

            return {

                "estado":
                    "OK",

                "cliente_id":
                    client_id,

                "cliente":
                    client["nombre"],

                "tipo_cliente":
                    client["tipo_cliente"],

                "producto_id":
                    product_id,

                "producto":
                    product["nombre"],

                "kg":
                    quantity,

                "precio_kg":
                    fixed_price,

                "tipo_aplicado":
                    "Oferta",

                "motivo":
                    offer.get(
                        "name"
                    )
                    or
                    "Oferta",

                "oferta_id":
                    offer.get(
                        "id"
                    ),

                "oferta_nombre":
                    offer.get(
                        "name"
                    ),

                "precio_especial_id":
                    None,

                "costo_unitario":
                    safe_float(
                        product.get(
                            "costo"
                        )
                    ),
            }

        discount = safe_float(
            offer.get(
                "discount"
            )
        )

        base_price = get_normal_price(
            product,
            client["tipo_cliente"]
        )

        if (
            discount is not None
            and
            base_price is not None
        ):

            final_price = (
                base_price
                *
                (
                    1
                    -
                    discount / 100
                )
            )

            return {

                "estado":
                    "OK",

                "cliente_id":
                    client_id,

                "cliente":
                    client["nombre"],

                "tipo_cliente":
                    client["tipo_cliente"],

                "producto_id":
                    product_id,

                "producto":
                    product["nombre"],

                "kg":
                    quantity,

                "precio_kg":
                    final_price,

                "tipo_aplicado":
                    "Oferta",

                "motivo":
                    offer.get(
                        "name"
                    )
                    or
                    "Oferta con descuento",

                "oferta_id":
                    offer.get(
                        "id"
                    ),

                "oferta_nombre":
                    offer.get(
                        "name"
                    ),

                "precio_especial_id":
                    None,

                "costo_unitario":
                    safe_float(
                        product.get(
                            "costo"
                        )
                    ),
            }

    # ========================================================
    # 2. ESPECIAL CLIENTE
    # ========================================================

    special = get_best_special_price(

        workbook,

        client_id,

        product_id,

        operation_date
    )

    if special is not None:

        fixed_price = safe_float(
            special.get(
                "price"
            )
        )

        if fixed_price is not None:

            return {

                "estado":
                    "OK",

                "cliente_id":
                    client_id,

                "cliente":
                    client["nombre"],

                "tipo_cliente":
                    client["tipo_cliente"],

                "producto_id":
                    product_id,

                "producto":
                    product["nombre"],

                "kg":
                    quantity,

                "precio_kg":
                    fixed_price,

                "tipo_aplicado":
                    "Especial",

                "motivo":
                    special.get(
                        "reason"
                    )
                    or
                    "Precio especial cliente",

                "oferta_id":
                    None,

                "oferta_nombre":
                    None,

                "precio_especial_id":
                    special.get(
                        "id"
                    ),

                "costo_unitario":
                    safe_float(
                        product.get(
                            "costo"
                        )
                    ),
            }

        discount = safe_float(
            special.get(
                "discount"
            )
        )

        base_price = get_normal_price(
            product,
            client["tipo_cliente"]
        )

        if (
            discount is not None
            and
            base_price is not None
        ):

            final_price = (
                base_price
                *
                (
                    1
                    -
                    discount / 100
                )
            )

            return {

                "estado":
                    "OK",

                "cliente_id":
                    client_id,

                "cliente":
                    client["nombre"],

                "tipo_cliente":
                    client["tipo_cliente"],

                "producto_id":
                    product_id,

                "producto":
                    product["nombre"],

                "kg":
                    quantity,

                "precio_kg":
                    final_price,

                "tipo_aplicado":
                    "Especial",

                "motivo":
                    special.get(
                        "reason"
                    )
                    or
                    "Descuento especial cliente",

                "oferta_id":
                    None,

                "oferta_nombre":
                    None,

                "precio_especial_id":
                    special.get(
                        "id"
                    ),

                "costo_unitario":
                    safe_float(
                        product.get(
                            "costo"
                        )
                    ),
            }

    # ========================================================
    # 3. MAYORISTA / 4. MINORISTA
    # ========================================================

    normal_price = get_normal_price(
        product,
        client["tipo_cliente"]
    )

    if normal_price is None:

        return {
            "estado":
                "PRECIO NO DISPONIBLE"
        }

    client_type = str(
        client["tipo_cliente"]
        or "Minorista"
    ).strip().casefold()

    applied_type = (
        "Mayorista"
        if client_type == "mayorista"
        else "Minorista"
    )

    return {

        "estado":
            "OK",

        "cliente_id":
            client_id,

        "cliente":
            client["nombre"],

        "tipo_cliente":
            client["tipo_cliente"],

        "producto_id":
            product_id,

        "producto":
            product["nombre"],

        "kg":
            quantity,

        "precio_kg":
            normal_price,

        "tipo_aplicado":
            applied_type,

        "motivo":
            "Precio base",

        "oferta_id":
            None,

        "oferta_nombre":
            None,

        "precio_especial_id":
            None,

        "costo_unitario":
            safe_float(
                product.get(
                    "costo"
                )
            ),
    }


# ============================================================
# VENTA COMPLETA
# ============================================================

def calculate_sale(
    workbook,
    client_name,
    product_name,
    operation_date,
    kg
):
    """
    Calcula precio + importe + costo + ganancia + margen.
    """

    result = calculate_price(

        workbook,

        client_name,

        product_name,

        operation_date,

        kg
    )

    if result.get(
        "estado"
    ) != "OK":

        return result

    quantity = float(
        result["kg"]
    )

    price = float(
        result["precio_kg"]
    )

    importe = (
        quantity
        * price
    )

    cost = result.get(
        "costo_unitario"
    )

    if cost is not None:

        costo_total = (
            quantity
            * cost
        )

        ganancia = (
            importe
            - costo_total
        )

        ganancia_pct = (
            ganancia / importe
            if importe
            else None
        )

    else:

        costo_total = None
        ganancia = None
        ganancia_pct = None

    return {

        **result,

        "importe":
            importe,

        "costo_total":
            costo_total,

        "ganancia":
            ganancia,

        "ganancia_pct":
            ganancia_pct,
    }


# ============================================================
# COMPATIBILIDAD
# ============================================================

def determinar_precio_venta(
    workbook,
    nombre_cliente,
    nombre_producto,
    fecha_venta,
    kg
):
    """
    Alias de compatibilidad con código anterior.

    Devuelve el formato esperado por ventas_excel.py.
    """

    return calculate_price(

        workbook,

        nombre_cliente,

        nombre_producto,

        fecha_venta,

        kg
    )