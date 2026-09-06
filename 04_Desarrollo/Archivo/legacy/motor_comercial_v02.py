from __future__ import annotations

from datetime import date, datetime


# ============================================================
# FECHAS
# ============================================================

def to_date(value):
    """
    Convierte datetime/date/string a date.
    Las vigencias comerciales se manejan por día.
    """

    if value in (None, ""):
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    if isinstance(value, str):

        for fmt in (
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%d-%m-%Y",
        ):

            try:
                return datetime.strptime(
                    value.strip(),
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
    Vigencia inclusiva.
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

    if start is not None and current < start:
        return False

    if end is not None and current > end:
        return False

    return True


# ============================================================
# VALORES
# ============================================================

def safe_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
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

        if (
            str(name)
            .strip()
            .casefold()
            == target
        ):

            return {
                "id": ws.cell(
                    row,
                    1
                ).value,

                "nombre": name,

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

        if (
            str(name)
            .strip()
            .casefold()
            == target
        ):

            return {
                "id": ws.cell(
                    row,
                    1
                ).value,

                "nombre": name,

                "costo": ws.cell(
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

def get_offer_candidates(
    workbook,
    product_id
):
    """
    Obtiene todas las ofertas del producto.
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

        row_product_id = ws.cell(
            row,
            2
        ).value

        if row_product_id != product_id:
            continue

        results.append({

            "id":
                ws.cell(row, 1).value,

            "product_id":
                row_product_id,

            "name":
                ws.cell(row, 4).value,

            "price":
                ws.cell(row, 5).value,

            "discount":
                ws.cell(row, 6).value,

            "start":
                ws.cell(row, 7).value,

            "end":
                ws.cell(row, 8).value,

            "min_kg":
                ws.cell(row, 9).value,

            "status":
                ws.cell(row, 10).value,

            "priority":
                ws.cell(row, 11).value,
        })

    return results


def get_best_offer(
    workbook,
    product_id,
    operation_date,
    kg
):
    """
    Devuelve la oferta válida de mayor prioridad.
    """

    quantity = safe_float(
        kg
    )

    if quantity is None:
        return None

    candidates = []

    for offer in get_offer_candidates(
        workbook,
        product_id
    ):

        if str(
            offer.get(
                "status",
                ""
            )
        ).strip().casefold() != "activo":

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
            and
            discount is None
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

def get_special_candidates(
    workbook,
    client_id,
    product_id
):
    """
    Obtiene todas las reglas especiales del cliente/producto.
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

        row_client_id = ws.cell(
            row,
            2
        ).value

        row_product_id = ws.cell(
            row,
            4
        ).value

        if (
            row_client_id != client_id
            or
            row_product_id != product_id
        ):
            continue

        results.append({

            "id":
                ws.cell(row, 1).value,

            "client_id":
                row_client_id,

            "product_id":
                row_product_id,

            "type":
                ws.cell(row, 6).value,

            "price":
                ws.cell(row, 7).value,

            "discount":
                ws.cell(row, 8).value,

            "start":
                ws.cell(row, 9).value,

            "end":
                ws.cell(row, 10).value,

            "status":
                ws.cell(row, 11).value,

            "priority":
                ws.cell(row, 12).value,

            "reason":
                ws.cell(row, 13).value,
        })

    return results


def get_best_special_price(
    workbook,
    client_id,
    product_id,
    operation_date
):
    """
    Devuelve la regla especial vigente de mayor prioridad.
    """

    candidates = []

    for rule in get_special_candidates(
        workbook,
        client_id,
        product_id
    ):

        if str(
            rule.get(
                "status",
                ""
            )
        ).strip().casefold() != "activo":

            continue

        if str(
            rule.get(
                "type",
                ""
            )
        ).strip().casefold() != "especial":

            continue

        if not date_is_valid(
            operation_date,
            rule.get("start"),
            rule.get("end")
        ):

            continue

        if (
            safe_float(rule.get("price"))
            is None
            and
            safe_float(rule.get("discount"))
            is None
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
    Mayorista y Especial tienen fallback a Minorista.
    """

    client_type = str(
        client_type or "Minorista"
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

        return (
            wholesale
            if wholesale is not None
            else retail
        )

    if client_type == "especial":

        return (
            special
            if special is not None
            else retail
        )

    return retail


# ============================================================
# MOTOR CENTRAL
# ============================================================

def calculate_price(
    workbook,
    client_name,
    product_name,
    operation_date,
    kg
):
    """
    Motor Comercial v0.2.

    PRIORIDAD:

        Oferta
        Especial cliente
        Mayorista
        Minorista
    """

    client = find_client(
        workbook,
        client_name
    )

    if client is None:
        return {
            "status":
                "CLIENTE NO ENCONTRADO"
        }

    if str(
        client.get(
            "estado",
            ""
        )
    ).strip().casefold() != "activo":

        return {
            "status":
                "CLIENTE INACTIVO"
        }

    product = find_product(
        workbook,
        product_name
    )

    if product is None:
        return {
            "status":
                "PRODUCTO NO ENCONTRADO"
        }

    if str(
        product.get(
            "estado",
            ""
        )
    ).strip().casefold() != "activo":

        return {
            "status":
                "PRODUCTO INACTIVO"
        }

    quantity = safe_float(
        kg
    )

    if (
        quantity is None
        or
        quantity <= 0
    ):
        return {
            "status":
                "KG INVALIDO"
        }

    # --------------------------------------------------------
    # 1. OFERTA
    # --------------------------------------------------------

    offer = get_best_offer(
        workbook,
        product["id"],
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
                "status":
                    "OK",

                "price":
                    fixed_price,

                "type":
                    "Oferta",

                "reason":
                    offer.get(
                        "name"
                    ),

                "offer_id":
                    offer.get(
                        "id"
                    ),

                "special_price_id":
                    None,

                "client_id":
                    client["id"],

                "product_id":
                    product["id"],

                "cost":
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
                "status":
                    "OK",

                "price":
                    final_price,

                "type":
                    "Oferta",

                "reason":
                    offer.get(
                        "name"
                    ),

                "offer_id":
                    offer.get(
                        "id"
                    ),

                "special_price_id":
                    None,

                "client_id":
                    client["id"],

                "product_id":
                    product["id"],

                "cost":
                    safe_float(
                        product.get(
                            "costo"
                        )
                    ),
            }

    # --------------------------------------------------------
    # 2. ESPECIAL CLIENTE
    # --------------------------------------------------------

    special = get_best_special_price(
        workbook,
        client["id"],
        product["id"],
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
                "status":
                    "OK",

                "price":
                    fixed_price,

                "type":
                    "Especial",

                "reason":
                    special.get(
                        "reason"
                    )
                    or
                    "Precio especial cliente",

                "offer_id":
                    None,

                "special_price_id":
                    special.get(
                        "id"
                    ),

                "client_id":
                    client["id"],

                "product_id":
                    product["id"],

                "cost":
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
                "status":
                    "OK",

                "price":
                    final_price,

                "type":
                    "Especial",

                "reason":
                    special.get(
                        "reason"
                    )
                    or
                    "Descuento especial cliente",

                "offer_id":
                    None,

                "special_price_id":
                    special.get(
                        "id"
                    ),

                "client_id":
                    client["id"],

                "product_id":
                    product["id"],

                "cost":
                    safe_float(
                        product.get(
                            "costo"
                        )
                    ),
            }

    # --------------------------------------------------------
    # 3/4. MAYORISTA / MINORISTA
    # --------------------------------------------------------

    normal_price = get_normal_price(
        product,
        client["tipo_cliente"]
    )

    if normal_price is None:

        return {
            "status":
                "PRECIO NO DISPONIBLE"
        }

    applied_type = (
        "Mayorista"
        if str(
            client["tipo_cliente"]
        ).strip().casefold()
        == "mayorista"
        else "Minorista"
    )

    return {

        "status":
            "OK",

        "price":
            normal_price,

        "type":
            applied_type,

        "reason":
            "Precio base",

        "offer_id":
            None,

        "special_price_id":
            None,

        "client_id":
            client["id"],

        "product_id":
            product["id"],

        "cost":
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
    Calcula precio, importe y rentabilidad.
    """

    price_result = calculate_price(
        workbook,
        client_name,
        product_name,
        operation_date,
        kg
    )

    if price_result.get("status") != "OK":

        return price_result

    quantity = float(
        kg
    )

    price = float(
        price_result["price"]
    )

    importe = (
        quantity
        * price
    )

    cost = price_result.get(
        "cost"
    )

    if cost is not None:

        costo_total = (
            quantity
            * cost
        )

        gain = (
            importe
            - costo_total
        )

        margin = (
            gain
            / importe
            if importe
            else None
        )

    else:

        costo_total = None
        gain = None
        margin = None

    return {
        **price_result,

        "kg":
            quantity,

        "importe":
            importe,

        "costo_total":
            costo_total,

        "ganancia":
            gain,

        "ganancia_pct":
            margin,
    }