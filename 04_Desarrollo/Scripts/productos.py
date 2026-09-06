def obtener_producto_por_nombre(
    workbook,
    nombre_producto
):

    hoja = workbook[
        "04_PRODUCTOS"
    ]

    if not nombre_producto:
        return None

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

                "id":
                    hoja.cell(
                        fila,
                        1
                    ).value,

                "nombre":
                    nombre,

                "categoria":
                    hoja.cell(
                        fila,
                        3
                    ).value,

                "costo":
                    hoja.cell(
                        fila,
                        4
                    ).value,

                "precio_minorista":
                    hoja.cell(
                        fila,
                        5
                    ).value,

                "precio_mayorista":
                    hoja.cell(
                        fila,
                        6
                    ).value,

                "precio_especial":
                    hoja.cell(
                        fila,
                        7
                    ).value,

                "estado":
                    hoja.cell(
                        fila,
                        8
                    ).value,

            }

    return None