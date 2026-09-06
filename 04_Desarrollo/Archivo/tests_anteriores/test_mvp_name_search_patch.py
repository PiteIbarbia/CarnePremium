# Añadir estas comprobaciones al test_mvp.py

sales = workbook["05_VENTAS"]

required_sales_headers = [
    "VentaID", "Fecha", "Hora", "Cliente", "TipoPrecio",
    "Producto", "Kg", "PrecioKg", "Importe", "Ganancia$",
    "Ganancia%", "MedioPago", "Vendedor", "Observaciones",
    "_ClienteID", "_ProductoID"
]

actual_sales_headers = [
    cell.value for cell in sales[1]
]

missing_sales_headers = [
    h for h in required_sales_headers
    if h not in actual_sales_headers
]

if missing_sales_headers:
    fail(
        f"05_VENTAS: faltan encabezados "
        f"{missing_sales_headers}"
    )

for cell in ["A2","B2","C2","O2","P2","H2","I2","J2","K2"]:
    value = sales[cell].value
    if not isinstance(value, str) or not value.startswith("="):
        fail(
            f"05_VENTAS!{cell} no contiene fórmula válida."
        )

print("Ventas por nombre: OK")
print("IDs internos técnicos: OK")
