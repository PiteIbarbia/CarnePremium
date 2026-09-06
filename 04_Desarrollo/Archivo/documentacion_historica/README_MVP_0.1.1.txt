MVP 0.1.1 — CLIENTES + PRODUCTOS + VENTAS

Copiar los dos archivos del ZIP reemplazando:
- 04_Desarrollo/Scripts/builder.py
- 04_Desarrollo/Pruebas/test_mvp.py

Luego ejecutar en Git Bash desde la raíz del proyecto:

python 04_Desarrollo/Scripts/builder.py
python 04_Desarrollo/Pruebas/test_mvp.py

Abrir 01_ERP/CarnePremiumERP.xlsx en Excel 2019.

Prueba manual:
1) 03_CLIENTES: cargar un cliente.
2) 04_PRODUCTOS: cargar un producto con costo y precios.
3) 05_VENTAS: seleccionar ClienteID, TipoPrecio y ProductoID; ingresar Kg.
4) Comprobar PrecioKg, Importe, Ganancia$ y Ganancia%.

No cargar todavía ofertas ni precios especiales por cliente: eso entra después del núcleo.
