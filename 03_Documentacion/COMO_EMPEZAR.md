# Cómo empezar

En PowerShell:

```powershell
cd "C:\Users\Pite\Documents\Trabajo\Carne Premium-Vacío\CarnePremiumERP"
py -3.13 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\pip.exe install -r .\04_Desarrollo\requirements.txt
.\.venv\Scripts\python.exe .\04_Desarrollo\Scripts\builder.py
.\.venv\Scripts\python.exe .\04_Desarrollo\Pruebas\test_mvp.py
```

Después abrí:
`01_ERP\CarnePremiumERP.xlsx`

No borres ni renombres las hojas del proyecto.


## Build 0.3.6 — Ventas, Ticket y Caja

El libro maestro sigue siendo `01_ERP/CarnePremiumERP.xlsx`. La lógica de negocio vive en módulos Python bajo `04_Desarrollo/Scripts`. No edites manualmente las fórmulas generadas.

1. Cerrá Excel.
2. Ejecutá el constructor desde PowerShell en la raíz del repositorio:
   `python .\04_Desarrollo\Scripts\builder.py`
3. Revisá la salida del constructor. Si termina con error, no continúes con datos reales.
4. Ejecutá las pruebas:
   `python .\04_Desarrollo\Pruebas\test_mvp.py`
   `python .\04_Desarrollo\Pruebas\test_ticket_caja.py`
5. Abrí el libro maestro y comprobá las hojas `05_VENTAS`, `06_TICKET` y `07_CAJA`.

Las columnas técnicas V y X de Ventas están ocultas. W (`TicketID`) queda visible: dejala vacía para generar un ticket por venta, o repetí el mismo identificador en varias filas para agrupar varios cortes en un ticket. El constructor actualiza el detalle del ticket y sincroniza cobros aprobados con Caja sin duplicar movimientos automáticos.

La cuenta corriente no se trata como dinero cobrado en Caja. Su integración con `09_CTACTE` queda como trabajo pendiente.

Build 0.3.6 no añade dependencias nuevas. La rama de desarrollo debe revisarse y probarse antes de incorporarse a `main`.
