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


## Build 0.4.0 — Ventas, ticket automático, cobros, Caja y Cuenta Corriente

El libro maestro sigue siendo `01_ERP/CarnePremiumERP.xlsx`. La lógica vive en módulos Python bajo `04_Desarrollo/Scripts`. No edites manualmente las fórmulas generadas.

1. Cerrá Excel.
2. Ejecutá el constructor desde la raíz del repositorio:
   `./.venv/Scripts/python.exe ./04_Desarrollo/Scripts/builder.py`
3. Ejecutá las pruebas:
   `./.venv/Scripts/python.exe ./04_Desarrollo/Pruebas/test_mvp.py`
   `./.venv/Scripts/python.exe ./04_Desarrollo/Pruebas/test_ticket_caja.py`
4. Abrí el libro y revisá `05_VENTAS`, `06_TICKET`, `07_CAJA` y `09_CTACTE`.

### Captura única desde 05_VENTAS

- Una fila con tipo vacío o `VENTA` representa un corte vendido: cliente, tipo de precio, producto, kilos y medio de pago.
- `PagoRecibido` (columna W) es el dinero efectivamente recibido por esa línea. No ingreses ahí el total de la venta, salvo que ese importe se haya cobrado realmente. Un cobro parcial deja un saldo pendiente; si se cobra más que el total, el saldo queda a favor.
- El ticket se agrupa automáticamente por cliente y fecha de calendario. No se carga `TicketID` manualmente. Al seleccionar un ticket en `06_TICKET` se muestran sus cortes, kilos e importe total.
- Para cobrar una deuda anterior en una fecha posterior, seleccioná `COBRO` en `TipoRegistro` (columna AA), completá cliente, fecha si corresponde, medio de pago y `PagoRecibido`; dejá producto y kilos vacíos. Esa fila no genera una línea de carne en el ticket, pero se refleja en Caja y reduce el saldo acumulado en Cuenta Corriente.
- `07_CAJA` registra solo cobros reales por medios cobrables; una venta a Cuenta Corriente sin dinero recibido no se contabiliza como ingreso de caja.
- `09_CTACTE` es una vista automática por cliente y fecha. El saldo acumulado positivo significa deuda pendiente; el negativo representa saldo a favor del cliente. No cargues manualmente en esta hoja: se reconstruye desde Ventas.

### Reglas de seguridad

Las columnas técnicas V y X están ocultas; W contiene el importe cobrado, Y el estado de cobro y Z el saldo del ticket. No elimines filas ni cambies encabezados. Ejecutá el constructor solo con Excel cerrado. Al migrar desde Build 0.3.6, los importes de cobro preexistentes en Caja se usan para poblar `PagoRecibido` cuando se identifican por referencia de venta; verificá ese campo en las ventas antiguas antes de continuar con datos reales.

Build 0.4.0 sigue en la rama de desarrollo y requiere validación en Excel 2019 antes de incorporarse a `main`. La hoja de Stock todavía necesita definir cómo se registrarán entradas, compras y existencias iniciales antes de poder calcular stock real de manera fiable.
