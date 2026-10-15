# Cómo empezar — Carne Premium ERP

## Entorno y archivo maestro

- Windows, Python 3.13, Git y VS Code.
- Archivo maestro único: `01_ERP/CarnePremiumERP.xlsx`.
- Plataforma operativa: Excel Professional Plus 2019.
- La lógica se mantiene en `04_Desarrollo/Scripts`; no edites manualmente las fórmulas generadas ni las columnas técnicas ocultas.

## Actualizar el ERP

Desde la raíz del repositorio, en PowerShell:

```powershell
.\.venv\Scripts\python.exe .\04_Desarrollo\Scripts\actualizar_erp.py
```

También se puede ejecutar `ACTUALIZAR_ERP.bat` con doble clic. Antes, guardá los cambios y cerrá el libro maestro en Excel. El proceso construye el libro, ejecuta las pruebas de integridad y las pruebas de Ticket/Caja/Cuenta Corriente. Si una prueba falla, no se debe continuar usando ese resultado como versión validada.

## Caja — Build 0.5.3

En `07_CAJA`:

- `Ingreso` (G): dinero realmente recibido.
- `Egreso` (H): dinero realmente entregado o pagado.
- `Saldo` (I): saldo acumulado cronológico de ingresos menos egresos; no incluye un saldo inicial de efectivo.
- `Importe original` y `Observación / ajuste`:  conservan la referencia y el motivo de movimientos automáticos que necesitan conciliación.
- Las filas `AJUSTE PENDIENTE` se revisan contra `05_VENTAS` y los respaldos; no representan por sí solas una entrada o salida de dinero.

Los primeros diez encabezados de Caja se conservan por compatibilidad con el resto del sistema. Los movimientos no se borran para arreglar diferencias.

## Carga operativa

1. Ingresá ventas en `05_VENTAS` usando los catálogos de clientes y productos.
2. Registrá solo el importe cobrado realmente en `PagoRecibido` (W); dejalo vacío si todavía no se recibió dinero.
3. Si el cobro es posterior a la venta, indicá `FechaCobro` (AB).
4. Revisá el ticket en `06_TICKET`, los movimientos y ajustes de `07_CAJA`,  y el saldo de `09_CTACTE`.
5. Antes de cargar un lote histórico grande, mantené una copia de respaldo y revisá cualquier ajuste pendiente.

## Compatibilidad con Google Sheets

El libro puede importarse para revisar y operar datos, pero el generador local Python y los archivos BAT no se ejecutan desde Google Sheets. La automatización en la nube será una etapa independiente y no debe darse por activa hasta tenerla implementada y probada.

## Siguiente etapa del proyecto

Después de validar Caja en Excel 2019, el siguiente módulo funcional será compras, entradas, mermas y existencias iniciales para que la hoja `08_STOCK` refleje inventario real.
