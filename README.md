# Carne Premium ERP

Proyecto local de gestión de carnicería premium.

- Plataforma operativa principal: Microsoft Excel Professional Plus 2019.
- Compatibilidad objetivo: Google Sheets, cuando sea técnicamente viable.
- Código, reglas y automatización: Python modular en `04_Desarrollo/Scripts`.
- Control de versiones: Git y GitHub.
- Archivo maestro único: `01_ERP/CarnePremiumERP.xlsx`.

## Build 0.5.3 — Caja clara y ajustes auditables

La venta sigue siendo el registro principal. Cobrar no transforma una venta en un registro `COBRO`; `TipoRegistro` (AA) es técnico y permanece oculto.

### Uso de Caja

- Las primeras diez columnas se mantienen para no romper las integraciones existentes.
- Se agregan `Importe original` y `Observación / ajuste` para conservar una referencia de los movimientos automáticos que requieren revisión.
- `Ingreso` y `Egreso` solo deben representar movimientos reales de dinero.
- Las anulaciones automáticas heredadas y los duplicados se conservan como `AJUSTE PENDIENTE`, con importe original y explicación. No se convierten en egresos ficticios ni se eliminan silenciosamente.
- Los movimientos se ordenan cronológicamente antes de recalcular `Saldo`.
- `Saldo` es un acumulado de los movimientos registrados (ingresos menos egresos). No incluye un saldo inicial de efectivo.
- Los ajustes pendientes se resaltan para conciliarlos contra `05_VENTAS` y las copias de respaldo. No se debe marcar como egreso una corrección técnica.

### Flujo de venta, ticket y cobro

1. En `05_VENTAS`, cargar cliente, tipo de precio, corte y kilos.
2. Al recibir dinero, completar `MedioPago` (L) y `PagoRecibido` (W) con el importe realmente recibido.
3. Si el cobro ocurre en otra fecha, indicar `FechaCobro` (AB).
4. Verificar el ticket en `06_TICKET`; caja y cuenta corriente se regeneran desde las ventas y cobros registrados.
5. No editar las columnas técnicas ocultas ni borrar movimientos de `07_CAJA` para corregir una diferencia. Usar las observaciones y revisar la fuente del movimiento.

La validación automatizada cubre el flujo venta → ticket → cobro, la conservación del histórico, la neutralización auditable de anulaciones automáticas y el control de movimientos duplicados.

## Operación segura

1. Cerrar el archivo maestro en Excel antes de ejecutar el proceso de actualización.
2. Ejecutar `ACTUALIZAR_ERP.bat`.
3. Confirmar que terminen correctamente el Build, `test_mvp.py` y `test_ticket_caja.py`.
4. Revisar `05_VENTAS`, `06_TICKET`, `07_CAJA` y `09_CTACTE` antes de seguir con cargas masivas.

El monitor local y el proceso manual no deben ejecutarse al mismo tiempo. Google Sheets no ejecuta el generador local ni los archivos BAT; su automatización requerirá una integración propia.

## Próximas etapas del plan

Tras estabilizar ventas, ticket, cobros y Caja, continuar con compras/stock real, precios históricos, dashboard profesional, recuperación ante desastres e integraciones futuras. No se considera listo el stock real hasta definir entradas, compras, mermas y existencias iniciales.
