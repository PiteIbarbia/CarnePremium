# Carne Premium ERP

Proyecto local del Carne Premium ERP.

Plataforma principal: Excel Professional Plus 2019.
Compatibilidad objetivo: Google Sheets cuando sea técnicamente posible.

Regla: existe un único archivo maestro en `01_ERP/CarnePremiumERP.xlsx`.
Las modificaciones estructurales se realizan mediante scripts en `04_Desarrollo/Scripts`.


## Flujo de venta, ticket y cobro — Build 0.5.2

La venta sigue siendo el registro principal. **No se cambia a `COBRO` cuando el cliente paga**; `TipoRegistro` (AA) es técnico y queda oculto.

1. En `05_VENTAS`, cargá cliente, tipo de precio, corte y kilos. La venta calcula precio e importe y queda disponible para generar/consultar su ticket aunque no se haya cobrado.
2. Al recibir dinero, completá `MedioPago` (L) e `Importe cobrado` / `PagoRecibido` (W). Podés dejarlo en blanco mientras la venta esté pendiente. `EstadoCobro` (Y) y `SaldoTicket` (Z) se recalculan a partir del importe total del ticket y los cobros registrados.
3. Si el dinero se recibe en una fecha posterior, completá `FechaCobro` (AB) con la fecha real. Si queda vacía, los procesos usan la fecha original de venta para conservar compatibilidad con los históricos previos.
4. El ticket permanece en `06_TICKET` después del cobro. `07_CAJA` registra el dinero efectivamente recibido y `09_CTACTE` conserva la venta en su fecha y refleja el cobro en la fecha real cuando se indicó.
5. Al actualizar el ERP, se conserva el precio histórico de las ventas aprobadas si no cambiaron fecha, cliente, tipo de precio, producto ni kilos. Una huella técnica oculta (`_FirmaEntrada`, AC) permite detectar cambios reales en los datos del pedido.

La prueba automática del Build 0.5.2 cubre el ciclo pedido → ticket → cobro posterior y verifica que cobrar no elimina ni transforma la venta.
