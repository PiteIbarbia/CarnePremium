# Operación diaria: ventas, cobros y Caja

## Cargar ventas

En `05_VENTAS`, elegí cliente en la columna D y producto en F desde los catálogos `03_CLIENTES` y `04_PRODUCTOS`. Completá fecha, tipo de precio, kilos y los datos de la operación. No edites `TipoRegistro` (AA), `_FirmaEntrada` (AC) ni otras columnas técnicas ocultas.

## Registrar un cobro

- `PagoRecibido` (W) contiene solo el dinero efectivamente recibido.
- `MedioPago` (L) identifica cómo ingresó ese dinero.
- `FechaCobro` (AB) permite imputar un cobro en su fecha real cuando sea posterior a la venta.
- Una venta a cuenta corriente sin dinero recibido no es un ingreso de Caja.
- Una fila separada de tipo `COBRO` se utiliza para cobrar saldos anteriores; debe dejar producto y kilos vacíos para no generar otra línea de carne en el ticket.

## Leer la hoja 07_CAJA

Los primeros diez encabezados se mantienen por compatibilidad. Se agregan dos columnas de auditoría: `Importe original` y `Observación / ajuste`.

- **Ingreso (G):** dinero realmente recibido.
- **Egreso (H):** dinero realmente pagado o entregado.
- **Saldo (I):** acumulado cronológico de ingresos menos egresos registrados. No incluye saldo inicial de caja.
- **Ajuste pendiente:** movimiento automático heredado, dudoso o duplicado que requiere conciliación. El importe original se conserva como referencia y el ajuste no afecta el saldo hasta confirmarse.
- **Origen (J):** indica si el movimiento procede de la sincronización automática o de otra fuente.

Las anulaciones automáticas antiguas que anteriormente podían aparecer como egresos se migran a `AJUSTE PENDIENTE`. No se crean salidas de dinero ficticias para neutralizar registros. Los duplicados quedan conservados para auditoría, pero no se cuentan como un segundo ingreso.

La hoja ordena los movimientos por fecha y vuelve a calcular el saldo. Si una fila está resaltada como ajuste pendiente, revisá la referencia contra `05_VENTAS` y los respaldos antes de decidir cómo conciliarla. No borres filas ni cambies manualmente el saldo para hacer desaparecer la diferencia.

## Actualizar y validar

1. Guardá y cerrá el libro maestro `01_ERP/CarnePremiumERP.xlsx`.
2. Ejecutá `ACTUALIZAR_ERP.bat` una sola vez.
3. Confirmá que finalicen correctamente las pruebas del Build y las pruebas de Ticket/Caja/Cuenta Corriente.
4. Revisá ventas, tickets, Caja y Cuenta Corriente antes de una carga masiva.

No ejecutes simultáneamente el monitor y el actualizador manual. Google Sheets no ejecuta los archivos BAT ni el generador local; esa automatización necesita una implementación independiente.

## Próxima etapa

Una vez comprobada Caja en Excel 2019, se avanzará con compras, entradas de mercadería, mermas y existencias iniciales para construir el stock real.
