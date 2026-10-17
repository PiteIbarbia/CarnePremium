# Operación diaria: ventas, cobros y Caja

## Cargar ventas

En `05_VENTAS`, elegí cliente en la columna D y producto en F desde los catálogos `03_CLIENTES` y `04_PRODUCTOS`. Completá fecha, tipo de precio, kilos y los datos de la operación. No edites `TipoRegistro` (AA), `_FirmaEntrada` (AC) ni otras columnas técnicas ocultas.

## Registrar un cobro

- `PagoRecibido` (W) contiene solo el dinero efectivamente recibido.
- `MedioPago` (L) identifica cómo ingresó ese dinero.
- `FechaCobro` (AB) permite imputar un cobro en su fecha real cuando sea posterior a la venta.
- Una venta a cuenta corriente sin dinero recibido no es un ingreso de Caja.
- Una fila separada de tipo `COBRO` se utiliza para cobrar saldos anteriores; debe dejar producto y kilos vacíos para no generar otra línea de carne en el ticket.

## Entender 09_CTACTE — quién debe y qué falta cobrar

La hoja `09_CTACTE` tiene tres zonas:

1. **Resumen superior:** muestra cuánto queda por cobrar a clientes, cuánto hay a favor de clientes y cuántos están debiendo, a favor o al día.
2. **Resumen por cliente:** consultá el nombre del cliente y la columna `Saldo actual`. Un saldo positivo significa que el cliente todavía debe dinero; un saldo negativo significa que tiene dinero a favor; cero significa que está al día.
3. **Movimientos por día:** muestra las ventas y los cobros registrados y el saldo acumulado que se arrastra entre fechas.

La cuenta se obtiene de las ventas históricas y de los importes realmente cobrados. Un cobro en una fecha posterior reduce el saldo en su fecha de cobro; no borra ni cambia la venta original. No edites manualmente esta hoja: el actualizador la vuelve a generar desde `05_VENTAS`.

**Importante:** esta vista responde quién debe dinero a la carnicería. Las deudas que la carnicería tenga con proveedores aún no se calculan porque falta integrar Compras/Proveedores y sus pagos; no se mostrarán importes inventados.

## Leer la hoja 07_CAJA

Los primeros diez encabezados se mantienen por compatibilidad. Se agregan dos columnas de auditoría: `Importe original` y `Observación / ajuste`.

A la derecha de la tabla se muestra un resumen automático de ingresos, egresos, saldo neto registrado y cantidad de ajustes pendientes. Ese saldo no incluye el efectivo existente al inicio del período; si no se registra un saldo inicial, el saldo neto puede diferir del dinero físico disponible.

- **Ingreso (G):** dinero realmente recibido.
- **Egreso (H):** dinero realmente pagado o entregado.
- **Saldo (I):** acumulado cronológico de ingresos menos egresos registrados. No incluye saldo inicial de caja.
- **Ajuste pendiente:** movimiento automático heredado, dudoso o duplicado que requiere conciliación. El importe original se conserva como referencia y el ajuste no afecta el saldo hasta confirmarse.
- **Origen (J):** indica si el movimiento procede de la sincronización automática o de otra fuente.

Las anulaciones automáticas antiguas que anteriormente podían aparecer como egresos se migran a `AJUSTE PENDIENTE`. No se crean salidas de dinero ficticias para neutralizar registros. Los duplicados quedan conservados para auditoría, pero no se cuentan como un segundo ingreso.

La hoja ordena los movimientos por fecha y vuelve a calcular el saldo. Si una fila está resaltada como ajuste pendiente, revisá la referencia contra `05_VENTAS` y los respaldos antes de decidir cómo conciliarla. No borres filas ni cambies manualmente el saldo para hacer desaparecer la diferencia.

## Actualizar y validar de forma segura

El actualizador construye el resultado en un directorio temporal y corre todas las pruebas contra esa copia. El modo predeterminado solo valida: no reemplaza el libro de entrada.

1. Para validar el maestro sin aplicarle cambios, cerrá Excel y ejecutá `python 04_Desarrollo/Scripts/actualizar_erp.py`.
2. Para validar el archivo de prueba sin reemplazarlo, ejecutá `python 04_Desarrollo/Scripts/actualizar_erp.py --archivo "01_ERP/CarnePremiumERP_prueba_036.xlsx"`.
3. Para aplicar el resultado probado al maestro, cerrá Excel y ejecutá `python 04_Desarrollo/Scripts/actualizar_erp.py --aplicar`, o usá `ACTUALIZAR_ERP.bat`.
4. El modo de aplicación crea y verifica un respaldo antes del reemplazo. Si falla una prueba, se cancela la aplicación.
5. Revisá ventas, tickets, Caja y Cuenta Corriente en Excel Professional Plus 2019 antes de hacer una carga masiva.

No ejecutes simultáneamente el monitor y el actualizador manual. El monitor utiliza la misma validación temporal. Esta protección tampoco sustituye una prueba visual real en Excel 2019.

## Próxima etapa

Una vez comprobada Caja en Excel 2019, se avanzará con compras, entradas de mercadería, mermas y existencias iniciales para construir el stock real.
