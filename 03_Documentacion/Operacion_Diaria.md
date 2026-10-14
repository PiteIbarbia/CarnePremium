# Operación diaria: ventas y actualización del ERP

## Seleccionar clientes y productos
En `05_VENTAS`, las celdas D2:D501 (Cliente) y F2:F501 (Producto) tienen listas desplegables basadas en los catálogos maestros:

- Clientes: `03_CLIENTES`, columna B.
- Productos: `04_PRODUCTOS`, columna B.

Elegí el nombre exacto de la lista. Podés comenzar a escribir el nombre si ya conocés el valor; las validaciones ayudan a evitar errores de escritura. Los registros nuevos deben cargarse primero en su catálogo.

## Actualizar el ERP con un doble clic (Excel en Windows)

1. Guardá las ventas y los cobros en `01_ERP/CarnePremiumERP.xlsx`.
2. Cerrá ese libro en Excel. No hace falta cerrar otros libros de Excel, salvo que estén usando el mismo archivo maestro.
3. En la carpeta raíz del proyecto, hacé doble clic en `ACTUALIZAR_ERP.bat`.
4. El proceso ejecuta el generador, valida el libro y corre las pruebas de Ticket/Caja/Cuenta Corriente.
5. Si todo termina correctamente, se abrirá el libro maestro para revisar `06_TICKET`, `07_CAJA`, `09_CTACTE` y `10_DASHBOARD`.

El generador crea una copia de respaldo antes de guardar el libro. Si alguna validación falla, el proceso avisa y no abre automáticamente el archivo resultante para revisión.

## Actualización automática local (Excel en Windows)

Para que el ERP se regenere automáticamente tras guardar los cambios, iniciá `INICIAR_MONITOR_ERP.bat` una sola vez y dejá abierta su ventana. Cargá ventas/cobros en Excel y guardá. El monitor espera a que se cierre el libro maestro y a que el archivo quede estable; después ejecuta el generador y las pruebas para actualizar tickets, Caja, Cuenta Corriente y reportes. No abre otro Excel automáticamente para no bloquear el siguiente lote. Cuando quieras detener el monitor, presioná Ctrl+C en su ventana.

No ejecutes el monitor y `ACTUALIZAR_ERP.bat` simultáneamente. Si preferís una ejecución manual de un solo paso, cerrá Excel y usá `ACTUALIZAR_ERP.bat`.

## Alcance y compatibilidad

El monitor automatiza el proceso local de Excel; no ejecuta Python dentro del libro ni puede regenerar informes mientras el archivo maestro siga abierto. El archivo debe guardarse y cerrarse antes de regenerarlo para evitar perder datos o sobrescribir cambios no guardados. Si el monitor detecta un error, lo muestra en la consola y no repetirá continuamente el mismo intento; tras corregir y guardar el archivo, volverá a evaluar el cambio.

Las listas desplegables usan rangos con nombre y cubren las filas operativas 2 a 501. Si importás el libro en Google Sheets, verificá las reglas de validación en D y F una vez: el diseño de las listas se conserva cuando la importación mantiene los rangos con nombre, pero Google Sheets no ejecuta el archivo BAT ni el generador local. La automatización del lado de Sheets requiere un mecanismo propio de Google Apps Script o una integración autenticada; no se debe asumir que el proceso local sincroniza una hoja en la nube.
