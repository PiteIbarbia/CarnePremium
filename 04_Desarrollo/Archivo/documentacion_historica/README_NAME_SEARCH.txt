# Cambio de interfaz — búsqueda por nombre

Se mantiene el ID interno, pero el usuario trabaja escribiendo:
- Cliente por nombre
- Producto por nombre

No se utilizan listas desplegables para Cliente/Producto en Ventas.

La hoja 05_VENTAS mantiene internamente ClienteID y ProductoID en columnas técnicas
ocultas, mientras que las columnas de trabajo muestran Cliente y Producto.

Importante:
- Excel 2019 no tiene XLOOKUP, por eso se usa INDEX/MATCH.
- Google Sheets puede interpretar estas fórmulas.
- La búsqueda por nombre devuelve la primera coincidencia. Para nombres duplicados,
  el siguiente paso será incorporar un buscador/selector más avanzado o un código visible.
