# MVP 0.1.2

Este paso conecta el motor de negocio con la estructura Excel sin modificar todavía
el maestro con datos de prueba.

## Archivos

Scripts/
- ventas_excel.py
- dashboard.py

Pruebas/
- test_mvp_012.py

## Instalación

Copiar los dos `.py` de Scripts en:
04_Desarrollo/Scripts/

Copiar `test_mvp_012.py` en:
04_Desarrollo/Pruebas/

## Requisito

Mantener:
04_Desarrollo/Scripts/ventas.py
04_Desarrollo/Scripts/clientes.py
04_Desarrollo/Scripts/productos.py
04_Desarrollo/Scripts/precios.py
04_Desarrollo/Scripts/config.py

## Prueba

Con `.venv` activo:

```bash
python 04_Desarrollo/Pruebas/test_mvp_012.py
```

Debe terminar en:

`OK: MVP 0.1.2 SUPERADO`

El test abre el Excel pero no guarda los datos de prueba.

## Objetivo posterior

Después de pasar esta prueba, el siguiente paso será integrar esta lógica al builder de forma controlada
y comenzar el motor comercial:
Oferta > Especial > Mayorista > Minorista,
sin exigir todavía que existan clientes mayoristas ni ofertas reales.
