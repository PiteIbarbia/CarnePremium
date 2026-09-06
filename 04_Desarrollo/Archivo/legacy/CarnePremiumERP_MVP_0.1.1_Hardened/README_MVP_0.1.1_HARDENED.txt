MVP 0.1.1 — Builder endurecido

Archivos:
- builder.py: construye el ERP y solo reemplaza el maestro después de validar el archivo nuevo.
- test_mvp.py: detecta automáticamente si CarnePremiumERP.xlsx es una versión vieja o si el último build no terminó correctamente.
- config.py: rutas y lista de hojas del proyecto.

Ejecución desde Git Bash, en la raíz CarnePremiumERP:

python 04_Desarrollo/Scripts/builder.py
python 04_Desarrollo/Pruebas/test_mvp.py

IMPORTANTE: cerrá Excel antes de ejecutar el builder.

El builder crea/actualiza 01_ERP/CarnePremiumERP.xlsx y genera 01_ERP/.build_state.json.
Antes de reemplazar un maestro existente, guarda un backup en 02_Backups/.

El test verifica:
- versión esperada;
- BuildToken;
- SHA-256 del maestro;
- hojas obligatorias;
- encabezados;
- fórmulas críticas de Ventas;
- validaciones de listas.
