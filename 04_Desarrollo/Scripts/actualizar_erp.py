from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]
MASTER_FILE = PROJECT_ROOT / "01_ERP" / "CarnePremiumERP.xlsx"
SCRIPTS_DIR = PROJECT_ROOT / "04_Desarrollo" / "Scripts"
TESTS_DIR = PROJECT_ROOT / "04_Desarrollo" / "Pruebas"


def main() -> int:
    """Recalcula el ERP y valida el resultado en una sola ejecución local."""
    lock_file = MASTER_FILE.parent / f"~${MASTER_FILE.name}"
    if lock_file.exists():
        print("NO SE ACTUALIZÓ EL ERP.")
        print("Excel parece tener abierto el archivo maestro:")
        print(f"  {MASTER_FILE}")
        print("Guardá los cambios, cerrá ese libro y volvé a ejecutar ACTUALIZAR_ERP.bat.")
        print("No se modificó el archivo porque podría estar bloqueado por Excel.")
        return 2

    steps = [
        ("Actualizando ventas, Caja, Cuenta Corriente, tickets y reportes", SCRIPTS_DIR / "builder.py"),
        ("Validando el libro generado", TESTS_DIR / "test_mvp.py"),
        ("Ejecutando pruebas de Ticket, Caja y Cuenta Corriente", TESTS_DIR / "test_ticket_caja.py"),
    ]

    for label, script in steps:
        print()
        print("=" * 72)
        print(label)
        print("=" * 72)
        result = subprocess.run([sys.executable, str(script)], cwd=PROJECT_ROOT)
        if result.returncode != 0:
            print()
            print(f"ERROR: el paso terminó con código {result.returncode}.")
            print("El libro no se abrirá automáticamente. Revisá el mensaje anterior.")
            return result.returncode or 1

    print()
    print("ERP ACTUALIZADO Y VALIDADO CORRECTAMENTE.")
    print("Se abrirá el libro maestro para que puedas revisar el resultado.")
    if os.name == "nt" and MASTER_FILE.exists():
        os.startfile(str(MASTER_FILE))  # type: ignore[attr-defined]
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
