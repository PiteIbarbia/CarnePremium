from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import time
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]
MASTER_FILE = PROJECT_ROOT / "01_ERP" / "CarnePremiumERP.xlsx"
STATE_FILE = MASTER_FILE.parent / ".build_state.json"
LOCK_FILE = MASTER_FILE.parent / f"~${MASTER_FILE.name}"
SCRIPTS_DIR = PROJECT_ROOT / "04_Desarrollo" / "Scripts"
TESTS_DIR = PROJECT_ROOT / "04_Desarrollo" / "Pruebas"
POLL_SECONDS = 2
STABLE_SECONDS = 5


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_built_hash() -> str | None:
    try:
        state = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        value = state.get("sha256")
        return value if isinstance(value, str) else None
    except (OSError, json.JSONDecodeError):
        return None


def signature(path: Path) -> tuple[int, int] | None:
    try:
        stat = path.stat()
        return stat.st_mtime_ns, stat.st_size
    except OSError:
        return None


def refresh() -> bool:
    steps = [
        ("Generando ventas, Caja, Cuenta Corriente, tickets y reportes", SCRIPTS_DIR / "builder.py"),
        ("Validando el libro generado", TESTS_DIR / "test_mvp.py"),
        ("Probando Ticket, Caja y Cuenta Corriente", TESTS_DIR / "test_ticket_caja.py"),
    ]
    for label, script in steps:
        print()
        print("=" * 72, flush=True)
        print(label, flush=True)
        print("=" * 72, flush=True)
        result = subprocess.run([sys.executable, str(script)], cwd=PROJECT_ROOT)
        if result.returncode != 0:
            print(
                f"ERROR: {script.name} terminó con código {result.returncode}. "
                "Corregí el problema y guardá el libro otra vez.",
                flush=True,
            )
            return False
    print("ERP ACTUALIZADO Y VALIDADO.", flush=True)
    return True


def main() -> int:
    print("=" * 72)
    print("CARNE PREMIUM ERP — MONITOR DE ACTUALIZACIÓN")
    print("=" * 72)
    print(f"Libro observado: {MASTER_FILE}")
    print("Dejá esta ventana abierta. Guardá y cerrá el libro Excel para activar la actualización.")
    print("Para detener el monitor, presioná Ctrl+C en esta ventana.")
    print("")

    last_signature = None
    stable_since = time.monotonic()
    attempted_hash = None

    try:
        while True:
            if not MASTER_FILE.exists():
                time.sleep(POLL_SECONDS)
                continue

            current_signature = signature(MASTER_FILE)
            if current_signature is None:
                time.sleep(POLL_SECONDS)
                continue

            if LOCK_FILE.exists():
                last_signature = current_signature
                stable_since = time.monotonic()
                time.sleep(POLL_SECONDS)
                continue

            if current_signature != last_signature:
                last_signature = current_signature
                stable_since = time.monotonic()
                time.sleep(POLL_SECONDS)
                continue

            if time.monotonic() - stable_since < STABLE_SECONDS:
                time.sleep(POLL_SECONDS)
                continue

            try:
                current_hash = sha256_file(MASTER_FILE)
            except OSError:
                time.sleep(POLL_SECONDS)
                continue

            built_hash = read_built_hash()
            if current_hash == built_hash:
                attempted_hash = None
                time.sleep(POLL_SECONDS)
                continue

            if current_hash == attempted_hash:
                time.sleep(POLL_SECONDS)
                continue

            print(
                f"\nCambio detectado ({time.strftime('%Y-%m-%d %H:%M:%S')}). "
                "Esperando archivo cerrado y estable: OK.",
                flush=True,
            )
            attempted_hash = current_hash
            refresh()
            last_signature = signature(MASTER_FILE)
            stable_since = time.monotonic()
            # Builder updates .build_state.json with the resulting file hash.
            # If validation failed, do not retry the identical input every loop.
    except KeyboardInterrupt:
        print("\nMonitor detenido por el usuario.", flush=True)
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
