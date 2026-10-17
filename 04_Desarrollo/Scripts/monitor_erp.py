from __future__ import annotations

import hashlib
import subprocess
import sys
import time
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
MASTER_FILE = PROJECT_ROOT / "01_ERP" / "CarnePremiumERP.xlsx"
LOCK_FILE = MASTER_FILE.parent / ("~$" + MASTER_FILE.name)
UPDATER = PROJECT_ROOT / "04_Desarrollo" / "Scripts" / "actualizar_erp.py"
POLL_SECONDS = 2
STABLE_SECONDS = 5


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def signature(path: Path) -> tuple[int, int] | None:
    try:
        stat = path.stat()
        return stat.st_mtime_ns, stat.st_size
    except OSError:
        return None


def refresh() -> bool:
    print()
    print("=" * 72, flush=True)
    print("Validando una copia temporal antes de aplicar cambios", flush=True)
    print("=" * 72, flush=True)
    result = subprocess.run(
        [sys.executable, str(UPDATER), "--aplicar"],
        cwd=PROJECT_ROOT,
        check=False,
    )
    if result.returncode != 0:
        print(
            f"ERROR: el actualizador seguro terminó con código {result.returncode}. "
            "El libro no se considera actualizado.",
            flush=True,
        )
        return False
    print("ERP actualizado después de superar todas las pruebas.", flush=True)
    return True


def main() -> int:
    print("=" * 72)
    print("CARNE PREMIUM ERP — MONITOR DE ACTUALIZACIÓN SEGURA")
    print("=" * 72)
    print(f"Libro observado: {MASTER_FILE}")
    print("El monitor valida en una copia temporal antes de reemplazar el maestro.")
    print("Guardá y cerrá Excel para que pueda procesar el libro.")
    print("Para detenerlo, presioná Ctrl+C.")

    last_signature = signature(MASTER_FILE)
    last_hash = None
    stable_since = time.monotonic()
    attempted_hash = None

    if MASTER_FILE.exists():
        try:
            last_hash = sha256_file(MASTER_FILE)
        except OSError:
            pass

    try:
        while True:
            if not MASTER_FILE.exists() or LOCK_FILE.exists():
                stable_since = time.monotonic()
                time.sleep(POLL_SECONDS)
                continue

            current_signature = signature(MASTER_FILE)
            if current_signature is None:
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

            if current_hash == last_hash:
                attempted_hash = None
                time.sleep(POLL_SECONDS)
                continue

            if current_hash == attempted_hash:
                time.sleep(POLL_SECONDS)
                continue

            attempted_hash = current_hash
            print(
                f"\nCambio detectado ({time.strftime('%Y-%m-%d %H:%M:%S')}).",
                flush=True,
            )
            if refresh():
                try:
                    last_hash = sha256_file(MASTER_FILE)
                except OSError:
                    last_hash = None
            else:
                # No reintenta automáticamente el mismo archivo indefinidamente.
                last_hash = current_hash
            last_signature = signature(MASTER_FILE)
            stable_since = time.monotonic()
    except KeyboardInterrupt:
        print("\nMonitor detenido por el usuario.", flush=True)
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
