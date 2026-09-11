#!/usr/bin/env python3
"""
Vigila la carpeta uploads/ y corre update_alarmas.py automaticamente
cada vez que agregas o modificas un archivo .xlsx.

Uso: dejalo corriendo mientras trabajas.
    python3 watch_uploads.py
(o haz doble clic en "Actualizar Panel.command")

Para detenerlo: Ctrl+C
"""
import time
from pathlib import Path

import update_alarmas

ROOT = Path(__file__).resolve().parent
UPLOADS_DIR = ROOT / "uploads"
POLL_SECONDS = 3


def snapshot():
    if not UPLOADS_DIR.exists():
        return {}
    files = list(UPLOADS_DIR.glob("*.xlsx")) + list(UPLOADS_DIR.glob("*.txt"))
    return {p.name: p.stat().st_mtime for p in files}


def main():
    print(f"Vigilando {UPLOADS_DIR} ... (Ctrl+C para detener)")
    last = snapshot()
    # correr una vez al iniciar, por si hay archivos pendientes
    update_alarmas.main()

    while True:
        time.sleep(POLL_SECONDS)
        current = snapshot()
        if current != last:
            print("\nCambio detectado en uploads/, actualizando base de datos...")
            try:
                update_alarmas.main()
            except Exception as exc:
                print(f"  ! Error actualizando: {exc}")
            last = current


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nDetenido.")
