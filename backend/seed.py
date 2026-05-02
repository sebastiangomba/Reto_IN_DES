"""
seed.py — Migración inicial
Facultad de Ingeniería · Universidad de La Sabana

Carga el Excel histórico completo (2016-2 → 2026-1) en facultad.db.
Ejecutar UNA sola vez. Las cargas futuras van por POST /cargar.

Uso:
    python backend/seed.py --excel PREGRADO_CONSOLIDADO_2016_2_2026_1.xlsx
    python backend/seed.py --excel datos.xlsx --db /ruta/personalizada/facultad.db
"""

import argparse
import sys
from pathlib import Path
import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
import db

EXCEL_DTYPE = {
    "Numero documento docente": str,
    "Departamento":             str,
    "ID Sección Combinada":     str,
    "Nº Clase":                 str,
    "Id profesor":              str,
    "Sección Clase":            str,
    "ID Evento":                str,
}


def main():
    parser = argparse.ArgumentParser(description="Migración inicial a SQLite — Facultad de Ingeniería")
    parser.add_argument("--excel", required=True, help="Ruta al Excel histórico completo")
    parser.add_argument("--db",    help="Ruta alternativa para facultad.db (opcional)")
    args = parser.parse_args()

    if args.db:
        db.DB_PATH = Path(args.db)

    ruta = Path(args.excel)
    if not ruta.exists():
        print(f"Error: no se encontró el archivo '{ruta}'")
        sys.exit(1)

    print(f"Inicializando base de datos en: {db.DB_PATH}")
    db.init_db()

    print(f"Leyendo {ruta} ...")
    df = pd.read_excel(ruta, dtype=EXCEL_DTYPE)
    print(f"  {len(df):,} filas · {len(df.columns)} columnas")

    faltantes = db.validar_columnas(df)
    if faltantes:
        print(f"\nError: faltan columnas requeridas en el Excel:")
        for c in faltantes:
            print(f"  - {c}")
        sys.exit(1)

    print("\nInsertando datos...")
    resultado = db.cargar_dataframe(df)

    print(f"\n{'='*50}")
    print(f"  Semestres cargados : {len(resultado['semestres_nuevos'])}")
    for s in resultado["semestres_nuevos"]:
        print(f"    + {s}")
    if resultado["semestres_omitidos"]:
        print(f"  Semestres omitidos : {len(resultado['semestres_omitidos'])} (ya existían)")
    print(f"  Filas insertadas   : {resultado['filas_insertadas']:,}")
    print(f"{'='*50}")
    print(f"\n✅ {resultado['mensaje']}")
    print(f"   Base de datos lista en: {db.DB_PATH}")


if __name__ == "__main__":
    main()
