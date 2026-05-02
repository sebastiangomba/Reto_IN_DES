"""
api_server.py — API REST
Facultad de Ingeniería · Universidad de La Sabana

Endpoints:
  GET  /health
  GET  /semestres
  GET  /profesores
  GET  /consolidado
  POST /cargar
"""

import os
import sys
import pandas as pd
from flask import Flask, jsonify, request
from werkzeug.utils import secure_filename
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import db

app = Flask(__name__)

UPLOAD_FOLDER = Path(__file__).parent.parent / "uploads"
UPLOAD_FOLDER.mkdir(exist_ok=True)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

EXCEL_DTYPE = {
    "Numero documento docente": str,
    "Departamento":             str,
    "ID Sección Combinada":     str,
    "Nº Clase":                 str,
    "Id profesor":              str,
    "Sección Clase":            str,
    "ID Evento":                str,
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def ok(data, **kwargs):
    return jsonify({"ok": True, "data": data, **kwargs})


def error(mensaje, status=400):
    return jsonify({"ok": False, "error": mensaje}), status


def _leer_excel(path: str | Path) -> pd.DataFrame:
    return pd.read_excel(path, dtype=EXCEL_DTYPE)


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.get("/health")
def health():
    """Verificar que el servidor está corriendo."""
    return ok({"status": "ok", "db": str(db.DB_PATH)})


# ------------------------------------------------------------------
@app.get("/semestres")
def semestres():
    """
    Lista todos los semestres cargados en la base de datos.

    Respuesta:
      [
        { "ciclo_lectivo": "PERIODO 2024-2", "filas_insertadas": 3200, "cargado_en": "..." },
        ...
      ]
    """
    return ok(db.listar_semestres())


# ------------------------------------------------------------------
@app.get("/profesores")
def profesores():
    """
    Lista profesores disponibles. Acepta búsqueda parcial.

    Query params:
      q  (opcional) — fragmento de nombre, insensible a mayúsculas

    Ejemplo:
      GET /profesores?q=bravo
    """
    q = request.args.get("q", "").strip() or None
    return ok(db.listar_profesores(busqueda=q))


# ------------------------------------------------------------------
@app.get("/consolidado")
def consolidado():
    """
    Devuelve el consolidado de horas de un profesor por semestre.

    Query params:
      profesor   (opcional) — fragmento de nombre
      semestres  (opcional) — uno o varios, separados por coma
                              formato: '2024-2,2025-1' o 'PERIODO 2024-2'

    Ejemplos:
      GET /consolidado?profesor=BRAVO BUITRAGO
      GET /consolidado?profesor=BRAVO&semestres=2024-2,2025-1,2025-2
      GET /consolidado?semestres=2025-1

    Respuesta:
      [
        {
          "nombre_profesor":   "JOHN EDISON BRAVO BUITRAGO",
          "documento_docente": "1072661695",
          "semestre":          "PERIODO 2024-2",
          "asignatura":        "CALCULO INTEGRAL",
          "departamento":      "MATEMATICAS FISICA Y ESTADISTICA",
          "componente":        "Clase",
          "fecha_inicio":      "22/07/2024",
          "fecha_fin":         "30/12/2024",
          "sesiones":          128
        },
        ...
      ]
    """
    profesor = request.args.get("profesor", "").strip() or None

    semestres_raw = request.args.get("semestres", "").strip()
    semestres = None
    if semestres_raw:
        # Normalizar: aceptar '2024-2' o 'PERIODO 2024-2', siempre guardar con prefijo
        semestres = [
            s.strip() if s.strip().startswith("PERIODO") else f"PERIODO {s.strip()}"
            for s in semestres_raw.split(",") if s.strip()
        ]

    if not profesor and not semestres:
        return error("Debes enviar al menos 'profesor' o 'semestres' como parámetro.")

    filas = db.consultar_consolidado(profesor=profesor, semestres=semestres)

    # Convertir sesiones a int (DuckDB devuelve float)
    for f in filas:
        if f.get("sesiones") is not None:
            f["sesiones"] = int(f["sesiones"])

    return ok(filas, total=len(filas))


# ------------------------------------------------------------------
@app.post("/cargar")
def cargar():
    """
    Carga un Excel con datos de un semestre nuevo.
    Usa multipart/form-data con campo 'archivo'.

    Reglas:
      - Si el semestre ya existe en la DB, se omite (no duplica).
      - Valida que el archivo tenga las columnas requeridas.

    Ejemplo con curl:
      curl -X POST http://localhost:5001/cargar \\
           -F "archivo=@PREGRADO_2026_2.xlsx"

    Respuesta exitosa:
      {
        "ok": true,
        "data": {
          "semestres_nuevos":   ["PERIODO 2026-2"],
          "semestres_omitidos": [],
          "filas_insertadas":   3450,
          "mensaje":            "1 semestre(s) cargado(s) con 3450 filas."
        }
      }
    """
    if "archivo" not in request.files:
        return error("Falta el campo 'archivo' en el form-data.")

    archivo = request.files["archivo"]
    if not archivo.filename:
        return error("Nombre de archivo vacío.")

    nombre = secure_filename(archivo.filename)
    if not nombre.endswith((".xlsx", ".xls")):
        return error("Solo se aceptan archivos .xlsx o .xls.")

    ruta = UPLOAD_FOLDER / nombre
    archivo.save(ruta)

    try:
        df = _leer_excel(ruta)
    except Exception as e:
        return error(f"No se pudo leer el archivo: {e}")

    faltantes = db.validar_columnas(df)
    if faltantes:
        return error(f"Columnas faltantes en el archivo: {', '.join(faltantes)}")

    try:
        resultado = db.cargar_dataframe(df)
    except Exception as e:
        return error(f"Error al insertar datos: {e}", status=500)

    return ok(resultado)


# ---------------------------------------------------------------------------
# Arranque
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    db.init_db()
    port = int(os.environ.get("PORT", 5001))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    print(f"API corriendo en http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=debug)
