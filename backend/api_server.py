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
from flask_cors import CORS
from werkzeug.utils import secure_filename
from pathlib import Path
from io import BytesIO
from xhtml2pdf import pisa
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

sys.path.insert(0, str(Path(__file__).parent))
import db

app = Flask(__name__)
CORS(app)

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


def generar_pdf_bytes(profesor, resultados):
    """Genera un PDF en memoria usando una plantilla HTML."""
    html_template = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Helvetica, Arial, sans-serif; padding: 40px; color: #333; }}
            .header {{ border-bottom: 4px solid #000; padding-bottom: 10px; margin-bottom: 30px; font-weight: bold; }}
            .footer {{ margin-top: 50px; color: #003B70; }}
            table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
            th, td {{ border: 1px solid #000; padding: 8px; text-align: left; font-size: 12px; }}
            th {{ background-color: #eee; }}
            .signature {{ margin-top: 40px; }}
        </style>
    </head>
    <body>
        <div class="header">Logística Ingeniería</div>
        <p>Buen Día, Cordial Saludo</p>
        <p>Apreciad@s, envío la información encontrada del profesor <strong>{profesor}</strong>:</p>
        
        <table>
            <thead>
                <tr>
                    <th>SEMESTRE</th>
                    <th>ASIGNATURA</th>
                    <th>SESIONES</th>
                    <th>DEPARTAMENTO</th>
                </tr>
            </thead>
            <tbody>
                {"".join([f'<tr><td>{r["semestre"]}</td><td>{r["asignatura"]}</td><td>{r["sesiones"]}</td><td>{r["departamento"]}</td></tr>' for r in resultados])}
            </tbody>
        </table>
        
        <p>Gracias por su amable atención. Sin otro particular,</p>
        
        <div class="signature">
            <p style="color: #003B70; font-weight: bold; font-size: 18px; margin: 0;">SANDRA TORRES</p>
            <p style="margin: 0;">Gestora Logística</p>
            <p style="margin: 0;">Facultad de Ingeniería</p>
            <p style="margin: 0;">Universidad de La Sabana</p>
        </div>
    </body>
    </html>
    """
    
    result = BytesIO()
    pisa_status = pisa.CreatePDF(html_template, dest=result)
    
    if pisa_status.err:
        return None
        
    return result.getvalue()


def generar_excel_bytes(profesor, resultados):
    """Genera un archivo Excel en memoria usando pandas."""
    df = pd.DataFrame(resultados)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Certificacion")
    return output.getvalue()


def enviar_correo_con_adjuntos(destinatario, profesor, pdf_data, excel_data):
    """Envía un correo con el PDF y el Excel adjuntos."""
    # --- CONFIGURACIÓN SMTP (Completar con datos reales) ---
    SMTP_SERVER = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
    SMTP_PORT = int(os.environ.get("SMTP_PORT", 587))
    SMTP_USER = os.environ.get("SMTP_USER", "tu-correo@gmail.com")
    SMTP_PASS = os.environ.get("SMTP_PASS", "tu-password-de-aplicacion")
    # -------------------------------------------------------

    mensaje = MIMEMultipart()
    mensaje["From"] = SMTP_USER
    mensaje["To"] = destinatario
    mensaje["Subject"] = f"Certificación Docente - {profesor}"

    cuerpo = f"Adjuntamos el certificado y el consolidado en Excel del profesor {profesor}."
    mensaje.attach(MIMEText(cuerpo, "plain"))

    # Adjuntar PDF
    parte_pdf = MIMEBase("application", "pdf")
    parte_pdf.set_payload(pdf_data)
    encoders.encode_base64(parte_pdf)
    parte_pdf.add_header("Content-Disposition", f"attachment; filename=Certificado_{profesor}.pdf")
    mensaje.attach(parte_pdf)

    # Adjuntar Excel
    parte_excel = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    parte_excel.set_payload(excel_data)
    encoders.encode_base64(parte_excel)
    parte_excel.add_header("Content-Disposition", f"attachment; filename=Consolidado_{profesor}.xlsx")
    mensaje.attach(parte_excel)

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.send_message(mensaje)
        return True
    except Exception as e:
        print(f"Error enviando correo: {e}")
        return False


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
@app.get("/opciones")
def opciones():
    """
    Devuelve las opciones de materias, departamentos y componentes 
    disponibles para un profesor y/o semestres específicos.
    """
    profesor = request.args.get("profesor", "").strip() or None
    semestres_raw = request.args.get("semestres", "").strip()
    
    semestres = None
    if semestres_raw:
        semestres = [
            s.strip() if s.strip().startswith("PERIODO") else f"PERIODO {s.strip()}"
            for s in semestres_raw.split(",") if s.strip()
        ]
        
    return ok(db.obtener_opciones(profesor=profesor, semestres=semestres))



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


# ------------------------------------------------------------------
@app.post("/enviar-certificado")
def enviar_certificado():
    """
    Recibe un PDF de certificación y simula el envío a un equipo de desarrollo.
    """
    if "archivo" not in request.files:
        return error("Falta el archivo del certificado.")

    archivo = request.files["archivo"]
    profesor = request.form.get("profesor", "Desconocido")

    # Guardar una copia local para auditoría
    nombre = secure_filename(f"envio_{profesor}_{archivo.filename}")
    ruta = UPLOAD_FOLDER / nombre
    archivo.save(ruta)

    # Aquí iría la lógica de smtplib para enviar el correo real
    print(f">>> SIMULACIÓN: Enviando certificado de {profesor} al equipo de desarrollo...")
    print(f">>> Archivo guardado en: {ruta}")

    return ok({
        "mensaje": f"Certificado de {profesor} enviado correctamente al equipo de desarrollo.",
        "archivo_id": nombre
    })


@app.post("/webhook/n8n")
def webhook_n8n():
    """
    Endpoint para ser llamado por n8n cuando llegue un correo de Gmail.
    Acepta el nombre directo en 'profesor' o lo extrae de 'body'.
    """
    data = request.json
    
    # 1. Intentar obtener el nombre directamente
    profesor_nombre = data.get("profesor")
    
    # 2. Si no está directo, buscar en el cuerpo del mensaje
    if not profesor_nombre:
        cuerpo = data.get("body", "")
        import re
        match = re.search(r"profesor:\s*(.*)", cuerpo, re.IGNORECASE)
        if match:
            profesor_nombre = match.group(1).strip()
    
    if not profesor_nombre:
        return error("No se encontró el nombre del profesor (usa el campo 'profesor' o 'body').")
        
    destinatario = data.get("from_email", "destinatario-por-defecto@gmail.com")
    
    # 3. Consultar datos en la DB
    try:
        resultados = db.consultar_consolidado(profesor=profesor_nombre)
        if not resultados:
            return error(f"No se encontraron datos para el profesor: {profesor_nombre}")
            
        # 2. Generar Archivos
        pdf_content = generar_pdf_bytes(profesor_nombre, resultados)
        excel_content = generar_excel_bytes(profesor_nombre, resultados)
        
        if not pdf_content or not excel_content:
            return error("Error generando los archivos en el servidor.")
            
        # 3. Enviar Correo con ambos archivos
        exito = enviar_correo_con_adjuntos(destinatario, profesor_nombre, pdf_content, excel_content)
        
        if exito:
            print(f">>> n8n AUTOMATION: Certificado y Excel enviados para {profesor_nombre}")
            return ok({"mensaje": f"Certificado y Excel enviados a {destinatario} para {profesor_nombre}"})
        else:
            return error("No se pudo enviar el correo, revisa la configuración SMTP.")
        
    except Exception as e:
        return error(f"Error en el proceso automático: {e}", status=500)


@app.post("/procesar-json")
def procesar_json():
    """
    Recibe el JSON estructurado desde n8n, genera los archivos y los devuelve 
    codificados en Base64 para que n8n pueda continuar el flujo.
    """
    import base64
    data = request.json
    profesor = data.get("profesor", "Desconocido")
    resultados = data.get("resultados") 

    if not resultados:
        return error("El JSON debe contener el campo 'resultados'.")

    try:
        # 1. Generar Archivos con los datos recibidos
        pdf_content = generar_pdf_bytes(profesor, resultados)
        excel_content = generar_excel_bytes(profesor, resultados)
        
        if not pdf_content or not excel_content:
            return error("Error generando los archivos en el servidor.")
            
        # 2. Codificar a Base64
        pdf_b64 = base64.b64encode(pdf_content).decode('utf-8')
        excel_b64 = base64.b64encode(excel_content).decode('utf-8')
        
        print(f">>> JSON API: Archivos base64 generados para {profesor}")
        
        return ok({
            "mensaje": f"Archivos de {profesor} generados con éxito.",
            "pdf_base64": pdf_b64,
            "excel_base64": excel_b64,
            "filename_pdf": f"Certificado_{profesor}.pdf",
            "filename_excel": f"Consolidado_{profesor}.xlsx"
        })
            
    except Exception as e:
        return error(f"Error procesando JSON: {e}", status=500)



# ---------------------------------------------------------------------------
# Arranque
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    db.init_db()
    port = int(os.environ.get("PORT", 5001))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    print(f"API corriendo en http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=debug)
