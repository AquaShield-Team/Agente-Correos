import os
import re
import sys
import uuid
import shutil
import webbrowser
import urllib.parse
import pythoncom
import openpyxl
import win32com.client
from flask import Flask, render_template, request, jsonify

# Configuracion de rutas (compatible con PyInstaller)
if getattr(sys, 'frozen', False):
    BUNDLE_DIR = sys._MEIPASS
    DIRECTORIO_ACTUAL = os.path.dirname(sys.executable)
else:
    BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))
    DIRECTORIO_ACTUAL = BUNDLE_DIR

ARCHIVO_DB = os.path.join(DIRECTORIO_ACTUAL, "Base_Clientes.xlsx")
UPLOAD_FOLDER = os.path.join(DIRECTORIO_ACTUAL, "uploads")

app = Flask(__name__,
            template_folder=os.path.join(BUNDLE_DIR, 'templates'),
            static_folder=os.path.join(BUNDLE_DIR, 'static'))
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def cargar_clientes():
    if not os.path.exists(ARCHIVO_DB):
        return []
    try:
        wb = openpyxl.load_workbook(ARCHIVO_DB, data_only=True, read_only=True)
    except PermissionError:
        temp_path = os.path.join(UPLOAD_FOLDER, "_temp_clientes.xlsx")
        shutil.copy2(ARCHIVO_DB, temp_path)
        wb = openpyxl.load_workbook(temp_path, data_only=True, read_only=True)

    ws = wb.active
    clientes = []

    def safe_cell(row, idx):
        if idx >= len(row) or row[idx] is None:
            return ""
        val = row[idx]
        # Normalizar IDs numericos: 1.0 -> "1"
        if isinstance(val, (int, float)):
            if val == int(val):
                return str(int(val))
        return str(val)

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 2 or not row[1]:
            continue
        clientes.append({
            "id": safe_cell(row, 0),
            "cliente": safe_cell(row, 1),
            "para": safe_cell(row, 2),
            "cc": safe_cell(row, 3),
            "asunto": safe_cell(row, 4).replace("\n", " "),
            "cuerpo": safe_cell(row, 5).replace("\n", "<br>")
        })
    wb.close()
    return clientes




def limpiar_parrafos_vacios(html):
    """Elimina <p class=MsoNormal> vacios que Outlook pone antes de la firma."""
    patron_p = re.compile(r'<p\b[^>]*>.*?</p>\s*', re.IGNORECASE | re.DOTALL)

    resultado = html
    for _ in range(10):
        match = patron_p.match(resultado)
        if not match:
            break
        texto = re.sub(r'<[^>]*>', '', match.group(0))
        texto = texto.replace('&nbsp;', '').replace('\r', '').replace('\n', '').strip()
        if texto:
            break
        resultado = resultado[match.end():]

    return resultado


# ── Rutas ─────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/clientes", methods=["GET"])
def api_clientes():
    return jsonify(cargar_clientes())


@app.route("/api/abrir_excel", methods=["GET"])
def abrir_excel():
    try:
        os.startfile(ARCHIVO_DB)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/generar_correo", methods=["POST"])
def generar_correo():
    try:
        data = request.form
        cliente_id = data.get("cliente_id")
        modo = data.get("modo", "clasico")  # "clasico" o "nuevo"

        clientes = cargar_clientes()
        cliente = next((c for c in clientes if str(c["id"]) == str(cliente_id)), None)

        if not cliente:
            return jsonify({"success": False, "error": "Cliente no encontrado"}), 404

        asunto = cliente.get("asunto", "").replace("[CLIENTE]", cliente["cliente"])
        cuerpo = cliente.get("cuerpo", "").replace("[CLIENTE]", cliente["cliente"])
        cuerpo = re.sub(r'(<br\s*/?>|\s)+$', '', cuerpo, flags=re.IGNORECASE)

        # ── Modo: Outlook Nuevo (mailto:) ─────────────────────
        if modo == "nuevo":
            para = cliente.get("para", "")
            cc = cliente.get("cc", "")
            # Limpiar HTML del cuerpo para mailto (texto plano)
            cuerpo_plano = re.sub(r'<br\s*/?>', '\n', cuerpo, flags=re.IGNORECASE)
            cuerpo_plano = re.sub(r'<[^>]*>', '', cuerpo_plano)

            params = {"subject": asunto, "body": cuerpo_plano}
            if cc:
                params["cc"] = cc

            mailto_url = f"mailto:{para}?{urllib.parse.urlencode(params, quote_via=urllib.parse.quote)}"
            webbrowser.open(mailto_url)

            # Nota sobre adjunto
            archivo_adjunto = request.files.get("archivo")
            msg = "Correo abierto en Outlook Nuevo."
            if archivo_adjunto and archivo_adjunto.filename:
                msg += " NOTA: Debes adjuntar el archivo manualmente en la ventana de Outlook."

            return jsonify({"success": True, "message": msg, "modo": "nuevo"})

        # ── Modo: Outlook Clasico (COM) ───────────────────────
        archivo_adjunto = request.files.get("archivo")
        ruta_archivo = None
        if archivo_adjunto and archivo_adjunto.filename:
            nombre_seguro = f"{uuid.uuid4().hex}_{archivo_adjunto.filename}"
            ruta_archivo = os.path.join(app.config['UPLOAD_FOLDER'], nombre_seguro)
            archivo_adjunto.save(ruta_archivo)

        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        def formatear_correos(correos_str):
            if not correos_str:
                return ""
            lista = [e.strip() for e in re.split(r'[,;]+', correos_str) if e.strip()]
            return "; ".join(lista)

        mail.Display()

        mail.To = formatear_correos(cliente.get("para", ""))
        mail.CC = formatear_correos(cliente.get("cc", ""))
        mail.Subject = asunto

        # Capturar HTML de Outlook (firma dentro de <html><body><div class=WordSection1>...)
        firma_html = mail.HTMLBody
        cuerpo_html = "<p class=MsoNormal style='margin:0;padding:0'><span style='font-family:Calibri,sans-serif;font-size:11.0pt'>" + cuerpo + "</span></p><br>"

        # Inyectar dentro de WordSection1, limpiando parrafos vacios
        ws_match = re.search(r'<div\s+class=\s*["\']?WordSection1["\']?\s*>', firma_html, re.IGNORECASE)
        if ws_match:
            antes = firma_html[:ws_match.end()]
            despues = firma_html[ws_match.end():]
            despues = limpiar_parrafos_vacios(despues)
            mail.HTMLBody = antes + cuerpo_html + despues
        else:
            body_match = re.search(r'<body[^>]*>', firma_html, re.IGNORECASE)
            if body_match:
                antes = firma_html[:body_match.end()]
                despues = firma_html[body_match.end():]
                despues = limpiar_parrafos_vacios(despues)
                mail.HTMLBody = antes + cuerpo_html + despues
            else:
                mail.HTMLBody = cuerpo_html + firma_html

        if ruta_archivo:
            mail.Attachments.Add(ruta_archivo)
            try:
                os.remove(ruta_archivo)
            except Exception as e:
                print(f"No se pudo eliminar el archivo temporal: {e}")

        return jsonify({"success": True, "message": "Correo abierto en Outlook Clasico", "modo": "clasico"})

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "error": str(e)}), 500


if __name__ == "__main__":
    print("Iniciando Agente de Correos en http://localhost:5055")
    app.run(port=5055, debug=True)