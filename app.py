import os
import re
import sys
import uuid
import shutil
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