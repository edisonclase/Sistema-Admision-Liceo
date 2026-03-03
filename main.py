"""
-------------------------------------------------------------------------
SISTEMA DE GESTIÓN DE ADMISIONES - INSTITUCIÓN EDUCATIVA
Autor: Edison Clase
Versión: 1.3.1 (Ajuste de Columnas Real y Concatenación)
Python: 3.14.3
-------------------------------------------------------------------------
"""
import os
import pandas as pd
import smtplib
import logging
import requests
from io import BytesIO
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv

# 1. CONFIGURACIÓN DE AUDITORÍA Y PERSISTENCIA
if not os.path.exists('logs'): os.makedirs('logs')

DB_ENVIADOS = 'logs/enviados.txt'

def cargar_enviados():
    if not os.path.exists(DB_ENVIADOS): return set()
    with open(DB_ENVIADOS, 'r') as f:
        return set(line.strip() for line in f)

def guardar_id_enviado(id_solicitud, tipo_correo):
    with open(DB_ENVIADOS, 'a') as f:
        f.write(f"{id_solicitud}_{tipo_correo}\n")

logging.basicConfig(filename='logs/envios.log', level=logging.INFO, format='%(asctime)s - %(message)s')
load_dotenv()

# 2. MOTOR DE PLANTILLAS
def obtener_plantilla(tipo, nombre_responsable, nombre_estudiante):
    link_wa = os.getenv("WHATSAPP_LINK")
    
    plantillas = {
        "REGISTRO": {
            "asunto": f"¡Registro Completado con Éxito! - Admisión: {nombre_estudiante}",
            "cuerpo": f"Saludos, {nombre_responsable},\n\nGracias por completar la solicitud de admisión para {nombre_estudiante}. Para asegurar una comunicación fluida, únase al Grupo de WhatsApp: {link_wa}\n\nDebe asistir a la reunión informativa el día ________ a las ______ en el centro educativo.\n\nAtentamente,\nDepartamento de Registro."
        },
        "ADMITIDO": {
            "asunto": f"¡Felicidades! Admitido(a) - {nombre_estudiante}",
            "cuerpo": f"Estimado(a) {nombre_responsable},\n\nNos complace informarle que {nombre_estudiante} ha sido ADMITIDO(A) para el año escolar 2026-2027. Favor pasar por el centro del ___ al ___ de ___ para completar la inscripción física.\n\n¡Bienvenidos!"
        },
        "REPETIR": {
            "asunto": f"Nueva Convocatoria de Evaluación - {nombre_estudiante}",
            "cuerpo": f"Saludos, {nombre_responsable},\n\nLe informamos que {nombre_estudiante} debe presentarse a una nueva evaluación el día ________ a las ______.\n\nNo es necesario un nuevo registro."
        },
        "NO_ADMITIDO": {
            "asunto": f"Resultado Proceso de Admisión - {nombre_estudiante}",
            "cuerpo": f"Estimado(a) {nombre_responsable},\n\nAgradecemos su interés. Por el momento no ha sido posible otorgar una plaza para {nombre_estudiante} debido a cupos limitados. Le deseamos éxito en su búsqueda académica."
        }
    }
    return plantillas.get(tipo)

def enviar_notificacion(correo_destino, nombre_responsable, nombre_estudiante, tipo_correo):
    remitente = os.getenv("EMAIL_USER")
    password = os.getenv("EMAIL_PASS")
    datos_correo = obtener_plantilla(tipo_correo, nombre_responsable, nombre_estudiante)
    
    if not datos_correo: return False

    msg = MIMEMultipart()
    msg['From'] = remitente
    msg['To'] = correo_destino
    msg['Subject'] = datos_correo["asunto"]
    msg.attach(MIMEText(datos_correo["cuerpo"], 'plain'))
    
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=30) as server:
            server.login(remitente, password)
            server.send_message(msg)
        return True
    except Exception as e:
        print(f"❌ Error SMTP: {e}")
        return False

# 3. LÓGICA DE PROCESAMIENTO
def ejecutar_proceso():
    url = os.getenv("EXCEL_LINK")
    enviados = cargar_enviados()
    
    try:
        response = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=60)
        df = pd.read_excel(BytesIO(response.content), engine='openpyxl')
        df.columns = df.columns.str.strip() 

        print(f"✅ Conexión exitosa. Filas encontradas: {len(df)}")

        for index, fila in df.iterrows():
            try:
                id_solicitud = str(fila.get('IdSolicitud', '')).strip()
                
                # CONSTRUCCIÓN DEL NOMBRE DEL ESTUDIANTE (Nombres + Primer Apellido)
                nombres = str(fila.get('NombresEstudiante', '')).strip()
                apellido = str(fila.get('PrimerApellido', '')).strip()
                nombre_estudiante = f"{nombres} {apellido}".title()
                
                # DATOS DEL RESPONSABLE
                nombre_responsable = str(fila.get('NombreResponsable', 'Tutor')).strip().title()
                correo_destino = str(fila.get('CorreoResponsable', '')).strip()
                
                # LÓGICA DE ESTADOS
                estado = str(fila.get('Estado', '')).strip().upper()
                resultado = str(fila.get('Resultado_Final', '')).strip().upper()

                tipo_a_enviar = None
                if estado == 'PENDIENTE': 
                    tipo_a_enviar = "REGISTRO"
                elif resultado in ["ADMITIDO", "REPETIR", "NO_ADMITIDO"]: 
                    tipo_a_enviar = resultado

                # Validar si ya se envió este tipo de correo para este ID
                if not tipo_a_enviar or f"{id_solicitud}_{tipo_a_enviar}" in enviados:
                    continue

                if "@" in correo_destino:
                    print(f"🚀 Procesando {tipo_a_enviar} para: {nombre_estudiante}...")
                    if enviar_notificacion(correo_destino, nombre_responsable, nombre_estudiante, tipo_a_enviar):
                        guardar_id_enviado(id_solicitud, tipo_a_enviar)
                        print(f"✅ Enviado con éxito.")
            
            except Exception as e:
                print(f"Error en fila {index}: {e}")
                continue
                    
    except Exception as e:
        print(f"⚠️ Error crítico: {e}")

if __name__ == "__main__":
    print(f"--- Iniciando Sistema | {datetime.now().strftime('%H:%M:%S')} ---")
    ejecutar_proceso()
    print("--- Proceso finalizado ---")