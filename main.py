"""
-------------------------------------------------------------------------
SISTEMA DE GESTIÓN DE ADMISIONES - INSTITUCIÓN EDUCATIVA
Autor: Edison Clase
Versión: 1.1.5
Python: 3.14.3
Descripción: Automatización de notificaciones vía SMTP (Gmail SSL) 
             cumpliendo con la Ley No. 172-13 y Ley No. 53-07 (Rep. Dom.)
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

# 1. CONFIGURACIÓN DE AUDITORÍA Y ENTORNO
if not os.path.exists('logs'):
    os.makedirs('logs')

logging.basicConfig(
    filename='logs/envios.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

load_dotenv()

def enviar_notificacion(correo_destino, nombre_responsable, nombre_estudiante):
    remitente = os.getenv("EMAIL_USER")
    password = os.getenv("EMAIL_PASS")
    link_wa = os.getenv("WHATSAPP_LINK")
    
    msg = MIMEMultipart()
    msg['From'] = remitente
    msg['To'] = correo_destino
    msg['Subject'] = f"¡Felicidades! Paso siguiente admisión: {nombre_estudiante}"

    cuerpo = f"""
    Estimado/a {nombre_responsable},
    
    Le felicitamos por completar el formulario para {nombre_estudiante}. 
    Para finalizar, únase al grupo oficial de WhatsApp:
    {link_wa}
    
    Atentamente,
    Edison Clase
    """
    msg.attach(MIMEText(cuerpo, 'plain'))
    
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=30) as server:
            server.login(remitente, password)
            server.send_message(msg)
        return True
    except Exception as e:
        print(f"❌ Error de envío: {e}")
        return False

# 3. LÓGICA DE PROCESAMIENTO
def ejecutar_proceso():
    """Descarga el Excel de SharePoint y procesa las filas pendientes."""
    url = os.getenv("EXCEL_LINK")
    
    if not url:
        print("⚠️ Error: No se encontró la URL del Excel en el archivo .env")
        return

    if "action=embedview" in url:
        url = url.replace("action=embedview", "action=download")
    
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(url, headers=headers, timeout=60)
        response.raise_for_status()
        
        df = pd.read_excel(BytesIO(response.content), engine='openpyxl')
        df.columns = df.columns.str.strip() 

        print(f"✅ Conexión exitosa. Filas encontradas: {len(df)}")

        for index, fila in df.iterrows():
            try:
                nombre_estudiante = str(fila.get('NombreCompleto', 'Estudiante')).strip()
                nombre_responsable = str(fila.get('NombreResponsable', 'Padre/Madre/Tutor')).strip()
                correo_destino = str(fila.get('CorreoResponsable', '')).strip()
                estado_solicitud = str(fila.get('Estado', '')).strip().upper()

                if "@" not in correo_destino:
                    continue

                if estado_solicitud == 'PENDIENTE':
                    print(f"Enviando correo a: {nombre_responsable}...")
                    exito = enviar_notificacion(correo_destino, nombre_responsable, nombre_estudiante)
                    
                    if exito:
                        logging.info(f"ÉXITO: Enviado a {nombre_responsable} ({correo_destino})")
                        print(f"✅ Notificación enviada correctamente.")
                    else:
                        print(f"❌ Falló el envío. Revisa logs/envios.log")

            except Exception as e:
                logging.error(f"Error procesando fila {index}: {e}")
                continue
                    
    except Exception as e:
        logging.critical(f"Error crítico: {e}")
        print(f"⚠️ Error: {e}")

# 4. PUNTO DE ENTRADA
if __name__ == "__main__":
    hora_actual = datetime.now().strftime('%H:%M:%S')
    print(f"--- Iniciando Sistema Admisión (Edison Clase) | {hora_actual} ---")
    ejecutar_proceso()
    print("--- Proceso finalizado ---")