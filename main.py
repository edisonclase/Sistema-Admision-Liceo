"""
-------------------------------------------------------------------------
SISTEMA DE GESTIÓN DE ADMISIONES - INSTITUCIÓN EDUCATIVA
Autor: Edison Clase
Versión: 1.4.0 (Lógica de Notas y Seguimiento)
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

def obtener_plantilla(tipo, nombre_responsable, nombre_estudiante):
    link_wa = os.getenv("WHATSAPP_LINK")
    
    plantillas = {
        "REGISTRO": {
            "asunto": f"¡Registro Completado con Éxito! - Proceso de Admisión: {nombre_estudiante}",
            "cuerpo": f"""¡Registro Completado con Éxito! 

Saludos, {nombre_responsable},

Gracias por completar la solicitud de admisión para {nombre_estudiante}. Para asegurar una comunicación fluida, siga estos pasos: 

1. Únase al Grupo Oficial de Seguimiento:
{link_wa}

2. Reunión informativa: 
El día ________ a las: ________, en el centro educativo.

3. Sobre la Documentación:
No necesita entregar documentos físicos el día de la prueba. En caso de ser admitido(a), se requerirá el expediente completo en el mes de ____________. 

4. Seguimiento del Proceso:
Los resultados de la prueba de admisión serán notificados exclusivamente a través de este correo electrónico. El sistema le enviará una notificación automática una vez que la evaluación sea calificada (Puntaje mínimo de aprobación: 50 puntos).

¡Nos vemos pronto!

Atentamente,
Departamento de Registro (CEJOMA)"""
        },
        "ADMITIDO": {
            "asunto": f"¡Felicidades! Admitido(a) - {nombre_estudiante}",
            "cuerpo": f"Estimado(a) {nombre_responsable},\n\nNos complace informarle que {nombre_estudiante} ha superado la prueba con éxito (50 pts mín.) y ha sido ADMITIDO(A) para el año escolar 2026-2027.\n\nFavor pasar por el centro para completar la inscripción física."
        },
        "REPETIR": {
            "asunto": f"Nueva Convocatoria de Evaluación - {nombre_estudiante}",
            "cuerpo": f"Saludos, {nombre_responsable},\n\nLe informamos que {nombre_estudiante} debe presentarse a una nueva evaluación el día ________.\n\nÁnimo, ¡esta es una nueva oportunidad para alcanzar los 50 puntos requeridos!"
        },
        "NO_ADMITIDO": {
            "asunto": f"Resultado Proceso de Admisión - {nombre_estudiante}",
            "cuerpo": f"Estimado(a) {nombre_responsable},\n\nAgradecemos su interés. Por el momento no ha sido posible otorgar una plaza para {nombre_estudiante}. Le deseamos éxito en su búsqueda académica."
        }
    }
    return plantillas.get(tipo)

def enviar_notificacion(correo_destino, nombre_responsable, nombre_estudiante, tipo_correo):
    remitente = os.getenv("EMAIL_USER")
    password = os.getenv("EMAIL_PASS")
    datos_correo = obtener_plantilla(tipo_correo, nombre_responsable, nombre_estudiante)
    if not datos_correo: return False
    msg = MIMEMultipart(); msg['From'] = remitente; msg['To'] = correo_destino; msg['Subject'] = datos_correo["asunto"]
    msg.attach(MIMEText(datos_correo["cuerpo"], 'plain'))
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=30) as server:
            server.login(remitente, password); server.send_message(msg)
        return True
    except: return False

def ejecutar_proceso():
    url = os.getenv("EXCEL_LINK")
    enviados = cargar_enviados()
    try:
        response = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=60)
        df = pd.read_excel(BytesIO(response.content), engine='openpyxl')
        df.columns = df.columns.str.strip() 

        for index, fila in df.iterrows():
            id_solicitud = str(fila.get('IdSolicitud', '')).strip()
            nombre_estudiante = f"{str(fila.get('NombresEstudiante', ''))} {str(fila.get('PrimerApellido', ''))}".title()
            nombre_responsable = str(fila.get('NombreResponsable', 'Tutor')).strip().title()
            correo_destino = str(fila.get('CorreoResponsable', '')).strip()
            
            estado = str(fila.get('Estado', '')).strip().upper()
            resultado = str(fila.get('Resultado_Final', '')).strip().upper()
            
            tipo_a_enviar = None
            if estado == 'PENDIENTE':
                tipo_a_enviar = "REGISTRO"
            elif resultado in ["ADMITIDO", "REPETIR", "NO_ADMITIDO"]:
                tipo_a_enviar = resultado

            if tipo_a_enviar and f"{id_solicitud}_{tipo_a_enviar}" not in enviados:
                if "@" in correo_destino:
                    if enviar_notificacion(correo_destino, nombre_responsable, nombre_estudiante, tipo_a_enviar):
                        guardar_id_enviado(id_solicitud, tipo_a_enviar)
                        print(f"✅ {tipo_a_enviar} enviado a {nombre_estudiante}")
    except Exception as e: print(f"Error: {e}")

if __name__ == "__main__":
    ejecutar_proceso()