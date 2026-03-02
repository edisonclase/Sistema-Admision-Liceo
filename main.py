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
    msg['Subject'] = f"¡Registro Completado con Éxito! - Admisión: {nombre_estudiante}"

    # CUERPO DEL MENSAJE ACTUALIZADO (MENSAJE OFICIAL)
    cuerpo = f"""
¡Registro Completado con Éxito! 

Estimado/a {nombre_responsable},

Gracias por completar la solicitud de admisión para {nombre_estudiante}. Para asegurar una comunicación fluida y que no se pierda ningún detalle importante de las próximas fases, siga estos pasos: 

1. Únase al Grupo Oficial de Seguimiento:
Haga clic en el siguiente enlace para ingresar al grupo de WhatsApp exclusivo para solicitantes:
{link_wa}

2. Debe asistir a la reunión informativa: 
El ______________________ tendremos la reunión informativa para todos los interesados, la hora de la reunión es: ___________________, en el centro educativo.

3. Información Importante sobre la Documentación:
Para el día de la prueba de admisión, su hijo(a) o representado(a) no necesita entregar ningún documento físico. La prioridad ese día es su desempeño en la evaluación. 

Sin embargo, para que puedan ir preparándose, les informamos que en caso de ser admitido(a), deberán presentar el expediente completo en un fólder durante el mes de ____________ (las fechas exactas de recepción se comunicarán oportunamente). 

Lista de requisitos para la inscripción definitiva: 
- Acta de nacimiento original y reciente (emisión para fines escolares).
- 2 fotografías 2x2. 
- Récord de calificaciones original del centro de procedencia.
- Certificación de sexto grado de primaria. 
- Copias de las cédulas de identidad de ambos padres y/o tutor(a).
- El Manual de Convivencia con las firmas requeridas. 
- Certificado médico y copia del carnet de seguro médico (si posee). 
- Formulario de inscripción completado (este se le entregará físicamente en el centro educativo en la fecha que le indiquemos). 

Toda la información proporcionada en este formulario está protegida bajo nuestros protocolos de Microsoft 365 y la Ley de Ciberseguridad, garantizando el uso exclusivo para fines académicos. 

¡Nos vemos pronto en la reunión informativa sobre nuestro centro educativo!

Atentamente,
Edison Clase
Liceo de Admisiones
    """
    msg.attach(MIMEText(cuerpo, 'plain'))
    
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=30) as server:
            server.login(remitente, password)
            server.send_message(msg)
        return True
    except Exception as e:
        print(f"❌ Error de envío: {e}")
        logging.error(f"Error SMTP: {e}")
        return False

# 3. LÓGICA DE PROCESAMIENTO
def ejecutar_proceso():
    """Descarga el Excel de SharePoint y procesa las filas pendientes."""
    url = os.getenv("EXCEL_LINK")
    
    if not url:
        print("⚠️ Error: No se encontró la URL del Excel en el archivo .env o Secrets")
        return

    # Limpieza de URL para forzar descarga directa
    if "action=embedview" in url:
        url = url.replace("action=embedview", "action=download")
    elif "viewer" in url:
         url = url.split('?')[0] + '?download=1'
    
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
                    print(f"Procesando admisión de: {nombre_estudiante} (Responsable: {nombre_responsable})...")
                    exito = enviar_notificacion(correo_destino, nombre_responsable, nombre_estudiante)
                    
                    if exito:
                        logging.info(f"ÉXITO: Enviado a {nombre_responsable} ({correo_destino})")
                        print(f"✅ Notificación enviada correctamente.")
                    else:
                        print(f"❌ Falló el envío para {nombre_responsable}.")

            except Exception as e:
                logging.error(f"Error procesando fila {index}: {e}")
                continue
                    
    except Exception as e:
        logging.critical(f"Error crítico en proceso: {e}")
        print(f"⚠️ Error: {e}")

# 4. PUNTO DE ENTRADA
if __name__ == "__main__":
    hora_actual = datetime.now().strftime('%H:%M:%S')
    print(f"--- Iniciando Sistema Admisión (Edison Clase) | {hora_actual} ---")
    ejecutar_proceso()
    print("--- Proceso finalizado ---")