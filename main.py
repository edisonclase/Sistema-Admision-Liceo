"""
-------------------------------------------------------------------------
SISTEMA DE GESTIÓN DE ADMISIONES - INSTITUCIÓN EDUCATIVA
Autor: Edison Clase
Versión: 1.2.0 (Protección contra Duplicados por ID)
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
if not os.path.exists('logs'):
    os.makedirs('logs')

DB_ENVIADOS = 'logs/enviados.txt' # Archivo para guardar IDs procesados

def cargar_enviados():
    if not os.path.exists(DB_ENVIADOS):
        return set()
    with open(DB_ENVIADOS, 'r') as f:
        return set(line.strip() for line in f)

def guardar_id_enviado(id_solicitud):
    with open(DB_ENVIADOS, 'a') as f:
        f.write(f"{id_solicitud}\n")

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

    cuerpo = f"""
¡Registro Completado con Éxito! 

Saludos, {nombre_responsable},

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
Departamento de Registro y Control Académico
Politécnico Prof. José Mercedes Alvino
Cejoma
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
    url = os.getenv("EXCEL_LINK")
    if not url:
        print("⚠️ Error: No se encontró la URL del Excel")
        return

    enviados = cargar_enviados()
    
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(url, headers=headers, timeout=60)
        df = pd.read_excel(BytesIO(response.content), engine='openpyxl')
        df.columns = df.columns.str.strip() 

        print(f"✅ Conexión exitosa. Filas encontradas: {len(df)}")

        for index, fila in df.iterrows():
            try:
                # Usamos IdSolicitud como llave maestra
                id_solicitud = str(fila.get('IdSolicitud', '')).strip()
                nombre_estudiante = str(fila.get('NombreCompleto', 'Estudiante')).strip()
                nombre_responsable = str(fila.get('NombreResponsable', 'Padre/Madre/Tutor')).strip()
                correo_destino = str(fila.get('CorreoResponsable', '')).strip()
                estado_solicitud = str(fila.get('Estado', '')).strip().upper()

                if not id_solicitud or id_solicitud in enviados:
                    continue # Saltar si ya se envió o no hay ID

                if estado_solicitud == 'PENDIENTE' and "@" in correo_destino:
                    print(f"🚀 Enviando a: {nombre_estudiante} (ID: {id_solicitud})...")
                    exito = enviar_notificacion(correo_destino, nombre_responsable, nombre_estudiante)
                    
                    if exito:
                        guardar_id_enviado(id_solicitud)
                        logging.info(f"ÉXITO: ID {id_solicitud} enviado a {correo_destino}")
                        print(f"✅ Notificación enviada.")
                    else:
                        print(f"❌ Falló ID {id_solicitud}.")

            except Exception as e:
                continue
                    
    except Exception as e:
        print(f"⚠️ Error: {e}")

if __name__ == "__main__":
    print(f"--- Iniciando Sistema Admisión | {datetime.now().strftime('%H:%M:%S')} ---")
    ejecutar_proceso()
    print("--- Proceso finalizado ---")