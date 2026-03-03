"""
-------------------------------------------------------------------------
SISTEMA DE GESTIÓN DE ADMISIONES - INSTITUCIÓN EDUCATIVA
Autor: Edison Clase
Versión: 1.3.2 (Plantilla de Registro Extendida y Concatenación)
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
            "asunto": f"¡Registro Completado con Éxito! - Proceso de Admisión: {nombre_estudiante}",
            "cuerpo": f"""¡Registro Completado con Éxito! 

Saludos, {nombre_responsable},

Gracias por completar la solicitud de admisión para {nombre_estudiante}. Para asegurar una comunicación fluida y que no se pierda ningún detalle importante de las próximas fases, siga estos pasos: 

1. Únase al Grupo Oficial de Seguimiento:
Haga clic en el siguiente enlace para ingresar al grupo de WhatsApp exclusivo para solicitantes:
{link_wa}

2. Debe asistir a la reunión informativa: 
El día ________ tendremos la reunión informativa para todos los interesados, a las: ________, en el centro educativo.

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
Politécnico Prof. José Mercedes Alvino (CEJOMA)"""
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