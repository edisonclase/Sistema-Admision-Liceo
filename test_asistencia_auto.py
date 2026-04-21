import os
import pandas as pd
import smtplib
import requests
import urllib.parse
from io import BytesIO
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv

load_dotenv()

def enviar_pases_final():
    # 1. Configuración de enlaces (Basado en tu último link)
    url_base = "https://docs.google.com/forms/d/e/1FAIpQLSeSrlYN2H2LsO1PpQrWexqQkk33l8OxgpN2ehjU-E7HFYZN5Q/viewform"
    
    # 2. Cargar datos del Excel de Microsoft 365
    url_excel = os.getenv("EXCEL_LINK")
    try:
        response = requests.get(url_excel, headers={'User-Agent': 'Mozilla/5.0'}, timeout=60)
        df = pd.read_excel(BytesIO(response.content), engine='openpyxl')
        df.columns = df.columns.str.strip() # Limpiar espacios en nombres de columnas
    except Exception as e:
        print(f"Error al cargar Excel: {e}")
        return

    # 3. Conexión al servidor de correo
    remitente = os.getenv("EMAIL_USER")
    password = os.getenv("EMAIL_PASS")
    
    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(remitente, password)
    except Exception as e:
        print(f"Error de login: {e}")
        return

    print("🚀 Iniciando envío de pases...")

    # 4. Procesar cada fila
    for index, fila in df.iterrows():
        # Extraemos los datos según tus columnas
        id_solicitud = str(fila.get('IdSolicitud', '')).strip()
        nombre_estudiante = str(fila.get('NombreCompleto', 'Estudiante')).strip().upper()
        nombre_responsable = str(fila.get('NombreResponsable', 'Tutor')).strip().title()
        correo_destino = str(fila.get('CorreoResponsable', '')).strip()

        # Validación básica de correo
        if "@" not in correo_destino:
            print(f"⚠️ Saltando a {nombre_estudiante}: Correo inválido ({correo_destino})")
            continue

        # Construir URL pre-llenada con los IDs que identificamos
        params = {
            "usp": "pp_url",
            "entry.1151100480": id_solicitud,       # ID_Solicitud
            "entry.441170025": nombre_estudiante   # Representante del alumno/a
        }
        url_final = f"{url_base}?{urllib.parse.urlencode(params)}"
        
        # Generar QR (Tamaño 300x300 para asegurar lectura rápida de IDs largos)
        qr_url = f"https://api.qrserver.com/v1/create-qr-code/?size=300x300&data={urllib.parse.quote(url_final)}"

        # Crear el correo HTML
        msg = MIMEMultipart()
        msg['From'] = remitente
        msg['To'] = correo_destino
        msg['Subject'] = f"Pase de Entrada - Reunión de Admisión: {nombre_estudiante}"

        cuerpo_html = f"""
        <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; background-color: #f9f9f9; padding: 20px;">
            <div style="max-width: 500px; margin: auto; background: white; border-radius: 15px; border: 1px solid #e0e0e0; padding: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                <div style="text-align: center; border-bottom: 2px solid #1a73e8; padding-bottom: 15px; margin-bottom: 20px;">
                    <h2 style="color: #1a73e8; margin: 0;">Pase de Entrada Digital</h2>
                    <p style="font-size: 14px; color: #666;">Politécnico Prof. José Mercedes Alvino</p>
                </div>
                
                <p>Estimado(a) <strong>{nombre_responsable}</strong>,</p>
                <p>Este es el pase de acceso para la reunión del <strong>24 de abril</strong> correspondiente al alumno(a):</p>
                
                <div style="background-color: #f1f3f4; padding: 10px; border-radius: 8px; text-align: center; font-weight: bold; margin: 15px 0;">
                    {nombre_estudiante}
                </div>

                <div style="text-align: center; margin: 25px 0;">
                    <img src="{qr_url}" alt="Código QR de Asistencia" style="border: 4px solid #fff; outline: 1px solid #ddd;">
                    <p style="font-family: monospace; font-size: 11px; color: #999; margin-top: 10px;">ID: {id_solicitud}</p>
                </div>

                <p style="font-size: 13px; line-height: 1.4;">
                    <strong>Instrucciones para la entrada:</strong><br>
                    1. Tenga este código listo en la pantalla de su celular al llegar.<br>
                    2. Nuestro personal lo escaneará para registrar su entrada de forma automática.
                </p>

                <div style="margin-top: 25px; padding-top: 15px; border-top: 1px solid #eee; text-align: center; font-size: 12px; color: #888;">
                    Coordinación de Registro y Control Académico
                </div>
            </div>
        </body>
        </html>
        """
        msg.attach(MIMEText(cuerpo_html, 'html'))
        
        try:
            server.send_message(msg)
            print(f"✅ Enviado: {nombre_estudiante}")
        except Exception as e:
            print(f"❌ Error con {nombre_estudiante}: {e}")

    server.quit()
    print("\n🏁 Proceso completado exitosamente.")

if __name__ == "__main__":
    enviar_pases_final()