"""
-------------------------------------------------------------------------
SISTEMA DE GESTIÓN DE ADMISIONES - POLITÉCNICO PROF. JOSÉ MERCEDES ALVINO
Autor: Edison Clase
Versión: 3.2.0 (Instrucciones de Llenado de Formulario y Envío desde Terminal)
-------------------------------------------------------------------------
"""
import os
import pandas as pd
import smtplib
import logging
import requests
from io import BytesIO
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv

if not os.path.exists('logs'): os.makedirs('logs')
DB_ENVIADOS = 'logs/enviados_cierre.txt'  # <-- NUEVO ARCHIVO EXCLUSIVO

def cargar_enviados():
    if not os.path.exists(DB_ENVIADOS): return set()
    with open(DB_ENVIADOS, 'r') as f:
        return set(line.strip() for line in f)

def guardar_id_enviado(id_solicitud, tipo_correo):
    with open(DB_ENVIADOS, 'a') as f:
        f.write(f"{id_solicitud}_{tipo_correo}\n")

logging.basicConfig(filename='logs/envios.log', level=logging.INFO, format='%(asctime)s - %(message)s')
load_dotenv()

def obtener_lista_documentos():
    return """
    <ol style="line-height: 1.6; text-align: left; max-width: 480px; margin: auto;">
        <li><strong>Formulario de Inscripción</strong> (completado, adjunto a este correo).</li>
        <li><strong>Acta de Nacimiento</strong> original para fines escolares (no requiere legalización).</li>
        <li><strong>Récord de Calificaciones</strong> original, debidamente firmado y sellado.</li>
        <li><strong>Certificación de Sexto Grado de Primaria</strong> original.</li>
        <li>Copia de la(s) <strong>Cédula(s)</strong> de los padres o tutores legales.</li>
        <li><strong>Certificado Médico</strong> oficial.</li>
        <li><strong>Historial Académico del SIGERD</strong> (sellado por el centro de procedencia).</li>
        <li>Copia del carnet del <strong>Seguro Médico</strong> (si posee).</li>
        <li><strong>Certificación escolar</strong> del último curso aprobado.</li>
        <li><strong>Dos (2) fotos 2x2 (Recientes)</strong> del estudiante, con su nombre completo e ID escrito con lapicero detrás.</li>
    </ol>
    <br>
    <div style="background-color: #fff3cd; border-left: 5px solid #ffc107; padding: 12px; font-size: 13px; color: #856404; display: inline-block; text-align: left;">
        <strong>NOTA CRÍTICA DE ENTREGA:</strong> Todos los documentos deben ser depositados en un <strong>folder</strong>. Es mandatorio que <strong>NO estén grapados</strong>. Por favor, retire todas las grapas antes de presentarse al centro educativo.
    </div>
    """

def obtener_plantilla_html(tipo, nombre_responsable, nombre_estudiante, fecha_entrega=""):
    documentos = obtener_lista_documentos()
    
    # 1. PLANTILLA ADMITIDO DIRECTO (SIN ACUERDO)
    if tipo == "ADMITIDO":
        asunto = f"¡Importante! Admisión Formal Confirmada - {nombre_estudiante}"
        cuerpo = f"""
        <div style="text-align: center;">
            <h1 style="color: #1a73e8; margin-bottom: 5px;">¡Felicidades!</h1>
            <h3 style="color: #5f6368; margin-top: 0; font-weight: normal;">Nos complace darle la bienvenida a nuestra familia educativa</h3>
        </div>
        <p>Estimado(a) <strong>{nombre_responsable}</strong>,</p>
        <p>Para la Coordinación Académica y el Departamento de Registro y Control Académico del <strong>Politécnico Prof. José Mercedes Alvino (CEJOMA)</strong>, es un honor y una profunda alegría informarle que el/la estudiante <strong>{nombre_estudiante}</strong> ha sido <strong>ADMITIDO(A)</strong> formalmente en nuestra institution para cursar el próximo año escolar.</p>
        
        <p>Nuestro compromiso a partir de hoy es acompañarle en el desarrollo de sus competencias académicas y humanas bajo nuestro lema: <em>"Formando con amor, seres justos y competentes"</em>.</p>
        
        <div style="background-color: #f1f3f4; border-radius: 8px; padding: 20px; border-left: 5px solid #1a73e8; margin: 20px 0;">
            <h4 style="margin-top: 0; color: #1a73e8; text-transform: uppercase;">Cronograma Obligatorio de Inscripción Presencial:</h4>
            <p style="font-size: 16px; margin: 5px 0;">Fecha Asignada: <strong>{fecha_entrega}</strong></p>
            <p style="font-size: 15px; margin: 5px 0;">Horario General: <strong>8:00 a.m. a 10:00 a.m.</strong> (Por orden de llegada)</p>
            <p style="font-size: 13px; color: #666; margin-top: 10px;"><em>* Este horario ha sido reservado de manera exclusiva para un grupo de familias con el fin de garantizar un proceso ágil y ordenado. Agradecemos su puntualidad.*</em></p>
        </div>

        <div style="background-color: #e8f0fe; border-radius: 8px; padding: 15px; border-left: 5px solid #1967d2; margin: 15px 0; font-size: 14px;">
            <strong>Información de la Asociación de Padres, Madres, Tutores y Amigos de Cejoma:</strong> Le informamos que el día de la entrega de documentos, los representantes de la <strong>Asociación de Padres, Madres, Tutores y Amigos de Cejoma</strong> de nuestro centro estarán presentes en la recepción para recibir el aporte anual correspondiente a las familias de nuevo ingreso.
        </div>

        <h4 style="color: #1a73e8;">Instrucción Importante sobre el Formulario Adjunto:</h4>
        <div style="background-color: #eaf2fd; border: 1px solid #b4d2ff; border-radius: 6px; padding: 15px; margin-bottom: 20px; font-size: 14px; line-height: 1.5;">
            Adjunto a este correo encontrará el <strong>Formulario de Inscripción Oficial</strong>. Es un requisito obligatorio que lo <strong>imprima y lo llene a mano utilizando lapicero de tinta azul</strong> antes de presentarse al centro el día de su cita. Por favor, <strong>deje completamente en blanco la casilla que dice "ID"</strong>, ya que este espacio es de uso exclusivo para el personal de Registro.
        </div>

        <h4 style="color: #1a73e8;">Requisitos y Documentación a Depositar:</h4>
        <p>El expediente físico debe ser depositado de manera presencial por el responsable o tutor del proceso. <strong>Para esta etapa no es necesaria la asistencia del alumno</strong>, a menos que se les informe que los uniformes del centro educativo estarán disponibles para la compra.</p>
        {documentos}
        """

    # 2. PLANTILLA CONDICIONADO (ACUERDO) - CONVOCATORIA A REUNIÓN DE COMPROMISO
    elif tipo == "ACUERDO":
        asunto = f"Aviso Importante: Proceso de Admisión y Convocatoria a Reunión - {nombre_estudiante}"
        cuerpo = f"""
        <div style="text-align: center;">
            <h2 style="color: #e67e22; margin-bottom: 5px;">Actualización de Proceso de Admisión</h2>
            <p style="color: #666; margin-top: 0;">Convocatoria Oficial a Firma de Compromiso Académico</p>
        </div>
        <p>Estimado(a) <strong>{nombre_responsable}</strong>,</p>
        <p>Le saludamos cordialmente desde la Coordinación Académica y el Departamento de Registro y Control Académico del <strong>Politécnico Prof. José Mercedes Alvino (CEJOMA)</strong>.</p>
        <p>Tras evaluar detalladamente las pruebas y entrevistas correspondientes al proceso de ingreso, informamos que la matriculación definitiva del estudiante <strong>{nombre_estudiante}</strong> estará sujeta y condicionada al establecimiento de un <strong>Acuerdo de Mejoras Académicas</strong> entre la institución y su familia.</p>
        <p>Esta medida responde al desempeño mostrado en las evaluaciones diagnósticas y tiene como único fin explicar un plan de apoyo psicopedagógico y de seguimiento continuo para asegurar la correcta nivelación e integración del estudiante en nuestro centro educativo.</p>
        
        <div style="background-color: #fdf5e6; border-radius: 8px; padding: 20px; border-left: 5px solid #e67e22; margin: 20px 0;">
            <h4 style="margin-top: 0; color: #e67e22; text-transform: uppercase;">Convocatoria Obligatoria a Reunión Institucional:</h4>
            <p style="font-size: 16px; margin: 5px 0;">Fecha: <strong>Martes, 23 de junio de 2026</strong></p>
            <p style="font-size: 16px; margin: 5px 0;">Horario: <strong>10:30 a.m.</strong></p>
            <p style="font-size: 14px; color: #666; margin-top: 5px;">Lugar: Cejoma.</p>
            <p style="font-size: 13px; color: #b75a00; margin-top: 10px;"><strong>Nota:</strong> En este encuentro formal conversaremos sobre los lineamientos técnicos del acuerdo, los compromisos de las partes y <strong>allí mismo se le asignará la fecha oficial de entrega para su expediente físico</strong>.</p>
        </div>

        <h4 style="color: #e67e22;">Instrucción Importante sobre el Formulario Adjunto:</h4>
        <div style="background-color: #fff9f2; border: 1px solid #ffe3c9; border-radius: 6px; padding: 15px; margin-bottom: 20px; font-size: 14px; line-height: 1.5;">
            Adjunto a este correo encontrará el <strong>Formulario de Inscripción Oficial</strong>. Le solicitamos que lo <strong>imprima y lo complete a mano utilizando lapicero de tinta azul</strong>. Recuerde traerlo ya completado el día de nuestra reunión. Por favor, <strong>deje en blanco la sección marcada como "ID"</strong>, ya que será asignada en la oficina de Control Académico.
        </div>

        <h4 style="color: #e67e22;">Documentación Requeridos para el Expediente (Vayan Preparando):</h4>
        <p>El tutor legal debe acudir a esta cita para pautar el plan de mejora. Los requisitos que conformarán el expediente final son los siguientes:</p>
        {documentos}
        """

    # 3. PLANTILLA REESTRUCTURADA PARA NO ADMITIDOS
    else:
        asunto = f"Información Importante: Estatus de Solicitud de Admisión - {nombre_estudiante}"
        cuerpo = f"""
        <div style="text-align: center;">
            <h2 style="color: #5f6368; margin-bottom: 5px;">Estatus de Solicitud de Admisión</h2>
            <p style="color: #888; margin-top: 0;">Actualización de Disponibilidad de Plazas</p>
        </div>
        <p>Estimado(a) <strong>{nombre_responsable}</strong>,</p>
        <p>Agradecemos sinceramente el alto interés y la confianza depositada por su familia en la propuesta formativa de nuestro <strong>Politécnico Prof. José Mercedes Alvino (CEJOMA)</strong>.</p>
        <p>Le informamos que debido a las limitaciones estrictas de infraestructura física y cupos máximos permitidos por aula en nuestras secciones, no será posible otorgarle una plaza regular al estudiante <strong>{nombre_estudiante}</strong> para el período escolar venidero en nuestra entidad.</p>
        
        <div style="background-color: #f8f9fa; border-radius: 8px; padding: 20px; border-left: 5px solid #7f8c8d; margin: 20px 0;">
            <h4 style="margin-top: 0; color: #2c3e50; text-transform: uppercase;">Canalización y Ubicación de Cupo a través del Ministerio:</h4>
            <p>Queremos asegurarle que el derecho a la educación de su hijo(a) está plenamente resguardado. Los datos de su solicitud de admisión han sido transferidos de manera automática a la <strong>Plataforma de Distribución de Cupos del Ministerio de Educación (MINERD)</strong>.</p>
            <p style="margin-top: 10px;">La Dirección del Distrito Educativo correspondiente le estará contactando próximamente utilizando las vías registradas para asistirle y asegurarle la asignación de una plaza en otro centro educativo de la zona que cuente con disponibilidad física inmediata.</p>
        </div>
        <p>Reiteramos nuestro agradecimiento por su esfuerzo durante las etapas evaluativas y deseamos el mayor de los éxitos en la trayectoria formativa del estudiante.</p>
        """

    plantilla_final = f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #333; background-color: #f4f6f9; padding: 20px;">
        <div style="max-width: 600px; margin: auto; background: white; border-radius: 12px; padding: 30px; box-shadow: 0 4px 10px rgba(0,0,0,0.05); border-top: 10px solid #1a73e8;">
            <div style="text-align: center; border-bottom: 1px solid #eee; padding-bottom: 15px; margin-bottom: 20px;">
                <span style="font-size: 12px; font-weight: bold; color: #888; text-transform: uppercase; letter-spacing: 1px;">Santiago, R. D. • CEJOMA</span>
            </div>
            {cuerpo}
            <div style="margin-top: 35px; padding-top: 20px; border-top: 1px solid #eee; text-align: center; font-size: 11px; color: #7f8c8d; line-height: 1.5;">
                <strong>Coordinación Académica & Coordinación de Registro y Control Académico</strong><br>
                Politécnico Prof. José Mercedes Alvino (CEJOMA)<br>
                Santiago, República Dominicana
            </div>
        </div>
    </body>
    </html>
    """
    return asunto, plantilla_final

def enviar_notificacion_v3(correo_destino, nombre_responsable, nombre_estudiante, tipo_correo, fecha_entrega=""):
    remitente = os.getenv("EMAIL_USER")
    password = os.getenv("EMAIL_PASS")
    
    asunto, html_cuerpo = obtener_plantilla_html(tipo_correo, nombre_responsable, nombre_estudiante, fecha_entrega)
    
    msg = MIMEMultipart()
    
    # 📑 CAMBIO AQUÍ: Esto define el nombre que verán los padres en su bandeja de entrada
    msg['From'] = f"Politécnico Cejoma <{remitente}>"
    
    msg['To'] = correo_destino
    msg['Subject'] = asunto
    msg.attach(MIMEText(html_cuerpo, 'html'))
    # ... (el resto de la función se queda exactamente igual)
    
    if tipo_correo in ["ADMITIDO", "ACUERDO"]:
        ruta_formulario = "documentos/formulario_inscripcion.pdf"
        if os.path.exists(ruta_formulario):
            with open(ruta_formulario, "rb") as adjunto:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(adjunto.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f"attachment; filename= Formulario_Inscripcion_CEJOMA.pdf")
                msg.attach(part)
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=30) as server:
            server.login(remitente, password)
            server.send_message(msg)
        return True
    except Exception as e:
        logging.error(f"Error enviando a {correo_destino}: {e}")
        return False

def ejecutar_proceso():
    # =========================================================================
    # 🛠️ PANEL DE CONTROL PRINCIPAL
    # =========================================================================
    FASE_ACTUAL = "CIERRE"           
    MODO_PRUEBA_ESTRICTO = False     
    CORREO_A_PROBAR = "clasedison@gmail.com"
    HOJA_RESULTADOS = "HojaPrueba"   
    # =========================================================================

    url = os.getenv("EXCEL_LINK")
    enviados = cargar_enviados()
    
    try:
        response = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=60)
        archivo_excel = BytesIO(response.content)
        
        if FASE_ACTUAL == "CIERRE":
            print(f"🚀 Procesando FASE 2: Distribución de Resultados y Asignación Logística...")
            df_resultados = pd.read_excel(archivo_excel, sheet_name=HOJA_RESULTS if 'HOJA_RESULTS' in locals() else HOJA_RESULTADOS, engine='openpyxl')
            df_resultados.columns = df_resultados.columns.str.strip()

            contador_admitidos_directos = 0
            contador_acuerdos = 0

            for index, fila in df_resultados.iterrows():
                id_solicitud = str(fila.get('No.', '')).strip()  
                
                nombre_estudiante = str(fila.get('NombreSolicitante', 'Solicitante')).strip().title()
                correo_destino = str(fila.get('CorreoResponsable', '')).strip().lower()
                estado_excel = str(fila.get('ESTADO', '')).strip().upper()  

                nombre_responsable = "Tutor/a" 

                if not correo_destino or "@" not in correo_destino:
                    continue

                tipo_envio = None
                fecha_asignada = ""

                if estado_excel == "ADMITIDO":
                    tipo_envio = "ADMITIDO"
                    if contador_admitidos_directos < 29:
                        fecha_asignada = "Lunes, 20 de julio de 2026"
                    elif contador_admitidos_directos < 58:
                        fecha_asignada = "Martes, 21 de julio de 2026"
                    else:
                        fecha_asignada = "Miércoles, 22 de julio de 2026"
                    
                    contador_admitidos_directos += 1

                elif estado_excel == "ACUERDO":
                    tipo_envio = "ACUERDO"
                    fecha_asignada = "CONVOCATORIA_REUNION" 
                    contador_acuerdos += 1

                elif estado_excel == "RECHAZADO" or "NO" in estado_excel:
                    tipo_envio = "NO_ADMITIDO"

                if tipo_envio and f"{id_solicitud}_{tipo_envio}" not in enviados:
                    if MODO_PRUEBA_ESTRICTO and correo_destino != CORREO_A_PROBAR.lower():
                        print(f"🚫 SIMULACIÓN: {tipo_envio} -> {nombre_estudiante}")
                        continue

                    if enviar_notificacion_v3(correo_destino, nombre_responsable, nombre_estudiante, tipo_envio, fecha_asignada):
                        if not MODO_PRUEBA_ESTRICTO: guardar_id_enviado(id_solicitud, tipo_envio)
                        print(f"🎯 ENVIADO [{tipo_envio}]: Correo de {nombre_estudiante} despachado con éxito a {correo_destino}.")

            print("\n" + "="*50)
            print(f"📈 RESUMEN DE PROCESAMIENTO LOGÍSTICO:")
            print(f"   * Admitidos directos distribuidos: {contador_admitidos_directos} estudiantes.")
            print(f"   * Convocados a firma de Acuerdo: {contador_acuerdos} estudiantes.")
            print(f"   * Total procesados en Fase Final: {contador_admitidos_directos + contador_acuerdos} de 99.")
            print("="*50)
            
    except Exception as e: 
        print(f"❌ Error General: {e}")

if __name__ == "__main__":
    ejecutar_proceso()