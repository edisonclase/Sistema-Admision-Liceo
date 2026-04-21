import pandas as pd
import requests
from io import BytesIO
import os
from dotenv import load_dotenv

load_dotenv()

def ejecutar_auditoria():
    url = os.getenv("EXCEL_LINK")
    try:
        # Descarga de datos
        response = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=60)
        df = pd.read_excel(BytesIO(response.content), engine='openpyxl')
        
        # Limpieza de nombres para comparación exacta
        df['Nombre_Completo'] = (df['NombresEstudiante'].astype(str).str.strip() + " " + 
                                 df['PrimerApellido'].astype(str).str.strip()).str.upper()
        
        # 1. DETECTAR DUPLICADOS
        # Consideramos duplicado si se repite el Nombre del Estudiante Y el Correo del Responsable
        total_respuestas = len(df)
        df_unicos = df.drop_duplicates(subset=['Nombre_Completo', 'CorreoResponsable'], keep='first')
        total_solicitantes_reales = len(df_unicos)
        
        # 2. ESTADÍSTICAS POR SEXO
        # Nota: Si la columna 'Sexo' no está en tu lista, asegúrate de que el Form la capture.
        # Si no está, el script te avisará.
        if 'Sexo' in df.columns:
            stats_sexo = df_unicos['Sexo'].value_counts()
        else:
            stats_sexo = "Columna 'Sexo' no encontrada en el Excel."

        # 3. RESULTADOS EN CONSOLA
        print("\n" + "="*40)
        print("REPORTE TÉCNICO DE DEPURACIÓN - CEJOMA")
        print("="*40)
        print(f"Total de registros recibidos:    {total_respuestas}")
        print(f"Registros duplicados eliminados: {total_respuestas - total_solicitantes_reales}")
        print(f"SOLICITANTES ÚNICOS (REALES):    {total_solicitantes_reales}")
        print("-" * 40)
        
        if isinstance(stats_sexo, pd.Series):
            print("DISTRIBUCIÓN POR SEXO:")
            for sexo, cant in stats_sexo.items():
                print(f" - {sexo}: {cant} ({ (cant/total_solicitantes_reales)*100 :.1f}%)")
        else:
            print(stats_sexo)
            
        print("="*40)

        # 4. EXPORTAR DATA LIMPIA
        if not os.path.exists('logs'): os.makedirs('logs')
        df_unicos.to_excel('logs/solicitantes_reales_depurados.xlsx', index=False)
        print("✅ Archivo 'logs/solicitantes_reales_depurados.xlsx' generado con éxito.")

    except Exception as e:
        print(f"❌ Error al procesar auditoría: {e}")

if __name__ == "__main__":
    ejecutar_auditoria()