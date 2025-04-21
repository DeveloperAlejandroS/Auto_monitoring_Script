import re
import pandas as pd
from openpyxl import load_workbook
from fuzzywuzzy import fuzz, process
from datetime import datetime
import pytz, time
import psutil
from timezonefinder import TimezoneFinder
import subprocess

# Función para cruzar información del proveedor entre archivos y obtener "Vendor" para cada "Estación"
def get_vendor(excel_path, aux_path, sheet_name):
    """
    Cruza información entre un archivo Excel principal y un archivo auxiliar para obtener 
    detalles del proveedor, como "Vendor", "Feed Index", "Channel" y "Feed" para cada "Estación".

    Parámetros:
    - excel_path (str): Ruta del archivo Excel principal donde se almacenarán los resultados.
    - aux_path (str): Ruta del archivo Excel auxiliar que contiene la información adicional.
    - sheet_name (str): Nombre de la hoja en la que se guardarán los datos cruzados.

    Proceso:
    1. Carga el archivo principal y el auxiliar.
    2. Realiza un merge entre ambas fuentes de datos, relacionando la columna "Estación" del 
       archivo principal con la columna "Estacion" del auxiliar.
    3. Extrae las columnas "Vendor", "Feed Index", "Channel" y "Feed".
    4. Guarda los datos cruzados en la hoja de trabajo especificada dentro del archivo principal.

    El archivo principal es modificado para incluir la nueva información sin sobrescribir 
    otras hojas de trabajo existentes.

    Retorna:
    - No retorna ningún valor, pero modifica el archivo Excel de entrada.
    """
    # Cargar los datos principales y auxiliares
    df_main = pd.read_excel(excel_path)
    df_auxiliar = pd.read_excel(aux_path, sheet_name='Channel Info Monitoria')
    
    # Realizar merge para obtener Vendor, Feed Index, Channel, Feed
    df_extract = df_main.merge(
        df_auxiliar[['Estacion', 'Vendor', 'Feed Index', 'Channel', 'Feed']],
        left_on='Estación',
        right_on='Estacion',
        how='left'
    )[['Vendor', 'Feed Index', 'Channel', 'Feed']]

    # Guardar la hoja con los datos cruzados
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df_extract.to_excel(writer, sheet_name=sheet_name, index=False)

# Función para formatear fechas y guardarlas en una hoja de trabajo
def format_date(excel_path, sheet_name):
    """
    Formatea las fechas en la columna B de un archivo Excel y las guarda en la columna E.

    Esta función lee los datos de la columna B, los convierte a formato de fecha (MM/DD/YYYY)
    y sobrescribe los valores en la columna E del mismo archivo.

    Parámetros:
    - excel_path (str): Ruta del archivo Excel donde se realizará la transformación.
    - sheet_name (str): Nombre de la hoja en la que se encuentra la información.

    Comportamiento:
    - Si la columna B contiene datos inválidos, estos se convertirán en NaT.
    - Si hay errores al escribir en el archivo, se capturan y se imprimen en la consola.

    Excepción:
    - Captura y muestra errores si la escritura en el archivo falla.

    """
    # Leer solo la columna B
    df = pd.read_excel(excel_path, usecols='B', names=['Fecha'])
    # Convertir columna a formato de fecha y aplicar formato MM/DD/YYYY
    df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce', dayfirst=True).dt.strftime('%m/%d/%Y')

    # Guardar la hoja de trabajo actualizada con las fechas formateadas en columna E
    try:
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=4, index=False)
    except Exception as e:
        print(f"Error al escribir en la hoja: {e}")

# Función para formatear hora al formato de 24 horas (HH:MM:SS)
def format_hour_column(excel_path, sheet_name):
    """
    Convierte los valores de la columna F de un archivo Excel a formato de 24 horas (HH:MM:SS) 
    y los guarda en la misma columna.

    Parámetros:
    - excel_path (str): Ruta del archivo Excel donde se realizará la transformación.
    - sheet_name (str): Nombre de la hoja en la que se encuentra la información.

    Comportamiento:
    - Lee los datos de la columna F del archivo Excel.
    - Convierte los valores a formato de 24 horas (HH:MM:SS).
    - Guarda los valores formateados en la columna F del mismo archivo.
    - Si un valor no es una hora válida, se convierte en NaT.

    Excepción:
    - Captura y muestra errores si la escritura en el archivo falla.

    """
    df = pd.read_excel(excel_path, usecols='F', names=['Horario'])
    
    # Convertir a tiempo en formato 24 horas
    df['Horario'] = pd.to_datetime(df['Horario'], format='%H:%M:%S', errors='coerce').dt.time

    # Guardar la columna formateada en el archivo
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=5, index=False)

# Función para agregar información de cantidad (CANTIDAD) y tipo de spot (Type Spot)
def fill_spot_info(excel_path, sheet_name):
    """
    Agrega información de cantidad (Cantidad) y tipo de spot (Type Spot) a un archivo Excel.

    Parámetros:
    - excel_path (str): Ruta del archivo Excel donde se realizará la modificación.
    - sheet_name (str): Nombre de la hoja en la que se agregarán las columnas.

    Comportamiento:
    - Carga la hoja de cálculo especificada.
    - Agrega una nueva columna "Cantidad" con un valor fijo de 1.
    - Agrega una nueva columna "Type Spot" con un valor fijo de 'Paid'.
    - Guarda las nuevas columnas en las posiciones H (columna 8) e I (columna 9) de la hoja.

    Excepción:
    - Si ocurre un error al escribir en el archivo, se lanzará una excepción.

    """
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    # Agregar columna CANTIDAD con valor 1 y Type Spot con 'Paid'
    df['Cantidad'] = 1
    df['Type Spot'] = 'Paid'

    # Guardar las columnas nuevas en la hoja transformada en la columna H y I
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df[['Cantidad', 'Type Spot']].to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=7, index=False)

# Función para obtener datos creativos de una lista auxiliar basándose en coincidencias aproximadas
def get_creatives_data(excel_path, aux_path, sheet_name):
    aux_sheet_name = 'Month_Rotation'
    aux_side_sheet_name = 'Ads DB'
    df_main = pd.read_excel(excel_path)
    df_aux = pd.read_excel(aux_path, sheet_name=aux_sheet_name)
    side_aux_df = pd.read_excel(aux_path, sheet_name=aux_side_sheet_name)
    threshold = 91  # Umbral de similitud

    # Función para normalizar texto: minúsculas y sin espacios
    def normalize(text):
        return re.sub(r'\d+', '', ''.join(text.lower().split()))

    # Mapeo de valores normalizados a su forma original
    original_mapping = {normalize(row['Creativo']): row['Creativo'] for _, row in df_aux.iterrows()}
    side_original_mapping = {normalize(row['Creativo']): row['Creativo'] for _, row in side_aux_df.iterrows()}

    # Limpiar columna "Versión"
    df_main['Versión'] = df_main['Versión'].replace(r'\(\d{2}s\)', '', regex=True)

    # Crear columnas auxiliares con creativos normalizados
    df_aux_lower = df_aux.copy()
    df_aux_lower['Creativo'] = df_aux_lower['Creativo'].apply(normalize)

    df_side_aux_lower = side_aux_df.copy()
    df_side_aux_lower['Creativo'] = df_side_aux_lower['Creativo'].apply(normalize)

    # Función de búsqueda fuzzy
    def get_best_match_info(value, choices_lower, side_choices_lower, original_mapping, side_original_mapping, threshold):
        normalized_value = normalize(value)
        match, score = process.extractOne(normalized_value, choices_lower['Creativo'].tolist())

        if score >= threshold:
            original_creativo = original_mapping.get(match, match)
            row_original = choices_lower[choices_lower['Creativo'] == match]
            return pd.Series([
                original_creativo,
                row_original['Duration'].values[0],
                row_original['Brand'].values[0],
                'Found'
            ])
        else:
            match, score = process.extractOne(normalized_value, side_choices_lower['Creativo'].tolist())
            if score >= threshold:
                original_creativo = side_original_mapping.get(match, match)
                row_original = side_choices_lower[side_choices_lower['Creativo'] == match]
                return pd.Series([
                    original_creativo,
                    row_original['Duration'].values[0],
                    row_original['Brand'].values[0],
                    'Found'
                ])
            return pd.Series([value, 0, 'N/F', 'Not Found'])

    # Aplicar fuzzy matching
    df_main[['Creativo', 'Duracion', 'Brand', 'Estado']] = df_main['Versión'].apply(
        lambda x: get_best_match_info(x, df_aux_lower, df_side_aux_lower, original_mapping, side_original_mapping, threshold)
    )

    # Guardar resultados en el Excel original
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df_main[['Duracion']].to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=6, index=False)
        df_main[['Creativo']].to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=9, index=False)
        df_main[['Estado']].to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=10, index=False)
        df_main[['Brand']].to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=11, index=False)
    
# Función para generar el certificado final con todas las transformaciones aplicadas
def generar_certificado_final(aux_path, excel_path, final_path, log_func=None):
    sheet_name = 'Archivo Final Play Logger'

    if log_func: log_func("🧩 Obteniendo información del proveedor...")
    get_vendor(excel_path, aux_path, sheet_name)

    if log_func: log_func("📅 Dando formato a las fechas...")
    format_date(excel_path, sheet_name)

    if log_func: log_func("🕓 Dando formato a la columna de hora...")
    format_hour_column(excel_path, sheet_name)

    if log_func: log_func("📝 Llenando información de spots...")
    fill_spot_info(excel_path, sheet_name)

    if log_func: log_func("🎨 Obteniendo datos de creativos...")
    get_creatives_data(excel_path, aux_path, sheet_name)

    if log_func: log_func("✅ Certificado final generado correctamente.")