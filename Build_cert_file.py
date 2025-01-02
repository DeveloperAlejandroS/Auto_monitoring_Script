import os
import pandas as pd
from openpyxl import load_workbook
from fuzzywuzzy import fuzz, process
from datetime import datetime
import pytz
from timezonefinder import TimezoneFinder

# Función para cruzar información del proveedor entre archivos y obtener "Vendor" para cada "Estación"
def get_vendor(excel_path, aux_path, sheet_name):
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
def formatDate(excel_path, sheet_name):
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

    df = pd.read_excel(excel_path, usecols='F', names=['Horario'])
    
    # Convertir a tiempo en formato 24 horas
    df['Horario'] = pd.to_datetime(df['Horario'], format='%H:%M:%S', errors='coerce').dt.time

    # Guardar la columna formateada en el archivo
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=5, index=False)

# Función para agregar información de cantidad (CANTIDAD) y tipo de spot (Type Spot)
def fill_spot_info(excel_path, sheet_name):

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
    df_main = pd.read_excel(excel_path)
    df_aux = pd.read_excel(aux_path, sheet_name=aux_sheet_name)
    threshold = 91  # Similarity threshold (91%)

    # Create a dictionary mapping lowercase to original case-sensitive values
    original_mapping = {row['Creativo'].lower(): row['Creativo'] for _, row in df_aux.iterrows()}

    # Clean the "Versión" column in the main DataFrame
    df_main['Versión'] = df_main['Versión'].replace(r'\(20s\)', '', regex=True)
    
    # Convert the auxiliary DataFrame to lowercase for comparison
    df_aux_lower = df_aux.copy()
    df_aux_lower['Creativo'] = df_aux['Creativo'].str.lower()

    # Function to perform fuzzy matching
    def get_best_match_info(value, choices_lower, original_mapping, threshold):
        match, score = process.extractOne(value, choices_lower['Creativo'].tolist())
        if score >= threshold:
            # Retrieve the original value using the dictionary
            original_creativo = original_mapping.get(match, match)
            row_original = choices_lower[choices_lower['Creativo'] == match]
            
            print(f"Matched: {value} with -> {match} (Score: {score})")
            return pd.Series([
                original_creativo,  # Original case-sensitive value
                row_original['Duration'].values[0],
                row_original['Brand'].values[0],
                'Found'
            ])
        return pd.Series([value, None, None, 'Not Found'])

    # Apply the fuzzy matching to the "Versión" column
    df_main[['Creativo', 'Duracion', 'Brand', 'Estado']] = df_main['Versión'].apply(
        lambda x: get_best_match_info(x, df_aux_lower, original_mapping, threshold)
    )

    # Save the results back to the Excel file
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df_main[['Duracion']].to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=6, index=False)
        df_main[['Creativo']].to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=9, index=False)
        df_main[['Estado']].to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=10, index=False)
        df_main[['Brand']].to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=11, index=False)

# Función para generar el certificado final con todas las transformaciones aplicadas
def generar_certificado_final(aux_path, excel_path, final_path):

    #HOJAS EXCEL
    sheet_name='Archivo Final Play Logger'

    get_vendor(excel_path, aux_path, sheet_name)
    formatDate(excel_path, sheet_name)
    format_hour_column(excel_path, sheet_name)
    fill_spot_info(excel_path, sheet_name)
    get_creatives_data(excel_path, aux_path, sheet_name)