import pandas as pd
from datetime import datetime
import dotenv
import os

dotenv.load_dotenv()

def crear_columna_date_time_zone(df):
    """
    Crea la columna 'Date Time Zone' unificando la fecha y la hora en formato 'mm/dd/yyyy hh:mm:ss'.
    """
    # Especificamos el formato para evitar la advertencia
    df['Date Time Zone'] = pd.to_datetime(
        df['Date'] + ' ' + df['Hour'], 
        format='%m/%d/%y %H:%M:%S'
    ).dt.strftime('%m/%d/%Y %H:%M:%S')
    return df

def crear_columna_date_full_day(df):
    """
    Crea la columna 'Date Full Day' tomando solo la fecha de 'Date Time Zone' y fijando la hora en '00:00'.
    """
    df['Date Full Day'] = pd.to_datetime(df['Date Time Zone']).dt.strftime('%m/%d/%Y') + ' 00:00'
    return df

def organizar_por_canal_y_fecha(df):
    """
    Ordena el DataFrame por las columnas 'Channel' y 'Date Time Zone' de menor a mayor.
    """
    df.sort_values(by=['Channel', 'Date Time Zone'], ascending=[True, True], inplace=True)
    return df

def split_repeated_spots(df):
    # Lista para almacenar las nuevas filas
    nuevos_registros = []
    
    # Iterar sobre las filas del DataFrame
    for _, fila in df.iterrows():
        spot = fila['Spots']
        
        # Si el spot es mayor que 1, agregamos las copias
        if spot > 1:
            for _ in range(spot - 1):  # Se crea 'spot' copias (incluyendo la original)
                nueva_fila = fila.copy()
                nuevos_registros.append(nueva_fila)
    
    # Crear el DataFrame final con los nuevos registros
    df_ajustado = pd.DataFrame(nuevos_registros)
    
    # Concatenar el DataFrame original con el ajustado
    df = pd.concat([df, df_ajustado], ignore_index=True)
    df['Spots'] = 1  # Ajustar 'spot' a 1 en todas las copias


    return df

def process_and_filter_data(bdd_path, aux_path, start_date, end_date, Month_dict):
    """
    Procesa los datos en el archivo bdd_path, filtra por VEN=1, agrega columnas 'Date Time Zone' y 'Date Full Day',
    organiza los datos, y guarda el resultado en un nuevo archivo Excel.
    """

    # Cargar datos de la hoja 'BDD Final'
    df_bdd = pd.read_excel(bdd_path, sheet_name='BDD Final', skiprows=1)  # Saltar a partir de la fila 2

    # Limpiar encabezados de columnas
    df_bdd.columns = df_bdd.columns.str.strip()

    # Filtrar por registros donde VEN = 1
    df_bdd = df_bdd[df_bdd['VEN'] == 1]

    #Convert string date input with format MM/DD/YYYY into date type in format MM/DD/YYYY
    start_date = datetime.strptime(start_date, '%m/%d/%Y')
    end_date = datetime.strptime(end_date, '%m/%d/%Y')

    # Convertir 'Date' a formato datetime para la comparación
    df_bdd['Date'] = pd.to_datetime(df_bdd['Date'], format='%m/%d/%y')
    
    # Filtrar por registros desde la fecha inicial hasta la fecha final
    df_bdd = df_bdd[(df_bdd['Date'] >= start_date) & (df_bdd['Date'] <= end_date)]

    # Convertir 'Date' a formato MM/DD/YY sin la hora y 'Hour' a cadena
    df_bdd['Date'] = pd.to_datetime(df_bdd['Date'], format='%B %d %Y').dt.strftime('%m/%d/%y')
    df_bdd['Hour'] = df_bdd['Hour'].astype(str)

    # Agregar columnas adicionales, ajustar spots multiples y organizar datos
    df_bdd = crear_columna_date_time_zone(df_bdd)
    df_bdd = crear_columna_date_full_day(df_bdd)
    df_bdd = split_repeated_spots(df_bdd)
    df_bdd = organizar_por_canal_y_fecha(df_bdd)

    # Seleccionar columnas requeridas desde la hoja Index Tablas de aux_path columna Monitoring_db_Index
    df_index = pd.read_excel(aux_path, sheet_name='Index Tablas')
    #Extraer columna Monitoring_db_Index y convertir a lista
    columnas_requeridas = df_index['Monitoring_db_Index'].tolist()
    
    df_bdd_filtrado = df_bdd[columnas_requeridas]

    # Guardar el archivo en un nuevo Excel en la carpeta raíz
    # Generar archivo Excel bajo el nombre ‘BDD Pauta MMM Dd to DD YYYY’ en formato .xlsx

    start_date = start_date.strftime('%B %d')
    end_date = end_date.strftime('%d %Y')
    end_day = end_date.split()[0]
    year = end_date.split()[1]
    output_path = f'BDD Pauta {start_date} to {end_day} {year}.xlsx'

    df_bdd_filtrado.to_excel(output_path, index=False)
    print(f"Archivo guardado en {output_path}")