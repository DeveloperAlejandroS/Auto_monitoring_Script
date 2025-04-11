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
    """
    Divide los registros con 'Spots' mayores a 1 en múltiples registros con 'Spots' igual a 1.
    """
    
     # Asegurarse de que 'Spots' es numérico
    df['Spots'] = pd.to_numeric(df['Spots'], errors='coerce')
    
    # Eliminar filas donde 'Spots' no es un número válido (NaN)
    df = df.dropna(subset=['Spots'])

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

def process_and_filter_data(full_bdd_path, aux_path, base_file, filtered_bdd_file, start_date, end_date, log_func=None):
    """Procesa y filtra los datos de la BDD completa según el rango de fechas y el Feed Index."""
    
    if log_func: log_func("📥 Cargando archivo base y extrayendo Feed Index...")
    df = pd.read_excel(base_file, sheet_name='Convertido y procesado')
    unique_feed_index = df['Feed Index'].dropna().unique()

    if log_func: log_func("📥 Cargando archivo BDD completo y limpiando encabezados...")
    df_bdd = pd.read_excel(full_bdd_path, sheet_name='BDD Final', skiprows=1, dtype=str)
    df_bdd.columns = df_bdd.columns.str.strip()

    if log_func: log_func("🔍 Filtrando por Feed Index común...")
    df_bdd = df_bdd[df_bdd['Feed Index'].isin(unique_feed_index)]

    if log_func: log_func("📅 Procesando fechas para filtrado...")
    start_date = datetime.strptime(start_date, '%m/%d/%Y')
    end_date = datetime.strptime(end_date, '%m/%d/%Y')
    df_bdd['Date'] = pd.to_datetime(df_bdd['Date'], format='%Y-%m-%d %H:%M:%S')
    df_bdd = df_bdd[(df_bdd['Date'] >= start_date) & (df_bdd['Date'] <= end_date)]

    if log_func: log_func("🗂️ Formateando fechas y horas...")
    df_bdd['Date'] = pd.to_datetime(df_bdd['Date'], format='%B %d %Y').dt.strftime('%m/%d/%y')
    df_bdd['Hour'] = df_bdd['Hour'].astype(str)

    if log_func: log_func("➕ Agregando columnas adicionales y organizando datos...")
    df_bdd = crear_columna_date_time_zone(df_bdd)
    df_bdd = crear_columna_date_full_day(df_bdd)
    df_bdd = split_repeated_spots(df_bdd)
    df_bdd = organizar_por_canal_y_fecha(df_bdd)

    if log_func: log_func("📊 Seleccionando columnas requeridas desde archivo auxiliar...")
    df_index = pd.read_excel(aux_path, sheet_name='Index Tablas')
    columnas_requeridas = df_index['Monitoring_db_Index'].dropna().tolist()
    df_bdd_filtrado = df_bdd[columnas_requeridas]

    if log_func: log_func("💾 Guardando archivo filtrado...")
    df_bdd_filtrado.to_excel(filtered_bdd_file, index=False)
    if log_func: log_func(f"✅ Archivo guardado")