import pandas as pd
from datetime import datetime
import pytz
import dotenv, os
import pytz

dotenv.load_dotenv()


def create_datarev(excel_path, sheet_name):
    # Cargar los datos del archivo principal
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    df = pd.read_excel(excel_path, sheet_name=sheet_name, usecols=['Fecha', 'Horario'])

    # Asegurarse de que las columnas tengan el formato correcto (eliminar espacios extra)
    df.columns = df.columns.str.strip()

    # Convertir ambas columnas a datetime, especificando el formato correcto de la fecha
    df['Fecha'] = pd.to_datetime(df['Fecha'], format='%m/%d/%Y', errors='coerce')
    df['Horario'] = pd.to_datetime(df['Horario'].astype(str), format='%H:%M:%S', errors='coerce')

    # Concatenar fecha y hora en una sola columna en el formato deseado
    def format_date_rev(row):
        if pd.notnull(row['Fecha']) and pd.notnull(row['Horario']):
            return f"{row['Fecha'].strftime('%m/%d/%Y')} {row['Horario'].strftime('%H:%M:%S')}"
        else:
            return None

    df['Date Rev'] = df.apply(format_date_rev, axis=1)

    # Guardar los resultados en la columna L del archivo Excel
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df[['Date Rev']].to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=11, index=False)

    print(f"Columna 'Date Rev' creada y guardada en la columna L de la hoja {sheet_name}.")

def get_time_zone_dict(aux_path):
    
    df_zonas = pd.read_excel(aux_path, sheet_name='Zona Horaria', usecols=['Country', 'Time zone'])
    df_zonas.columns = df_zonas.columns.str.strip()  # Limpia espacios en los encabezados

    # Crea el diccionario a partir de las columnas 'country' y 'Time Zone'
    zonas_horarias_map = dict(zip(df_zonas['Country'], df_zonas['Time zone']))
    return zonas_horarias_map

def convert_time_venezuela(excel_path, aux_path, sheet_name):
    # Cargar los datos de los archivos Excel
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    aux_df = pd.read_excel(aux_path, sheet_name='Channel Info Monitoria')
    zonas_horarias_map = get_time_zone_dict(aux_path)

    # Crear una nueva columna para la conversión
    df['Date Time Zone'] = None

    # Iterar para realizar la conversión por fila
    for index, row in df.iterrows():
        try:
            vendor = row['Vendor']
            channel = row['Channel']
            feed = row['Feed']
            date_rev = row['Date Rev']

            # Obtener la zona horaria correspondiente
            time_zone_io = aux_df[
                (aux_df['Vendor'] == vendor) & 
                (aux_df['Channel'] == channel) & 
                (aux_df['Feed'] == feed)
            ]["Time Zone IO's"].values

            if len(time_zone_io) > 0:
                time_zone_io = time_zone_io[0]
                if time_zone_io in zonas_horarias_map:
                    # Realizar la conversión de la fecha y hora
                    original_time = pd.to_datetime(date_rev).tz_localize('America/Caracas')
                    converted_time = original_time.tz_convert(pytz.timezone(zonas_horarias_map[time_zone_io]))
                    df.at[index, 'Date Time Zone'] = converted_time.strftime('%m/%d/%Y %H:%M:%S')
        except Exception as e:
            print(f"Error en la fila {index}: {e}")
    
    # Guardar la nueva columna en el archivo Excel
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df[['Date Time Zone']].to_excel(writer, sheet_name='Archivo Final Play Logger', startrow=0, startcol=12, index=False)

    print(f"Columna 'Date Time Zone' creada y guardada correctamente.")

def sort_by_date_and_channel(excel_path, sheet_name):

    # Cargar los datos del archivo Excel
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    # Ordenar los datos por 'Channel' y 'Date Time Zone'
    df.sort_values(by=['Channel', 'Date Time Zone'], ascending=[True, True], inplace=True)
    
    # Guardar los resultados en el archivo Excel
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"Datos ordenados por 'Channel' y 'Date Time Zone' en la hoja {sheet_name}.")

def gen_date_plus_minutes(excel_path, aux_path, sheet_name):
    import pandas as pd
    from datetime import datetime
    
    # Leer los datos del archivo principal y el auxiliar
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    df_aux = pd.read_excel(aux_path, sheet_name='Channel Info Monitoria')

    # Combinar los archivos para obtener la columna 'Condition +/-'
    df_merged = df.merge(df_aux[['Vendor', 'Channel', 'Feed', 'Condition +/-']], 
                         on=['Vendor', 'Channel', 'Feed'], how='left')[['Condition +/-']]

    # Crear una nueva columna con la fecha ajustada por los minutos de 'Condition +/-'
    df['Date Time Zone + minutes'] = pd.to_datetime(df['Date Time Zone']) + pd.to_timedelta(df_merged['Condition +/-'], unit='m')

    # Convertir los minutos y segundos a 00
    df['Date Time Zone + minutes'] = df['Date Time Zone + minutes'].apply(lambda x: x.replace(minute=0, second=0))

    # Convertir la fecha al formato MM/DD/YYYY HH:MM:SS
    df['Date Time Zone + minutes'] = df['Date Time Zone + minutes'].dt.strftime('%m/%d/%Y %H:%M:%S')

    # Seleccionar la nueva columna
    df = df['Date Time Zone + minutes']

    # Guardar la nueva columna en el archivo Excel
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=13, index=False)
    
    print('Date Time Zone + minutes column created and saved in the Excel file.')

def gen_date_pminus_minutes(excel_path, aux_path, sheet_name):
    #Must get Condition +/- from aux_path, getting onli the value of the existing registers in the main file comparung with the aux file with vendor, channel and feed
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    df_aux = pd.read_excel(aux_path, sheet_name='Channel Info Monitoria')

    # Merge the main file with the aux file to get the 'Condition +/-' column
    df_merged = df.merge(df_aux[['Vendor', 'Channel', 'Feed', 'Condition +/-']], 
                         on=['Vendor', 'Channel', 'Feed'], how='left')[['Condition +/-']]

    # Create a new column with the date minus the minutes from 'Condition +/-', Keeping format MM/DD/YYYY HH:MM:SS, getting clear that de Condition +/- is a integer number of minutes and it will be converted
    df['Date Time Zone - minutes'] = pd.to_datetime(df['Date Time Zone']) - pd.to_timedelta(df_merged['Condition +/-'], unit='m')

    # Convertir los minutos y segundos a 00
    df['Date Time Zone - minutes'] = df['Date Time Zone - minutes'].apply(lambda x: x.replace(minute=0, second=0))

    #conver date from YY-MM-DD HH:MM:SS to MM/DD/YYYY HH:MM:SS
    df['Date Time Zone - minutes'] = df['Date Time Zone - minutes'].dt.strftime('%m/%d/%Y %H:%M:%S')

    df = df['Date Time Zone - minutes']
    #save the new column in the excel file
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=15, index=False)
    print('Date Time Zone - column created and saved in the excel file')

def gen_date_equal_minutes(excel_path, aux_path, sheet_name):
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    df['Date Time Zone = minutes'] = pd.to_datetime(df['Date Time Zone'])
    df['Date Time Zone = minutes'] = df['Date Time Zone = minutes'].apply(
        lambda x: x.replace(minute=0, second=0) if not pd.isnull(x) else x
    )
    df['Date Time Zone = minutes'] = df['Date Time Zone = minutes'].dt.strftime('%m/%d/%Y %H:%M:%S')

    df = df['Date Time Zone = minutes']

    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=14 ,index=False)
    print('Date Time Zone = minutes column created and saved in the excel file')

def create_time_plus(excel_path, aux_path):
    # Open Excel file on the sheet 'Archivo Final Play Logger'
    df = pd.read_excel(excel_path, sheet_name='Archivo Final Play Logger')
    df_aux = pd.read_excel(aux_path, sheet_name='Channel Info Monitoria')

    # Merge df with df_aux to include time range columns from aux file
    df_merged = df.merge(
        df_aux[['Vendor', 'Channel', 'Feed', 
                'Start - Madrugada', 'End - Madrugada', 
                'Start - Morning', 'End - Morning', 
                'Start - Afternoon', 'End - Afternoon', 
                'Start - Prime Time', 'End - Prime Time']], 
        on=['Vendor', 'Channel', 'Feed'], 
        how='left'
    )

    # Extract HH:MM:SS from the 'Date Time Zone + minutes' column
    df['Time'] = pd.to_datetime(df['Date Time Zone + minutes'], errors='coerce').dt.time

    # Add 'Time' column to df_merged for comparison
    df_temp = pd.concat([df_merged, df['Time']], axis=1)

    # Drop rows where 'Time' is NaN
    df_temp = df_temp.dropna(subset=['Time'])

    # Define a function to classify times into time ranges
    def classify_time(row):
        if row['Start - Madrugada'] <= row['Time'] <= row['End - Madrugada']:
            return 'Madrugada'
        elif row['Start - Morning'] <= row['Time'] <= row['End - Morning']:
            return 'Morning'
        elif row['Start - Afternoon'] <= row['Time'] <= row['End - Afternoon']:
            return 'Afternoon'
        elif row['Start - Prime Time'] <= row['Time'] <= row['End - Prime Time']:
            return 'Prime Time'
        return 'Other'

    # Apply the function to each row
    df_temp['Franja + minutes'] = df_temp.apply(classify_time, axis=1)

    # Add the 'Franja + minutes' column back to the original dataframe
    df = df_temp['Franja + minutes']

    #save the new column in the excel file
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name='Archivo Final Play Logger', startrow=0, startcol=16, index=False)
    
    print('Franja + minutes column created and saved in the excel file')

def create_time_equal(excel_path, aux_path):
    # Open Excel file on the sheet 'Archivo Final Play Logger'
    df = pd.read_excel(excel_path, sheet_name='Archivo Final Play Logger')
    df_aux = pd.read_excel(aux_path, sheet_name='Channel Info Monitoria')

    # Merge df with df_aux to include time range columns from aux file
    df_merged = df.merge(
        df_aux[['Vendor', 'Channel', 'Feed', 
                'Start - Madrugada', 'End - Madrugada', 
                'Start - Morning', 'End - Morning', 
                'Start - Afternoon', 'End - Afternoon', 
                'Start - Prime Time', 'End - Prime Time']], 
        on=['Vendor', 'Channel', 'Feed'], 
        how='left'
    )

    # Extract HH:MM:SS from the 'Date Time Zone + minutes' column
    df['Time'] = pd.to_datetime(df['Date Time Zone = minutes'], errors='coerce').dt.time

    # Add 'Time' column to df_merged for comparison
    df_temp = pd.concat([df_merged, df['Time']], axis=1)

    # Drop rows where 'Time' is NaN
    df_temp = df_temp.dropna(subset=['Time'])

    # Define a function to classify times into time ranges
    def classify_time(row):
        if row['Start - Madrugada'] <= row['Time'] <= row['End - Madrugada']:
            return 'Madrugada'
        elif row['Start - Morning'] <= row['Time'] <= row['End - Morning']:
            return 'Morning'
        elif row['Start - Afternoon'] <= row['Time'] <= row['End - Afternoon']:
            return 'Afternoon'
        elif row['Start - Prime Time'] <= row['Time'] <= row['End - Prime Time']:
            return 'Prime Time'
        return 'Other'

    # Apply the function to each row
    df_temp['Franja = minutes'] = df_temp.apply(classify_time, axis=1)

    # Add the 'Franja + minutes' column back to the original dataframe
    df = df_temp['Franja = minutes']

    #save the new column in the excel file
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name='Archivo Final Play Logger', startrow=0, startcol=17, index=False)
    
    print('Franja = minutes column created and saved in the excel file')

def create_time_minus(excel_path, aux_path):
    # Open Excel file on the sheet 'Archivo Final Play Logger'
    df = pd.read_excel(excel_path, sheet_name='Archivo Final Play Logger')
    df_aux = pd.read_excel(aux_path, sheet_name='Channel Info Monitoria')

    # Merge df with df_aux to include time range columns from aux file
    df_merged = df.merge(
        df_aux[['Vendor', 'Channel', 'Feed', 
                'Start - Madrugada', 'End - Madrugada', 
                'Start - Morning', 'End - Morning', 
                'Start - Afternoon', 'End - Afternoon', 
                'Start - Prime Time', 'End - Prime Time']], 
        on=['Vendor', 'Channel', 'Feed'], 
        how='left'
    )

    # Extract HH:MM:SS from the 'Date Time Zone + minutes' column
    df['Time'] = pd.to_datetime(df['Date Time Zone - minutes'], errors='coerce').dt.time

    # Add 'Time' column to df_merged for comparison
    df_temp = pd.concat([df_merged, df['Time']], axis=1)

    # Drop rows where 'Time' is NaN
    df_temp = df_temp.dropna(subset=['Time'])

    # Define a function to classify times into time ranges
    def classify_time(row):
        if row['Start - Madrugada'] <= row['Time'] <= row['End - Madrugada']:
            return 'Madrugada'
        elif row['Start - Morning'] <= row['Time'] <= row['End - Morning']:
            return 'Morning'
        elif row['Start - Afternoon'] <= row['Time'] <= row['End - Afternoon']:
            return 'Afternoon'
        elif row['Start - Prime Time'] <= row['Time'] <= row['End - Prime Time']:
            return 'Prime Time'
        return 'Other'

    # Apply the function to each row
    df_temp['Franja - minutes'] = df_temp.apply(classify_time, axis=1)

    # Add the 'Franja + minutes' column back to the original dataframe
    df = df_temp['Franja - minutes']

    #save the new column in the excel file
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name='Archivo Final Play Logger', startrow=0, startcol=18, index=False)
    
    print('Franja - minutes column created and saved in the excel file')

def Generate_additional_columns(excel_path, aux_path, start_date, end_date, Month_dict):

    sheet_name = 'Archivo Final Play Logger'

    create_datarev(excel_path, sheet_name)
    convert_time_venezuela(excel_path, aux_path, sheet_name)
    gen_date_plus_minutes(excel_path, aux_path, sheet_name)
    gen_date_equal_minutes(excel_path, aux_path, sheet_name)
    gen_date_pminus_minutes(excel_path, aux_path, sheet_name)
    create_time_plus(excel_path, aux_path)
    create_time_equal(excel_path, aux_path)
    create_time_minus(excel_path, aux_path)
    sort_by_date_and_channel(excel_path, sheet_name)

    start_date = datetime.strptime(start_date, '%m/%d/%Y')
    end_date = datetime.strptime(end_date, '%m/%d/%Y')

    start_date = start_date.strftime('%B %d')
    end_date = end_date.strftime('%d %Y')
    month = Month_dict[start_date.split()[0]]
    day = start_date.split()[1]
    end_day = end_date.split()[0]
    year = end_date.split()[1]
    output_path = f'Archivo Final Play Logger {start_date} to {end_day} {year}.xlsx'

    final_path = f'./certs/{output_path}'

    #save the final file
    with pd.ExcelWriter(final_path, engine='openpyxl') as writer:
        pd.DataFrame().to_excel(writer, sheet_name='Archivo Final Play Logger', index=False)

    print(f"Archivo guardado en {final_path}")