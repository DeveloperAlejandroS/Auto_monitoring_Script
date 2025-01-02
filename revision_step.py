import pandas as pd
from datetime import timedelta
from openpyxl import load_workbook
import time

def delete_outdated_rows(final_path, start_date, end_date, sheet_name):
    # Leer la hoja específica del archivo Excel
    df = pd.read_excel(final_path, sheet_name=sheet_name)

    # Convertir 'Date Time Zone' a formato datetime (incluyendo hora)
    df['Date Time Zone'] = pd.to_datetime(df['Date Time Zone'])

    # Crear una columna auxiliar con solo la fecha (sin hora)
    df['aux_date'] = df['Date Time Zone'].dt.date

    # Convertir start_date y end_date a formato datetime.date
    start_date = pd.to_datetime(start_date).date()
    end_date = pd.to_datetime(end_date).date()

    # Filtrar filas dentro del rango de fechas
    df = df[(df['aux_date'] >= start_date) & (df['aux_date'] <= end_date)]

    # Eliminar la columna auxiliar
    df.drop(columns=['aux_date'], inplace=True)
    
    df['Date Time Zone'] = pd.to_datetime(df['Date Time Zone']).dt.strftime('%m/%d/%Y %H:%M:%S')

    # Guardar los datos actualizados en la hoja específica
    with pd.ExcelWriter(final_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

def b2bV2(final_path, sheet_name):
    # Leer archivo Excel y hoja
    df = pd.read_excel(final_path, sheet_name=sheet_name)

    # Convertir 'Date Time Zone' a tipo datetime
    df['Date Time Zone'] = pd.to_datetime(df['Date Time Zone'], errors='coerce')

    # Ordenar datos por 'Feed Index' y 'Date Time Zone'
    df.sort_values(by=['Feed Index', 'Date Time Zone'], inplace=True)
    df.reset_index(drop=True, inplace=True)

    # Inicializar columna 'Back to back'
    df['Back to back'] = ''

    # Variables para almacenar los valores previos
    previous_feed_index = None
    previous_date = None
    previous_duration = None

    # Iterar sobre las filas del DataFrame
    for idx, row in df.iterrows():
        current_feed_index = row['Feed Index']
        current_date = row['Date Time Zone']
        current_duration = row['Duracion']

        # Validar datos nulos o inválidos
        if pd.isna(current_date) or pd.isna(current_duration):
            df.at[idx, 'Back to back'] = 'Error: Datos incompletos'
            continue
        if not isinstance(current_duration, (int, float)):
            df.at[idx, 'Back to back'] = 'Error: Duración inválida'
            continue

        if idx == 0:  # Primera fila
            df.at[idx, 'Back to back'] = 'Ok'
        else:
            # Calcular la fecha límite para "Back to back"
            date_with_seconds = previous_date + timedelta(seconds=previous_duration + 2)

            # Verificar si cumple la condición de "Back to back"
            if current_feed_index == previous_feed_index and current_date <= date_with_seconds:
                df.at[idx, 'Back to back'] = 'Back to back'
                print(f'Fila {idx}: Back to back - Fecha previa: {previous_date}, Fecha actual: {current_date}')
            else:
                df.at[idx, 'Back to back'] = 'Ok'

        # Actualizar las variables previas
        previous_feed_index = current_feed_index
        previous_date = current_date
        previous_duration = current_duration

    # Guardar el DataFrame actualizado en el archivo Excel
    with pd.ExcelWriter(final_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

        
def back_to_back_rev(final_path, sheet_name):
    df = pd.read_excel(final_path, sheet_name=sheet_name)

    #get Date Time Zone column from final_path
    df['Date Time Zone'] = pd.to_datetime(df['Date Time Zone'])
    
    df['Back to back'] = ''
    
    df.sort_values(by=['Feed Index', 'Date Time Zone'], inplace=True)
    
    for i in range(len(df)):
        if i == 0:
            df.loc[i, 'Back to back'] = 'Ok'
            
        else:
            current_feed_index = df.loc[i, 'Feed Index']
            previous_feed_index = df.loc[i-1, 'Feed Index']
            
            current_date_time_zone = df.loc[i, 'Date Time Zone']
            previous_date_time_zone = df.loc[i-1, 'Date Time Zone']
                  
            if pd.isna(df.loc[i,'Duracion']):
                df.loc[i, 'Back to back'] = 'Ok'
            else:
                duration_seconds = timedelta(seconds=int(df.loc[i-1, 'Duracion'])) + timedelta(seconds=2)
                      
                if current_feed_index != previous_feed_index:
                    df.loc[i, 'Back to back'] = 'Ok'
                else:
                    previous_plus_seconds = previous_date_time_zone + duration_seconds
                    
                    print(current_date_time_zone)
                    print(previous_date_time_zone)
                    print(previous_plus_seconds)
                    
                    if current_date_time_zone <= previous_plus_seconds:
                        df.loc[i, 'Back to back'] = 'Back to back'
                    else:
                        df.loc[i, 'Back to back'] = 'Ok'
    
    df = df['Back to back']
    
    #add df to Z column in final_path, in the oly one sheet that is already there
    with pd.ExcelWriter(final_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=28, index=False, header=True)
    print('Back to back revision completed')
        
def rev_spots_vs_pauta(final_path, filtered_bdd_path, sheet_name):            
    # Leer los datos desde los archivos Excel
    df = pd.read_excel(final_path, sheet_name=sheet_name)
    df_bdd = pd.read_excel(filtered_bdd_path)
    
    # Crear y preparar columnas necesarias
    df_bdd['Spot status'] = ''  # Inicializar el estado de los spots
    df['Revision type'] = pd.to_numeric(df['Revision type'], errors='coerce').fillna(0).astype(int)
    df['Rev vs pauta'] = ''
    df['Spot Observation'] = ''
    
    # Asegurarse de que las fechas están en formato datetime
    df_bdd['Date Time Zone'] = pd.to_datetime(df_bdd['Date Time Zone'], format='%m/%d/%Y %H:%M:%S')
    
    for idx, row in df.iterrows():
        current_feed_index = row['Feed Index']
        current_brand = row['Brand']
        
        # Filtrar registros correspondientes en df_bdd
        matches = df_bdd[(df_bdd['Feed Index'] == current_feed_index) & (df_bdd['Brand'] == current_brand)]
        
        if not matches.empty:
            for idx2, row2 in matches.iterrows():
                # Verificar si el registro ya fue procesado
                if row2['Spot status'] == 'Ok':
                    continue
                
                if row['Revision type'] == 1:
                    # Procesamiento para "Revision type" 1
                    current_minus_date = pd.to_datetime(row['Date Time Zone - Minutes'], errors='coerce')
                    current_equal_date = pd.to_datetime(row['Date Time Zone = Minutes'], errors='coerce')
                    current_plus_date = pd.to_datetime(row['Date Time Zone + Minutes'], errors='coerce')
                    comparer_date = pd.to_datetime(row2['Date Time Zone'], errors='coerce')
                    
                    if current_minus_date == comparer_date:
                        df.at[idx, 'Rev vs pauta'] = f'Ok - {current_minus_date.strftime("%m/%d/%Y %H:%M:%S")}'
                        df.at[idx, 'Spot Observation'] = 'Spot Correcto'
                        df_bdd.at[idx2, 'Spot status'] = 'Ok'
                        break
                    if current_equal_date == comparer_date:
                        df.at[idx, 'Rev vs pauta'] = f'Ok - {current_equal_date.strftime("%m/%d/%Y %H:%M:%S")}'
                        df.at[idx, 'Spot Observation'] = 'Spot Correcto'
                        df_bdd.at[idx2, 'Spot status'] = 'Ok'
                        break
                    if current_plus_date == comparer_date:
                        df.at[idx, 'Rev vs pauta'] = f'Ok - {current_plus_date.strftime("%m/%d/%Y %H:%M:%S")}'
                        df.at[idx, 'Spot Observation'] = 'Spot Correcto'
                        df_bdd.at[idx2, 'Spot status'] = 'Ok'
                        break
                elif row['Revision type'] == 2:
                    # Procesamiento para "Revision type" 2
                    current_minus_dp = row['Day Part - Minutes'].strip().lower()
                    current_equal_dp = row['Day Part = Minutes'].strip().lower()
                    current_plus_dp = row['Day Part + Minutes'].strip().lower()
                    
                    current_minus_full_day = pd.to_datetime(row['Full Day - Minutes'], errors='coerce')
                    current_equal_full_day = pd.to_datetime(row['Full Day = Minutes'], errors='coerce')
                    current_plus_full_day = pd.to_datetime(row['Full Day + Minutes'], errors= 'coerce')
                    
                    comparer_date_full_day = pd.to_datetime(row2['Date Full Day'], errors='coerce')
                    comparer_day_part = row2['Franja'].strip().lower()
                    
                    if current_minus_dp == comparer_day_part and current_minus_full_day == comparer_date_full_day:
                        df.at[idx, 'Rev vs pauta'] = f'Ok - {current_minus_dp}'
                        df.at[idx, 'Spot Observation'] = 'Spot Correcto'
                        df_bdd.at[idx2, 'Spot status'] = 'Ok'
                        break
                    
                    if current_equal_dp == comparer_day_part and current_equal_full_day == comparer_date_full_day:
                        df.at[idx, 'Rev vs pauta'] = f'Ok - {current_equal_dp}'
                        df.at[idx, 'Spot Observation'] = 'Spot Correcto'
                        df_bdd.at[idx2, 'Spot status'] = 'Ok'
                        break
                    if current_plus_dp == comparer_day_part and current_plus_full_day == comparer_date_full_day:
                        df.at[idx, 'Rev vs pauta'] = f'Ok - {current_plus_dp}'
                        df.at[idx, 'Spot Observation'] = 'Spot Correcto'
                        df_bdd.at[idx2, 'Spot status'] = 'Ok'
                        break

                elif row['Revision type'] == 3:
                    # Procesamiento para "Revision type" 3
                    current_minus_full_day = pd.to_datetime(row['Full Day - Minutes'], format='%m/%d/%Y %H:%M')
                    current_equal_full_day = pd.to_datetime(row['Full Day = Minutes'], format='%m/%d/%Y %H:%M')
                    current_plus_full_day = pd.to_datetime(row['Full Day + Minutes'], format='%m/%d/%Y %H:%M')
                    comparer_date_full_day = pd.to_datetime(row2['Date Full Day'], errors='coerce')
                    
                    if current_minus_full_day == comparer_date_full_day:
                        df.at[idx, 'Rev vs pauta'] = f'Ok - {current_minus_full_day.strftime("%m/%d/%Y %H:%M")}'
                        df.at[idx, 'Spot Observation'] = 'Spot Correcto'
                        df_bdd.at[idx2, 'Spot status'] = 'Ok'
                        break
                    elif current_equal_full_day == comparer_date_full_day:
                        df.at[idx, 'Rev vs pauta'] = f'Ok - {current_equal_full_day.strftime("%m/%d/%Y %H:%M")}'
                        df.at[idx, 'Spot Observation'] = 'Spot Correcto'
                        df_bdd.at[idx2, 'Spot status'] = 'Ok'
                        break
                    elif current_plus_full_day == comparer_date_full_day:
                        df.at[idx, 'Rev vs pauta'] = f'Ok - {current_plus_full_day.strftime("%m/%d/%Y %H:%M")}'
                        df.at[idx, 'Spot Observation'] = 'Spot Correcto'
                        df_bdd.at[idx2, 'Spot status'] = 'Ok'
                        break
        
        # Si no se encontró un registro válido
        if df.at[idx, 'Rev vs pauta'] == '':
            df.at[idx, 'Rev vs pauta'] = 'No'
            df.at[idx, 'Spot Observation'] = 'Spot Incorrecto - No encontrado en pauta'
    
    df_bdd['Spot status'] = df_bdd['Spot status'].replace('', 'No')
    
    wb = load_workbook(final_path)
    sn = 'BDD Final Revisada'
    
    if sn in wb.sheetnames:
        wb.remove(wb[sn])
    
    ws = wb.create_sheet(title=sn)
    
    df_bdd['Date Time Zone']= df_bdd['Date Time Zone'].dt.strftime('%m/%d/%Y %H:%M:%S')
    df['Date Time Zone'] = df['Date Time Zone'].dt.strftime('%m/%d/%Y %H:%M:%S')
    
    with pd.ExcelWriter(final_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        df_bdd.to_excel(writer, sheet_name=sn, index=False)
                                      
def full_revision(final_path, filtered_bdd_path, start_date, end_date, sheet_name):
    
    start_time = time.time()

    delete_outdated_rows(final_path, start_date, end_date, sheet_name)
    print('Starting back to back revision')
    b2bV2(final_path, sheet_name)
    print('Back to back revision completed')
    print('Starting Spot revision')
    rev_spots_vs_pauta(final_path, filtered_bdd_path, sheet_name)
    print('Spot revision completed')
    final_time = time.time() - start_time
    print(f'Time elapsed: {final_time} seconds')