import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment, NamedStyle, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
import os

def generate_report(aux_path):
    
    print('creando columnas de reporte resumen')
        
    aux_data_df = pd.read_excel(aux_path, sheet_name='Index Tablas')
    
    report_resumen_array = aux_data_df['Report_Index'].dropna().to_list()
    report_datails_array = aux_data_df['Report_details_index'].dropna().to_list()
    rotation_report_array = aux_data_df['Rotation_report_index'].dropna().to_list()
    required_columns_data = aux_data_df['Required_final_columns'].dropna().to_list()
    
    print('columnas de reporte resumen creadas')
    
    return report_resumen_array, report_datails_array, rotation_report_array, required_columns_data
    
def insert_data(final_path, required_columns_data, final_report_file):
    print('Moviendo datos a reporte final')

    # Cargar los datos de las hojas necesarias
    data_df = pd.read_excel(final_path, sheet_name='Archivo Final Play Logger', converters={'Id Rev % Ads': str})
    bdd_data_df = pd.read_excel(final_path, sheet_name='BDD Final Revisada')

    # Filtrar columnas necesarias
    data_df = data_df[required_columns_data]

    sheet_name = 'Detalle Revision'
    bdd_sheet_name = 'BDD Revision'

    if not os.path.exists(final_report_file):
        # Si el archivo no existe, crearlo con ambas hojas
        with pd.ExcelWriter(final_report_file, engine='openpyxl') as writer:
            data_df.to_excel(writer, sheet_name=sheet_name, index=False)
            bdd_data_df.to_excel(writer, sheet_name=bdd_sheet_name, index=False)
        print(f"Archivo creado: {final_report_file} con hojas {sheet_name} y {bdd_sheet_name}")
    else:
        # Si el archivo ya existe, asegurarnos de que las hojas existan antes de escribir
        with pd.ExcelWriter(final_report_file, engine='openpyxl', mode='a') as writer:
            # Cargar el archivo para verificar sus hojas
            workbook = load_workbook(final_report_file)

            # Si las hojas existen, eliminarlas antes de escribir
            if sheet_name in workbook.sheetnames:
                del workbook[sheet_name]
            if bdd_sheet_name in workbook.sheetnames:
                del workbook[bdd_sheet_name]
            workbook.save(final_report_file)

            # Ahora escribir los nuevos datos
            data_df.to_excel(writer, sheet_name=sheet_name, index=False)
            bdd_data_df.to_excel(writer, sheet_name=bdd_sheet_name, index=False)
        
        print(f"Datos actualizados en {final_report_file}")
    
    # Verificar hojas finales
    print(f"Hojas definitivas del archivo final: {pd.ExcelFile(final_report_file).sheet_names}")

def get_vendor_mapping():
    vendor_mapping = {
            'CC MEDIOS USA LLC': 'CC Medios',
            'NBCUniversal Networks International Spanish Latin America LLC': 'NBC Universal',
            'Sony Pictures Television Advertising Sales Company': 'Sony Pictures',
            'VC MEdios Latin America, LLC   (IntL)': 'VC Medios',
            'INVERCORP LIMITED': 'Invercorp',
            'Buena Vista International, INC': 'Buena Vista International'
        }
    return vendor_mapping

def apply_styles_to_sheet(workbook, sheet_name, table_positions):
    """Aplica estilos a las tablas dentro de una hoja específica, según las posiciones dadas."""
    ws = workbook[sheet_name]
    
    # Estilos para encabezados
    resumen_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")  # Fondo negro
    resumen_font = Font(color="FFFFFF", bold=True)  # Texto blanco
    detalles_fill = PatternFill(start_color="D35400", end_color="D35400", fill_type="solid")  # Fondo naranja oscuro
    detalles_font = Font(color="FFFFFF", bold=True)  # Texto blanco

    # Estilos para registros
    registros_texto = Font(color="000000")  # Texto negro
    registros_alineacion = Alignment(horizontal="center", vertical="center")  # Centrado
    registros_fondo_blanco = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # Fondo blanco

    # Formatos numéricos
    numero_format = "#,##0.00"  # Formato con miles y 2 decimales
    usd_format = '"$"#,##0.00'  # Formato de moneda USD

    # Aplicar estilos a cada tabla en la hoja
    for start_row, end_row in table_positions:
        # Aplicar estilos al encabezado (Fila start_row, desde columna B hasta antes de U)
        for col_num, cell in enumerate(ws[start_row], 1):  
            col_letter = get_column_letter(col_num)
            if "B" <= col_letter < "T":  
                cell.fill = resumen_fill
                cell.font = resumen_font

        # Aplicar estilos al encabezado de Detalles (Fila start_row, desde columna U en adelante)
        for col_num, cell in enumerate(ws[start_row], 1):  
            col_letter = get_column_letter(col_num)
            if col_letter >= "U":  
                cell.fill = detalles_fill
                cell.font = detalles_font

        # Aplicar estilos a los registros de la tabla de Resumen
        for row in ws.iter_rows(min_row=start_row + 1, max_row=end_row, min_col=2, max_col=19):
            for cell in row:
                cell.font = registros_texto
                cell.fill = registros_fondo_blanco  # Fondo blanco

        # Aplicar estilos a los registros de la tabla de Detalles
        for row in ws.iter_rows(min_row=start_row + 1, max_row=end_row, min_col=21, max_col=ws.max_column):
            for cell in row:
                cell.font = registros_texto
                cell.alignment = registros_alineacion
                cell.fill = registros_fondo_blanco  # Fondo blanco

        # Aplicar formato numérico y moneda solo en las columnas correspondientes
        for row in ws.iter_rows(min_row=start_row + 1, max_row=end_row):
            for cell in row:
                if cell.value and isinstance(cell.value, (int, float)):
                    col_letter = get_column_letter(cell.column)
                    if col_letter == "P":  # P = Formato numérico con miles
                        cell.number_format = numero_format
                    elif col_letter == "S":  # S = Formato de moneda USD
                        cell.number_format = usd_format

def generate_columns(final_report_file, report_resumen_array, report_details_array, rotation_report_array):
    print('Generando columnas de resumen con datos de reporte final')
    # Crear DataFrames vacíos con los encabezados requeridos
    df_details = pd.DataFrame(columns=report_details_array)
    df_rotation = pd.DataFrame(columns=rotation_report_array)
    
    #imprimir las hoja del  archivo final
    print(f'Las hojas del archivo final son: {pd.ExcelFile(final_report_file).sheet_names}')
    
    # Leer datos de la hoja "Detalle Revision"
    df_data = pd.read_excel(final_report_file, sheet_name='Detalle Revision')
    df_bdd = pd.read_excel(final_report_file, sheet_name='BDD Revision')
    
    unique_vendor_array = df_data['Vendor'].dropna().unique()
    vendor_mapping = get_vendor_mapping()
    
    def create_vendor_sheet(unique_vendor_array, final_report_file, vendor_mapping):
        print('Creando hojas de resumen por proveedor')
        wb = load_workbook(final_report_file)
        
        for vendor in unique_vendor_array:
            vendor_name = vendor_mapping.get(vendor, vendor)
            if vendor_name in wb.sheetnames:
                del wb[vendor_name]  # Elimina la hoja si ya existe
            wb.create_sheet(vendor_name)  # Crea la nueva hoja
        
        wb.save(final_report_file)
        print('Hojas de resumen por proveedor creadas')

    def generate_vendor_resume(unique_vendor_array, df_data, BDD_file, final_report_file, vendor_mapping):
        print('Generando resumen por proveedor')
        vendor_set = set(df_data['Vendor'].values)
        
        table_positions_dict = {}
        
        with pd.ExcelWriter(final_report_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            for vendor in unique_vendor_array:
                if vendor in vendor_set:
                    vendor_name = vendor_mapping.get(vendor, vendor)
                    vendor_df = df_data[df_data['Vendor'] == vendor]
                    
                    unique_brand_array = vendor_df['Brand'].dropna().unique()
                    
                    start_row = 1  # Fila inicial
                    table_positions = []
                    
                    for brand in unique_brand_array:
                        
                        brand_df = vendor_df[vendor_df['Brand'] == brand]
                        
                        resumen_list, details_list = create_resumen_list(brand_df, BDD_file)
                        df_resumen = pd.DataFrame(resumen_list)
                        df_details = pd.DataFrame(details_list)
                        
                        end_row = start_row + max(len(df_resumen), len(df_details))  # Calcular fin de la tabla
                        table_positions.append(((start_row+1), (end_row+1)))  # Guardar posiciones de la tabla

                        df_resumen.to_excel(writer, sheet_name=vendor_name, startrow=start_row, startcol=1, index=False)
                        #añadir la columna de detalles desde la fila 2 columna
                        df_details.to_excel(writer, sheet_name=vendor_name, startrow=start_row, startcol=20, index=False)
                        
                        start_row += max(len(df_resumen), len(df_details)) + 3
                    
                    table_positions_dict[vendor_name] = table_positions
        
        wb = load_workbook(final_report_file)
        for vendor, positions in table_positions_dict.items():
            apply_styles_to_sheet(wb, vendor, positions)
        wb.save(final_report_file)
        
        print(f"Resumen guardado en {final_report_file}")

    def create_resumen_list(vendor_df, BDD_file):
        unique_feed_index_array = vendor_df['Feed Index'].dropna().unique()
        resumen_list = []
        details_list = []
        unique_brand_array = vendor_df['Brand'].dropna().unique()

        for brand in unique_brand_array:
            for feed_index in unique_feed_index_array:
                sub_df = vendor_df[vendor_df['Feed Index'] == feed_index]
                unique_brand_array = sub_df['Brand'].dropna().unique()
                unique_duration_array = sub_df['Duracion'].dropna().unique()

                for duration in unique_duration_array:
                    feed_country_value = vendor_df.loc[vendor_df['Feed Index'] == feed_index, 'Feed'].dropna().unique()
                    feed_country = ', '.join(feed_country_value.astype(str))
                    sumary, details = create_resumen_row(vendor_df, BDD_file, feed_index, brand, duration, feed_country)
                    resumen_list.append(sumary)
                    details_list.append(details)
                    
        return resumen_list, details_list

    def create_resumen_row(vendor_df, BDD_file, feed_index, brand, duration, feed_country):
        sumary = {
            'BRAND': brand,
            'FEED INDEX': feed_index,
            'CHANNEL NAME': ', '.join(vendor_df[vendor_df['Feed Index'] == feed_index]['Channel'].dropna().unique().astype(str)),
            'FEED COUNTRY': feed_country,
            'VENDOR (NETSUIT)': ', '.join(vendor_df[vendor_df['Feed Index'] == feed_index]['Vendor'].dropna().unique().astype(str)),
            'DURATION': duration,
            'PAID SPOTS IO': len(BDD_file[(BDD_file['Feed Index'] == feed_index) & (BDD_file['Brand'] == brand) & (BDD_file['Duration'] == duration) & (BDD_file['Type Spot'] == 'Paid')]),
            'BONUS SPOTS IO': len(BDD_file[(BDD_file['Feed Index'] == feed_index) & (BDD_file['Brand'] == brand) & (BDD_file['Duration'] == duration) & (BDD_file['Type Spot'] == 'Bonus')]),
            'SPOT PAID TRANSMITTED': len(vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Duracion'] == duration) & (vendor_df['Type Spot'] == 'Paid')]),
            'SPOTS BONUS TRANSMITTED': len(vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Duracion'] == duration) & (vendor_df['Type Spot'] == 'Bonus')]),
            'SPOTS PAID RECOGNIZED': len(vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Duracion'] == duration) & (vendor_df['Type Spot'] == 'Paid') & (vendor_df['Final Result'] == 'Ok')]),
            'SPOTS BONUS RECOGNIZED': len(vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Duracion'] == duration) & (vendor_df['Type Spot'] == 'Bonus') & (vendor_df['Final Result'] == 'Ok')]),
            'SPOT PAID NOT RECOGNIZED': len(vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Duracion'] == duration) & (vendor_df['Type Spot'] == 'Paid') & (vendor_df['Final Result'] == 'No')]),
            'SPOT BONUS NOT RECOGNIZED': len(vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Duracion'] == duration) & (vendor_df['Type Spot'] == 'Bonus') & (vendor_df['Final Result'] == 'No')]),
            'SPEND LOCAL CURRENT': vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Duracion'] == duration) & (vendor_df['Final Result'] == 'Ok')]['Rate'].sum(),
            'TYPE SPOT': 'Paid' if 'Paid' in vendor_df['Type Spot'].values or 'Bonus' in vendor_df['Type Spot'].values else ('Cost Zero' if 'Cost Zero' in vendor_df['Type Spot'].values else 'Bonus'),
            'OBSERVATIONS': feed_country,
            'TOTAL SPEND DOLARIZED': vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Final Result'] == 'Ok')]['Rate'].sum(),

            }
        
        details = {
            'SPOT DUPLICADO': len(vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Duracion'] == duration)  & (vendor_df['Spot Observation'] == 'Spot Duplicado')]),
            'SPOT NO SOLICITADO': len(vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Duracion'] == duration) & (vendor_df['Spot Observation'] == 'Spot No solicitado')]),
            'CREATIVO INCORRECTO': len(vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Duracion'] == duration) & (vendor_df['Creative observation'] == 'Creativo Incorrecto')]),
            'CREATIVO TRANSMITIDO INCORRECTAMENTE': len(vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Duracion'] == duration) & (vendor_df['Creative observation'] == 'Creativo transmitido incorrectamente')]),
            'BACK TO BACK': len(vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Duracion'] == duration) & (vendor_df['Back to back'] == 'Back to back')]),
        }
        
        return sumary, details
    
    create_vendor_sheet(unique_vendor_array, final_report_file, vendor_mapping)
    generate_vendor_resume(unique_vendor_array, df_data, df_bdd, final_report_file, vendor_mapping)

def generate_rotation_tables(final_revisionpath, aux_path):
    vendor_mapping = get_vendor_mapping()
    
    month_rotation_df = pd.read_excel(aux_path, sheet_name='Month_Rotation')
    revision_details_df = pd.read_excel(final_revisionpath, sheet_name='Detalle Revision')
    
    revision_details_df['Date Time Zone'] = pd.to_datetime(revision_details_df['Date Time Zone'], errors='coerce')
    month_rotation_df['Start date'] = pd.to_datetime(month_rotation_df['Start date'], errors='coerce')
    month_rotation_df['End date'] = pd.to_datetime(month_rotation_df['End date'], errors='coerce')
        
    unique_vendors = revision_details_df['Vendor'].unique()
    
    
    with pd.ExcelWriter(final_revisionpath, mode='a', if_sheet_exists='overlay') as writer:
        for vendor in unique_vendors:
            
            vendor_filtered_df = revision_details_df[revision_details_df['Vendor'] == vendor]
            sheet_name = vendor_mapping.get(vendor, vendor)
            
            #abrir la hoja y detectar el tamaño de la tabla resumen para el vendor, y agurdar la longitud en una variable
            workbook = load_workbook(final_revisionpath)
            worksheet = workbook[sheet_name]
            max_row = worksheet.max_row
            start_row = max_row + 3
            workbook.close()
            
            unique_feed_index = vendor_filtered_df['Feed Index'].unique()
            
            for feed_index in unique_feed_index:
                channel_name = vendor_filtered_df.loc[vendor_filtered_df['Feed Index'] == feed_index, 'Channel'].values[0]
                feed_index_filtered_df = vendor_filtered_df[vendor_filtered_df['Feed Index'] == feed_index]          
                feed_country_value = next(iter(vendor_filtered_df.loc[vendor_filtered_df['Feed Index'] == feed_index, 'Feed'].dropna().unique()), "")                
                
                table_data = []
                
                for id_rev in month_rotation_df['Id Rev % Ads'].unique():
                    fecha_ads_id = month_rotation_df.loc[month_rotation_df['Id Rev % Ads'] == id_rev, 'Id Fecha Ads'].values[0]
                    start_date = month_rotation_df.loc[month_rotation_df['Id Rev % Ads'] == id_rev, 'Start date'].values[0]
                    end_date = month_rotation_df.loc[month_rotation_df['Id Rev % Ads'] == id_rev, 'End date'].values[0]
                    current_ad_brand = month_rotation_df.loc[month_rotation_df['Id Rev % Ads'] == id_rev, 'Brand'].values[0]
                    
                    date_filtered_df= feed_index_filtered_df[feed_index_filtered_df['Date Time Zone'].between(start_date, end_date) & (feed_index_filtered_df['Brand'] == current_ad_brand)]
                    
                    #Filtrar el dataframe por el id rev % ads
                    relevant_ads = date_filtered_df[(date_filtered_df['Id Rev % Ads'] == id_rev) & (date_filtered_df['Final Result'] == 'Ok')]
                    
                    total_ads = date_filtered_df[date_filtered_df['Final Result'] == 'Ok']
                    
                    expected_percentage = (month_rotation_df.loc[month_rotation_df['Id Rev % Ads'] == id_rev, 'Percentage'].values[0])*100
                    real_percentage = (relevant_ads.shape[0] / total_ads.shape[0])*100 if total_ads.shape[0] > 0 else 0
                    
                    #leave only 1 decimak value on real percentage
                    real_percentage = round(real_percentage, 1)
                    diff_pp = real_percentage - expected_percentage
                    diff_pp = round(diff_pp, 1)
                    
                    row_data = {
                        'Id Fechas Ads': fecha_ads_id,
                        'Id Rev % Ads': id_rev,
                        'Start Date': pd.Timestamp(start_date).strftime('%m/%d/%Y %H:%M:%S') if pd.notna(start_date) else '',
                        'End Date': pd.Timestamp(end_date).strftime('%m/%d/%Y %H:%M:%S') if pd.notna(end_date) else '',
                        'Ad': month_rotation_df.loc[month_rotation_df['Id Rev % Ads'] == id_rev, 'Creativo'].values[0],
                        'Brand': month_rotation_df.loc[month_rotation_df['Id Rev % Ads'] == id_rev, 'Brand'].values[0],
                        '# Ads': total_ads[(total_ads['Id Rev % Ads'] == id_rev)].shape[0],
                        '% Esperado': expected_percentage,
                        '% Real': real_percentage,
                        'Diff p.p': f"{diff_pp}pp"
                    }
                    table_data.append(row_data)
                
                table_df = pd.DataFrame(table_data)
                table_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, startcol=1, index=False)
                
                worksheet = writer.book[sheet_name]
                # Merge cells for title
                worksheet.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=11)
                # Aplicar formato al título (Channel Name - Feed Index)
                title_cell = worksheet.cell(row=start_row, column=2, value=f"{channel_name} - {feed_country_value} - {feed_index}")
                title_cell.fill = PatternFill(start_color='000000', fill_type='solid')  # Negro
                title_cell.font = Font(color='FFFFFF', bold=True)  # Blanco y negrita
                
                start_row += 1
                
                # Aplicar formato a encabezados
                header_fill = PatternFill(start_color='000080', fill_type='solid')  # Azul oscuro
                header_font = Font(color='FFFFFF', bold=True)
                for col in range(2, 12):
                    cell = worksheet.cell(row=start_row, column=col)
                    cell.fill = header_fill
                    cell.font = header_font
                    
                start_row += 1
                
                # Aplicar formato a registros
                border = Border(bottom=Side(style='dotted', color='000000'))  # Línea divisoria punteada negra
                
                previous_brand = None
                for index, row in table_df.iterrows():
                    row_fill = PatternFill(start_color='D3D3D3' if row['Id Fechas Ads'] % 2 != 0 else 'FFFFFF', fill_type='solid')
                    
                    for col, value in enumerate(row, start=2):
                        cell = worksheet.cell(row=start_row, column=col, value=value)
                        cell.fill = row_fill
                        cell.alignment = Alignment(horizontal='center')  # Centrar desde '# Ads'
                    
                    if previous_brand and previous_brand != row['Brand']:
                        for col in range(2, 11):
                            worksheet.cell(row=start_row - 1, column=col+1).border = border
                    previous_brand = row['Brand']
                    start_row += 1
                start_row += 2
    
def set_column_widths(file_path):
    # Abre el archivo
    workbook = load_workbook(file_path)

    # Hojas a excluir
    excluded_sheets = {"Detalle Revision", "BDD Revision"}

    # Recorre todas las hojas, excepto las excluidas
    for sheet_name in workbook.sheetnames:
        if sheet_name in excluded_sheets:
            worksheet = workbook[sheet_name]
            # Auto-adjust column width
            for col in worksheet.columns:
                max_length = 0
                col_letter = col[0].column_letter  # Get column letter (A, B, C, etc.)

                for cell in col:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass

        worksheet = workbook[sheet_name]
        
        # Ajusta el ancho de las columnas
        for col_idx in range(1, worksheet.max_column + 1):  
            col_letter = get_column_letter(col_idx)
            if col_letter in ['A', 'T']:  
                worksheet.column_dimensions[col_letter].width = 3
            else:
                worksheet.column_dimensions[col_letter].width = 20

    # Guarda los cambios en el archivo
    workbook.save(file_path)
    workbook.close()

def full_report(aux_path, final_path, final_report_file):
    
    #delete exsitent final report file
    if os.path.exists(final_report_file):
        os.remove(final_report_file)
    
    report_resumen_array, report_datails_array, rotation_report_array, required_columns_data = generate_report(aux_path)
    insert_data(final_path, required_columns_data, final_report_file)
    generate_columns(final_report_file, report_resumen_array, report_datails_array, rotation_report_array)
    generate_rotation_tables(final_report_file, aux_path)
    set_column_widths(final_report_file)
