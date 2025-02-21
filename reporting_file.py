import pandas as pd
from openpyxl import load_workbook
import os

aux_path = 'G:/Unidades compartidas/Marketing Team/Offline Marketing/04. Operations/07. Monitoring/Ejecutable/Auxiliar y Reglas/BDD Auxiliar y Reglas.xlsx'
final_path = 'G:/Unidades compartidas/Marketing Team/Offline Marketing/04. Operations/07. Monitoring/2025/02. Febrero/PlayLogger[Revision February 01 to 11 2025/Archivo Final Play Logger February 01 to 11 2025.xlsx'
final_report_file = './Reporte_Final.xlsx'


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

def generate_columns(final_report_file, report_resumen_array, report_details_array, rotation_report_array):
    print('Generando columnas de resumen con datos de reporte final')
    # Crear DataFrames vacíos con los encabezados requeridos
    df_resumen = pd.DataFrame(columns=report_resumen_array)
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
        
        with pd.ExcelWriter(final_report_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            for vendor in unique_vendor_array:
                if vendor in vendor_set:
                    vendor_name = vendor_mapping.get(vendor, vendor)
                    vendor_df = df_data[df_data['Vendor'] == vendor]
                    resumen_list = create_resumen_list(vendor_df, BDD_file)
                    df_resumen = pd.DataFrame(resumen_list)
                    df_resumen.to_excel(writer, sheet_name=vendor_name, startrow=1, startcol=1, index=False)
                    
        print(f"Resumen guardado en {final_report_file}")

    def create_resumen_list(vendor_df, BDD_file):
        unique_feed_index_array = vendor_df['Feed Index'].dropna().unique()
        resumen_list = []

        for feed_index in unique_feed_index_array:
            sub_df = vendor_df[vendor_df['Feed Index'] == feed_index]
            unique_brand_array = sub_df['Brand'].dropna().unique()
            unique_duration_array = sub_df['Duracion'].dropna().unique()

            for brand in unique_brand_array:
                for duration in unique_duration_array:
                    feed_country_value = vendor_df.loc[vendor_df['Feed Index'] == feed_index, 'Feed'].dropna().unique()
                    feed_country = ', '.join(feed_country_value.astype(str))
                    row = create_resumen_row(vendor_df, BDD_file, feed_index, brand, duration, feed_country)
                    resumen_list.append(row)
        return resumen_list

    def create_resumen_row(vendor_df, BDD_file, feed_index, brand, duration, feed_country):
        return {
            'BRAND': brand,
            'FEED INDEX': feed_index,
            'CHANNEL NAME': ', '.join(vendor_df[vendor_df['Feed Index'] == feed_index]['Channel'].dropna().unique().astype(str)),
            'FEED COUNTRY': feed_country,
            'VENDOR (NETSUIT)': ', '.join(vendor_df[vendor_df['Feed Index'] == feed_index]['Vendor'].dropna().unique().astype(str)),
            'DURATION': duration,
            'PAID SPOTS IO': len(BDD_file[(BDD_file['Feed Index'] == feed_index) & (BDD_file['Brand'] == brand) & (BDD_file['Type Spot'] == 'Paid')]),
            'BONUS SPOTS IO': len(BDD_file[(BDD_file['Feed Index'] == feed_index) & (BDD_file['Brand'] == brand) & (BDD_file['Type Spot'] == 'Bonus')]),
            'SPOT PAID TRANSMITTED': len(vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Type Spot'] == 'Paid')]),
            'SPOTS BONUS TRANSMITTED': len(vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Type Spot'] == 'Bonus')]),
            'SPOTS PAID RECOGNIZED': len(vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Type Spot'] == 'Paid') & (vendor_df['Final Result'] == 'Ok')]),
            'SPOTS BONUS RECOGNIZED': len(vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Type Spot'] == 'Bonus') & (vendor_df['Final Result'] == 'Ok')]),
            'SPOT PAID NOT RECOGNIZED': len(vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Type Spot'] == 'Paid') & (vendor_df['Final Result'] == 'No')]),
            'SPOT BONUS NOT RECOGNIZED': len(vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Type Spot'] == 'Bonus') & (vendor_df['Final Result'] == 'No')]),
            'SPEND LOCAL CURRENT': vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Final Result'] == 'Ok')]['Rate'].sum(),
            'TYPE SPOT': 'Paid' if 'Paid' in vendor_df['Type Spot'].values else ('Cost Zero' if 'Cost Zero' in vendor_df['Type Spot'].values else 'Bonus'),
            'OBSERVATIONS': feed_country,
            'TOTAL SPEND DOLARIZED': vendor_df[(vendor_df['Feed Index'] == feed_index) & (vendor_df['Brand'] == brand) & (vendor_df['Final Result'] == 'Ok')]['Rate'].sum(),
        }
    
    create_vendor_sheet(unique_vendor_array, final_report_file, vendor_mapping)
    generate_vendor_resume(unique_vendor_array, df_data, df_bdd, final_report_file, vendor_mapping)
    
def full_report(aux_path, final_path, final_report_file):
    
    #delete exsitent final report file
    if os.path.exists(final_report_file):
        os.remove(final_report_file)
    
    report_resumen_array, report_datails_array, rotation_report_array, required_columns_data = generate_report(aux_path)
    insert_data(final_path, required_columns_data, final_report_file)
    generate_columns(final_report_file, report_resumen_array, report_datails_array, rotation_report_array)
    
    os.system(f'start excel "{final_report_file}"')
    
full_report(aux_path, final_path, final_report_file)