import pandas as pd
from openpyxl import load_workbook
import os

aux_path = 'G:/Unidades compartidas/Marketing Team/Offline Marketing/04. Operations/07. Monitoring/Ejecutable/Auxiliar y Reglas/BDD Auxiliar y Reglas.xlsx'
final_path = 'G:/Unidades compartidas/Marketing Team/Offline Marketing/04. Operations/07. Monitoring/2025/02. Febrero/PlayLogger[Revision February 01 to 11 2025/Archivo Final Play Logger February 01 to 11 2025.xlsx'
final_report_file = './Reporte_Final.xlsx'

def generate_report(aux_path):
        
    aux_data_df = pd.read_excel(aux_path, sheet_name='Index Tablas')
    
    report_resumen_array = aux_data_df['Report_Index'].dropna().to_list()
    report_datails_array = aux_data_df['Report_details_index'].dropna().to_list()
    rotation_report_array = aux_data_df['Rotation_report_index'].dropna().to_list()
    required_columns_data = aux_data_df['Required_final_columns'].dropna().to_list()
    
    return report_resumen_array, report_datails_array, rotation_report_array, required_columns_data
    
def insert_data(final_path, required_columns_data, final_report_file):
    
    data_df = pd.read_excel(final_path, sheet_name='Archivo Final Play Logger', converters={'Id Rev % Ads': str}) 
    print(data_df)
    ##filter the columns that are required
    
    data_df = data_df[required_columns_data]
    
    sheet_name = 'Detalle Revision'
    
    if not os.path.exists(final_report_file):
        data_df.to_excel(final_report_file, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(final_report_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            data_df.to_excel(writer, sheet_name=sheet_name, index=False)

def genrate_columns(final_report_file ,report_resumen_array, report_datails_array, rotation_report_array):
    
    #create datafframe with the required columns
    df_resumen = pd.DataFrame(columns=report_resumen_array)
    df_details = pd.DataFrame(columns=report_datails_array)
    df_rotation = pd.DataFrame(columns=rotation_report_array)
    
    dfs_to_add = []
    
    df_data = pd.read_excel(final_report_file, sheet_name='Detalle Revision')
    
    unique_vendor_array = df_data['Vendor'].dropna().unique()
    
    def create_vendor_sheet(unique_vendor_array, final_report_file):
        
        wb = load_workbook(final_report_file)
        
        for vendor in unique_vendor_array:
            
            if vendor in wb.sheetnames:
                del wb[vendor]
                
            wb.create_sheet(vendor)
        wb.save(final_report_file)
    
    def genrate_vendor_resume(df_resumen, unique_vendor_array, df_data):
        
        for vendor in unique_vendor_array:
            
            if vendor in df_data['Vendor']:
                
                vendor_df = df_data[df_data['Vendor'] == vendor]
                
                unique_feed_index_array = vendor_df['Feed Index'].dropna().unique()
                unique_brand_array = vendor_df['Brand'].dropna().unique()
                unique_durtion_array = vendor_df['Duration'].dropna().unique()
                
                for feed_index in unique_feed_index_array:
                    
                    for brand in unique_brand_array:
                        for duration in unique_durtion_array:
                            df_resumen['BRAND'] = brand
                            df_resumen['FEED INDEX'] = feed_index
                            df_resumen['CHANNEL NAME'] = vendor_df[vendor_df['Feed Index'] == feed_index]['Channel'].dropna().unique()
                            df_resumen['FEED COUNTRY'] = vendor_df[vendor_df['Feed Index'] == feed_index]['Feed'].dropna().unique()
                            df_resumen['VENDOR (NETSUIT)'] = vendor_df[vendor_df['Feed Index'] == feed_index]['Vendor'].dropna().unique()
                            df_resumen['DURATION'] = duration
                
            vendor_name = vendor
            
            #add the table in the corresponding sheet staritng in the cell B2 en every case
            df_resumen.to_excel(final_report_file, sheet_name=vendor_name, startrow=1, startcol=1, index=False)
            
            
                
    def genrate_vendor_details(df_details, unique_vendor_array, df_data):
        pass
    
def full_report(aux_path, final_path, final_report_file):
    
    report_resumen_array, report_datails_array, rotation_report_array, required_columns_data = generate_report(aux_path)
    insert_data(final_path, required_columns_data, final_report_file)
    
full_report(aux_path, final_path, final_report_file)