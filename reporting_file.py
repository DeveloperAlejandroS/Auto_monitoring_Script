import pandas as pd

aux_path = 'G:/Unidades compartidas/Marketing Team/Offline Marketing/04. Operations/07. Monitoring/Ejecutable/Auxiliar y Reglas/BDD Auxiliar y Reglas.xlsx'
final_path = 'G:/Unidades compartidas/Marketing Team/Offline Marketing/04. Operations/07. Monitoring/2025/02. Febrero/PlayLogger[Revision February 01 to 06 2025/Archivo Final Play Logger February 01 to 06 2025.xlsx'


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
    
    with pd.ExcelWriter(final_report_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        data_df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    
def full_report(aux_path, final_path, final_report_file):
    
    report_resumen_array, report_datails_array, rotation_report_array, required_columns_data = generate_report(aux_path)
    insert_data(final_path, required_columns_data, final_report_file)