from customtkinter import *
from tkcalendar import DateEntry
from datetime import datetime
import os
from get_BDD import process_and_filter_data
from fix_file_format import apply_transformations_to_excel_file
from Build_cert_file import generar_certificado_final
from gen_additional_columns import fetch_additional_columns
from revision_step import full_revision
from tkinter import filedialog
import shutil

#---------------------------------#


# Diccionario de meses
Month_dict = {
    "January": "Enero",
    "February": "Febrero",
    "March": "Marzo",
    "April": "Abril",
    "May": "Mayo",
    "June": "Junio",
    "July": "Julio",
    "August": "Agosto",
    "September": "Septiembre",
    "October": "Octubre",
    "November": "Noviembre",
    "December": "Diciembre"
}

excel_path = ''
bdd_filtered_path = ''
aux_path = 'G:/Unidades compartidas/Marketing Team/Offline Marketing/04. Operations/07. Monitoring/Ejecutable/Auxiliar y Reglas/BDD Auxiliar y Reglas.xlsx'
resources_folder = 'G:/Unidades compartidas/Marketing Team/Offline Marketing/04. Operations/07. Monitoring'


def Get_dates():
    """Obtiene las fechas de inicio y fin desde los DateEntry y las convierte a formato mm/dd/yyyy"""
    start_date = start_date_entry.get_date()
    end_date = end_date_entry.get_date()

    # Convertir fechas de yyyy-mm-dd a mm/dd/yyyy
    start_date = start_date.strftime("%m/%d/%Y")
    end_date = end_date.strftime("%m/%d/%Y")

    print(f"Start Date: {start_date}, End Date: {end_date}")

    return start_date, end_date

def build_file_name(start_date, end_date):
    """Genera los nombres de archivo basados en las fechas de inicio y fin"""
    start_date = datetime.strptime(start_date, "%m/%d/%Y")
    end_date = datetime.strptime(end_date, "%m/%d/%Y")

    month_index = f"{start_date.month:02d}"
    
    start_date = start_date.strftime('%B %d')
    end_date = end_date.strftime('%d %Y')
    end_day = end_date.split()[0]
    year = end_date.split()[1]
    month_name = start_date.split()[0]
    
    raw_playlogger_file_name = f'Descarga Play Logger {start_date} to {end_day} {year}.xlsx'
    final_playlogger_file_name = f'Archivo Final Play Logger {start_date} to {end_day} {year}.xlsx'
    filtered_bdd_file_name = f'BDD Pauta {start_date} to {end_day} {year}.xlsx'
    full_bdd_path = f'G:/Unidades compartidas/Marketing Team/Offline Marketing/04. Operations/05. Orders BDD/Año {year}/{month_index}-{month_name}/01. Orders BDD/BDD {month_name} {year} v1.xlsm'

    return raw_playlogger_file_name, final_playlogger_file_name, filtered_bdd_file_name, full_bdd_path

def createFolders(start_date, end_date):
    start_date = datetime.strptime(start_date, "%m/%d/%Y")
    end_date = datetime.strptime(end_date, "%m/%d/%Y")
    
    #Crear la carpeta de recursos
    resources_path = f"{resources_folder}/{start_date.strftime('%Y')}/{start_date.strftime("%m")}. {Month_dict[start_date.strftime('%B')]}/PlayLogger[Revision {start_date.strftime('%B')} {start_date.strftime('%d')} to {end_date.strftime('%d')} {end_date.strftime('%Y')}/Recursos"
    final_rev_path = f"{resources_folder}/{start_date.strftime('%Y')}/{start_date.strftime("%m")}. {Month_dict[start_date.strftime('%B')]}/PlayLogger[Revision {start_date.strftime('%B')} {start_date.strftime('%d')} to {end_date.strftime('%d')} {end_date.strftime('%Y')}"
    
    if not os.path.exists(resources_path):
        os.makedirs(resources_path, exist_ok=True)
    if not os.path.exists(final_rev_path):
        os.makedirs(final_rev_path, exist_ok=True)
    
    return resources_path, final_rev_path

def gen_full_file_path(raw_playlogger_file_name, final_playlogger_file_name, filtered_bdd_file_name, resources_path, final_rev_path):
    base_file = f'{resources_path}/{raw_playlogger_file_name}'
    final_file = f'{final_rev_path}/{final_playlogger_file_name}'
    filtered_bdd_file = f'{resources_path}/{filtered_bdd_file_name}'
    
    return base_file, final_file, filtered_bdd_file
    

def generate_required_files():
    """Genera los archivos necesarios y realiza las transformaciones y filtrados"""
    start_date, end_date = Get_dates()
    
    # Generar los nombres de archivo
    raw_playlogger_file_name, final_playlogger_file_name, filtered_bdd_file_name, full_bdd_path = build_file_name(start_date, end_date)
    
    resources_path, final_rev_path = createFolders(start_date, end_date)
    
    #Geenrar losa rchivos enrutados
    base_file, final_file, filtered_bdd_file = gen_full_file_path(raw_playlogger_file_name, final_playlogger_file_name, filtered_bdd_file_name, resources_path, final_rev_path)
    
    #Imprimir todas las rutas para verificar
    print(base_file)
    print(final_file)
    print(filtered_bdd_file)
    print(full_bdd_path)
    
    #Open file from location and move it to the resources folder
    excel_path = filedialog.askopenfilename()
    print(excel_path)
    if excel_path:
        shutil.move(excel_path, base_file)
        print(f"File moved to: {base_file}")
    else:
        print("No file selected")
        return
    
    if not os.path.exists(base_file) or base_file == '':
        print(f"Excel file does not exist: {base_file}")
        return
    else:
        sheet_name = 'Archivo Final Play Logger'
        if not os.path.exists(base_file):
            print(f"Excel file does not exist: {base_file}")
            return
        else:
            # Aplicar transformaciones y generar certificados
            apply_transformations_to_excel_file(base_file)
            generar_certificado_final(aux_path, base_file, final_file)
            fetch_additional_columns(base_file, aux_path, final_file, sheet_name)
            
            # Filtrar los datos BDD
            process_and_filter_data(full_bdd_path, aux_path, base_file , filtered_bdd_file, start_date, end_date)
            
            full_revision(final_file, filtered_bdd_file, start_date, end_date, sheet_name)
            
            #Open final folder
            os.startfile(final_rev_path)
            os.startfile(final_file)

#---------------------------------#
app = CTk()
app.title("Auto-Monitoria v1")
app.geometry("500x550")
app.resizable(False, False)

app.config(bg="white")

frame = CTkFrame(
    master=app, 
    width=356, 
    height=438, 
    fg_color="white", 
    bg_color="white", 
    corner_radius=15, 
    border_width=1,
    border_color="#0084ff"
)

frame.place(relx=0.5, rely=0.5, anchor="s", y=100)

progress_text_field = CTkTextbox(
    master=app,
    width=450, 
    height=120, 
    font=("Roboto", 12), 
    bg_color="white", 
    fg_color="#fcbd92", 
    corner_radius=15, 
    border_width=1, 
    border_color="#0084ff",
    text_color="#141414",
    
)

progress_text_field.configure(state="disabled")

progress_text_field.place(
    relx=0.5, 
    rely=0.5, 
    anchor="center", 
    y=190
)

# Start Date Label y DateEntry
title_label = CTkLabel(
    master=app, 
    text="Auto-Monitoria", 
    font=("Roboto", 40, "bold"), 
    text_color="#0084ff", 
    bg_color="white"
)
start_date_label = CTkLabel(
    master=frame, 
    text="Start Date:", 
    font=("Roboto", 12, "bold"), 
    text_color="#0084ff"
)
start_date_entry = DateEntry(
    master=frame, 
    font=("Roboto", 12), 
    date_pattern='mm/dd/yyyy'
)

end_date_label = CTkLabel(
    master=frame, 
    text="End Date:", 
    font=("Roboto", 12, "bold"), 
    text_color="#0084ff"
)
end_date_entry = DateEntry(
    master=frame, 
    font=("Roboto", 12), 
    date_pattern='mm/dd/yyyy'
)

process_button = CTkButton(
    master=frame, 
    text="Generar Archivos", 
    corner_radius=15, fg_color="#f60", 
    hover_color="#0084ff", 
    text_color="white", 
    width=200, 
    height=50, 
    command=generate_required_files
)

title_label.pack(anchor="n", pady=5, padx=10)
start_date_label.pack(anchor="s", pady=5, padx=10)
start_date_entry.pack(anchor="s", pady=5, padx=10)
end_date_label.pack(anchor="s", pady=5, padx=10)
end_date_entry.pack(anchor="s", pady=5, padx=10)
process_button.pack(anchor="s", pady=5, padx=20)

app.mainloop()