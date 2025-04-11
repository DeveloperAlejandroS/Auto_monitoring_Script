# Import de librerías
import customtkinter as ctk
from tkinter import BooleanVar, filedialog
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import os, shutil
import pandas as pd
import threading
import time

#Import de librerías de terceros 
from get_BDD import process_and_filter_data
from fix_file_format import apply_transformations_to_excel_file
from Build_cert_file import generar_certificado_final
from gen_additional_columns import fetch_additional_columns
from revision_step import full_revision
from reporting_file import full_report

#Import de funciones

#==============================================================================#
#                                    LÓGICA                                    #
#==============================================================================#

# Diccionario para traducir los nombres de los meses
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

# Botones del menú con efecto hover y selección persistente
menu_buttons = []
menu_options = ["Monitoria", "Cierres", "Horarios"]

# Rutas de los archivos necesarios (FIJAS)
excel_path = ''
bdd_filtered_path = ''
aux_path = 'G:/Unidades compartidas/Marketing Team/Offline Marketing/04. Operations/07. Monitoring/Ejecutable/Auxiliar y Reglas/BDD Auxiliar y Reglas.xlsx'
resources_folder = 'G:/Unidades compartidas/Marketing Team/Offline Marketing/04. Operations/07. Monitoring'

def open_aux_rules():
    aux_path_container = aux_path.replace('/BDD Auxiliar y Reglas.xlsx', '')
    os.startfile(aux_path_container)

# Función para manejar la selección de botones en el menú
def select_button(btn_selected, option):
    global selected_button
    if selected_button:
        selected_button.configure(fg_color="transparent")  # Restaurar el anterior
    btn_selected.configure(fg_color=ORANGE_COLOR)  # Resaltar el seleccionado
    selected_button = btn_selected
    update_content(option)

# Función para generar los nombres de los archivos a partir de las fechas
def generate_filenames(start_date, end_date):
    success = False
    
    month_lettered = start_date.strftime("%B")
    end_month_lettered = end_date.strftime("%B")
    start_day_numered = start_date.strftime("%d")
    end_day_numered = end_date.strftime("%d")
    year = start_date.strftime("%Y")
    end_date_year = end_date.strftime("%Y")
    month_index = f"{start_date.month:02d}"
    
    #Crear los nombres de los archivos a partir de las fechas 
    raw_pl_filename = f'Descarga PlayLogger {month_lettered} {start_day_numered} to {end_day_numered} {year}.xlsx'
    final_data_pg_filename = f'Archivo Final PlayLogger {month_lettered} {start_day_numered} to {end_day_numered} {year}.xlsx'
    filtered_bdd_filename = f'BDD Filtrada pauta {month_lettered} {start_day_numered} to {end_day_numered} {year}.xlsx'
    full_bdd_path = f'G:/Unidades compartidas/Marketing Team/Offline Marketing/04. Operations/05. Orders BDD/Año {year}/{month_index}-{month_lettered}/01. Orders BDD/BDD {month_lettered} {year} v1.xlsm'
    pg_final_report_filename = f'Reporte Final PlayLogger {month_lettered} {start_day_numered} to {end_day_numered} {year}.xlsx'
    
    #Generar la ruta a la  carpetera de los archivos
    resources_path = f"{resources_folder}/{start_date.strftime('%Y')}/{start_date.strftime("%m")}. {Month_dict[start_date.strftime('%B')]}/PlayLogger[Revision {start_date.strftime('%B')} {start_date.strftime('%d')} to {end_date.strftime('%d')} {end_date.strftime('%Y')}/Recursos"
    final_rev_path = f"{resources_folder}/{start_date.strftime('%Y')}/{start_date.strftime("%m")}. {Month_dict[start_date.strftime('%B')]}/PlayLogger[Revision {start_date.strftime('%B')} {start_date.strftime('%d')} to {end_date.strftime('%d')} {end_date.strftime('%Y')}"
    
    base_file = f'{resources_path}/{raw_pl_filename}'
    final_file = f'{final_rev_path}/{final_data_pg_filename}'
    filtered_bdd_file = f'{resources_path}/{filtered_bdd_filename}'
    final_report_file = f'{final_rev_path}/{pg_final_report_filename}'
    
    if (year != end_date_year) or( month_lettered != end_month_lettered) or (start_date > end_date):
        custom_alert_trigger('DateSetError')
        success = False
        return success, "", "", "", "", "", "", "", "", "", "", ""
    else:
        if not os.path.exists(resources_path):
            os.makedirs(resources_path, exist_ok=True)
        if not os.path.exists(final_rev_path):
            os.makedirs(final_rev_path, exist_ok=True)
        success = True
        return success, full_bdd_path, base_file, final_file, filtered_bdd_file, final_report_file

def validate_dates(excel_path, start_date, end_date):
    
    
    generate_filenames(start_date, end_date)
    df = pd.read_excel(excel_path, usecols=["Fecha"], dtype={"Fecha": str})  # Solo cargar la columna "Fecha"
        
    # Reemplazar caracteres no válidos
    df["Fecha"] = df["Fecha"].str.replace("\u00A0", "", regex=True)  # Eliminar U+00A0
    
    # Filtrar valores inválidos
    df = df.dropna(subset=["Fecha"])  # Eliminar nulos
    df = df[~df["Fecha"].str.contains("conteo", case=False, na=False)]  # Eliminar "conteo"
    
    # Convertir a formato datetime (solo fechas válidas)
    df["Fecha"] = pd.to_datetime(df["Fecha"], format="%d/%m/%Y", errors="coerce")
    df = df.dropna(subset=["Fecha"])  # Eliminar filas con fechas no válidas
    
    # Convertir formato a "YYYY-MM-DD"
    df["Fecha"] = df["Fecha"].dt.strftime("%Y-%m-%d")
    
    # Extraer fecha mínima y máxima
    min_date = df["Fecha"].min()
    max_date = df["Fecha"].max()

    # Convertir start_date y end_date a string en formato YYYY-MM-DD para comparar
    start_date_str = start_date.strftime("%Y-%m-%d")
    end_date_str = end_date.strftime("%Y-%m-%d")

    # Comparar si las fechas son diferentes a las ingresadas
    if (min_date != start_date_str) or (max_date != end_date_str):
        custom_alert_trigger('DateMismatchError')
        success = False
    else:
        success = True
        
    return success

# Función para cargar y monitorear los archivos, PRINCIPAL
def upload_and_monitor(start_date, end_date):
    print("Iniciando proceso de monitoreo...")
    boton_monitorear.configure(state="disabled")
    
    # Cargar los archivos
    excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    #Evaluar si el archivo existe y no es nulo
    if excel_path and os.path.exists(excel_path):
        # Verificar si las fechas son válidas y están en el mismo mes y año
        # y si el rango de días coincide con el del certificado
        date_status = validate_dates(excel_path, start_date, end_date)
        success, full_bdd_path, base_file, final_file, filtered_bdd_file, final_report_file = generate_filenames(start_date, end_date)
        # Si las fechas son válidas, generar los nombres de los archivos, asi mismo se verifica si el rango de días coincide con el del certificado
        if date_status and success:
            sheet_name = 'Archivo Final Play Logger'
            shutil.move(excel_path, base_file)
            print(f"File moved to: {base_file}")
            start_time = time.time()
            
            apply_transformations_to_excel_file(base_file, escribir_estado)
            generar_certificado_final(aux_path, base_file, final_file, escribir_estado)
            time.sleep(3)
            fetch_additional_columns(base_file, aux_path, final_file, sheet_name, escribir_estado)
            
            process_and_filter_data(full_bdd_path, aux_path, base_file , filtered_bdd_file, start_date.strftime('%m/%d/%Y'), end_date.strftime('%m/%d/%Y'), escribir_estado)
            full_revision(final_file, filtered_bdd_file, aux_path, start_date, end_date, sheet_name, escribir_estado)
            full_report(aux_path, final_file, final_report_file, escribir_estado)
            
            os.startfile(final_report_file)
            os.startfile(final_file)
            
            final_time = time.time() - start_time
            escribir_estado(f"Proceso finalizado en {final_time:.2f} segundos.")
            boton_monitorear.configure(state="normal")
        else:
            #Detener el proceso, no se da alerta ya que ya se dio anteriormente 
            return
    else:
        custom_alert_trigger('NullFilepath')
    

#==============================================================================#
#                                    VISTA                                     #
#==============================================================================#

# Configuración de la ventana principal
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

root = ctk.CTk()

# Obtener tamaño de la pantalla
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Tamaño de la ventana
window_width = 800
window_height = 500

# Posicionar en el centro de la pantalla
x = (screen_width // 2) - (window_width // 2)
y = (screen_height // 2) - (window_height // 2)

# Configurar ventana
root.geometry(f"{window_width}x{window_height}+{x}+{y}")
root.resizable(False, False)
root.title("Monitoreo de Pauta")

# Colores personalizados
BLUE_COLOR = "#0078FF"
LIGHT_BLUE = "#B3D7FF"  # Azul más claro para el text field de estado
ORANGE_COLOR = "#FF6600"
FONT_INTER = ("Inter", 14)  # Fuente oficial de la empresa
FONT_INTER_BOLD = (FONT_INTER[0], FONT_INTER[1], "bold")

# Panel lateral (menú de hamburguesa) sin esquinas redondeadas a la derecha
menu_frame = ctk.CTkFrame(root, width=275, height=500, fg_color=BLUE_COLOR, corner_radius=0)
menu_frame.pack_propagate(False)
menu_frame.pack(side="left", fill="y")

# Logo de Open English
logo_label = ctk.CTkLabel(menu_frame, text="Open English", font=("Inter", 20, "bold"), text_color="white")
logo_label.pack(pady=20)

# Variable para rastrear el botón seleccionado
selected_button = None

# Panel derecho (contenido principal) sin esquinas redondeadas a la izquierda
main_frame = ctk.CTkFrame(root, width=525, height=500, fg_color="white", corner_radius=0)
main_frame.pack_propagate(False)
main_frame.pack(side="right", fill="both", expand=True)

def custom_alert_trigger(type):
    alert_frame = ctk.CTkToplevel()
    alert_frame.geometry("350x200")
    alert_frame.resizable(False, False)
    alert_frame.configure(fg_color="white")  # Fondo blanco

    # Posicionar en el centro de la pantalla
    screen_width = alert_frame.winfo_screenwidth()
    screen_height = alert_frame.winfo_screenheight()
    window_width = 350
    window_height = 200

    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)

    alert_frame.geometry(f"{window_width}x{window_height}+{x}+{y}")

    # Diccionario de mensajes de error
    error_messages = {
        "DateSetError": ("Error en rango de fechas", 
                         "Las fechas deben estar en el mismo mes y año. La inicial no puede ser mayor que la final."),
        "DateMismatchError": ("Error en rango de fechas", 
                             "El rango de días seleccionado debe coincidir con el del certificado."),
        "NullTrmValue": ("Error en TRM", 
                         "Debe ingresar un valor para la TRM."),
        "NullFilepath": ("Error en archivo",
                         "Debe seleccionar un archivo valido para continuar."),
    }

    # Obtener título y mensaje, o usar un valor por defecto
    title, message_text = error_messages.get(type, ("Error", "Se ha producido un error inesperado"))
    
    alert_frame.title(title)

    # Contenedor principal centrado
    content_frame = ctk.CTkFrame(alert_frame, fg_color="white")
    content_frame.pack(fill="both", expand=True, padx=20, pady=20)

    # Label con texto centrado y ajuste automático
    message = ctk.CTkLabel(
        content_frame, 
        text=message_text, 
        font=("Arial", 14), 
        wraplength=300,  
        justify="center"  
    )
    message.pack(pady=10, padx=10, fill="both", expand=True)

    # Botón centrado
    close_button = ctk.CTkButton(
        content_frame, 
        text="Cerrar", 
        fg_color="#007BFF", 
        hover_color="#0056b3", 
        command=alert_frame.destroy
    )
    close_button.pack(pady=15)

    alert_frame.grab_set()

def escribir_estado(mensaje):
    estado_texto.configure(state="normal")
    estado_texto.insert("end", mensaje + "\n")
    estado_texto.see("end")  # Para hacer scroll automático al final
    estado_texto.configure(state="disabled")

# Contenedor para actualizar contenido
def update_content(option):
    global content_frame, boton_monitorear, estado_texto
    
    for widget in main_frame.winfo_children():
        widget.destroy()
    
    content_frame = ctk.CTkFrame(main_frame, fg_color="white")
    content_frame.place(relx=0.5, rely=0.5, anchor="center")
    
    title_label = ctk.CTkLabel(content_frame, text=option, font=("Inter", 20, "bold"), text_color="black")
    title_label.pack(pady=10)
    
    if option == "Monitoria":
        
        default_end_date = datetime.today() - timedelta(days=1)
        default_start_date = datetime.today().replace(day=1)
        
        fecha_inicio_label = ctk.CTkLabel(
            content_frame, 
            text="Fecha Inicio", 
            text_color="black", 
            font=FONT_INTER
        )
        fecha_inicio_label.pack()
        fecha_inicio = DateEntry(
            content_frame, 
            width=30, 
            background=BLUE_COLOR, 
            foreground='white', 
            borderwidth=2, 
            justify='center', 
            day=default_start_date.day  # Establecer el primer día del mes actual
            )
        
        fecha_inicio.pack(pady=5, ipady=10)
        
        fecha_fin_label = ctk.CTkLabel(
            content_frame, 
            text="Fecha Final", 
            text_color="black", 
            font=FONT_INTER
        )
        fecha_fin_label.pack()
        fecha_fin = DateEntry(
            content_frame, 
            width=30, 
            background=BLUE_COLOR, 
            foreground='white', 
            borderwidth=2, 
            justify='center',
            maxdate=default_end_date
        )
        fecha_fin.pack(pady=5, ipady=10)
        
        boton_monitorear = ctk.CTkButton(
            content_frame, 
            text="Cargar y Monitorear", 
            fg_color=BLUE_COLOR, 
            hover_color=ORANGE_COLOR, 
            width=274, height=50, 
            font=FONT_INTER,
            command=(lambda: threading.Thread(target=upload_and_monitor, args=(fecha_inicio.get_date(), fecha_fin.get_date())).start())
        )
        boton_monitorear.pack(pady=20)
        
        estado_texto = ctk.CTkTextbox(
            content_frame, 
            width=350, 
            height=150, 
            fg_color=LIGHT_BLUE, 
            text_color="black", 
            font=FONT_INTER, 
            corner_radius=10, 
            state="disabled", 
            border_color=BLUE_COLOR, 
            border_width=2
            )
        estado_texto.pack(pady=10)
    
    else:
        proveedor_label = ctk.CTkLabel(content_frame, text="Proveedor", text_color="black", font=FONT_INTER)
        proveedor_label.pack()
        
        proveedor_selector = ctk.CTkComboBox(content_frame, values=["Proveedor A", "Proveedor B", "Proveedor C"], font=FONT_INTER)
        proveedor_selector.pack(pady=5)
        
        if option == "Cierres":
            trm_var = BooleanVar()
            
            def toggle_trm():
                if trm_var.get():
                    trm_entry.configure(state="normal")
                else:
                    trm_entry.configure(state="disabled")
            
            trm_checkbox = ctk.CTkCheckBox(content_frame, text="Usar TRM personalizada", variable=trm_var, command=toggle_trm, font=FONT_INTER)
            trm_checkbox.pack(pady=5)
            
            trm_entry = ctk.CTkEntry(content_frame, font=FONT_INTER, width=200, state="disabled")
            trm_entry.pack(pady=5)

for text in menu_options:
    button = ctk.CTkButton(menu_frame, text=text, fg_color="transparent", hover_color=ORANGE_COLOR,
                           corner_radius=0, border_width=0, height=50, width=275, font=FONT_INTER)
    button.pack(fill="both", pady=2)
    button.configure(command=lambda b=text, btn=button: select_button(btn, b))
    menu_buttons.append(button)
    
aux_button = ctk.CTkButton(
    menu_frame, text="Ruta Aux y Reglas", 
    fg_color="transparent", 
    hover_color=ORANGE_COLOR,
    corner_radius=0, 
    height=50, 
    width=275, 
    font=FONT_INTER_BOLD,  
    command=open_aux_rules
    )
aux_button.place(x=0, rely=1.0, y=-55, relwidth=1.0)
    
# Inicializar con Monitoria seleccionada
select_button(menu_buttons[0], "Monitoria")

root.mainloop()