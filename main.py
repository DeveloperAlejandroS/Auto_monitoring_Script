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
from Monitoria.fix_file_format import apply_transformations_to_excel_file
from Monitoria.Build_cert_file import generar_certificado_final
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
    """
    Abre la carpeta que contiene el archivo auxiliar y las reglas asociadas.

    Esta función toma la ruta del archivo auxiliar (almacenada en `aux_path`), elimina la parte que
    hace referencia al archivo y abre la carpeta contenedora para facilitar la navegación o acceso
    a otros archivos relevantes.

    Dependencias:
    -------------
    - `aux_path`: Variable global que contiene la ruta completa al archivo auxiliar (por ejemplo, 
      "C:/ruta/a/BDD Auxiliar y Reglas.xlsx"). Se debe asegurar que esta ruta esté definida correctamente 
      en el entorno del programa.
    
    Comportamiento:
    ---------------
    La función utiliza la ruta `aux_path`, reemplaza el nombre del archivo con una cadena vacía (para 
    obtener la ruta del directorio) y luego abre la carpeta contenedora mediante el comando `os.startfile`.
    """
    
    # Obtener la ruta de la carpeta contenedora del archivo auxiliar
    aux_path_container = aux_path.replace('/BDD Auxiliar y Reglas.xlsx', '')
    
    # Abrir la carpeta contenedora en el explorador de archivos
    os.startfile(aux_path_container)


# Función para manejar la selección de botones en el menú
def select_button(btn_selected, option):
    """
    Resalta el botón seleccionado en la interfaz gráfica y actualiza el contenido del panel principal 
    en función de la opción seleccionada.

    Esta función es útil para implementar un sistema de botones interactivos, donde cada botón 
    cambia de color al ser seleccionado, y también puede actualizar la interfaz con un nuevo contenido 
    relacionado con la opción correspondiente.

    Parámetros:
    ----------
    btn_selected : customtkinter.CTkButton
        El botón que ha sido seleccionado. Es un botón de la interfaz gráfica que cambiará su color 
        para reflejar su estado de selección.

    option : str
        La opción que se debe cargar o actualizar en el panel principal. Este parámetro puede ser una 
        clave que determina qué información o vista mostrar en la interfaz, como un panel de configuración, 
        un gráfico, una lista de resultados, etc.

    Dependencias:
    -------------
    - `selected_button`: Variable global que mantiene una referencia al último botón seleccionado.
    - `update_content(option)`: Función que actualiza el contenido en el panel principal de la interfaz 
      con base en la opción proporcionada.
    - `ORANGE_COLOR`: Constante global que define el color utilizado para resaltar el botón seleccionado.
    """
    
    global selected_button  # Hace referencia al botón previamente seleccionado
    
    # Si ya hay un botón seleccionado previamente, restaurar su color original
    if selected_button:
        selected_button.configure(fg_color="transparent")  # Restaurar el color anterior al estado "no seleccionado"
    
    # Cambiar el color del botón seleccionado a un color de resaltado (naranja)
    btn_selected.configure(fg_color=ORANGE_COLOR)
    
    # Actualizar la variable global que mantiene el último botón seleccionado
    selected_button = btn_selected
    
    # Llamar a la función que actualiza el contenido del panel principal según la opción seleccionada
    update_content(option)


# Función para generar los nombres de los archivos a partir de las fechas
def generate_filenames(start_date, end_date):
    """
    Genera los nombres de los archivos requeridos para el flujo de procesamiento, 
    así como las rutas necesarias donde serán guardados.

    Esta función construye dinámicamente los nombres y rutas de:
    - El archivo Excel crudo descargado de PlayLogger
    - El archivo final procesado
    - La BDD filtrada para pauta
    - El reporte final generado
    - La BDD maestra del mes correspondiente

    También asegura que las carpetas destino existen, creándolas si es necesario.

    Parámetros:
    ----------
    start_date : datetime.date
        Fecha de inicio del rango ingresado por el usuario.
    end_date : datetime.date
        Fecha de fin del rango ingresado por el usuario.

    Retorna:
    -------
    tuple[str, str, str, str, str]
        Una tupla con los siguientes elementos en orden:
        - full_bdd_path: Ruta absoluta al archivo BDD maestro correspondiente al mes.
        - base_file: Ruta completa del archivo descargado de PlayLogger que será procesado.
        - final_file: Ruta completa del archivo final generado tras aplicar transformaciones.
        - filtered_bdd_file: Ruta completa del archivo de BDD filtrada por fechas y pauta.
        - final_report_file: Ruta completa del reporte final generado a partir del archivo procesado.

    Dependencias externas:
    ----------------------
    - `resources_folder`: Ruta base definida como constante global para almacenar todos los recursos.
    - `Month_dict`: Diccionario global que traduce nombres de meses en inglés a formato personalizado (por ejemplo, "March" -> "03. Marzo").

    Estructura de carpetas generadas:
    ---------------------------------
    Dentro de `resources_folder`, se sigue la estructura:
    [AÑO]/[MM]. [MES]/PlayLogger[Revision MES DD to DD YYYY]/
        ├── Recursos/                  → Contiene: archivo base (raw), y BDD filtrada
        └── (raíz del folder)         → Contiene: archivo final y reporte final
    """
    
    # Formatos auxiliares para nombres
    month_lettered = start_date.strftime("%B")         # Ej: March
    start_day_numered = start_date.strftime("%d")      # Ej: 01
    end_day_numered = end_date.strftime("%d")          # Ej: 07
    year = start_date.strftime("%Y")                   # Ej: 2025
    month_index = f"{start_date.month:02d}"            # Ej: 03

    # Nombres de archivos basados en fechas
    raw_pl_filename = f'Descarga PlayLogger {month_lettered} {start_day_numered} to {end_day_numered} {year}.xlsx'
    final_data_pg_filename = f'Archivo Final PlayLogger {month_lettered} {start_day_numered} to {end_day_numered} {year}.xlsx'
    filtered_bdd_filename = f'BDD Filtrada pauta {month_lettered} {start_day_numered} to {end_day_numered} {year}.xlsx'
    pg_final_report_filename = f'Reporte Final PlayLogger {month_lettered} {start_day_numered} to {end_day_numered} {year}.xlsx'

    # Ruta a la BDD compartida (ruta absoluta en red interna)
    full_bdd_path = (
        f'G:/Unidades compartidas/Marketing Team/Offline Marketing/04. Operations/05. Orders BDD/'
        f'Año {year}/{month_index}-{month_lettered}/01. Orders BDD/BDD {month_lettered} {year} v1.xlsm'
    )

    # Rutas internas del proyecto donde se guardarán archivos procesados
    resources_path = (
        f"{resources_folder}/{year}/{month_index}. {Month_dict[start_date.strftime('%B')]}/"
        f"PlayLogger[Revision {month_lettered} {start_day_numered} to {end_day_numered} {year}]/Recursos"
    )

    final_rev_path = (
        f"{resources_folder}/{year}/{month_index}. {Month_dict[start_date.strftime('%B')]}/"
        f"PlayLogger[Revision {month_lettered} {start_day_numered} to {end_day_numered} {year}]"
    )

    # Rutas completas de cada archivo
    base_file = f'{resources_path}/{raw_pl_filename}'
    final_file = f'{final_rev_path}/{final_data_pg_filename}'
    filtered_bdd_file = f'{resources_path}/{filtered_bdd_filename}'
    final_report_file = f'{final_rev_path}/{pg_final_report_filename}'

    # Crear las carpetas si no existen
    if not os.path.exists(resources_path):
        os.makedirs(resources_path, exist_ok=True)
    if not os.path.exists(final_rev_path):
        os.makedirs(final_rev_path, exist_ok=True)

    return full_bdd_path, base_file, final_file, filtered_bdd_file, final_report_file


def validate_dates(excel_path, start_date, end_date):
    """
    Valida la coherencia entre las fechas proporcionadas por el usuario y las fechas contenidas en el archivo Excel.

    Esta función cumple con dos propósitos principales:
    1. Validar que el rango de fechas ingresado por el usuario (start_date a end_date) sea coherente:
        - Ambas fechas deben pertenecer al mismo mes y año.
        - La fecha de inicio no puede ser posterior a la fecha de fin.
    2. Validar que el archivo Excel contenga únicamente fechas dentro del rango ingresado.
        - Se lee exclusivamente la columna "Fecha" del archivo.
        - Se eliminan caracteres inválidos, filas vacías, y entradas que contengan la palabra "conteo".
        - Se verifica que las fechas mínimas y máximas encontradas coincidan exactamente con el rango proporcionado.

    Si cualquiera de estas validaciones falla, se muestra una alerta personalizada y se detiene el flujo principal.

    Parámetros:
    ----------
    excel_path : str
        Ruta completa del archivo Excel que contiene la columna "Fecha".
    start_date : datetime.date
        Fecha de inicio del rango a validar.
    end_date : datetime.date
        Fecha de fin del rango a validar.

    Retorna:
    -------
    bool
        True si todas las validaciones son exitosas.
        False si ocurre algún error o si las fechas no coinciden.

    Posibles errores manejados:
    ---------------------------
    - 'DateSetError': El rango ingresado no pertenece al mismo mes/año o está invertido.
    - 'MissingDateColumn': La columna "Fecha" no existe o no se pudo leer.
    - 'EmptyDateColumn': Después de limpieza, no se encontró ninguna fecha válida.
    - 'DateMismatchError': Las fechas mínimas y máximas del archivo no coinciden con el rango ingresado.
    """
    
    # Validar que el rango ingresado tenga mismo mes y año, y que start <= end
    if (start_date.month != end_date.month) or (start_date.year != end_date.year) or (start_date > end_date):
        custom_alert_trigger('DateSetError')
        return False

    try:
        df = pd.read_excel(excel_path, usecols=["Fecha"], dtype={"Fecha": str})  # Cargar solo columna "Fecha"
    except Exception as e:
        custom_alert_trigger('MissingDateColumn')
        return False

    # Limpiar datos
    df["Fecha"] = df["Fecha"].str.replace("\u00A0", "", regex=True)  # Eliminar caracteres no visibles como NO-BREAK SPACE
    df = df.dropna(subset=["Fecha"])  # Eliminar valores nulos
    df = df[~df["Fecha"].str.contains("conteo", case=False, na=False)]  # Eliminar filas que contienen "conteo"

    # Convertir a fechas válidas
    df["Fecha"] = pd.to_datetime(df["Fecha"], format="%d/%m/%Y", errors="coerce")
    df = df.dropna(subset=["Fecha"])  # Eliminar conversiones fallidas

    if df.empty:
        custom_alert_trigger('EmptyDateColumn')
        return False

    # Convertir fechas a formato uniforme para comparación
    df["Fecha"] = df["Fecha"].dt.strftime("%Y-%m-%d")
    min_date = df["Fecha"].min()
    max_date = df["Fecha"].max()

    start_date_str = start_date.strftime("%Y-%m-%d")
    end_date_str = end_date.strftime("%Y-%m-%d")

    # Validar que las fechas del archivo coincidan exactamente con las fechas ingresadas
    if (min_date != start_date_str) or (max_date != end_date_str):
        custom_alert_trigger('DateMismatchError')
        return False

    return True

# Función para cargar y monitorear los archivos, PRINCIPAL
def upload_and_monitor(start_date, end_date):
    print("Iniciando proceso de monitoreo...")
    boton_monitorear.configure(state="disabled")
    
    try:
        excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        
        if not (excel_path and os.path.exists(excel_path)):
            custom_alert_trigger('NullFilepath')
            return
        
        # Validar el nombre del archivo
        file_name = os.path.splitext(os.path.basename(excel_path))[0]
        if file_name != "Ads_Played_Log_per_Day" and "Descarga PlayLogger" not in file_name:
            custom_alert_trigger('NullFilepath')
            return
        
        # Validar fechas y generar nombres
        date_status = validate_dates(excel_path, start_date, end_date)
        full_bdd_path, base_file, final_file, filtered_bdd_file, final_report_file = generate_filenames(start_date, end_date)
        
        if not (date_status):
            return

        sheet_name = 'Archivo Final Play Logger'
        shutil.move(excel_path, base_file)
        print(f"File moved to: {base_file}")
        start_time = time.time()

        apply_transformations_to_excel_file(base_file, escribir_estado)
        generar_certificado_final(aux_path, base_file, final_file, escribir_estado)
        time.sleep(3)
        fetch_additional_columns(base_file, aux_path, final_file, sheet_name, escribir_estado)
        process_and_filter_data(full_bdd_path, aux_path, base_file, filtered_bdd_file, start_date.strftime('%m/%d/%Y'), end_date.strftime('%m/%d/%Y'), escribir_estado)
        full_revision(final_file, filtered_bdd_file, aux_path, start_date, end_date, sheet_name, escribir_estado)
        full_report(aux_path, final_file, final_report_file, escribir_estado)

        os.startfile(final_report_file)
        os.startfile(final_file)

        final_time = time.time() - start_time
        escribir_estado(f"Proceso finalizado en {final_time:.2f} segundos.")
    
    except Exception as e:
        escribir_estado(f"Error inesperado: {str(e)}")
    
    finally:
        boton_monitorear.configure(state="normal")
    

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