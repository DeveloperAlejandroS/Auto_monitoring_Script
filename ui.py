from customtkinter import *
from tkcalendar import DateEntry
from datetime import datetime
import dotenv
import os
from get_BDD import process_and_filter_data
from fix_file_format import apply_transformations_to_excel_file
from Build_cert_file import generar_certificado_final
from gen_additional_columns import Generate_additional_columns
from revision_step import delete_outdated_rows
#---------------------------------#

dotenv.load_dotenv()

Month_dict = {
    "January": "01",
    "February": "02",
    "March": "03",
    "April": "04",
    "May": "05",
    "June": "06",
    "July": "07",
    "August": "08",
    "September": "09",
    "October": "10",
    "November": "11",
    "December": "12"
}

aux_path = os.getenv("AUX_FILE_PATH")
excel_path = ''
bdd_path = os.getenv("BDD_PATH")

def Get_dates():

    start_date = start_date_entry.get_date()
    end_date = end_date_entry.get_date()

    # converting dates from yyyy-mm-dd to mm/dd/yyyy
    start_date = start_date.strftime("%m/%d/%Y")
    end_date = end_date.strftime("%m/%d/%Y")
    
    print(f"Start Date: {start_date}, End Date: {end_date}")

    return start_date, end_date

def build_file_name(start_date, end_date):
    start_date = datetime.strptime(start_date, "%m/%d/%Y")
    end_date = datetime.strptime(end_date, "%m/%d/%Y")

    start_date = start_date.strftime('%B %d')
    end_date = end_date.strftime('%d %Y')
    end_day = end_date.split()[0]
    year = end_date.split()[1]
    download_file_name = f'Descarga Play Logger {start_date} to {end_day} {year}.xlsx'
    playlogger_file_name = f'Archivo Final Play Logger {start_date} to {end_day} {year}.xlsx'

    return download_file_name, playlogger_file_name

def generate_required_files():
    start_date, end_date = Get_dates()
    download_file_name, playlogger_file_name = build_file_name(start_date, end_date)

    excel_path = f"./certs/{download_file_name}"
    final_path = f"./certs/{playlogger_file_name}"

    process_and_filter_data(bdd_path, aux_path, start_date, end_date, Month_dict)

    #Evalue if excel path file exist
    if not os.path.exists(excel_path):
        print("Excel file does not exist")
        return
    else:
        apply_transformations_to_excel_file(excel_path)
        generar_certificado_final(aux_path, excel_path)
        Generate_additional_columns(excel_path, aux_path, final_path)
        delete_outdated_rows(final_path, start_date, end_date)
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
    text="Start", 
    corner_radius=15, fg_color="#f60", 
    hover_color="#0084ff", 
    text_color="white", 
    width=200, 
    height=50, 
    command=generate_required_files
)

title_label.pack(anchor="n", pady=10, padx=10)
start_date_label.pack(anchor="s", pady=10, padx=10)
start_date_entry.pack(anchor="s", pady=10, padx=10)
end_date_label.pack(anchor="s", pady=10, padx=10)
end_date_entry.pack(anchor="s", pady=10, padx=10)
process_button.pack(anchor="n", pady=30, padx=20)

app.mainloop()