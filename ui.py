import sys
import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter import filedialog
from datetime import datetime
import dotenv
import os
from get_BDD import process_and_filter_data
from fix_file_format import apply_transformations_to_excel_file
from Build_cert_file import generar_certificado_final
from gen_additional_columns import Generate_additional_columns

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

aux_path = os.getenv("AUX_PATH")
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
    output_path = f'Descarga Play Logger {start_date} to {end_day} {year}.xlsx'

    return output_path

def generate_required_files():
    start_date, end_date = Get_dates()

    excel_path = f"./certs/{build_file_name(start_date, end_date)}"

    process_and_filter_data(bdd_path, aux_path, start_date, end_date, Month_dict)

    apply_transformations_to_excel_file(excel_path)
    generar_certificado_final(aux_path, excel_path)
    Generate_additional_columns(excel_path, aux_path, start_date, end_date, Month_dict)
    

root = tk.Tk()
root.title("Auto monitoring")

frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

ttk.Label(frame, text="Start Date:").grid(row=0, column=0, padx=5, pady=5)
start_date_entry = DateEntry(frame, width=12, background='darkblue', foreground='white', borderwidth=2)
start_date_entry.grid(row=0, column=1, padx=5, pady=5)

ttk.Label(frame, text="End Date:").grid(row=1, column=0, padx=5, pady=5)
end_date_entry = DateEntry(frame, width=12, background='darkblue', foreground='white', borderwidth=2)
end_date_entry.grid(row=1, column=1, padx=5, pady=5)

build_button = ttk.Button(frame, text="Process", command=generate_required_files)
build_button.grid(row=2, column=0, columnspan=2, pady=10)



root.mainloop()