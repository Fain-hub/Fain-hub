from dbfread import DBF
import openpyxl
from openpyxl.utils import get_column_letter
import re
import os
import unicodedata
import tkinter as tk
from tkinter import filedialog

def clean_value(value):
    if isinstance(value, str):
        value = unicodedata.normalize('NFKC', value)
        value = ''.join(c for c in value if unicodedata.category(c) != 'Cc')
        return re.sub(r'[^\x00-\x7F]+', '', value)
    return value

def dbf_to_xlsx(dbf_file, xlsx_file, encoding='latin-1', file_type='xlsx'):
    table = DBF(dbf_file, encoding=encoding)
    xlsx_workbook = openpyxl.Workbook()
    xlsx_sheet = xlsx_workbook.active

    for i, field in enumerate(table.fields):
        xlsx_sheet.cell(row=1, column=i + 1, value=field.name)

    for row_idx, row in enumerate(table):
        for col_idx, value in enumerate(row.values()):
            xlsx_sheet.cell(row=row_idx + 2, column=col_idx + 1, value=clean_value(value))

    for column in xlsx_sheet.columns:
        max_length = 0
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except TypeError:
                pass
        adjusted_column = (max_length + 2) * 1.1
        xlsx_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_column

    xlsx_workbook.save(xlsx_file)

def select_dbf_file():
    root = tk.Tk()
    root.withdraw()
    dbf_file = filedialog.askopenfilename(title="Select DBF file", filetypes=[("DBF Files", "*.dbf")])
    return dbf_file

def get_xlsx_filename():
    root = tk.Tk()
    root.withdraw()
    xlsx_filename = filedialog.asksaveasfilename(title="Save XLSX file as", defaultextension=".xlsx", filetypes=[("XLSX Files", "*.xlsx")])
    return xlsx_filename

dbf_file = select_dbf_file()
if not dbf_file:
    print("No DBF file selected.")
    exit()

xlsx_filename = get_xlsx_filename()
if not xlsx_filename:
    print("No XLSX filename specified.")
    exit()

try:
    dbf_to_xlsx(dbf_file, xlsx_filename, file_type='xlsx')
except Exception as e:
    print(f"Error saving as strict XLSX: {e}")
    dbf_to_xlsx(dbf_file, xlsx_filename)

print("Finished")
