import os
import pandas as pd
from openpyxl import load_workbook

def prepare_file(download_path):
    files = os.listdir(download_path)
    for file in files:
        if file.endswith('.xlsx'):
            target_file = download_path + "\\" + file
    wb = load_workbook(target_file)
    sheet = wb['Sheet1']
    sheet.delete_rows(1)
    wb.save(target_file)
    correct_file(target_file)


def correct_file(file):
    # Replace NaN with 0 and convert all numerbs in float
    df = pd.read_excel(file)
    df = df.fillna(0)
    df['Imp. Total'] = df['Imp. Total'].apply(lambda x: float(x))
    # Convert to negative all credit notes
    columns = ['Imp. Neto Gravado', 'Imp. Neto No Gravado', 'Imp. Op. Exentas', 'IVA', 'Imp. Total']
    targets = ['11 - Nota de Credito A', '11 - Nota de Credito B', '11 - Nota de Credito C', '11 - Nota de Credito E']
    for column in columns:
        for target in targets:
            df.loc[df['Tipo'] == target, column] = df.loc[df['Tipo'] == target, column] * -1

