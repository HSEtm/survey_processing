from os import listdir
from os.path import isfile, join
from openpyxl import load_workbook
import pandas as pd


# sheet_names = ('Блок 1. Конкурентоспособность', 'Блок 2. НТ уровень', 'Блок 3. Импортозамещение', 'Блок 4. ИиР')
# frames = process_folder('data', sheet_names)
# dics = process_dics('dics', 'Лист1')
# frames[0][0][frames[0][0].columns[0]]
# df = pd.read_excel(io='data/Анкета_фотоника.xls', sheetname='Блок 1. Конкурентоспособность', header=1)
# df[~df[df.columns[1]].isnull()]

def process_folder(folder_path, sheet_names):
    paths = [f for f in listdir(folder_path) if isfile(join(folder_path, f))]
    frames = []
    for path in paths:
        filename = join(folder_path, path)
        sheets = []
        for sheet_name in sheet_names:
            df = pd.read_excel(io=filename, sheetname=sheet_name, header=1)
            df = df[df.columns[0:4]][~df[df.columns[1]].isnull()]
            df.columns = (1, 2, 3, 4)
            sheets.append(df)
        frames.append(sheets)
    return frames


def process_dics(folder_path, sheet_name):
    paths = [f for f in listdir(folder_path) if isfile(join(folder_path, f))]
    dics = []
    for path in paths:
        filename = join(folder_path, path)
        df = pd.read_excel(io=filename, sheetname=sheet_name, header=0)
        df.columns = (1, 2)
        dics.append(df)
    return dics


def export_to_file(file_path, frames, dics, sheet_names):
    # http://stackoverflow.com/questions/20219254/how-to-write-to-an-existing-excel-file-without-overwriting-data-using-pandas
    book = load_workbook(file_path)
    writer = pd.ExcelWriter(path=file_path, engine='xlsxwriter')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    for frame in frames:
        for x in range(0, len(sheet_names)):
            frame[x].to_excel(excel_writer=writer, sheet_name=sheet_names[x], startrow=0)
    writer.save()
