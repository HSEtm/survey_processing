from os import listdir, makedirs
from os.path import isfile, join, splitext, exists
import pandas as pd


# Method for processing initial Excel survey data
def process_folder(folder_path, sheet_ns):
    paths = [f for f in listdir(folder_path) if isfile(join(folder_path, f))]
    d = dict()
    for path in paths:
        filename = join(folder_path, path)
        for sheet_n in sheet_ns:
            dataf = pd.read_excel(io=filename, sheetname=sheet_n, skiprows=range(0, 3))
            dataf = dataf[dataf.columns[0:2]][~dataf[dataf.columns[1]].isnull()]
            dataf.columns = ('product', 'indicator')
            # TODO add other column processing in future
            try:
                d[sheet_n].append(dataf)
            except:
                d[sheet_n] = []
                d[sheet_n].append(dataf)
    return d


# Method for processing dictionaries
def process_dics(folder_path, sheet_n):
    paths = [f for f in listdir(folder_path) if isfile(join(folder_path, f))]
    d = dict()
    for path in paths:
        filename = join(folder_path, path)
        dataf = pd.read_excel(io=filename, sheetname=sheet_n, header=0)
        dataf.columns = ('product', 'group')
        d[splitext(path)[0]] = dataf
    return d


# Method for scoring results based on mean value
def scoring(x, **kwargs):
    if x <= 1.75:
        return kwargs['score_types'][0]
    elif x < 2.5:
        return kwargs['score_types'][1]
    elif x < 3.5:
        return kwargs['score_types'][2]
    elif x < 4.5:
        return kwargs['score_types'][3]
    else:
        return kwargs['score_types'][4]


# Excel sheets of survey that need to be processed
sheet_names = ('Блок 1. Конкурентоспособность', 'Блок 2. НТ уровень', 'Блок 3. Импортозамещение', 'Блок 4. ИиР')
# Scoring values
scores = dict()
scores['Блок 1. Конкурентоспособность'] = ['1 - низкая', '2 - средняя', '3 - высокая']
scores['Блок 2. НТ уровень'] = ['1 - низкий', '2 - средний', '3 - высокий']
scores['Блок 3. Импортозамещение'] = ['1 - отсутствуют и не могут быть созданы в ближайшие 5 лет',
                                      '2 - отсутствуют, но могут быть созданы в ближайшие 5 лет',
                                      ('3 - отдельные элементы технологической цепочки производства, ' +
                                       'производство полного цикла может быть создано в ближайшие 1-2 года'),
                                      '4 - полная технологическая цепочка для производства импортозамещающей продукции',
                                      '5 - объем производимой в РФ продукции соизмерим или превосходит объем импорта']
scores['Блок 4. ИиР'] = ['1 - «Белые пятна»', '2 - Зоны кооперации', '3 - Лидерство']
# Folder with data
data = process_folder('data', sheet_names)
# Dictionaries with description of objects and their categories, file names must be equal to survey sheets
dics = process_dics('dics', 'Лист1')
# Folder and file name for results
output_folder = 'output'
writer = pd.ExcelWriter(path=join(output_folder, 'photonics_results.xlsx'), engine='openpyxl')

# Loop for processing very similar Excel survey sheets. If data structure is different, additional work is needed
for sheet_name in sheet_names:
    df = pd.concat(data[sheet_name])
    df_merged = pd.merge(left=df, right=dics[sheet_name], how='inner', on='product')
    df_merged['count'] = 1
    df_merged['measure'] = df_merged['indicator'].str[0].astype(int)
    df_grouped = df_merged.groupby(['group', 'product', 'indicator', 'measure'])[
        'count'].count().to_frame().reset_index()
    df_grouped['experts'] = df_grouped.groupby(['group', 'product'])['count'].transform('sum')
    df_grouped['score'] = df_grouped['measure'] * (df_grouped['count'] / df_grouped['experts'])
    df_result = df_grouped.groupby(['group', 'product', 'experts']).score.sum().to_frame().reset_index()
    df_result['result'] = df_result['score'].apply(scoring, score_types=scores[sheet_name])
    df_result.to_excel(excel_writer=writer, sheet_name=sheet_name, startrow=0)

if not exists(output_folder):
    makedirs(output_folder)
writer.save()
