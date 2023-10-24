import os
import json
import pandas as pd


result = []
new_df = pd.DataFrame(
    columns=[
        'Папка',
        'Файл',
        'Значение',
        'Сумма'
    ]
)
itog_df = pd.DataFrame()


def open_json():
    with open('config.json', 'r', encoding='utf-8') as f:
        global text
        global path
        data = json.load(f)
        text = data['Значение']
        path = data['Путь к папке']
    try:
        os.remove(f'Итоги по файлам{text}.xlsx')
        os.remove(f'Итоги по папкам{text}.xlsx')
    except Exception:
        pass
    find_excel(path)


def find_excel(path):
    for i in text:
        itog_df[i.upper()] = ''
    for folder in os.walk(path):
        folder_name = folder[0]
        for list_files in folder:
            for file in list_files:
                if file.lower().endswith('.xls') or file.lower().endswith('.xlsx'):
                    full_path_file = f'{folder[0]}/{file}'
                    check_sheet_excel(full_path_file, file, folder_name)
    new_df.to_excel(f'Итоги по файлам{text}.xlsx', sheet_name='Итог')
    table_summ_sort_by_folder(new_df)
    itog_df.to_excel(f'Итоги по папкам{text}.xlsx', sheet_name='Итог')


def check_sheet_excel(full_path_file, file, folder_name):
    try:
        workbook = pd.ExcelFile(full_path_file)
        if 'Смета по ТСН-2001' in list(workbook.sheet_names):
            # print(workbook.sheet_names)
            process_file(workbook, full_path_file, file, folder_name)
        else:
            process_file_empty(workbook, full_path_file, file, folder_name)
    except Exception:
        pass


def process_file_empty(workbook, full_path_file, file, folder_name):
    global new_df
    frames = [new_df]
    try:
        for i in text:
            name_file = []
            search_item = []
            name_file.append((full_path_file))
            search_item.append(i.upper())
            record_df = pd.DataFrame({
                'Папка': folder_name,
                'Файл': file,
                'Значение': search_item,
                'Сумма': 0,
            })
            frames.append(record_df)
        new_df = pd.concat(frames)
    except Exception as e:
        print(e)


def process_file(workbook, full_path_file, file, folder_name):
    global new_df
    frames = [new_df]
    try:
        df = workbook.parse('Смета по ТСН-2001')
        if 'Unnamed: 2' in df.columns:
            df.rename(columns={'Unnamed: 2': 'ColC'}, inplace=True)
            # df['ColC'].str.strip()
        if 'Unnamed: 10' in df.columns:
            df.rename(columns={'Unnamed: 10': 'ColK'}, inplace=True)
        else:
            if 'Форма № 1б' in df.columns:
                df.rename(columns={'Форма № 1б': 'ColK'}, inplace=True)
        df['ColC'] = df['ColC'].str.strip()
        df['ColC'] = df['ColC'].str.upper()
        a = df['ColC'].tolist()
        print(a)
        for i in text:
            name_file = []
            search_item = []
            final_sum = []
            list_s = (df.query(f'ColC == "{i.upper()}"')['ColK']).tolist()
            sum_list = round(sum(list_s), 2)
            name_file.append((full_path_file))
            search_item.append(i.upper())
            final_sum.append(sum_list)
            record_df = pd.DataFrame({
                'Папка': folder_name,
                'Файл': file,
                'Значение': search_item,
                'Сумма': final_sum,
            })
            frames.append(record_df)
        new_df = pd.concat(frames)
    except Exception as e:
        # print(full_path_file)
        print(e)


def table_summ_sort_by_folder(table):
    global itog_df
    frames = [itog_df]
    for i in table.Папка.unique():
        values = {}
        folder_name = []
        folder_name.append(i)
        for j in text:
            sum_list = round(table.loc[(table['Папка'] == i) & (table['Значение'] == j.upper()), 'Сумма'].sum(), 2)
            values[j] = sum_list

        record_df = pd.DataFrame(values, index=[i])
        frames.append(record_df)
        itog_df = pd.concat(frames)


open_json()
