import pandas as pd
import xlsxwriter
import re
import datetime
import sys


file_1 = sys.argv[1]
file_2 = sys.argv[2]
file_info = 'ISBN.xlsx'

date_regex = re.compile(r'(\d\d).(\d\d).(\d\d)')
date_1 = datetime.datetime.strptime(date_regex.search(file_1).group(), '%d.%m.%y').date().strftime('%d.%m.%Y')
date_2 = datetime.datetime.strptime(date_regex.search(file_2).group(), '%d.%m.%y').date().strftime('%d.%m.%Y')

xl_1 = pd.ExcelFile(file_1)
xl_2 = pd.ExcelFile(file_2)
xl_info = pd.ExcelFile(file_info)

df_1 = xl_1.parse(header=None, names=['name', 'count_1'], skiprows=6, skipfooter=1, usecols=[0,4], converters={0: str, 1: int})
df_2 = xl_2.parse(header=None, names=['name', 'count_2'], skiprows=6, skipfooter=1, usecols=[0,4], converters={0: str, 1: int})
df_info = xl_info.parse(header=None, names=['name', 'published', 'ISBN'], skiprows=1)

df_info.fillna({'ISBN': ''}, inplace=True)

df_res = df_1.merge(df_2, on='name', how='outer')
df_res.fillna(0, inplace=True)
df_res = df_res.merge(df_info, on='name', how='left')

df_res['diff'] = df_res.count_1 - df_res.count_2

with xlsxwriter.Workbook('test.xlsx') as workbook:
    worksheet = workbook.add_worksheet()

    workbook.formats[0].set_font_name('Arial')
    workbook.formats[0].set_font_size(8)
    h1_format = workbook.add_format(
        {'bold': True, 'font_name': 'Arial', 'font_color': 'green', 'font_size': 18, 'valign': 'vcenter'})
    h2_format = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 10})
    name_format = workbook.add_format({'font_name': 'Arial', 'font_size': 8, 'indent': 2})
    number_format = workbook.add_format({'num_format': '# ##0.000'})
    h2_number_format = workbook.add_format(
        {'bold': True, 'font_name': 'Arial', 'font_size': 10, 'num_format': '# ##0.000'})
    text_format = workbook.add_format({'num_format': '@'})
    was_nan_format = workbook.add_format({'font_name': 'Arial', 'font_size': 10, 'bg_color': '#addff7'})
    negative_format = workbook.add_format({'font_name': 'Arial', 'font_size': 10, 'bg_color': '#d8adf7'})
    diff_more_than_format = workbook.add_format({'font_name': 'Arial', 'font_size': 10, 'bg_color': '#adf7b0'})
    diff_less_than_format = workbook.add_format({'font_name': 'Arial', 'font_size': 10, 'bg_color': '#f7adad'})

    worksheet.set_column(0, 0, 60)
    worksheet.set_column(1, 1, 20)
    worksheet.set_column(2, 2, 25)
    worksheet.set_column(3, 6, 15)
    worksheet.set_row(1, 30)

    worksheet.write(1, 0, 'Остатки на складах', h1_format)

    worksheet.write(3, 0, 'Название', h2_format)
    worksheet.write(3, 1, 'Издательство', h2_format)
    worksheet.write(3, 2, 'ISBN', h2_format)
    worksheet.write(3, 3, date_1, h2_format)
    worksheet.write(3, 4, date_2, h2_format)
    worksheet.write(3, 5, 'Разница', h2_format)

    from_row = row_num = 4
    col_num = 0
    for row_index, row in df_res.iterrows():
        worksheet.write(row_num, col_num, row['name'], name_format)
        worksheet.write(row_num, col_num + 1, row['published'], text_format)
        worksheet.write(row_num, col_num + 2, row['ISBN'], text_format)
        worksheet.write(row_num, col_num + 3, row['count_1'], number_format)
        worksheet.write(row_num, col_num + 4, row['count_2'], number_format)
        worksheet.write(row_num, col_num + 5, row['diff'], number_format)

        row_num += 1

    worksheet.write(from_row + len(df_res), 0, 'Итого', h2_format)
    worksheet.write_formula(f'D{from_row + len(df_res) + 1}',
                            f'=SUM(D{from_row + 1}:D{from_row + len(df_res)})',
                            h2_number_format)
    worksheet.write_formula(f'E{from_row + len(df_res) + 1}',
                    f'=SUM(E{from_row + 1}:E{from_row + len(df_res)})',
                    h2_number_format)
    worksheet.write_formula(f'F{from_row + len(df_res) + 1}',
                            f'=SUM(F{from_row + 1}:F{from_row + len(df_res)})',
                            h2_number_format)

    worksheet.conditional_format(f'E5:E{from_row + len(df_res)}', {'type': 'cell',
                                                                   'criteria': '==',
                                                                   'value': 0,
                                                                   'format': was_nan_format})

    worksheet.conditional_format(f'D5:E{from_row + len(df_res)}', {'type': 'cell',
                                                                   'criteria': '<',
                                                                   'value': 0,
                                                                   'format': negative_format})

    worksheet.conditional_format(f'F5:F{from_row + len(df_res)}', {'type': 'cell',
                                                                   'criteria': '>=',
                                                                   'value': 10,
                                                                   'format': diff_more_than_format})
    worksheet.conditional_format(f'F5:F{from_row + len(df_res)}', {'type': 'cell',
                                                                   'criteria': '<',
                                                                   'value': 3,
                                                                   'format': diff_less_than_format})