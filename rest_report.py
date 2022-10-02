import os
import pandas as pd
import xlsxwriter

LAST_N_DAYS = 14

def parse_file_name(file_name):
    spl = file_name.split('.')[0].split('_')

    return {
        'file_name': file_name,
        'type': spl[0],
        'number': int(spl[1])
    }


cwd = os.getcwd()
# os.chdir("reports/")
# os.chdir("..")
files = list(map(parse_file_name, [x for x in os.listdir('reports/') if x.endswith('.xlsx')]))
sales_files = sorted(list(filter(lambda x: x['type'] == 'продажи', files)), key=lambda x: x['number'])
remains_files = list(filter(lambda x: x['type'] == 'остатки', files))
file_info = 'isbn/ISBN.xlsx'

sales_length = len(sales_files)


xl_sales = []
for s_f in sales_files:
    xl_sales.append(pd.ExcelFile(f"reports/{s_f['file_name']}"))

remains_file = [x for x in remains_files if x['number'] == len(xl_sales)][0]['file_name']
xl_remains = pd.ExcelFile(f"reports/{remains_file}")
xl_info = pd.ExcelFile(file_info)


def normalize_name(name):
    index = name.find(', шт')
    n = name[:index] if index != -1 else name
    return n.strip()


dfs_sales = []

for s_index, xl_s in enumerate(xl_sales):
    s_index += 1
    df = xl_s.parse(header=None,
                    names=['name', f"sales_count_{s_index}", f"sales_amount_{s_index}",
                           f"sales_discount_amount_{s_index}", f"avg_price_{s_index}", f"avg_discount_price_{s_index}"],
                    skiprows=15, skipfooter=1, usecols=[0, 3, 4, 5, 6, 7],
                    converters={0: str, 1: int, 2: float, 3: float, 4: float, 5: float})
    df.name = df.name.apply(normalize_name)
    dfs_sales.append(df)

df_remains = xl_remains.parse(header=None, names=['name', 'remains_count'], skiprows=6, skipfooter=1, usecols=[0, 4],
                              converters={0: str, 1: int})
df_remains.name = df_remains.name.apply(normalize_name)
df_info = xl_info.parse(header=None, names=['name', 'published', 'ISBN'], skiprows=1)
df_info.name = df_info.name.apply(normalize_name)


def merge_sales_dfs(dfs_ss):
    l = len(dfs_ss)

    df_sales = dfs_ss[0]
    if l > 1:
        for i in range(1, len(dfs_ss)):
            df_sales = df_sales.merge(dfs_ss[i], on='name', how='outer')

    return df_sales.fillna(0)


df_sales = merge_sales_dfs(dfs_sales)
df_sales = df_sales.merge(df_remains, on='name', how='outer')
df_sales.fillna(0, inplace=True)
df_info.fillna('', inplace=True)
df_sales = df_sales.merge(df_info, on='name', how='left')
df_sales = df_sales.reindex(
    columns=['comments', 'name', 'published', 'ISBN']
            + [f"sales_count_{x + 1}" for x in range(sales_length)]
            + ['average_sales_count', 'remains_count']
)
df_sales.fillna('', inplace=True)
df_sales.drop_duplicates(inplace=True)

def to_letter_number(row, col):
    return f'{xlsxwriter.utility.xl_col_to_name(col)}{row + 1}'


font_name = {'font_name': 'Arial'}
bold = {'bold': True}
font_s = font_name | {'font_size': 12}
font_m = font_name | {'font_size': 14}
font_l = font_name | {'font_size': 16}
v_align = {'valign': 'vcenter'}
align_center = {'align': 'center'}
align_left = {'align': 'left'}
align_right = {'align': 'right'}
indent_1 = {'indent': 1}
text_wrap = {'text_wrap': True}
number = {'num_format': '###,###,###,##0'}
border_top = {'top': 1}
border_bottom = {'bottom': 1}
border_left = {'left': 1}
border_right = {'right': 1}


with xlsxwriter.Workbook(f'created_reports/Отчет {sales_length} недели.xlsx') as workbook:
    worksheet = workbook.add_worksheet()

    logo_header_format = workbook.add_format(
        font_l | bold | align_left  | {'num_format': '@'})
    caption_format = workbook.add_format(
        font_m | align_left | {'num_format': '@'})
    header_first_cell_format = workbook.add_format(v_align | text_wrap | indent_1 | border_top | border_bottom | border_left)
    header_mid_cell_format = workbook.add_format(border_top | border_bottom)
    header_last_cell_format = workbook.add_format(border_top | border_bottom | border_right)
    table_header_format = workbook.add_format(font_m | v_align | align_center | text_wrap | {'num_format': '@'})
    text_format = workbook.add_format(font_s | v_align | align_left | indent_1 | {'num_format': '@'})
    number_format = workbook.add_format(font_s | v_align | align_center | number)
    avg_sales_format = workbook.add_format(font_s | bold | v_align | align_center | number | {'bg_color': 'yellow'})
    remains_format = workbook.add_format(font_s | bold | v_align | align_center | number | {'bg_color': '#91bbd9'})
    remains_weeks_format = workbook.add_format(
        font_s | bold | v_align | align_center | number | {'bg_color': '#b5d991'})
    # was_nan_format = workbook.add_format({'font_name': 'Arial', 'font_size': 10, 'bg_color': '#addff7'})
    # negative_format = workbook.add_format({'font_name': 'Arial', 'font_size': 10, 'bg_color': '#d8adf7'})
    # diff_more_than_format = workbook.add_format({'font_name': 'Arial', 'font_size': 10, 'bg_color': '#adf7b0'})
    # diff_less_than_format = workbook.add_format({'font_name': 'Arial', 'font_size': 10, 'bg_color': '#f7adad'})

    t_from_row = 3
    t_from_col = 0


    sales_columns = []

    for x in range(sales_length):
        sales_columns.append(
            {'header': f'Продажи\n{x + 1} нед',
             'format': number_format,
             'total_function': 'sum'}
        )

    table_options = {
        'data': df_sales.sort_values(by=['name']).to_numpy(),
        'name': 'Rest',
        'total_row': True,
        'columns': [
            {'header': 'Комментарии',
             'format': text_format},
            {'header': 'Название',
             'format': text_format,
             'total_string': 'Тотал'},
            {'header': 'Издательство',
             'format': text_format},
            {'header': 'ISBN',
             'format': text_format},
            *sales_columns,
            {'header': 'Cредние\nпродажи',
             'format': avg_sales_format,
             'total_function': 'sum'},
            {'header': 'Количество\nна складе',
             'format': remains_format,
             'total_function': 'sum'},
            {'header': 'Остаток\n(нед.)',
             'format': remains_weeks_format,
             'total_function': 'std_dev'}
        ],
        'banded_rows': False,
        'style': 'Table Style Light 15'
    }

    columns_count = len(table_options['columns'])

    avg_sales_count_col = [i for i, x in enumerate(table_options['columns']) if x['header'] == 'Cредние\nпродажи'][
        0]

    worksheet.add_table(
        f'{to_letter_number(3, 0)}:{to_letter_number(len(df_sales.values) + 3, len(table_options["columns"]) - 1)}',
        table_options)

    # avg sales
    for x in range(3 + 1, len(df_sales.values) + 3):
        worksheet.write_formula(to_letter_number(x, avg_sales_count_col),
                                f"=ROUND(AVERAGE({to_letter_number(x, 4 + sales_length - LAST_N_DAYS)}:{to_letter_number(x, 4 + sales_length - 1)}), 0)",
                                avg_sales_format)
    # remains weeks
    for x in range(3 + 1, len(df_sales.values) + 3):
        worksheet.write_formula(to_letter_number(x, avg_sales_count_col + 2),
                                f"=IFERROR(FLOOR({to_letter_number(x, avg_sales_count_col + 1)}/{to_letter_number(x, avg_sales_count_col)}, 1), {to_letter_number(x, avg_sales_count_col + 1)})",
                                remains_weeks_format)


    for x in range(3, len(df_sales.values) + 3 + 1):
        worksheet.set_row(x, 30)

    # Header row
    worksheet.insert_image(to_letter_number(1, 0), 'logo.png')
    worksheet.write_rich_string(1, 1, logo_header_format, 'ООО «Поляндрия NoAge»\n', caption_format, 'www.polyandria.ru', header_first_cell_format)
    for x in range(2, columns_count - 2):
        worksheet.write(1, x, '', header_mid_cell_format)
    worksheet.write(1, columns_count - 2, '', header_last_cell_format)

    worksheet.set_row(1, 49.5)
    worksheet.set_row(3, 40, table_header_format)

    worksheet.set_column(0, 0, 30)
    worksheet.set_column(1, 1, 60)
    worksheet.set_column(2, 3, 20)
    worksheet.set_column(4, 4 + sales_length - LAST_N_DAYS - 1, None, None, {'hidden': True})
    worksheet.set_column(4 + sales_length - LAST_N_DAYS, len(table_options['columns']) - 1, 15)



    # worksheet.conditional_format(f'E5:E{from_row + len(df_res)}', {'type': 'cell',
    #                                                                'criteria': '==',
    #                                                                'value': 0,
    #                                                                'format': was_nan_format})
    #
    # worksheet.conditional_format(f'D5:E{from_row + len(df_res)}', {'type': 'cell',
    #                                                                'criteria': '<',
    #                                                                'value': 0,
    #                                                                'format': negative_format})
    #
    # worksheet.conditional_format(f'F5:F{from_row + len(df_res)}', {'type': 'cell',
    #                                                                'criteria': '>=',
    #                                                                'value': 10,
    #                                                                'format': diff_more_than_format})
    # worksheet.conditional_format(f'F5:F{from_row + len(df_res)}', {'type': 'cell',
    #                                                                'criteria': '<',
    #                                                                'value': 3,
    #                                                                'format': diff_less_than_format})