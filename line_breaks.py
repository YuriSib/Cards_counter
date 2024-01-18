import openpyxl
import os


def line_breaks(dir_path_):
    files = os.listdir(dir_path_)
    list_tables = [file for file in files if ".xlsx" in file]
    for table in list_tables:
        wb = openpyxl.load_workbook(f'{dir_path_}/{table}')
        ws = wb.active

        col_num = -1
        have_to_desc = False
        for col in ws.iter_cols(min_row=1, max_row=1, values_only=True):
            col_num += 1
            if 'писание' in col[0]:
                have_to_desc = True
                break
        if not have_to_desc:
            continue

        max_row = ws.max_row
        for row in range(2, max_row+1):
            old_value = ws.cell(row=row, column=col_num+1).value
            if not old_value:
                continue
            new_value = old_value.replace('•', ' <br> •')
            new_value = new_value[5:]
            # new_value = old_value[1:]
            ws.cell(row=row, column=col_num+1).value = new_value

        wb.save(rf'{dir_path_}\{table}')


def union_table(dir_path_):
    union_table_path = 'Объединение таблиц.xlsx'
    files = os.path.exists(rf'{dir_path_}\{union_table_path}')
    if not files:
        union_wb = openpyxl.Workbook()
    else:
        union_wb = openpyxl.load_workbook(rf'{dir_path_}\{union_table_path}')
    union_ws = union_wb.active

    files = os.listdir(dir_path_)
    list_tables = [file for file in files if ".xlsx" in file]

    row_to_start = union_ws.max_row + 1
    for table in list_tables:
        union_table_cols_list = [col[0] for col in union_ws.iter_cols(min_row=1, max_row=1, values_only=True) if col[0]]
        quantity_cols = len(union_table_cols_list)+1

        wb = openpyxl.load_workbook(f'{dir_path_}/{table}')
        ws = wb.active
        curr_table_cols_list = []
        for col in ws.iter_cols(min_row=1, max_row=1, values_only=True):
            if col[0]:
                column_name = col[0]
                if column_name[-1] == ' ':
                    column_name = column_name[:-1]
                curr_table_cols_list.append(column_name)



        for column_name in curr_table_cols_list:
            if column_name[-1] == ' ':
                column_name = column_name[:-1]
            if column_name in union_table_cols_list:
                continue
            else:
                union_ws.cell(row=1, column=quantity_cols).value = column_name
                union_table_cols_list.append(column_name)
                quantity_cols += 1

        for column_name in union_table_cols_list:
            if column_name in curr_table_cols_list:
                column_num = curr_table_cols_list.index(column_name)+1

                curr_row = row_to_start
                for values in ws.iter_cols(min_row=2, max_row=ws.max_row, min_col=column_num, max_col=column_num):
                    for value in values:
                        if value.value:
                            union_ws.cell(row=curr_row, column=column_num).value = value.value
                        curr_row += 1

        union_wb.save(rf'{dir_path_}\{union_table_path}')


if __name__ == "__main__":
    # line_breaks(r'C:\Users\Administrator\Desktop\Таблицы\Готовые к загрузке')
    union_table(r'C:\Users\Administrator\Desktop\Таблицы\Готовые к загрузке')
