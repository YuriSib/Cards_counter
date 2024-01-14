import openpyxl
import os


def br_set(dir_path_):
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
            break

        max_row = ws.max_row
        for row in range(2, max_row+1):
            old_value = ws.cell(row=row, column=col_num+1).value
            if not old_value:
                continue
            new_value = old_value.replace('•', '<br>•')
            ws.cell(row=row, column=col_num+1).value = new_value
        wb.save(rf'{dir_path_}\{table}')


if __name__ == "__main__":
    br_set(r'C:\Users\Administrator\Desktop\Таблицы\Готовые к загрузке')
