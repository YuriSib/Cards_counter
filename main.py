import os
import openpyxl
from datetime import datetime


def card_counter(path):
    path_to_count = r'C:\Users\Administrator\Desktop\Таблицы\учет карточек.xlsx'
    table_list = [table for table in os.listdir(path) if '.xlsx']

    wb_cnt = openpyxl.load_workbook(path_to_count)
    ws_cnt = wb_cnt.active

    num_rows_in_cnt = ws_cnt.max_row
    category_list = [ws_cnt.cell(row=row+1, column=1).value for row in range(1, num_rows_in_cnt)]

    for table in table_list:
        if '~$' in table or '.ini' in table:
            continue
        path_to_table = rf'{path}\{table}'
        wb = openpyxl.load_workbook(path_to_table)
        ws = wb.active

        column_num = 1
        for column in range(1, 8):
            desc_column = ws.cell(row=1, column=column).value
            if 'писание' in desc_column:
                column_num = column
                break

        num_rows = ws.max_row

        row_num = 0
        for row in range(2, num_rows):
            desc_row = ws.cell(row=row, column=column_num).value
            if desc_row:
                row_num += 1

        category_name = table.replace('.xlsx', '')
        if 'Олеся' in path:
            cnt_column = 1
        elif 'Вика' in path:
            cnt_column = 2

        last_modified_time = os.stat(rf'{path}\{table}').st_mtime
        date = datetime.fromtimestamp(last_modified_time).strftime('%d.%m.%Y')
        if category_name in category_list:
            idx = category_list.index(category_name) + 2
            ws_cnt[idx][cnt_column].value = row_num
            ws_cnt[idx][3].value = date
        else:
            idx = num_rows_in_cnt + 1
            ws_cnt[idx][cnt_column].value = row_num
            ws_cnt[idx][0].value = category_name
            num_rows_in_cnt += 1
            ws_cnt[idx][3].value = date

        wb_cnt.save(path_to_count)


if __name__ == "__main__":
    path_to_cards_1 = r'C:\Users\Administrator\Desktop\Общая папа\Олеся'
    path_to_cards_2 = r'C:\Users\Administrator\Desktop\Общая папа\Вика'
    card_counter(path_to_cards_1)
    card_counter(path_to_cards_2)
