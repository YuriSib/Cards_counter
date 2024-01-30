from datetime import datetime
import os

path = fr'C:\Users\Administrator\Desktop\Общая папа\Вика'
date_format = "%d.%m.%y"


table_list = [table for table in os.listdir(path) if '.xlsx' in table and '~$' not in table]
date_table_list = [(table, datetime.strptime(table.split('(')[1].replace(").xlsx", '').replace(" ", ''), date_format)) for table in table_list]

date_string_list = [table.split('(')[1].replace(").xlsx", '').replace(" ", '') for table in table_list]
date_list = [datetime.strptime(date_string, date_format).date() for date_string in date_string_list]

sorted_date_list = sorted(date_list)
sorted_date_str_list = [date.strftime(date_format) for date in sorted_date_list]
sorted_unique_list = []
for item in sorted_date_str_list:
    if item not in sorted_unique_list:
        sorted_unique_list.append(item)


print(date_list)
print(sorted_date_list)
