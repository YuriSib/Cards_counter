import openpyxl

wb = openpyxl.load_workbook('Оружие.xlsx')
ws = wb.active

result = ws.iter_rows()
print(result)
