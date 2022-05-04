import openpyxl
import datetime
# читаем excel-файл
wb = openpyxl.load_workbook('schedule.xlsx')

week = {0: (6, 12), 1: (12, 18), 2: (18, 24), 3: (24, 30), 4: (30, 360), 5: (36, 42), 6: (42, 48)}
# печатаем список листов
sheets = wb.sheetnames
sheet = wb[sheets[0]]
cort = week[datetime.datetime.today().weekday()]
for i in range(*cort):
    cell = sheet[f'P{str(i)}']
    if str(cell.value) == 'None':
        print(f'Пара №{abs(cort[0] - i) + 1} -----------')
    else:
        print(f'Пара №{abs(cort[0] - i) + 1}', str(cell.value))