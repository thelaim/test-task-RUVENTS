from openpyxl import load_workbook
from datetime import datetime
import calendar

wb = load_workbook('task_support.xlsx')
sheet_ranges = wb['Tasks']

num = list(range(3, 1003))

b = 0
for cell in num:
    if int(sheet_ranges[f'B{cell}'].value) % 2 == 0:
        b += 1
c = 0
for cell in num:
    n = int(sheet_ranges[f'C{cell}'].value)
    counter = 0
    for i in range(1, n + 1):
        if n % i == 0:
            counter += 1
    if counter == 2:
        c += 1

d = 0
for cell in num:
    if int(sheet_ranges[f'D{cell}'].value.replace(' ', '')[2]) < 5:
        d += 1

e = 0
for cell in num:
    if sheet_ranges[f'E{cell}'].value[0:3] == 'Tue':
        e += 1

f = 0
for cell in num:
    if datetime.strptime(f"{sheet_ranges[f'F{cell}'].value[0:10]}", "%Y-%m-%d").isoweekday() == 2:
        f += 1
    

print('Четных чисел:', b)
print('Простых чисел:', c)
print('Чисел меньше 0.5:', d)
print('Число втоников в E:', e)
print('Число втоников в F:', f)
