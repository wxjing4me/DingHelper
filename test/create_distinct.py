# 根据省市区行政规划，生成省市区列表
import os
import xlrd

filepath= os.path.join('C:', 'Users', '哇咔咔', 'Desktop')
filename = os.path.join(filepath, '省市区.xlsx')

excel = xlrd.open_workbook(filename)
table = excel.sheet_by_index(0)
nrows = table.nrows
province = ''
city = ''
district = ''
data = []
for i in range(nrows):
    col1 = table.cell(i, 0).value
    col2 = table.cell(i, 1).value
    col3 = table.cell(i, 2).value
    if col1 != '':
        province = col1
        city = ''
        district = ''
    if col2 != '':
        city = col2
        district = ''
    if col3 != '':
        district = col3
    data.append(f'{province}{city}{district}')

print(data)