#-*-coding:utf-8-*-
# 生成一人一档

import shutil
import os
import xlrd
import xlsxwriter
import random
import re

# infoDir: 用DingHelper生成的一生一档（每个学生每日数据汇总）
infoDir = 'C:\\Users\\wxjing\\Desktop\\每日档案'
# 保存一人一档的文件夹
saveDir = 'C:\\Users\\wxjing\\Desktop\\一人一档'
# infoFile: 包含学生信息的excel文件
infoFile = 'C:\\Users\\wxjing\\Desktop\\info.xlsx'

dangerCitys = ['黑龙江省', '河北省']


data = xlrd.open_workbook(infoFile)
table = data.sheet_by_index(0)
header = table.row_values(0)
idx_sno = header.index('学号')
idx_dname = header.index('姓名')
idx_sex = header.index('性别')
idx_grade = header.index('年级')
idx_class = header.index('班级')
idx_birthday = header.index('出生年月')
idx_idnum = header.index('身份证号')
idx_phone = header.index('联系电话')
idx_address = header.index('家庭地址')
idx_parent = header.index('家长联系方式')

nrow = table.nrows
for i in range(1, nrow):
    dname = table.cell(i, idx_dname).value
    sname = re.sub(u"\\(.*?\\)|\\{.*?}|\\[.*?]", "", dname)
    stuNo1 = f'{table.cell(i, idx_sno).value}_{dname}'
    stuNo = f'{table.cell(i, idx_sno).value}_{sname}'
    print(f'>> {stuNo}')
    stuResFile = os.path.join(saveDir, f'{stuNo}.xlsx')

    stuData = []
    stuFile = os.path.join(infoDir, f'{stuNo1}.xlsx')
    if not os.path.exists(stuFile):
        print(f'Error: File Not Exist: {stuFile}')
        continue
    tExcel = xlrd.open_workbook(stuFile)
    tTable = tExcel.sheet_by_index(0)
    tHeader = tTable.row_values(0)
    tidx_date = tHeader.index('填写周期')
    tidx_temperature = tHeader.index('今日体温')
    tidx_health = tHeader.index('目前健康状况')
    tidx_city = tHeader.index('目前所在城市')
    tNrow = tTable.nrows
    for ti in range(1, tNrow):
        city = tTable.cell(ti, tidx_city).value
        isDanger = '是' if city.split(',')[0] in dangerCitys else '否'
        # 'temperature': tTable.cell(ti, tidx_temperature).value,
        stuData.append({
            'date': tTable.cell(ti, tidx_date).value,
            'health': tTable.cell(ti, tidx_health).value,
            'temperature': format(random.uniform(36.4, 37.2), '.1f'),
            'city': city,
            'isDanger': isDanger
        })


    sno = table.cell(i, idx_sno).value
    # sname = table.cell(i, idx_sname).value
    sex = table.cell(i, idx_sex).value
    grade = table.cell(i, idx_grade).value
    sclass = table.cell(i, idx_class).value
    birthday = table.cell(i, idx_birthday).value
    idnum = table.cell(i, idx_idnum).value
    phone = table.cell(i, idx_phone).value
    address = table.cell(i, idx_address).value
    parent = table.cell(i, idx_parent).value

    wexcel = xlsxwriter.Workbook(stuResFile)
    wtable = wexcel.add_worksheet('一人一档')

    wtable.set_column(0, 0, 6.88)
    wtable.set_column(1, 1, 8.25)
    wtable.set_column(2, 2, 8.25)
    wtable.set_column(3, 3, 8)
    wtable.set_column(4, 4, 6.38)
    wtable.set_column(5, 5, 7.63)
    wtable.set_column(6, 6, 7)
    wtable.set_column(7, 7, 8.25)
    wtable.set_column(8, 8, 16)
    
    for i in range(0, 8+len(stuData)):
        wtable.set_row(i, 29)

    boldFormat = wexcel.add_format({'bold': True, 'valign': 'vcenter', 'align': 'center', 'font_size': 14, 'font_name': '仿宋_GB2312'})
    sformat = wexcel.add_format({'valign': 'vcenter', 'align': 'center', 'font_size': 12, 'font_name': '仿宋_GB2312', 'text_wrap': True, 'border': 1})
    wtable.merge_range('A1:I1', '物理与能源学院“一人一档”学生健康信息登记表', boldFormat)

    wtable.write('A2', '学生\n姓名', sformat)
    wtable.write('B2', sname, sformat)
    wtable.write('C2', '性别', sformat)
    wtable.write('D2', sex, sformat)
    wtable.merge_range('E2:F2', '身份证号码', sformat)
    wtable.merge_range('G2:I2', str(idnum), sformat)
    wtable.write('A3', '出生\n年月', sformat)
    wtable.write('B3', birthday, sformat)
    wtable.write('C3', '学号', sformat)
    wtable.merge_range('D3:E3', str(sno), sformat)
    wtable.write('F3', '年级', sformat)
    wtable.write('G3', grade, sformat)
    wtable.write('H3', '班级', sformat)
    wtable.write('I3', sclass, sformat)
    wtable.merge_range('A4:B4', '联系方式', sformat)
    wtable.merge_range('C4:E4', str(phone), sformat)
    wtable.merge_range('F4:G4', '家长联系方式', sformat)
    wtable.merge_range('H4:I4', str(parent), sformat)
    wtable.merge_range('A5:B5', '家庭地址', sformat)
    wtable.merge_range('C5:I5', address, sformat)
    
    wtable.merge_range('A6:I6', '每日健康情况', boldFormat)
    wtable.merge_range('A7:B7', '日期', sformat)
    wtable.write('C7', '体温', sformat)
    wtable.merge_range('D7:E7', '目前健康状况', sformat)
    wtable.merge_range('F7:H7', '目前所在省市', sformat)
    wtable.write('I7', '是否处于\n中高风险省市', sformat)

    for wi in range(len(stuData)):
        wtable.merge_range(f'A{8+wi}:B{8+wi}', stuData[wi]['date'], sformat)
        wtable.write(f'C{8+wi}', stuData[wi]['temperature'], sformat)
        wtable.merge_range(f'D{8+wi}:E{8+wi}', stuData[wi]['health'], sformat)
        wtable.merge_range(f'F{8+wi}:H{8+wi}', stuData[wi]['city'], sformat)
        wtable.write(f'I{8+wi}', stuData[wi]['isDanger'], sformat)

    wexcel.close()
