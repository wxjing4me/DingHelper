# -*-coding:utf-8-*-
# 生成一人一档
import re
import json
import random
import xlsxwriter
import xlrd
import os
import shutil
import requests
import sys
sys.path.append("..")
from configure.administrative_division import ADMINIS_DIVISION
from configure.config_values import AMAP_API_URL_LL2Address
from configure.danger_place import DANGER_PLACES

OPERATE_DIR = os.path.join('C:', 'Users', '哇咔咔', 'Desktop', '生成一人一档')

# dingDir: 钉钉每日健康打卡的数据
dingDir = os.path.join(OPERATE_DIR, '1-钉钉健康打卡数据')
# infoFile: 包含学生信息的excel文件【学号，姓名，年级，专业班级，身份证号码，联系电话，家庭地址，家长联系方式】
infoFile = os.path.join(OPERATE_DIR, '2-学生信息.xlsx')
# 保存一人一档的文件夹
saveDir = os.path.join(OPERATE_DIR, '3-一人一档')

def getCity(address):
    city = '-'
    addr = '-'
    if address != '' and address != '-':
        _, longitude, latitude, addr, _ = eval(address)
        flag = False
        for dist in ADMINIS_DIVISION:
            length = len(dist)
            if addr[:length] == dist:
                city = dist
                flag = True
        if not flag:
            print(f'>> {addr} NOT FOUND')
            try:
                response = requests.get(
                    f'{AMAP_API_URL_LL2Address}location={longitude},{latitude}&key={apiKey}')
                if response.status_code == 200:
                    res = json.loads(response.text)
                    if int(res['status']) == 1:
                        address = res['regeocode']['addressComponent']
                        city = f"{address['province']}{address['city']}{address['district']}"
            except:
                city = addr
    return city

def getStuInfo(infoFile):

    stuInfos = {}

    print(f'>> 处理学生信息文件: {infoFile}')
    excel = xlrd.open_workbook(infoFile)
    table = excel.sheet_by_index(0)
    header = table.row_values(0)
    idx_sno = header.index('学号')
    idx_sname = header.index('姓名')
    idx_grade = header.index('年级')
    idx_class = header.index('专业班级')
    idx_idnum = header.index('身份证号码')
    idx_phone = header.index('联系电话')
    idx_address = header.index('家庭地址')
    idx_parent = header.index('家长联系方式')

    nrow = table.nrows
    for i in range(1, nrow):

        aStuInfo = {}

        sno = table.cell(i, idx_sno).value
        aStuInfo['sname'] = table.cell(i, idx_sname).value
        aStuInfo['grade'] = table.cell(i, idx_grade).value
        aStuInfo['sclass'] = table.cell(i, idx_class).value
        idnum = table.cell(i, idx_idnum).value.strip()
        aStuInfo['idnum'] = idnum
        aStuInfo['gender'] = '男' if int(idnum[-2]) % 2 else '女'
        aStuInfo['birthday'] = f'{idnum[6:10]}.{idnum[10:12]}'
        aStuInfo['phone'] = table.cell(i, idx_phone).value
        aStuInfo['address'] = table.cell(i, idx_address).value
        aStuInfo['parent'] = table.cell(i, idx_parent).value

        stuInfos[sno] = aStuInfo

    return stuInfos


def getDatasFromDing(dingDir, stuInfos):

    datas = {}
    dates = []
    if not os.path.exists(dingDir):
        print(f'路径不存在：{dingDir}')
        return dates, datas

    excel_list = []
    for filename in os.listdir(dingDir):
        excel_list.append(os.path.join(dingDir, filename))

    count = 0
    total = len(excel_list)
    for excel_path in excel_list:
        count += 1
        excel = xlrd.open_workbook(excel_path)
        print(f'正在处理 {count}/{total}：{excel_path}')
        table = excel.sheet_by_index(0)
        nrows = table.nrows
        header = table.row_values(0)
        
        idx_sno = header.index('工号')
        idx_sname = header.index('提交人')
        idx_date = header.index('填写周期')
        idx_temperature = header.index('今日体温')
        if '今日有无以下症状' in header:
            idx_health = header.index('今日有无以下症状')
        else:
            idx_health = header.index('近两日有无以下症状')
        idx_address = header.index('当前时间,当前地点')

        for i in range(1, nrows):
            sno = table.cell(i, idx_sno).value
            if sno not in datas.keys():
                datas[sno] = {}
            date = table.cell(i, idx_date).value
            if date not in dates:
                dates.append(date)
            address = table.cell(i, idx_address).value
            temperature = table.cell(i, idx_temperature).value
            # temperature = format(random.uniform(36.4, 37.2), '.1f')
            city = getCity(address)
            isDanger = '是' if city in DANGER_PLACES else '否'

            aStuDateData = {
                'temperature': temperature,
                'health': table.cell(i, idx_health).value,
                'city': city,
                'isDanger': isDanger
            }

            datas[sno][date] = aStuDateData

    # 补充遗漏的人员数据
    for sno in stuInfos.keys():
        if sno not in datas.keys():
            print(f"缺少人员: {sno} {stuInfos[sno]['sname']}")
            datas[sno] = {}

    # 补充遗漏的日期数据
    undo = '-'
    undoData = {
        'temperature': undo,
        'health': undo,
        'city': undo,
        'isDanger': undo
    }
    for sno, sData in datas.items():
        for date in dates:
            if date not in sData.keys():
                datas[sno][date] = undoData

    dates = sorted(dates)

    return dates, datas


def generateEveryoneProfile(saveDir, stuInfos, dates, stuDatas):

    for sno, sinfo in stuInfos.items():
        print(f">> {sno} {sinfo['sname']}")
        stuResFile = os.path.join(saveDir, f"{sno}-{sinfo['sname']}.xlsx")

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

        studata = stuDatas[sno]  # {date: {...} }
        for i in range(0, 8+len(studata.keys())):
            wtable.set_row(i, 29)

        boldFormat = wexcel.add_format(
            {'bold': True, 'valign': 'vcenter', 'align': 'center', 'font_size': 14, 'font_name': '仿宋_GB2312'})
        sformat = wexcel.add_format({'valign': 'vcenter', 'align': 'center',
                                     'font_size': 12, 'font_name': '仿宋_GB2312', 'text_wrap': True, 'border': 1})
        wtable.merge_range('A1:I1', '物理与能源学院“一人一档”学生健康信息登记表', boldFormat)

        wtable.write('A2', '姓名', sformat)
        wtable.write('B2', sinfo['sname'], sformat)
        wtable.write('C2', '性别', sformat)
        wtable.write('D2', sinfo['gender'], sformat)
        wtable.merge_range('E2:F2', '身份证号码', sformat)
        wtable.merge_range('G2:I2', str(sinfo['idnum']), sformat)
        wtable.write('A3', '出生\n年月', sformat)
        wtable.write('B3', sinfo['birthday'], sformat)
        wtable.write('C3', '学号', sformat)
        wtable.merge_range('D3:E3', str(sno), sformat)
        wtable.write('F3', '年级', sformat)
        wtable.write('G3', sinfo['grade'], sformat)
        wtable.write('H3', '班级', sformat)
        wtable.write('I3', sinfo['sclass'], sformat)
        wtable.merge_range('A4:B4', '联系方式', sformat)
        wtable.merge_range('C4:E4', str(sinfo['phone']), sformat)
        wtable.merge_range('F4:G4', '家长联系方式', sformat)
        wtable.merge_range('H4:I4', str(sinfo['parent']), sformat)
        wtable.merge_range('A5:B5', '家庭地址', sformat)
        wtable.merge_range('C5:I5', sinfo['address'], sformat)

        wtable.merge_range('A6:I6', '每日健康情况', boldFormat)
        wtable.merge_range('A7:B7', '日期', sformat)
        wtable.write('C7', '体温', sformat)
        wtable.merge_range('D7:E7', '目前健康状况', sformat)
        wtable.merge_range('F7:H7', '目前所在省市', sformat)
        wtable.write('I7', '是否处于\n中高风险省市', sformat)

        wi = 8
        for date in dates:
            wtable.merge_range(f'A{wi}:B{wi}', str(date), sformat)
            wtable.write(f'C{wi}', str(studata[date]['temperature']), sformat)
            wtable.merge_range(
                f'D{wi}:E{wi}', studata[date]['health'], sformat)
            wtable.merge_range(f'F{wi}:H{wi}', studata[date]['city'], sformat)
            wtable.write(f'I{wi}', studata[date]['isDanger'], sformat)
            wi += 1

        wexcel.close()


if __name__ == '__main__':
    if os.path.exists(saveDir):
        shutil.rmtree(saveDir)
    os.makedirs(saveDir)

    stuInfos = getStuInfo(infoFile)
    dates, stuDatas = getDatasFromDing(dingDir, stuInfos)
    generateEveryoneProfile(saveDir, stuInfos, dates, stuDatas)
