#-*-coding:utf-8-*-
from xlrd import open_workbook as xlrd_open_workbook
import xlwings as xw
from PyQt5.QtCore import QThread, QObject, pyqtSignal, pyqtSlot

# 第1行默认为表头
START_ROW = 1 # default:1

'''
判断Excel格式是否符合【每日健康打卡】
'''
def testRawExcel(excel_path):
    res = {}
    res['code'] = 0
    res['msg'] = ''
    try:
        excel = xlrd_open_workbook(excel_path)
        sht_names = excel.sheet_names()
        for sht_name in sht_names:
            table = excel.sheet_by_name(sht_name)
            header = table.row_values(0)
            needed = ['提交人', '工号', '填写周期', '当前时间,当前地点']
            if len(header) != 0 and set(header) >= set(needed):
                res['code'] = 1
            elif len(header) == 0:
                res['msg'] = f'请删除空表：{excel_path}<{sht_name}>后重试'
            else:
                res['msg'] = f'{excel_path}<{sht_name}>中缺少{needed}的部分字段'
    except Exception as e:
        res['msg'] = f'测试表格出错了：{e}'
    finally:
        return res

def testSpectExcel(excel_path):
    res = {}
    res['code'] = 0
    res['msg'] = ''
    try:
        excel = xlrd_open_workbook(excel_path)
        sht_names = excel.sheet_names()
        for sht_name in sht_names:
            table = excel.sheet_by_name(sht_name)
            date_header = table.row_values(0)[1:]
            #TODO:判断表头格式符合【x月x日】的格式
            if len(date_header) > 1:
                res['code'] = 1
            elif len(date_header) == 0:
                res['msg'] = f'请删除空表：{excel_path}<{sht_name}>后重试'
            else:
                res['msg'] = f'{excel_path}<{sht_name}>中格式出错'
    except Exception as e:
        res['msg'] = f'测试表格出错了：{e}'
    finally:
        return res

def readExcel(excel_path):
    stuDatas = []
    try:
        data = xlrd_open_workbook(excel_path)
        table = data.sheet_by_index(0)
    except:
        return stuDatas
    nrow = table.nrows
    header = table.row_values(0)
    # print(header)
    for i in range(START_ROW, nrow):
        stu = {}
        stu['info'] = table.cell(i,0).value
        stuData = {}
        for j in range(1, len(header)):
            value = table.cell(i,j).value
            try:
                value = eval(value)
                value = value[1:4]
            except Exception as e:
                print(f'eval转换出错：{e}')
            stuData[header[j]] = value
        stu['data'] = stuData
        stuDatas.append(stu)
    return stuDatas

'''
将多个Excel中的位置信息整合为一个Excel
'''
class MergeExcelWorker(QObject):

    _finished = pyqtSignal()
    _signal = pyqtSignal(str)

    def __init__(self, excelList=[], excelOutput=''):
        super().__init__()
        self.excelList = excelList
        self.excelOutput = excelOutput

    @pyqtSlot()
    def work(self):
        datas, stuInfos = self.handleExcels(self.excelList)
        self.generateNewExcel(datas, stuInfos, self.excelOutput)

    def handleExcels(self, excel_list):
        xwApp = xw.App(visible=False, add_book=False)

        datas = {}
        stuInfos = []
        count = 0
        total = len(excel_list)
        for excel_path in excel_list:
            count += 1
            try:
                excel = xwApp.books.open(excel_path)
                self._signal.emit(f'正在处理 {count}/{total}：{excel_path}')
                tables = excel.sheets
                for table in tables:
                    header = table.range('A1').expand('horizontal').value
                    # 钉钉打卡导出结果中必有【工号】【提交人】【填写周期】【当前时间,当前地点】
                    idx_sno = header.index('工号') + 1
                    idx_sname = header.index('提交人') + 1
                    idx_date = header.index('填写周期') + 1
                    idx_location = header.index('当前时间,当前地点') + 1
                    last_row = len(table.range(2, idx_sname).expand('vertical').value) + 2
                    for i in range(2, last_row):
                        sinfo = f'{table.range(i, idx_sno).value} {table.range(i, idx_sname).value}'
                        if sinfo not in stuInfos:
                            stuInfos.append(sinfo)
                        sdate = table.range(i, idx_date).value
                        str_location = table.range(i, idx_location).value
                        if sdate not in datas:
                            datas[sdate] = {}
                        datas[sdate][sinfo] = str_location
            except Exception as e:
                print(f'读取Excel出错：{e}')
            finally:
                try:
                    excel.close()
                    xwApp.quit()
                except Exception as e:
                    print(f'关闭Excel出错：{e}')
        # 补充遗漏的学生数据
        for date, dateData in datas.items():
            for stu in stuInfos:
                if stu not in datas[date]:
                    datas[date][stu] = '-'
    
        return datas, stuInfos


    def generateNewExcel(self, datas, stuInfos, output_excel):
        # 写入一个新的Excel表中
        resHeader = ['提交人'] + sorted(datas.keys())
        dataToWrite = []
        for i in range(len(stuInfos)):
            stu = stuInfos[i]
            stu_locs = [stu]
            for j in range(1, len(resHeader)):
                stu_locs.append(datas[resHeader[j]][stu])
            dataToWrite.append(stu_locs)

        self._signal.emit('正在整合Excel...')
        xwApp = xw.App(visible=False, add_book=True)
        wb = xwApp.books(1)
        sht = wb.sheets(1)
        colcnt = len(resHeader)
        # 处理表头
        for i in range(1, len(resHeader)):
            m, d = [int(i) for i in resHeader[i][-5:].split('-')]
            resHeader[i] = f"'{m}月{d}日"
        lastCol = chr(ord("A")+colcnt-1) 
        try:
            sht.range(f'A1:{lastCol}1').column_width = 31
            sht.range(f'A1:{lastCol}1').api.Font.Bold = True
            sht.range('A1').value = resHeader
            sht.range('A2').value = dataToWrite
            sht.range('A2').rows.autofit()
            sht.range(f'A:{lastCol}').api.VerticalAlignment = -4130 #自动换行
            sht.range(f'A:{lastCol}').api.Font.Name = '微软雅黑'
            sht.range(f'A:{lastCol}').api.Font.Size = 10
            wb.save(output_excel)
            self._finished.emit()
        except Exception as e:
            self._signal.emit(f'表格创建失败:{e}')
        finally:
            wb.close()
            xwApp.quit()

    def handleExcelsByXlwt(self, excel_list):
        datas = {}
        stuInfos = []
        for excel_path in excel_list:
            excel = xlrd_open_workbook(excel_path)
            tables = excel.sheets()
            for table in tables:
                header = table.row_values(0)
                # 钉钉打卡导出结果中必有【工号】【提交人】【填写周期】
                idx_sno = header.index('工号')
                idx_sname = header.index('提交人')
                idx_date = header.index('填写周期')
                idx_location = header.index('当前时间,当前地点')
                for i in range(1, table.nrows):
                    sinfo = f'{table.cell(i, idx_sno).value} {table.cell(i, idx_sname).value}'
                    if sinfo not in stuInfos:
                        stuInfos.append(sinfo)
                    sdate = table.cell(i, idx_date).value
                    str_location = table.cell(i, idx_location).value
                    if sdate not in datas:
                        datas[sdate] = {}
                    datas[sdate][sinfo] = str_location
        # print(datas)

        return datas, stuInfos


if __name__ == '__main__':
    excel_list = ['../excels/每日健康打卡(28).xlsx', '../excels/每日健康打卡(27).xlsx', '../excels/每日健康打卡(29).xlsx']
    output_excel = '../excels/每日健康打卡位置汇总.xlsx'
    handleExcel(excel_list, output_excel)