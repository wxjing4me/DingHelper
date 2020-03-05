#-*-coding:utf-8-*-
from xlrd import open_workbook as xlrd_open_workbook
from xlsxwriter import Workbook as xlsxwriter_Workbook
from PyQt5.QtCore import QThread, QObject, pyqtSignal, pyqtSlot
from functions.logging_setting import Log
from time import sleep as time_sleep

# 第1行默认为表头
START_ROW = 1 # default:1

SPLIT_CHAR = '='  # 工号=提交人

log = Log(__name__).getLog()

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
            log.debug(f'Excel:{excel_path}, Sheet:{sht_name}')
            table = excel.sheet_by_name(sht_name)
            header = table.row_values(0)
            needed = ['提交人', '工号', '当前时间,当前地点']
            if len(header) != 0 and set(header) >= set(needed):
                res['code'] = 1
            elif len(header) == 0:
                res['msg'] = f'请删除空表：{excel_path}<{sht_name}>后重试'
            else:
                res['msg'] = f'{excel_path}<{sht_name}>中缺少{needed}的部分字段'
    except Exception as e:
        res['msg'] = f'Excel表格格式有误！'
        log.error(f'Excel表格格式有误！ - {e}', exc_info=True)
    finally:
        if res['msg'] != '':
            res['code'] = 0
            log.info(res['msg'])
        return res

def testSpectExcel(excel_path):
    res = {}
    res['code'] = 0
    res['msg'] = ''
    try:
        excel = xlrd_open_workbook(excel_path)
        table = excel.sheet_by_index(0)
        # 只用到第一个工作表
        nrows = table.nrows
        ncols = table.ncols
        if nrows == 0 and ncols == 0:
            res['msg'] = f'该Excel为空表，请重新选择文件！'
        elif nrows >= 2 and ncols >= 4:
            reFlag = True
            for i in range(1, nrows):
                for j in range(2, ncols):
                    val = table.cell(i, j).value
                    if val != '' and val != '-':
                        val = eval(val)
                        if not isinstance(val, list):
                            reFlag = False
                            break
            if reFlag:
                res['code'] = 1
        else:
            res['msg'] = '该Excel表格数据量不足，请重新选择文件！'
    except Exception as e:
        res['msg'] = f'该Excel表格格式有误，请重新选择文件！'
        log.error(f"{res['msg']} - {e}", exc_info=True)
    finally:
        if res['msg'] != '':
            log.info(res['msg'])
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
        stu['info'] = SPLIT_CHAR.join((table.cell(i,0).value, table.cell(i, 1).value))
        stuData = {}
        for j in range(2, len(header)):
            value = table.cell(i,j).value
            try:
                value = eval(value)
                value = value[1:4]
            except Exception as e:
                log.debug(f'eval转换出错：{e}', exc_info=True)
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
        datas = {}
        stuInfos = []
        count = 0
        total = len(excel_list)
        for excel_path in excel_list:
            count += 1
            try:
                excel = xlrd_open_workbook(excel_path)
                self._signal.emit(f'正在处理 {count}/{total}：{excel_path}')
                time_sleep(0.5)
                tables = excel.sheets()
                for table in tables:
                    header = table.row_values(0)
                    # 钉钉打卡导出结果中必有【工号】【提交人】【当前时间,当前地点】
                    idx_sno = header.index('工号')
                    idx_sname = header.index('提交人')
                    idx_location = header.index('当前时间,当前地点')
                    for i in range(1, len(table.col_values(0))):
                        sinfo = f'{table.cell(i, idx_sno).value}{SPLIT_CHAR}{table.cell(i, idx_sname).value}'
                        if sinfo not in stuInfos:
                            stuInfos.append(sinfo)
                        str_location = table.cell(i, idx_location).value
                        if str_location != '-':
                            try:
                                sdate = eval(str_location)[0][:10]
                                if sdate not in datas:
                                    datas[sdate] = {}
                                datas[sdate][sinfo] = str_location
                            except Exception as e:
                                log.warn(f'{e}', exc_info=True)
            except Exception as e:
                log.warn(f'读取{excel_path}出错: {e}', exc_info = True)
        # 补充遗漏的学生数据
        for date, dateData in datas.items():
            log.info(f'日期: {date}, 人员数: {len(datas[date])}')
            for stu in stuInfos:
                if stu not in datas[date]:
                    datas[date][stu] = '-'
        return datas, stuInfos


    def generateNewExcel(self, datas, stuInfos, output_excel):
        self._signal.emit('正在整合Excel...')
        time_sleep(0.5)
        try:
            excel = xlsxwriter_Workbook(output_excel)
            table = excel.add_worksheet()
            dates = sorted(list(datas.keys()))
            # 设置格式
            table.set_column(0, 1, 15)
            table.set_column(2, len(dates)+1, 31)
            format_header = excel.add_format({'bold': True, 'font_name': '微软雅黑', 'font_size': 10, 'text_wrap': True, 'valign': 'top'})
            format_cell = excel.add_format({'font_name': '微软雅黑', 'font_size': 10, 'text_wrap': True, 'valign': 'top'})
            # 写入表头
            table.write(0, 0, '工号', format_header)
            table.write(0, 1, '提交人', format_header)
            for ki in range(len(datas.keys())):
                m, d = [int(i) for i in dates[ki][-5:].split('-')]
                dateStr = f"{m}月{d}日"
                table.write(0, ki+2, dateStr, format_header)
            # 依次写入提交人及位置
            for si in range(len(stuInfos)):
                info = stuInfos[si]
                sno, sname = info.split(SPLIT_CHAR)
                table.write(si+1, 0, sno, format_cell)
                table.write(si+1, 1, sname, format_cell)
                for di in range(len(dates)):
                    table.write(si+1, di+2, datas[dates[di]][info], format_cell)
        except Exception as e:
            self._signal.emit(f'表格创建失败:{e}')
            log.error(f'创建{output_excel}失败: {e}', exc_info=True)
        finally:
            excel.close()
            self._finished.emit()


if __name__ == '__main__':
    excel_list = ['../excels/每日健康打卡(28).xlsx', '../excels/每日健康打卡(27).xlsx', '../excels/每日健康打卡(29).xlsx']
    output_excel = '../excels/每日健康打卡位置汇总.xlsx'
    handleExcel(excel_list, output_excel)