#-*-coding:utf-8-*-
from xlrd import open_workbook as xlrd_open_workbook
from xlsxwriter import Workbook as xlsxwriter_Workbook
from PyQt5.QtCore import QThread, QObject, pyqtSignal, pyqtSlot
from time import sleep as time_sleep
from os.path import join as os_path_join

from configure.logging_action import Log
from configure.config_values import *
from configure import config_action as confAct

log = Log(__name__).getLog()

def testRawExcel(excel_path):
    '''判断是否符合【每日健康打卡】导出的Excel文件格式
    '''
    res = {}
    res['code'] = 0
    res['msg'] = ''
    try:
        excel = xlrd_open_workbook(excel_path)
        sht_names = excel.sheet_names()
        for sht_name in sht_names:
            log.debug(f'Doing...Excel: {excel_path}, Sheet: {sht_name}')
            table = excel.sheet_by_name(sht_name)
            header = table.row_values(0)
            if len(header) == 0:
                res['msg'] = f'请删除空表：{excel_path}<{sht_name}>后重试'
            elif set(header) >= set(HEADER_REQUIRED):
                res['code'] = 1
            else:
                res['msg'] = f'{excel_path}<{sht_name}>中缺少{HEADER_REQUIRED}的部分字段'
            if confAct.ONLY_FIRST_SHEET:
                break
    except Exception as e:
        res['msg'] = f'Excel表格{excel_path}格式有误！'
        log.error(f'Excel表格格式有误！ - {e}', exc_info=True)
    finally:
        if res['msg'] != '':
            res['code'] = 0
            log.info(res['msg'])
        return res

def testSpectExcel(excel_path):
    '''判断是否符合位置分析的Excel文件格式
    '''
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
                    if val != '' and val != STR_UNDO:
                        val = eval(val)
                        if not isinstance(val, list):
                            reFlag = False
                            break
            if reFlag:
                res['code'] = 1
            else:
                res['msg'] = f'该Excel表格格式有误，请重新选择文件！'
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
    for i in range(confAct.START_ROW, nrow):
        stu = {}
        stu['info'] = SPLIT_CHAR.join((table.cell(i,0).value, table.cell(i, 1).value))
        stuData = {}
        for j in range(2, len(header)):
            value = table.cell(i,j).value
            try:
                value = eval(value)
                value = value[1:4]
            except Exception as e:
                log.debug(f'eval转换出错：{e} - value={value}')
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
                self._signal.emit(f'正在处理 {count}/{total}：<br>{excel_path}')
                time_sleep(0.5)
                tables = excel.sheets()
                for table in tables:
                    header = table.row_values(0)
                    # 钉钉打卡导出结果中必有【工号】【提交人】【当前时间,当前地点】
                    idx_sno = header.index(STR_SNO)
                    idx_sname = header.index(STR_NAME)
                    idx_location = header.index(STR_TIME_LOC)
                    for i in range(1, len(table.col_values(0))):
                        sinfo = f'{table.cell(i, idx_sno).value}{SPLIT_CHAR}{table.cell(i, idx_sname).value}'
                        if sinfo not in stuInfos:
                            stuInfos.append(sinfo)
                        str_location = table.cell(i, idx_location).value
                        if str_location != STR_UNDO:
                            try:
                                sdate = eval(str_location)[0][:10]
                                if sdate not in datas:
                                    datas[sdate] = {}
                                datas[sdate][sinfo] = str_location
                            except Exception as e:
                                log.warn(f'{e}', exc_info=True)
                    if confAct.ONLY_FIRST_SHEET:
                        break
            except Exception as e:
                log.warn(f'读取{excel_path}出错: {e}', exc_info = True)
        # 补充遗漏的学生数据
        for date, dateData in datas.items():
            log.info(f'日期: {date}, 人员数: {len(datas[date])}')
            for stu in stuInfos:
                if stu not in datas[date]:
                    datas[date][stu] = STR_UNDO
        stuInfos = sorted(stuInfos)
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
            format_header = excel.add_format({'bold': True, 'font_name': FONT_NAME_YAHEI, 'font_size': FONT_SIZE, 'text_wrap': True, 'valign': 'top'})
            format_cell = excel.add_format({'font_name': FONT_NAME_YAHEI, 'font_size': FONT_SIZE, 'text_wrap': True, 'valign': 'top'})
            # 写入表头
            table.write(0, 0, STR_SNO, format_header)
            table.write(0, 1, STR_NAME, format_header)
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

'''
将多个Excel根据学号姓名整合为一人一档
'''
class CreateProfilesWorker(QObject):

    _finished = pyqtSignal()
    _signal = pyqtSignal(str)

    def __init__(self, excelsList=[], profilesDir=''):
        super().__init__()
        self.excelsList = excelsList
        self.profilesDir = profilesDir

    @pyqtSlot()
    def work(self):
        datas, stuInfos = self.handleExcels(self.excelsList)
        self.generateProfiles(datas, stuInfos, self.profilesDir)

    def handleExcels(self, excel_list):
        datas = {}
        stuInfos = []
        stuDates = []
        count = 0
        total = len(excel_list)
        for excel_path in excel_list:
            count += 1
            try:
                excel = xlrd_open_workbook(excel_path)
                self._signal.emit(f'正在处理 {count}/{total}：{excel_path}')
                tables = excel.sheets()
                for table in tables:
                    header = table.row_values(0)
                    # 钉钉打卡导出结果中必有【工号】【提交人】【当前时间,当前地点】
                    idx_sno = header.index(STR_SNO)
                    idx_sname = header.index(STR_NAME)
                    idx_date = header.index(STR_DATE)
                    header_profile_idxs = []
                    for colName in HEADER_PROFILE[2:]:
                        if colName in header:
                            header_profile_idxs.append(header.index(colName))
                        else:
                            header_profile_idxs.append(-1)
                            log.error(f'EXCEL: {excel_path} - 缺少字段: {colName}')
                    for i in range(1, len(table.col_values(0))):
                        sinfo = f'{table.cell(i, idx_sno).value}{SPLIT_CHAR}{table.cell(i, idx_sname).value}'
                        if sinfo not in stuInfos:
                            datas[sinfo] = {}
                            stuInfos.append(sinfo)
                        try:
                            sdate = table.cell(i, idx_date).value
                            if sdate not in stuDates:
                                stuDates.append(sdate)
                            temp_data = []
                            for idx in header_profile_idxs:
                                if idx != -1:
                                    temp_data.append(table.cell(i, idx).value)
                                else:
                                    temp_data.append(STR_UNDO)
                            datas[sinfo][sdate] = temp_data
                        except Exception as e:
                            log.warn(f'{e}', exc_info=True)
                    if confAct.ONLY_FIRST_SHEET:
                        break
            except Exception as e:
                log.warn(f'读取{excel_path}出错: {e}', exc_info = True)
        
        # 补充遗漏的日期数据
        undoData = [STR_UNDO] * len(HEADER_PROFILE[2:])
        for sinfo, sData in datas.items():
            for date in stuDates:
                if date not in sData:
                    datas[sinfo][date] = undoData
                    log.debug(f'人员: {sinfo} - 缺少日期: {date}')
        
        return datas, stuInfos

    def generateProfiles(self, datas, stuInfos, output_dir):
        total = len(stuInfos)
        for i in range(total):
            sinfo = stuInfos[i]
            sno, sname = sinfo.split(SPLIT_CHAR)
            self._signal.emit(f'生成中: {i} / {total} {sno} {sname}')
            excel_path = os_path_join(output_dir, f'{sno}_{sname}.xlsx')
            aStuData = datas[sinfo]
            dataLen = len(HEADER_PROFILE)
            try:
                excel = xlsxwriter_Workbook(excel_path)
                table = excel.add_worksheet()
                dates = sorted(list(aStuData.keys()))
                format_header = excel.add_format({'bold': True, 'font_name': FONT_NAME_YAHEI, 'font_size': FONT_SIZE, 'text_wrap': True, 'valign': 'top'})
                format_cell = excel.add_format({'font_name': FONT_NAME_YAHEI, 'font_size': FONT_SIZE, 'valign': 'top'})
                # 写入表头
                for hi in range(dataLen):
                    table.write(0, hi, HEADER_PROFILE[hi], format_header)
                # 依次写入日期及当日数据
                for di in range(len(dates)):
                    table.write(di+1, 0, sno, format_cell)
                    table.write(di+1, 1, sname, format_cell)
                    for ki in range(dataLen-2):
                        table.write(di+1, ki+2, aStuData[dates[di]][ki], format_cell)
            except Exception as e:
                self._signal.emit(f'表格创建失败:{e}')
                log.error(f'创建{excel_path}失败: {e}', exc_info=True)
            finally:
                excel.close()
        self._finished.emit()
