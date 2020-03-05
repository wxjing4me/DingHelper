#-*-coding:utf-8-*-
from PyQt5.QtCore import QThread, QObject, pyqtSignal, pyqtSlot, QDateTime
from functions.excel_action import readExcel

from requests import get as requests_get
from json import loads as json_loads
from time import sleep as time_sleep
from functions.logging_setting import Log

DEFAULT_Address = {'nation': '未知'}
DEFAULT_Location = [20, 80, "未知"]
FJNU_LL = '26.0271200000,119.2099200000'

API_URL_LL2Address = 'https://apis.map.qq.com/ws/geocoder/v1/?location='
API_URL_DISTANCE = 'https://apis.map.qq.com/ws/distance/v1/?'

STR_STAY = "无变化"
STR_FUJIAN_IN = "【省外入闽】"
STR_FUZHOU_IN = "【外地返榕】"
STR_FUZHOU_ELSE = "【“榕”随便走走】"
STR_FUJIAN_ELSE = "【“闽”随便走走】"
STR_OUTSIDE = "【“省外”随便走走】"
STR_FUJIAN_OUT = "【离闽】"
STR_FUZHOU_OUT = "【离榕】"
STR_ELSE = "【系统不清楚】"
STR_FUJIAN = '福建省'
STR_FUZHOU = '福州市'
STR_FUZHOU_LIST = ['闽侯县', '鼓楼区', '台江区', '晋安区', '仓山区', '马尾区']
MAX_CNT_PER_SEC = 4

global REQ_CNT
global REQ_DIS_CNT

SPLIT_CHAR = '='  # 工号=提交人

log = Log(__name__).getLog()

class AnalyseWorker(QObject):
    def __init__(self, apiKey='', excelPath=''):
        super().__init__()
        self.apiKey = apiKey
        self.excelPath = excelPath
        self.stopFlag = False

    _finished = pyqtSignal()
    _signal = pyqtSignal(str)

    @pyqtSlot()
    def work(self):
        global REQ_CNT
        global REQ_DIS_CNT
        REQ_CNT, REQ_DIS_CNT = 1, 1
        stu_count = 1
        stuDatas = readExcel(self.excelPath)
        self._signal.emit('共 %d 位学生' % len(stuDatas))
        for aStuData in stuDatas:
            self._signal.emit('- %d / %d ' % (stu_count, len(stuDatas)) + '-' * 100)
            # print(aStuData)
            self.aStuAnalyse(aStuData)
            stu_count += 1
            if self.stopFlag:
                break
        self._finished.emit()

    @pyqtSlot()
    def stop(self):
        self.stopFlag = True

    '''
    逆地址解析
    利用经纬度（数字）获得具体地点（字典）
    '''
    def getAddressByLL(self, location):
        address = DEFAULT_Address
        global REQ_CNT
        try:
            longitude, latitude, _ = location
        except Exception as e:
            log.warn(f'location={location}, 可能原因: 用户未填')
            return address
        if longitude == '' or latitude == '':
            self._signal.emit('提示：该位置是手动输入，非自动定位')
            return address
        try:
            response = requests_get('%s%s,%s&key=%s' % (API_URL_LL2Address, latitude, longitude, self.apiKey))
            REQ_CNT += 1
        except Exception as e:
            # self._signal.emit('ERROR: request请求错误: %s' % e)
            log.error(f'ERROR: request请求错误: {e}', exc_info=True)
        if response.status_code != 200:
            self._signal.emit('ERROR: %d 获取腾讯地图地址失败' % response.status_code)
            log.error(f'获取腾讯地图地址失败: status_code={response.status_code}', exc_info=True)
        else:
            res = json_loads(response.text)
            if res['status'] != 0:
                log.warn(f"腾讯地图API错误（逆地址解析）: {res['message']}！location={location}")
            else:
                address = res['result']['address_component']
        if REQ_CNT % MAX_CNT_PER_SEC == 0:
            time_sleep(1)
        return address

    def compareAdress(self, yesterAddress, todayAddress):
        res = {}
        try:
            yesterAddressBrief = '%s%s%s' % (yesterAddress['province'], yesterAddress['city'], yesterAddress['district'])
        except:
            yesterAddressBrief = yesterAddress['nation']
        try:
            todayAddressBrief = '%s%s%s' % (todayAddress['province'], todayAddress['city'], todayAddress['district']) 
        except:
            todayAddressBrief = todayAddress['nation']
        res['amap_msg'] = '%s -> %s' % (yesterAddressBrief, todayAddressBrief)
        if yesterAddressBrief == todayAddressBrief:
            res['type'] = STR_STAY
            return res
        else:
            try:
                if todayAddress['city'] == STR_FUZHOU and todayAddress['district'] in STR_FUZHOU_LIST:
                    if yesterAddress['city'] == STR_FUZHOU and yesterAddress['district'] in STR_FUZHOU_LIST:
                        # “榕”随便走走
                        res['type'] = STR_FUZHOU_ELSE
                    else:
                        # 外地返榕
                        res['type'] = STR_FUZHOU_IN
                elif todayAddress['province'] == STR_FUJIAN:
                    if yesterAddress['province'] != STR_FUJIAN:
                        # 省外入闽
                        res['type'] = STR_FUJIAN_IN
                    elif yesterAddress['city'] == STR_FUZHOU and yesterAddress['district'] in STR_FUZHOU_LIST:
                        # 离榕
                        res['type'] = STR_FUZHOU_OUT
                    else:
                        # “闽”随便走走
                        res['type'] = STR_FUJIAN_ELSE
                elif todayAddress['province'] != STR_FUJIAN:
                    if yesterAddress['city'] == STR_FUZHOU and yesterAddress['district'] in STR_FUZHOU_LIST:
                        # 离榕
                        res['type'] = STR_FUZHOU_OUT
                    elif yesterAddress['province'] == STR_FUJIAN:
                        # 离闽
                        res['type'] = STR_FUJIAN_OUT
                    else:
                        # “省外”随便走走
                        res['type'] = STR_OUTSIDE
                else:
                    res['type'] = STR_ELSE
            except:
                # 系统不清楚
                res['type'] = STR_ELSE
            return res

    def aStuAnalyse(self, aStuData):
    
        sno, sname = aStuData['info'].split(SPLIT_CHAR)
        sdata = aStuData['data']

        self._signal.emit(f'{sno} {sname}')

        yesterDate, yesterAddress, yesterLL = '', '', ''
        for date, location in sdata.items():
            if yesterDate == '' and yesterAddress == '':
                yesterDate = date
                try:
                    yesterLocation = location[2]
                    yesterLL = ','.join((str(location[1]), str(location[0])))
                except:
                    yesterLocation = DEFAULT_Location[2]
                    yesterLL = ','.join((str(DEFAULT_Location[1]), str(DEFAULT_Location[0])))
                yesterAddress = self.getAddressByLL(location)
                continue
            else:
                todayDate = date
                try:
                    todayLocation = location[2]
                    todayLL = ','.join((str(location[1]), str(location[0])))
                except:
                    todayLocation = DEFAULT_Location[2]
                    todayLL = ','.join((str(DEFAULT_Location[1]), str(DEFAULT_Location[0])))
                todayAddress = self.getAddressByLL(location)
            cRes = self.compareAdress(yesterAddress, todayAddress)
            if cRes['type'] != STR_STAY:
                self._signal.emit(f">> {yesterDate} - {todayDate} <span style='color:red'>{cRes['type']}</span><br>腾讯位置：{cRes['amap_msg']}<br>钉钉位置：{yesterLocation} -> {todayLocation}")
            else:
                dRes = self.calculateDistance(yesterLL, todayLL)
                self._signal.emit(f'>> {yesterDate} - {todayDate} {dRes}<br>钉钉位置：{yesterLocation} -> {todayLocation}')
            yesterDate = todayDate
            yesterAddress = todayAddress
            yesterLocation = todayLocation
            yesterLL = todayLL

    def calculateDistance(self, fromLL, toLL):
        global REQ_DIS_CNT
        disRes, disStr = '相距', '未知'
        disResF, disStrF = '距离福师大', '未知'
        try:
            response = requests_get(f'{API_URL_DISTANCE}from={fromLL}&to={toLL};{FJNU_LL}&key={self.apiKey}')
            REQ_DIS_CNT += 1
            if response.status_code != 200:
                log.error(f'获取腾讯地图地址失败: status_code={response.status_code}', exc_info=True)
            else:
                res = json_loads(response.text)
                if res['status'] != 0:
                    log.error(f"腾讯地图API错误（距离计算）: {res['message']}！from={fromLL}&to={toLL};{FJNU_LL}")
                else:
                    dist = res['result']['elements'][0]['distance']
                    distF = res['result']['elements'][1]['distance']
                    if dist > 1000:
                        disStr = f'{round(dist/1000, 2)}公里'
                    else:
                        disStr = f'{dist}米'
                    if distF > 1000:
                        disStrF = f'{round(distF/1000, 2)}公里'
                    else:
                        disStrF = f'{distF}米'
                    if distF < 10000:
                        disStrF = f"<span style='color:blue'>{disStrF}</span>"
            if REQ_DIS_CNT % MAX_CNT_PER_SEC == 0:
                time_sleep(1)
        except Exception as e:
            log.error(f'计算距离出错：{e}', exc_info=True)
        return f'（{disRes}{disStr}, {disResF}{disStrF}）'