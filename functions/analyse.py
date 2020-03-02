from PyQt5.QtCore import QThread, QObject, pyqtSignal, pyqtSlot, QDateTime
from functions.excel_action import readExcel

import requests
import json
import time

DEFAULT_Address = {'nation': '未知'}
DEFAULT_Location = [20, 80, "未知"]

API_URL_LL2Address = 'https://apis.map.qq.com/ws/geocoder/v1/?location='

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

# 第1行默认为表头
START_ROW = 1 # default:1

# 第1,2列默认为学号、姓名
START_COL = 2 # default:2

global REQ_CNT


class AnalyseWorker(QObject):
    def __init__(self, apiToken='', excelPath=''):
        super().__init__()
        self.apiToken = apiToken
        self.excelPath = excelPath

    _finished = pyqtSignal()
    _signal = pyqtSignal(str)

    @pyqtSlot()
    def work(self):
        global REQ_CNT
        REQ_CNT = 1
        stu_count = 1
        stuDatas = readExcel(self.excelPath)
        self._signal.emit('共 %d 位学生' % len(stuDatas))
        for aStuData in stuDatas:
            self._signal.emit('- %d / %d ' % (stu_count, len(stuDatas)) + '-' * 100)
            # print(aStuData)
            self.aStuAnalyse(aStuData)
            stu_count += 1
        self._finished.emit()

    '''
    逆地址解析
    利用经纬度（数字）获得具体地点（字典）
    '''
    def getAddressByLL(self, location):
        global REQ_CNT
        try:
            # print(location)
            longitude, latitude, _ = location
            try:
                response = requests.get('%s%s,%s&key=%s' % (API_URL_LL2Address, latitude, longitude, self.apiToken))
                REQ_CNT += 1
            except Exception as e:
                self._signal.emit('ERROR: request请求错误: %s' % e)
                exit()
            if response.status_code != 200:
                self._signal.emit('ERROR: %d 获取腾讯地图地址失败' % response.status_code)
            else:
                res = json.loads(response.text)
                if res['status'] != 0:
                    self._signal.emit('ERROR: 腾讯地图API错误（逆地址解析）：%s' % res['message'])
                    exit()
                else:
                    address = res['result']['address_component']
            if REQ_CNT % MAX_CNT_PER_SEC == 0:
                time.sleep(1)
        except Exception as e:
            print(e)
            address = DEFAULT_Address
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
    
        sno, sname = aStuData['info']
        sdata = aStuData['data']

        self._signal.emit('%s %s' % (sno, sname))

        yesterDate, yesterAddress = '', ''
        for date, location in sdata.items():
            if yesterDate == '' and yesterAddress == '':
                yesterDate = date
                try:
                    yesterLocation = location[2]
                except:
                    yesterLocation = DEFAULT_Location[2]
                yesterAddress = self.getAddressByLL(location)
                continue
            else:
                todayDate = date
                try:
                    todayLocation = location[2]
                except:
                    todayLocation = DEFAULT_Location[2]
                todayAddress = self.getAddressByLL(location)
                cRes = self.compareAdress(yesterAddress, todayAddress)
                if cRes['type'] != STR_STAY:
                    self._signal.emit('>> %s - %s %s\n高德：%s \n钉钉：%s -> %s' % (yesterDate, todayDate, cRes['type'], cRes['amap_msg'], yesterLocation, todayLocation))
                yesterDate = todayDate
                yesterAddress = todayAddress

