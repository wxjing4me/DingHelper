# -*-coding:utf-8-*-
from PyQt5.QtCore import QThread, QObject, pyqtSignal, pyqtSlot, QDateTime
from requests import get as requests_get, exceptions
from json import loads as json_loads
from time import sleep as time_sleep

from common.excel_action import readExcel
from configure.logging_action import Log
from configure.config_values import *
from configure.danger_place import *
from configure import config_action as confAct

global REQ_CNT
global REQ_DIS_CNT

log = Log(__name__).getLog()


class AnalyseWorker(QObject):
    def __init__(self, apiKey='', excelPath='', mtype='AMAP'):
        super().__init__()
        self.apiKey = apiKey
        self.excelPath = excelPath
        self.mtype = mtype
        self.stopFlag = False

    _finished = pyqtSignal()
    _signal = pyqtSignal(str)

    @pyqtSlot()
    def work(self):
        global REQ_CNT
        global REQ_DIS_CNT
        REQ_CNT, REQ_DIS_CNT = 1, 1
        stu_count = confAct.START_ROW
        stuDatas = readExcel(self.excelPath)
        end_count = len(stuDatas)+confAct.START_ROW-1
        for aStuData in stuDatas:
            self._signal.emit('- %d / %d ' %
                              (stu_count, end_count) + '-' * 100)
            # print(aStuData)
            self.aStuAnalyse(aStuData)
            stu_count += 1
            if self.stopFlag:
                break
        self._finished.emit()

    @pyqtSlot()
    def stop(self):
        self.stopFlag = True

    def getAddressByLL(self, location):
        '''逆地址解析: 利用经纬度（数字）获得具体地点（字典）
        '''
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
            if self.mtype == 'AMAP':
                response = requests_get(
                    f'{AMAP_API_URL_LL2Address}location={longitude},{latitude}&key={self.apiKey}')
            elif self.mtype == 'QQ':
                response = requests_get(
                    f'{QQ_API_URL_LL2Address}location={latitude},{longitude}&key={self.apiKey}')
            REQ_CNT += 1
            if response.status_code != 200:
                self._signal.emit('ERROR: %d 获取地址失败' % response.status_code)
                log.error(
                    f'获取地址失败: mtype={self.mtype}, status_code={response.status_code}', exc_info=True)
            else:
                res = json_loads(response.text)
                if self.mtype == 'AMAP' and int(res['status']) == 0:
                    log.warn(
                        f"高德地图API错误（逆地理编码）: {res['info']}！location={location}")
                elif self.mtype == 'AMAP' and int(res['status']) == 1:
                    address = res['regeocode']['addressComponent']
                elif self.mtype == 'QQ' and res['status'] != 0:
                    log.warn(
                        f"腾讯地图API错误（逆地址解析）: {res['message']}！location={location}")
                elif self.mtype == 'QQ' and res['status'] == 0:
                    address = res['result']['address_component']
            if REQ_CNT % eval(self.mtype+'_MAX_CNT_PER_SEC') == 0:
                time_sleep(1)
        except exceptions.ConnectTimeout:
            #FIXME: 无效
            self._signal.emit('提示: ConnectTimeout请求超时')
        except exceptions.Timeout:
            self._signal.emit('提示: Timeout请求超时')
        except Exception as e:
            # self._signal.emit('ERROR: request请求错误: %s' % e)
            log.error(f'ERROR: request请求错误: {e}', exc_info=True)
        return address

    def compareAddress(self, yesterAddress, todayAddress):
        res = {}
        try:
            yesterAddressBrief = '%s%s%s' % (
                yesterAddress['province'], yesterAddress['city'], yesterAddress['district'])
        except:
            yesterAddressBrief = yesterAddress['nation']
        try:
            todayAddressBrief = '%s%s%s' % (
                todayAddress['province'], todayAddress['city'], todayAddress['district'])
        except:
            todayAddressBrief = todayAddress['nation']
        res['amap_msg'] = '%s -> %s' % (yesterAddressBrief, todayAddressBrief)
        # 中高风险地区
        res['is_danger'] = ''
        for danger_place in DANGER_PLACES:
            if todayAddressBrief[:len(danger_place)] == danger_place:
                res['is_danger'] = f'{danger_place}是中高风险地区'
                log.info(f"{todayAddress['province']}, {todayAddress['city']}")
        # 今日是否在“福建省福州市”
        res['todayInFZ'] = False
        log.debug(f'todayAddress={todayAddress}')
        try:
            if todayAddress['province'] == LOC_NAME_FUJIAN and todayAddress['city'] == LOC_NAME_FUZHOU:
                res['todayInFZ'] = True
        except Exception:
            log.error(f'todayAddress={todayAddress}')
        # 位置异动
        if yesterAddressBrief == '未知' or todayAddressBrief == '未知':
            res['type'] = LOC_TYPE_ELSE
        elif yesterAddressBrief == todayAddressBrief:
            res['type'] = LOC_TYPE_STAY
        else:
            try:
                if todayAddress['city'] == LOC_NAME_FUZHOU and todayAddress['district'] in LOC_NAME_FUZHOU_LIST:
                    if yesterAddress['city'] == LOC_NAME_FUZHOU and yesterAddress['district'] in LOC_NAME_FUZHOU_LIST:
                        # “榕”随便走走
                        res['type'] = LOC_TYPE_FUZHOU_ELSE
                    else:
                        # 外地返榕
                        res['type'] = LOC_TYPE_FUZHOU_IN
                elif todayAddress['province'] == LOC_NAME_FUJIAN:
                    if yesterAddress['province'] != LOC_NAME_FUJIAN:
                        # 省外入闽
                        res['type'] = LOC_TYPE_FUJIAN_IN
                    elif yesterAddress['city'] == LOC_NAME_FUZHOU and yesterAddress['district'] in LOC_NAME_FUZHOU_LIST:
                        # 离榕
                        res['type'] = LOC_TYPE_FUZHOU_OUT
                    else:
                        # “闽”随便走走
                        res['type'] = LOC_TYPE_FUJIAN_ELSE
                elif todayAddress['province'] != LOC_NAME_FUJIAN:
                    if yesterAddress['city'] == LOC_NAME_FUZHOU and yesterAddress['district'] in LOC_NAME_FUZHOU_LIST:
                        # 离榕
                        res['type'] = LOC_TYPE_FUZHOU_OUT
                    elif yesterAddress['province'] == LOC_NAME_FUJIAN:
                        # 离闽
                        res['type'] = LOC_TYPE_FUJIAN_OUT
                    else:
                        # “省外”随便走走
                        res['type'] = LOC_TYPE_OUTSIDE
                else:
                    res['type'] = LOC_TYPE_ELSE
            except:
                # 系统不清楚
                res['type'] = LOC_TYPE_ELSE
        return res

    def aStuAnalyse(self, aStuData):

        sno, sname = aStuData['info'].split(SPLIT_CHAR)
        sdata = aStuData['data']

        self._signal.emit(f'{sno} {sname}')

        yesterDate, yesterAddress, yesterLat, yestlng = '', '', '', ''
        for date, location in sdata.items():
            if yesterDate == '' and yesterAddress == '':
                yesterDate = date
                try:
                    yesterLat, yesterLng, yesterLocation = location[:3]
                except:
                    yesterLat, yesterLng, yesterLocation = DEFAULT_Location[:3]
                yesterAddress = self.getAddressByLL(location)
                continue
            else:
                todayDate = date
                try:
                    todayLat, todayLng, todayLocation = location[:3]
                except:
                    todayLat, todayLng, todayLocation = DEFAULT_Location[:3]
                todayAddress = self.getAddressByLL(location)
            cRes = self.compareAddress(yesterAddress, todayAddress)
            if cRes['type'] != LOC_TYPE_STAY:
                self._signal.emit(
                    f">> {yesterDate} - {todayDate} <span style='color:red'>{cRes['type']}</span> <span style='background-color: yellow'>{cRes['is_danger']}</span><br>{eval('LOC_'+self.mtype)}：{cRes['amap_msg']}<br>{LOC_DING}：{yesterLocation} -> {todayLocation}")
            else:
                if confAct.SHOW_DISTANCE and cRes['todayInFZ']:
                    dRes = self.calculateDistance(
                        yesterLat, yesterLng, todayLat, todayLng)
                    self._signal.emit(
                        f">> {yesterDate} - {todayDate} {dRes} <span style='background-color: yellow'>{cRes['is_danger']}</span><br>{LOC_DING}：{yesterLocation} -> {todayLocation}")
                else:
                    self._signal.emit(
                        f">> {yesterDate} - {todayDate} <span style='background-color: yellow'>{cRes['is_danger']}</span><br>{LOC_DING}：{yesterLocation} -> {todayLocation}")

            yesterDate = todayDate
            yesterAddress = todayAddress
            yesterLocation = todayLocation
            yesterLat, yesterLng = todayLat, todayLng

    def calculateDistance(self, fromLat, fromLng, toLat, toLng):
        global REQ_DIS_CNT
        disRes, disStr = '相距', '未知'
        disResF, disStrF = '距离福师大', '未知'
        try:
            if self.mtype == 'QQ':
                response = requests_get(
                    f'{QQ_API_URL_DISTANCE}from={fromLng},{fromLat}&to={toLng},{toLat};{FJNU_Lng},{FJNU_Lat}&key={self.apiKey}')
            elif self.mtype == 'AMAP':
                response = requests_get(
                    f'{AMAP_API_URL_DISTANCE}origins={fromLat},{fromLng}|{FJNU_Lat},{FJNU_Lng}&destination={toLat},{toLng}&type=0&key={self.apiKey}')
            REQ_DIS_CNT += 1
            if response.status_code != 200:
                log.error(
                    f'获取地址失败: status_code={response.status_code}', exc_info=True)
            else:
                dist, distF = -1, -1
                res = json_loads(response.text)
                if self.mtype == 'QQ' and res['status'] != 0:
                    log.error(
                        f"腾讯地图API错误（距离计算）: {res['message']}！from={fromLng},{fromLat}&to={toLng},{toLat};{FJNU_Lng},{FJNU_Lat}")
                elif self.mtype == 'QQ' and res['status'] == 0:
                    dist = res['result']['elements'][0]['distance']
                    distF = res['result']['elements'][1]['distance']
                elif self.mtype == 'AMAP' and int(res['status']) == 0:
                    log.error(
                        f"高德地图API错误（距离计算）: {res['info']}！origins={fromLat},{fromLng}|{FJNU_Lat},{FJNU_Lng}&destination={toLat},{toLng}")
                elif self.mtype == 'AMAP' and int(res['status']) == 1:
                    # log.debug(res['results'])
                    dist = int(res['results'][0]['distance'])
                    distF = int(res['results'][1]['distance'])
                # 两日距离
                if dist > 1000:
                    disStr = f"<span style='color:blue'>{round(dist/1000, 2)}公里</span>"
                elif dist >= 0:
                    disStr = f'{dist}米'
                # 距离福师大
                if distF > 1000:
                    disStrF = f'{round(distF/1000, 2)}公里'
                elif 0 < distF < 1000:
                    disStrF = f'{distF}米'
                if 0 < distF < 10000:
                    disStrF = f"<span style='color:blue'>{disStrF}</span>"
            if REQ_DIS_CNT % eval(self.mtype+'_MAX_CNT_PER_SEC') == 0:
                time_sleep(1)
        except exceptions.ConnectTimeout:
            #FIXME: 无效
            self._signal.emit('提示: ConnectTimeout请求超时')
        except exceptions.Timeout:
            self._signal.emit('提示: Timeout请求超时')
        except Exception as e:
            log.error(f'计算距离出错：{e}', exc_info=True)
        return f'（{disRes}{disStr}, {disResF}{disStrF}）'
