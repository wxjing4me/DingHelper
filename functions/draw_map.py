#-*-coding:utf-8-*-
from pyecharts.charts import Geo
from pyecharts import options as opts
from pyecharts.globals import ChartType
from os.path import join as ospath_join

from PyQt5.QtCore import QThread, QObject, pyqtSignal, pyqtSlot
from functions.excel_action import readExcel
from functions.logging_setting import Log

log = Log(__name__).getLog()

DEFAULT_Location = [20, 80, "未知"]

SPLIT_CHAR = '='  # 工号=提交人

class DrawMapWorker(QObject):

    _finished = pyqtSignal()
    _signal = pyqtSignal(str)
    
    def __init__(self, excelPath='', mapsDir=''):
        super().__init__()
        self.excelPath = excelPath
        self.mapsDir = mapsDir

    @pyqtSlot()
    def work(self):
        stuDatas = readExcel(self.excelPath)
        self._signal.emit('请稍等...%d张地图生成中...' % len(stuDatas))
        for aStuData in stuDatas:
            self.drawGeoMap(aStuData)
        self._finished.emit()


    def drawGeoMap(self, aStuData):

        sno, sname = aStuData['info'].split(SPLIT_CHAR)
        sdata = aStuData['data']

        try:
            geo = Geo()

            def add_data(date, location):
                # print(location)
                try:
                    longitude, latitude, address = location
                except:
                    longitude, latitude, address = DEFAULT_Location
                geo.add_coordinate(date, longitude, latitude)
                geo.add(date, [(date, address)], type_=ChartType.EFFECT_SCATTER)

            geo.add_schema(maptype="china")

            # add data
            for date, data in sdata.items():
                add_data(date, data)

            title = f'{sno} {sname} 位置动态'

            geo.set_series_opts(label_opts=opts.LabelOpts(is_show=False))
            geo.set_global_opts(legend_opts=opts.LegendOpts(orient='vertical', pos_left='left', pos_top='10%'), title_opts=opts.TitleOpts(title=title))

            html_path = ospath_join(self.mapsDir, f'{sno}_{sname}.html')
            geo.render(html_path)
        except Exception as e:
            log.error(f'地图生成失败: {sno} {sname} - 可能是包含了文件名不可用的特殊字符', exc_info=True)
            self._signal.emit(f'错误：{sno} {sname} 地图生成失败！可能原因：工号或提交人姓名中包含文件名不可用的特殊字符，如：*。请修改后重试！')