from pyecharts.charts import Geo
from pyecharts import options as opts
from pyecharts.faker import Faker
from pyecharts.globals import ChartType, SymbolType
import os

from PyQt5.QtCore import QThread, QObject, pyqtSignal, pyqtSlot
from functions.excel_action import readExcel

DEFAULT_Location = [20, 80, "未知"]

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

        sno, sname = aStuData['info']
        sdata = aStuData['data']

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

        title = '%s %s 位置动态' % (sno, sname)

        geo.set_series_opts(label_opts=opts.LabelOpts(is_show=False))
        geo.set_global_opts(legend_opts=opts.LegendOpts(orient='vertical', pos_left='left', pos_top='10%'), title_opts=opts.TitleOpts(title=title))

        html_path = os.path.join(self.mapsDir, '%s_%s.html' % (sno, sname))
        geo.render(html_path)