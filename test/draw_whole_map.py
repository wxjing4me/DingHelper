from pyecharts.charts import Geo
from pyecharts import options as opts
from pyecharts.globals import ThemeType
from pyecharts.faker import Faker
from pyecharts.globals import ChartType, SymbolType

import xlrd
import json
import os

DEFAULT_Location = [20, 80, "未知"]

dir_path = os.path.join('C:','Users', '哇咔咔', 'Desktop')

def drawGeoMap(excel_path, mtype='china'):
    data = xlrd.open_workbook(excel_path)
    table = data.sheet_by_index(0)
    nrow = table.nrows
    dates = table.row_values(0)
    # print(dates)
    allData = {}
    for i in range(1, nrow):
        aNameData = {}
        sname = f"{table.cell(i,0).value} {table.cell(i,1).value}"
        for j in range(2, len(dates)):
            value = table.cell(i,j).value
            try:
                value = json.loads(value)
                value = value[1:4]
            except Exception as e:
                print(f"error: {e}")
                pass
            aNameData[dates[j]] = value
        allData[sname] = aNameData

    # print(f'学生数：{len(allData)}')

    title = '2020年7月29日能仔位置'
    geo = Geo(init_opts=opts.InitOpts(width='1000px', height='580px', page_title=title, theme=ThemeType.LIGHT))

    geo.add_schema(maptype=mtype)

    for name, data in allData.items():
        for date, location in data.items():
            aNameLocation = []
            try:
                longitude, latitude, address = location
            except Exception as e:
                print(f"Error: {e}")
                longitude, latitude, address = DEFAULT_Location
            # aNameLocation.append(('%s %s\n' % (name, date), address))
            # geo.add_coordinate('%s %s' % (name, date), longitude, latitude)
            sno, sname = name.split(' ')
            # print(sno, sname)
            aNameLocation.append(('%s %s' % (sno, sname), date))
            geo.add_coordinate('%s %s' % (sno, sname), longitude, latitude)
            geo.add(name, aNameLocation, type_=ChartType.EFFECT_SCATTER)

    geo.set_series_opts(label_opts=opts.LabelOpts(is_show=False))
    geo.set_global_opts(legend_opts=opts.LegendOpts(type_='scroll', orient='vertical', pos_left='left', pos_top='10%', is_show=False), title_opts=opts.TitleOpts(title=title))

    map_path = os.path.join(dir_path, '2020年7月29日能仔位置地图.html')
    geo.render(map_path)

    # os.system(f'"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" {os.path.abspath(map_path)}')

    print('done.')

if __name__ == '__main__':
    excel_path = os.path.join(dir_path, '每日健康打卡_2020-07-28.xlsx')

    mtype = 'china' # china # 福建
    drawGeoMap(excel_path, mtype)
    