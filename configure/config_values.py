#-*-coding:utf-8-*-
import os

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
USER_SETTINGS_JSON = os.path.join(ROOT_DIR, 'settings', 'settings_user.json')
DEFAULT_SETTINGS_JSON = os.path.join(ROOT_DIR, 'settings', 'settings_default.json')

EXCEL_LOCATION = '生成位置文件'
EXCEL_PROFILE = '生成一人一档'

APP_NAME = 'DingHelper'
VERSION = 'v1.2.0'
AUTHOR = '@wxjing'
AUTHOR_GITHUB_URL = 'https://github.com/wxjing4me'
APP_GITHUB_URL = 'https://github.com/wxjing4me/DingHelper'
AUTHOR_TIP = '你发现了什么？点击有惊喜哦~'
APP_HELP_URL = 'https://docs.qq.com/doc/DZmJTeUZqYll2U2tR'
APP_ICON_PATH = 'images/favicon.ico'
A_URL_STYLE = 'text-decoration:none;color:black;'

DEFAULT_Address = {'nation': '未知'}
DEFAULT_Location = [20, 80, "未知"]

MAP_NAMES = {'AMAP': '高德', 'QQ': '腾讯'}

FJNU_Lat = '119.209920'
FJNU_Lng = '26.027120'

QQ_KEYX = '77NBZ-AAKCF-TXIJR-JMWWH-JZ6QS-UWBIT'
QQ_API_URL_LL2Address = 'https://apis.map.qq.com/ws/geocoder/v1/?'
QQ_API_URL_DISTANCE = 'https://apis.map.qq.com/ws/distance/v1/?'
QQ_API_URL_IP = 'https://apis.map.qq.com/ws/location/v1/ip?'
QQ_API_FORMAT = 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX-XXXXX'

AMAP_KEYX = 'x5i7xq2wt348qtax81iw54iqt78616x2'
AMAP_API_URL_IP = 'https://restapi.amap.com/v3/ip?'
AMAP_API_URL_LL2Address = 'https://restapi.amap.com/v3/geocode/regeo?'
AMAP_API_URL_Address2LL = 'https://restapi.amap.com/v3/geocode/geo?'
AMAP_API_URL_DISTANCE = 'https://restapi.amap.com/v3/distance?'
AMAP_API_FORMAT = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'


LOC_TYPE_STAY = "无变化"
LOC_TYPE_FUJIAN_IN = "【省外入闽】"
LOC_TYPE_FUJIAN_OUT = "【离闽】"
LOC_TYPE_FUZHOU_IN = "【外地返榕】"
LOC_TYPE_FUZHOU_OUT =  "【离榕】"
LOC_TYPE_FUZHOU_ELSE = "【“榕”随便走走】"
LOC_TYPE_FUJIAN_ELSE = "【“闽”随便走走】"
LOC_TYPE_OUTSIDE = "【“省外”随便走走】"
LOC_TYPE_ELSE = "【系统不清楚】"

LOC_NAME_FUJIAN = '福建省'
LOC_NAME_FUZHOU = '福州市'
LOC_NAME_FUZHOU_LIST = ['闽侯县', '鼓楼区', '台江区', '晋安区', '仓山区', '马尾区']

QQ_MAX_CNT_PER_SEC = 4
AMAP_MAX_CNT_PER_SEC = 49

SPLIT_CHAR = '='  # 工号=提交人

#----------------------------------------------------------
# 关于Excel
#----------------------------------------------------------


FONT_NAME_YAHEI = '微软雅黑'
FONT_SIZE = 10

STR_SNO = '工号'
STR_NAME = '提交人'
STR_TIME_LOC = '当前时间,当前地点'
STR_DATE = '填写周期'

STR_SID = '学号'
STR_SNAME = '姓名'

STR_UNDO = '-'

HEADER_REQUIRED = [STR_SNO, STR_NAME, STR_TIME_LOC]

# 一人一档表头
# HEADER_PROFILE = [STR_SID, STR_SNAME, '填写周期', '提交时间', '今日体温', '今日位置', '当前时间,当前地点', '今日有无以下症状', '今日个人动向', '今日接触人员情况', '其他情况说明', '目前健康状况', '目前所在城市']

HEADER_PROFILE = [STR_SID, STR_SNAME, '填写周期', '今日体温', '目前健康状况', '目前所在城市', '当前时间,当前地点']

LOC_QQ = '腾讯位置'
LOC_AMAP = '高德位置'
LOC_DING = '钉钉位置'
