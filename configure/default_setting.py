#-*-coding:utf-8-*-
import os

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

AUTHOR = '@wxjing'
AUTHOR_GITHUB_URL = 'https://github.com/wxjing4me'
AUTHOR_TIP = '你发现了什么？点击有惊喜哦~'
APP_HELP_URL = 'https://docs.qq.com/doc/DZmJTeUZqYll2U2tR'
APP_ICON_PATH = 'images/favicon.ico'
APP_KEYX = '77NBZ-AAKCF-TXIJR-JMWWH-JZ6QS-UWBIT'

DEFAULT_Address = {'nation': '未知'}
DEFAULT_Location = [20, 80, "未知"]

FJNU_LL = '26.0271200000,119.2099200000'

API_URL_LL2Address = 'https://apis.map.qq.com/ws/geocoder/v1/?'
API_URL_DISTANCE = 'https://apis.map.qq.com/ws/distance/v1/?'
API_URL_IP = 'https://apis.map.qq.com/ws/location/v1/ip?'

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
LOC_NAME_FUZHOU_LIST = ['闽侯县' '鼓楼区', '台江区', '晋安区', '仓山区', '马尾区']

MAX_CNT_PER_SEC = 4

SPLIT_CHAR = '='  # 工号=提交人

#----------------------------------------------------------
# 关于Excel
#----------------------------------------------------------
# 第1行默认为表头
START_ROW = 1 # default:1

# 仅处理Excel中的第一个工作表
ONLY_FIRST_SHEET = True # 默认为False
FONT_NAME_YAHEI = '微软雅黑'
FONT_SIZE = 10

STR_SNO = '工号'
STR_NAME = '提交人'
STR_TIME_LOC = '当前时间,当前地点'

STR_UNDO = '-'

HEADER_REQUIRED = [STR_SNO, STR_NAME, STR_TIME_LOC]

LOC_TENCENT = '腾讯位置'
LOC_DING = '钉钉位置'