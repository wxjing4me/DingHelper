from os.path import exists as os_path_exists, join as os_path_join
from json import dump as json_dump, load as json_load

from configure.logging_action import Log
from configure.config_values import *

log = Log(__name__).getLog()

DATA_DIR = os_path_join(ROOT_DIR, 'excels')
# 所使用的地图API类型: QQ, AMAP
MAP_TYPE = 'AMAP' 
# 第1行默认为表头, default:1
START_ROW = 1 
# 功能类型： clockIn: 打卡； SignOn: 签到
FUNC_TYPE = 'clockIn' 
# 是否显示距离，默认为True
SHOW_DISTANCE = True
# 仅处理Excel中的第一个工作表，默认为True
ONLY_FIRST_SHEET = True

def loadDefaultConfig():
    conf = {}
    with open(DEFAULT_SETTINGS_JSON, 'r', encoding='utf-8') as f:
        conf = json_load(f)
    if "DATA_DIR" not in conf:
        conf["DATA_DIR"] = DATA_DIR
        with open(DEFAULT_SETTINGS_JSON, 'w+', encoding='utf-8') as f:
            json_dump(conf, f)
    return conf

def saveConfig(config):
    with open(USER_SETTINGS_JSON, 'w+', encoding='utf-8') as f:
        json_dump(config, f)
        return True
    return False

def loadUserConfig():
    conf = {}
    with open(USER_SETTINGS_JSON, 'r', encoding='utf-8') as f:
        conf = json_load(f)
    return conf

def updateSettings():
    global MAP_TYPE, FUNC_TYPE, START_ROW, ONLY_FIRST_SHEET, SHOW_DISTANCE, DATA_DIR
    config = {}
    if not os_path_exists(USER_SETTINGS_JSON):
        config = loadDefaultConfig()
        log.debug('load Default Config')
    else:
        config = loadUserConfig()
        log.debug('load User Config')
    MAP_TYPE = config['MAP_TYPE'] if 'MAP_TYPE' in config else MAP_TYPE
    FUNC_TYPE = config['FUNC_TYPE'] if 'FUNC_TYPE' in config else FUNC_TYPE
    START_ROW = config['START_ROW'] if 'START_ROW' in config else START_ROW
    ONLY_FIRST_SHEET = config['HANDLE_SHEET']=='first' if 'HANDLE_SHEET' in config else ONLY_FIRST_SHEET
    SHOW_DISTANCE = config['SHOW_DISTANCE']=='show' if 'SHOW_DISTANCE' in config else SHOW_DISTANCE
    DATA_DIR = config['DATA_DIR'] if 'DATA_DIR' in config else DATA_DIR
    return config
    