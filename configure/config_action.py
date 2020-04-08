from os.path import exists as os_path_exists, join as os_path_join
from json import dump as json_dump, load as json_load, dumps as json_dumps

from configure.logging_action import Log
from configure.config_values import *

log = Log(__name__).getLog()

# 配置项默认值 -------------------------------
# DATA_DIR: 默认数据存放路径
DATA_DIR = os_path_join(ROOT_DIR, 'excels')
# BROWSER_PATH: 默认浏览器路径
BROWSER_PATH = os_path_join('C:', 'Program Files (x86)', 'Google', 'Chrome', 'Application', 'chrome.exe')
# MAP_TYPE: 所使用的地图API类型, 可选: QQ, AMAP
MAP_TYPE = 'AMAP' 
# START_ROW: 第1行默认为表头, default: 1
START_ROW = 1 
# FUNC_TYPE: 功能类型, 可选: clockIn: 打卡； SignOn: 签到
FUNC_TYPE = 'clockIn' 
# SHOW_DISTANCE: 是否显示距离, default: True
SHOW_DISTANCE = True
# ONLY_FIRST_SHEET: 仅处理Excel中的第一个工作表, default: True
ONLY_FIRST_SHEET = True
# ----------------------------------------------

def initDefaultConfig():
    '''初始化默认设置文件
    '''
    initConfig = {}
    initConfig['MAP_TYPE'] = MAP_TYPE
    initConfig['SHOW_DISTANCE'] = 'show' if SHOW_DISTANCE==True else 'hide'
    initConfig['FUNC_TYPE'] = FUNC_TYPE
    initConfig['HANDLE_SHEET'] = 'first' if ONLY_FIRST_SHEET==True else 'all'
    initConfig['START_ROW'] = START_ROW
    initConfig['BROWSER_PATH'] = BROWSER_PATH
    initConfig['DATA_DIR'] = DATA_DIR
    string = json_dumps(initConfig)
    with open(DEFAULT_SETTINGS_JSON, 'w', encoding='utf-8') as f:
        f.write(string)
    log.debug('init Default Config')

def loadDefaultConfig():
    conf = {}
    if not os_path_exists(DEFAULT_SETTINGS_JSON):
        initDefaultConfig()
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
    global MAP_TYPE, FUNC_TYPE, START_ROW, ONLY_FIRST_SHEET, SHOW_DISTANCE, DATA_DIR, BROWSER_PATH
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
    BROWSER_PATH = config['BROWSER_PATH'] if 'BROWSER_PATH' in config else BROWSER_PATH
    DATA_DIR = config['DATA_DIR'] if 'DATA_DIR' in config else DATA_DIR
    return config
    