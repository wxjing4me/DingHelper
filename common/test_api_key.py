#-*-coding:utf-8-*-
from json import loads as json_loads
from requests import get as requests_get

from configure.logging_action import Log
from configure.config_values import *

log = Log(__name__).getLog()

def testApiKey(key, mtype):
    '''测试API Key是否可用
    高德地图：status=1表示成功；status=0表示失败
    腾讯地图：status=0表示成功；status!=0表示失败
    '''
    key = key.strip()
    testRes = {}
    if len(key) == 0:
        testRes['code'] = 0
        testRes['msg'] = '提示：Key值不能为空'
        return testRes
    elif (mtype == 'AMAP' and len(key) != 32) or (mtype == 'QQ' and len(key) != 35):
        testRes['code'] = 0
        testRes['msg'] = '提示：Key格式有误！'
        return testRes
    url = f"{eval(mtype+'_API_URL_IP')}key={key}"
    try:
        response = requests_get(url)
    except:
        testRes['code'] = 0
        testRes['msg'] = '提示：网络不可用，请检查网络'
        return testRes
    if response.status_code != 200:
        print('response.status_code=%d' % response.status_code)
        testRes['code'] = 0
        testRes['msg'] = '请求服务失败！错误码：%d' % response.status_code
        return testRes
    else:
        res = json_loads(response.text)
        # print(res)
        if mtype == 'QQ' and res['status'] != 0:
            testRes['code'] = 0
            testRes['msg'] = f"API Key设置错误：{res['message']}；请查询文档后再试"
        elif mtype == 'QQ' and res['status'] == 0:
            testRes['code'] = 1
            testRes['msg'] = '提示：腾讯地图API Key设置成功！'
        elif mtype == 'AMAP' and int(res['status']) == 0:
            testRes['code'] = 0
            testRes['msg'] = f"API Key设置错误：{res['info']}；请查询文档后再试"
        elif mtype == 'AMAP' and int(res['status']) == 1:
            testRes['code'] = 1
            testRes['msg'] = '提示：高德地图API Key设置成功！'
        else:
            testRes['code'] = 0
            testRes['msg'] = '提示：未知错误！'
            log.error(f"type={mtype}, res['status']={res['status']}")
        return testRes