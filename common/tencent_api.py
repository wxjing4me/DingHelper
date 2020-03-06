#-*-coding:utf-8-*-
from json import loads as json_loads
from requests import get as requests_get

def testApiKey(key):
    key = key.strip()
    testRes = {}
    if len(key) == 0:
        testRes['code'] = 0
        testRes['msg'] = '提示：Key值不能为空'
        return testRes
    elif len(key) != 35:
        testRes['code'] = 0
        testRes['msg'] = '提示：Key格式有误！'
        return testRes
    url = 'https://apis.map.qq.com/ws/location/v1/ip?key=%s' % key
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
        if res['status'] != 0:
            testRes['code'] = 0
            testRes['msg'] = '腾讯地图API错误：%s；请查询文档后再试' % res['message']
            return testRes
        else:
            testRes['code'] = 1
            testRes['msg'] = '提示：Key设置成功！'
            return testRes