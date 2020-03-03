from json import loads as json_loads
from requests import get as requests_get

def testApiToken(token):
    token = token.strip()
    testRes = {}
    if len(token == 0:
        testRes['code'] = 0
        testRes['msg'] = 'Token值不能为空'
        return testRes
    elif len(token) != 36:
        testRes['code'] = 0
        testRes['msg'] = 'Token格式有误！'
        return testRes
    url = 'https://apis.map.qq.com/ws/location/v1/ip?key=%s' % token
    try:
        response = requests_get(url)
    except:
        testRes['code'] = 0
        testRes['msg'] = '网络不可用，请检查网络'
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
            testRes['msg'] = 'TOKEN设置成功，请点击【选择文件】吧~'
            return testRes