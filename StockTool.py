import urllib3
from datetime import datetime

def createManager():
    return urllib3.PoolManager()

def getPrice(http, code):
    r = http.request('GET', 'http://hq.sinajs.cn/list=sz' + code)
    if r.status == 200:
        s = r.data.decode("gb2312")
        s= s.split('\"')[1]
        s = s.split(',')
        d = datetime.strptime(s[30] + ' ' + s[31], "%Y-%m-%d %H:%M:%S")
        print(code, d, s[3])

http = createManager()
getPrice(http, '002063')