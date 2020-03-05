# DingHelper

> 钉钉位置分析小程序

## 界面

### 主界面

![主界面](https://github.com/wxjing4me/DingHelper/blob/master/docs/page1.png)

### 生成位置文件界面

![生成位置文件界面](https://github.com/wxjing4me/DingHelper/blob/master/docs/page2.png)

## 功能

1. 利用经纬度判断今昨两日的位置变动是否属于【省外入闽】、【外地返榕】、【离闽】、【离榕】等。

   注：此处“榕”指的是：福州市鼓楼区、台江区、晋安区、仓山区、马尾区、闽侯县。

2. 计算当天位置与前一天位置的距离，计算当天位置与福师大的距离

3. 在地图上绘制每个人员随时间变化的地点变动

4. 将钉钉健康打卡导出的数据进行整合汇总

## 缺陷

* 需自行申请腾讯地图API的开发者密钥（即Key/Token）

> 申请入口：https://lbs.qq.com/dev/console/key/add

* 受制于腾讯地图API的每秒并发量（5次/秒）

* 受制于腾讯地图API的次数限制

* 运行在Windows下最佳

## Development

### Dependencies

- Python 3.6
- pyecharts
- PyQt5 5.14
- requests
- xlrd / xlsxwriter

### Installation

```
~$ pip install pyecharts
~$ pip install PyQt5
~$ pip install requests
~$ pip install xlrd
~$ pip install XlsxWriter
```

* PyQt5-5.14.1-5.14.1-cp35.cp36.cp37.cp38-none-win_amd64.whl
> https://pypi.org/project/PyQt5/#files

* pywin32-227-cp36-cp36m-win_amd64.whl
> https://pypi.org/project/pywin32/#files

### Run

```
> python app.py
```

### Package into exe

* Use pyinstaller

```
> pyinstaller app.spec
```

## 帮助文档

> https://docs.qq.com/doc/DZmJTeUZqYll2U2tR
