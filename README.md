# DingMinProgram

> 钉钉位置分析小工具

## 界面

![界面](https://github.com/wxjing4me/DingHelper/blob/master/docs/page1.png)

## 功能

1. 利用经纬度判断今昨两日的位置变动是否属于【省外入闽】、【外地返榕】、【离闽】、【离榕】等。

   注：此处“榕”指的是：福州市鼓楼区、台江区、晋安区、仓山区、马尾区、闽侯县。

2. 在地图上绘制每个人员随时间变化的地点变动

## 缺陷

* 需自行申请腾讯地图API的开发者密钥（即Key/Token）

> 申请入口：https://lbs.qq.com/dev/console/key/add

* 受制于腾讯地图API的每秒并发量（5次/秒）

* 受制于腾讯地图API的次数限制

## Dependencies

- Python 3.6
- pyecharts
- PyQt5

#### 安装

```
~$ pip install pyecharts
~$ pip install PyQt5

```

## 帮助文档

> https://docs.qq.com/doc/DZmJTeUZqYll2U2tR

## Run

```
> python app.py
```
