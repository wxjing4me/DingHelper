from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QTextBrowser, QFileDialog
from PyQt5.QtGui import QIcon, QFont, QDesktopServices
from PyQt5.QtCore import QUrl, Qt, QThread, QDateTime

import os
import time
import sys
import shutil

from functions.tencent_api import *
from functions.excel_action import *
from functions.draw_map import *
from functions.analyse import AnalyseWorker
from functions.draw_map import DrawMapWorker

class MainWindow(QMainWindow):
    def __init__(self):
       super().__init__()

       self.ApiTokenOK = False
       self.ApiToken = ''
       self.MapsDir = ''
       self.excelPath = ''

       self.initUI()

    def initUI(self):

        font_Yahei = QFont("Microsoft YaHei")

        self.setWindowTitle('钉钉数据分析小程序')
        self.setWindowIcon(QIcon('images/favicon.ico'))
        self.setGeometry(300, 100, 800, 500)

        widget_main = QWidget()
        self.setCentralWidget(widget_main)
        self.setFont(font_Yahei)
        
        layout_main = QVBoxLayout()
        widget_main.setLayout(layout_main)
        
        widget_setToken = QWidget()

        layout_setToken = QHBoxLayout()

        label_setToken = QLabel('请填入API Token值：')
        self.input_setToken = QLineEdit(self)
        self.input_setToken.setPlaceholderText('XXXXX-XXXXX-XXXXX-XXXXX-XXXXX-XXXXX')
        
        self.btn_setToken = QPushButton('设置TOKEN')
        self.btn_setToken.clicked.connect(self.clickBtn_setToken)
        
        widget_setToken.setLayout(layout_setToken)

        layout_setToken.addWidget(label_setToken)
        layout_setToken.addWidget(self.input_setToken)
        layout_setToken.addWidget(self.btn_setToken)

        layout_main.addWidget(widget_setToken)

        widget_opt = QWidget()
        layout_opt = QHBoxLayout()
        widget_opt.setLayout(layout_opt)

        widget_buttons = QWidget()
        layout_buttons = QVBoxLayout()
        widget_buttons.setLayout(layout_buttons)

        btn_downloadExcel = QPushButton('下载Excel模板')
        btn_downloadExcel.clicked.connect(self.clickBtn_downloadExcel)
        layout_buttons.addWidget(btn_downloadExcel)

        self.btn_selectExcel = QPushButton('选择Excel文件')
        self.btn_selectExcel.clicked.connect(self.clickbtn_selectExcel)
        layout_buttons.addWidget(self.btn_selectExcel)

        self.btn_mergeExcel = QPushButton('整合Excel文件')
        # self.btn_mergeExcel.clicked.connect(self.clickbtn_mergeExcel)
        layout_buttons.addWidget(self.btn_mergeExcel)

        self.btn_startAnalyse = QPushButton('开始分析')
        self.btn_startAnalyse.clicked.connect(self.clickbtn_startAnalyse)
        self.btn_startAnalyse.setEnabled(False)
        layout_buttons.addWidget(self.btn_startAnalyse)

        self.btn_drawMap = QPushButton('生成地图')
        self.btn_drawMap.clicked.connect(self.clickBtn_drawMap)
        self.btn_drawMap.setEnabled(False)
        layout_buttons.addWidget(self.btn_drawMap)

        self.btn_output = QPushButton('导出结果')
        self.btn_output.clicked.connect(self.clickBtn_output)
        layout_buttons.addWidget(self.btn_output)

        btn_help = QPushButton('帮助')
        btn_help.clicked.connect(self.clickBtn_help)
        layout_buttons.addWidget(btn_help)

        layout_opt.addWidget(widget_buttons)

        self.output = QTextBrowser(self)
        layout_opt.addWidget(self.output)
        
        layout_main.addWidget(widget_opt)

        widget_status = QWidget()
        layout_status = QHBoxLayout()
        widget_status.setLayout(layout_status)

        self.status_label = QLabel('Tips: 不知道要怎么做？点击【帮助】看看吧~')
        
        author_label = QLabel('@wxjing')

        author_label.setAlignment(Qt.AlignRight)
        author_label.setText('<a href="https://github.com/wxjing4me" style="text-decoration:none;color:black">@wxjing</a>')
        author_label.setOpenExternalLinks(True)
        author_label.setToolTip('你发现了什么？点击有惊喜哦~')

        layout_status.addWidget(self.status_label)
        layout_status.addWidget(author_label)

        layout_main.addWidget(widget_status)

    def clickBtn_setToken(self):
        self.ApiToken = self.input_setToken.text()
        print('设置个Token值：%s' % self.ApiToken)
        res = testApiToken(self.ApiToken)
        if res['code'] == 1:
            self.ApiTokenOK = True
            self.updateOutput(res['msg'])
            self.input_setToken.setEnabled(False)
            self.btn_setToken.setEnabled(False)
            self.btn_selectExcel.setEnabled(True)
        else:
            self.updateOutput(res['msg'])

    def clickBtn_downloadExcel(self):
        templet_excel_path, _ = QFileDialog.getSaveFileName(self, '保存模板文件', 'templets/example.xlsx')
        if len(templet_excel_path.strip()) != 0:
            print(f'下载模板文件：{templet_excel_path}')
            try:
                shutil.copyfile('templets/example.xlsx', templet_excel_path)
                temp_excel_dir, _ = os.path.split(templet_excel_path)
                temp_excel_dir = os.path.normpath(temp_excel_dir)
                os.system("explorer.exe %s" % temp_excel_dir)
            except Exception as e:
                print('下载模板文件出错：%s' % e)
        
    def clickbtn_selectExcel(self):
        self.excelPath, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", os.getcwd(), "Excel files(*.xlsx , *.xls)")
        print('选择文件路径：%s' % self.excelPath)
        if len(self.excelPath.strip()) != 0:
            if testExcel(self.excelPath):
                self.updateOutput('选择文件路径：%s，请点击【开始分析】或【生成地图】' % self.excelPath)
                self.btn_startAnalyse.setEnabled(True)
                self.btn_drawMap.setEnabled(True)
            else:
                self.updateOutput('Excel格式有误，请重新选择文件')

    def clickbtn_startAnalyse(self):

        if self.ApiTokenOK:
            self.updateOutput('开始分析...', True)
            self.startAnalyse()
        else:
            self.updateOutput('请先设置Token值~')

    def clickBtn_output(self):
        now = time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime())
        log_file, _ = QFileDialog.getSaveFileName(self, '保存日志文件', '%s.txt' % now)
        print(log_file)
        if len(log_file.strip()) != 0:
            with open(log_file, 'w+') as f:
                f.write(self.output.toPlainText())

    def clickBtn_drawMap(self):
        self.MapsDir = QFileDialog.getExistingDirectory(self, '选择保存地图的文件夹路径', './')
        if self.MapsDir != '':
            self.btn_startAnalyse.setEnabled(False)
            self.drawMap()
            self.btn_startAnalyse.setEnabled(True)

    def clickBtn_help(self):
        QDesktopServices.openUrl(QUrl('https://docs.qq.com/doc/DZmJTeUZqYll2U2tR'))

    def startAnalyse(self):
        self.analyseWorker = AnalyseWorker(self.ApiToken, self.excelPath) 
        self.analyseThread = QThread()

        self.analyseWorker._signal.connect(self.updateOutput)
        self.analyseWorker.moveToThread(self.analyseThread)
        self.analyseWorker._finished.connect(self.analyseThread.quit)
        self.analyseThread.started.connect(self.analyseWorker.work)
        self.analyseThread.finished.connect(lambda: self.updateOutput('分析结束！可以【导出结果】存档~', True))

        self.analyseThread.start()

    def drawMap(self):
        
        self.updateOutput('地图将保存在该文件夹路径下：%s ' % self.MapsDir)

        self.drawMapWorker = DrawMapWorker(self.excelPath, self.MapsDir)  # no 
        self.drawMapThread = QThread()
        self.drawMapWorker._signal.connect(self.updateOutput)
        self.drawMapWorker.moveToThread(self.drawMapThread)
        self.drawMapWorker._finished.connect(self.drawMapThread.quit)
        self.drawMapThread.started.connect(self.drawMapWorker.work)
        self.drawMapThread.finished.connect(self.drawMapDone)

        self.drawMapThread.start()

    def drawMapDone(self):
        # 打开地图文件夹
        self.MapsDir = os.path.normpath(self.MapsDir)
        os.system("explorer.exe %s" % self.MapsDir)
        self.updateOutput('生成地图结束！', True)

    def updateOutput(self, msg, showTime=False):
        if showTime:
            now = QDateTime.currentDateTime().toString('yyyy-MM-dd hh:mm:ss')
            self.output.append(now)
        self.output.append(msg)