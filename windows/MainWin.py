#-*-coding:utf-8-*-
from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QTextBrowser, QFileDialog, QMessageBox
from PyQt5.QtGui import QIcon, QFont, QDesktopServices
from PyQt5.QtCore import QUrl, Qt, QThread, QDateTime

from os import getcwd as os_getcwd, system as os_system
from os.path import split as ospath_split, normpath as ospath_normpath
from time import strftime as time_strftime, localtime as time_localtime

from functions.tencent_api import testApiKey
from functions.excel_action import testSpectExcel
from functions.analyse import AnalyseWorker
from functions.draw_map import DrawMapWorker

from functions.logging_setting import Log

log = Log(__name__).getLog()

class MainWindow(QMainWindow):
    def __init__(self):
       super().__init__()

       self.ApiKeyOK = False
       self.ApiKey = ''
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
        
        widget_setKey = QWidget()

        layout_setKey = QHBoxLayout()

        label_setKey = QLabel('腾讯地图API Key：')
        self.input_setKey = QLineEdit(self)
        self.input_setKey.setPlaceholderText('XXXXX-XXXXX-XXXXX-XXXXX-XXXXX-XXXXX')
        self.input_setKey.returnPressed.connect(self.label_Enter)
        
        self.btn_setKey = QPushButton('设置Key')
        self.btn_setKey.clicked.connect(self.clickBtn_setKey)
        
        widget_setKey.setLayout(layout_setKey)

        layout_setKey.addWidget(label_setKey)
        layout_setKey.addWidget(self.input_setKey)
        layout_setKey.addWidget(self.btn_setKey)

        layout_main.addWidget(widget_setKey)

        widget_opt = QWidget()
        layout_opt = QHBoxLayout()
        widget_opt.setLayout(layout_opt)

        widget_buttons = QWidget()
        layout_buttons = QVBoxLayout()
        widget_buttons.setLayout(layout_buttons)

        self.btn_mergeExcel = QPushButton('生成位置文件')
        layout_buttons.addWidget(self.btn_mergeExcel)

        self.btn_selectExcel = QPushButton('选择文件')
        self.btn_selectExcel.clicked.connect(self.clickbtn_selectExcel)
        self.btn_selectExcel.setToolTip("<b>Excel格式如下：</b><br><img src='./images/excel_tip.png'>")
        layout_buttons.addWidget(self.btn_selectExcel)

        self.btn_startAnalyse = QPushButton('分析位移')
        self.btn_startAnalyse.clicked.connect(self.clickbtn_startAnalyse)
        self.btn_startAnalyse.setEnabled(False)
        layout_buttons.addWidget(self.btn_startAnalyse)

        self.btn_stopAnalyse = QPushButton('停止分析')
        self.btn_stopAnalyse.clicked.connect(self.clickbtn_stopAnalyse)
        self.btn_stopAnalyse.setEnabled(False)
        layout_buttons.addWidget(self.btn_stopAnalyse)

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
        self.output.append('提示：从【生成位置文件】开始旅程吧~')
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

    def clickBtn_setKey(self):
        self.ApiKey = self.input_setKey.text()
        res = testApiKey(self.ApiKey)
        if res['code'] == 1:
            self.ApiKeyOK = True
            self.updateOutput(res['msg'])
            self.input_setKey.setEnabled(False)
            self.btn_setKey.setEnabled(False)
            self.btn_selectExcel.setEnabled(True)
        else:
            self.input_setKey.setText('')
            self.updateOutput(res['msg'])

    def clickbtn_selectExcel(self):
        self.excelPath, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", os_getcwd(), "Excel files(*.xlsx , *.xls)")
        # print('选择文件路径：%s' % self.excelPath)
        if len(self.excelPath.strip()) != 0:
            testRes = testSpectExcel(self.excelPath)
            if testRes['code'] == 1:
                self.updateOutput('选择文件路径：%s\n提示：点击【分析位移】或【生成地图】' % self.excelPath)
                self.btn_startAnalyse.setEnabled(True)
                self.btn_drawMap.setEnabled(True)
            else:
                self.showMessageBox(testRes['msg']).exec()

    def clickbtn_startAnalyse(self):

        if self.ApiKeyOK and self.excelPath != '':
            self.updateOutput(f'开始分析...{self.excelPath}', True)
            self.btn_startAnalyse.setEnabled(False)
            self.btn_stopAnalyse.setEnabled(True)
            self.btn_mergeExcel.setEnabled(False)
            self.btn_selectExcel.setEnabled(False)
            self.btn_drawMap.setEnabled(False)
            self.startAnalyse()
        else:
            self.updateOutput('提示：请先设置腾讯地图API开发者密钥（Key）~')

    def clickbtn_stopAnalyse(self):
        self.analyseWorker.stop()
        self.analyseThread.quit()

    def clickBtn_output(self):
        now = time_strftime("%Y-%m-%d-%H-%M-%S", time_localtime())
        log_file, _ = QFileDialog.getSaveFileName(self, '保存结果文件', '%s.html' % now)
        if len(log_file.strip()) != 0:
            with open(log_file, 'w+') as f:
                f.write(self.output.toHtml())

    def clickBtn_drawMap(self):
        self.MapsDir = QFileDialog.getExistingDirectory(self, '选择保存地图的文件夹路径', './')
        if self.MapsDir != '':
            self.btn_startAnalyse.setEnabled(False)
            self.drawMap()
            self.btn_startAnalyse.setEnabled(True)

    def clickBtn_help(self):
        QDesktopServices.openUrl(QUrl('https://docs.qq.com/doc/DZmJTeUZqYll2U2tR'))

    def startAnalyse(self):
        self.analyseWorker = AnalyseWorker(self.ApiKey, self.excelPath) 
        self.analyseThread = QThread()

        self.analyseWorker._signal.connect(self.updateOutput)
        self.analyseWorker.moveToThread(self.analyseThread)
        self.analyseWorker._finished.connect(self.analyseThread.quit)
        self.analyseThread.started.connect(self.analyseWorker.work)
        self.analyseThread.finished.connect(self.startAnalyseEnd)

        self.analyseThread.start()

    def startAnalyseEnd(self):
        self.analyseThread.quit()
        self.updateOutput('分析结束！可以点击【导出结果】存档~', True)
        self.btn_stopAnalyse.setEnabled(False)
        self.btn_startAnalyse.setEnabled(True)
        self.btn_mergeExcel.setEnabled(True)
        self.btn_selectExcel.setEnabled(True)
        self.btn_drawMap.setEnabled(True)

    def drawMap(self):
        
        self.updateOutput('地图将保存在该文件夹路径下：%s ' % self.MapsDir)

        self.drawMapWorker = DrawMapWorker(self.excelPath, self.MapsDir)
        self.drawMapThread = QThread()
        self.drawMapWorker._signal.connect(self.updateOutput)
        self.drawMapWorker.moveToThread(self.drawMapThread)
        self.drawMapWorker._finished.connect(self.drawMapThread.quit)
        self.drawMapThread.started.connect(self.drawMapWorker.work)
        self.drawMapThread.finished.connect(self.drawMapDone)

        self.drawMapThread.start()

    def drawMapDone(self):
        # 打开地图文件夹
        self.MapsDir = ospath_normpath(self.MapsDir)
        os_system("explorer.exe %s" % self.MapsDir)
        self.updateOutput('生成地图结束！')

    def updateOutput(self, msg, showTime=False):
        if showTime:
            now = QDateTime.currentDateTime().toString('yyyy-MM-dd hh:mm:ss')
            self.output.append(f"<span style='color:blue'>{now}</span>")
        self.output.append(msg)

    def label_Enter(self):
        x = self.input_setKey.text().strip()
        if len(x) == 6:
            key = '77NBZ-AAKCF-TXIJR-JMWWH-JZ6QS-UWBIT'
            code = str.maketrans(x.upper(), 'LOVEPY')
            tip = key.translate(code)
            self.input_setKey.setText(tip)

    def showMessageBox(self, msg):
        msgBox = QMessageBox()
        msgBox.setWindowTitle('错误')
        msgBox.setWindowIcon(QIcon('images/favicon.ico'))
        msgBox.setText(f"友情提示：{msg}\n建议使用【生成位置文件】进行生成")
        return msgBox