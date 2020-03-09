#-*-coding:utf-8-*-
from PyQt5.QtWidgets import QMainWindow, QMenuBar, QMenu, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QTextBrowser, QFileDialog, QMessageBox, QAction
from PyQt5.QtGui import QIcon, QFont, QDesktopServices
from PyQt5.QtCore import QUrl, Qt, QThread, QDateTime

from os import getcwd as os_getcwd, system as os_system
from os.path import split as ospath_split, normpath as ospath_normpath
from time import strftime as time_strftime, localtime as time_localtime

from common.test_api_key import testApiKey
from common.excel_action import testSpectExcel
from common.analyse import AnalyseWorker
from common.draw_map import DrawMapWorker

from configure.logging_action import Log
from configure.config_values import *
from configure import config_action as confAct

log = Log(__name__).getLog()

class MainWindow(QMainWindow):
    def __init__(self):
       super().__init__()

       self.ApiKeyOK = False
       self.ApiKey = ''
       self.MapsDir = ''
       self.excelPath = ''

       self.initUI()
       self.refreshUI()

    def initUI(self):
        
        font_Yahei = QFont(FONT_NAME_YAHEI)

        self.setWindowTitle('钉钉数据分析小程序')
        self.setWindowIcon(QIcon(APP_ICON_PATH))
        self.setGeometry(300, 100, 800, 500)

        widget_main = QWidget()
        self.setCentralWidget(widget_main)
        self.setFont(font_Yahei)
        
        layout_main = QVBoxLayout()
        widget_main.setLayout(layout_main)
        
        widget_setKey = QWidget()

        layout_setKey = QHBoxLayout()

        self.label_setKey = QLabel()
        self.input_setKey = QLineEdit(self)
        self.input_setKey.returnPressed.connect(self.label_Enter)
        
        self.btn_setKey = QPushButton('设置Key')
        self.btn_setKey.clicked.connect(self.clickBtn_setKey)
        
        widget_setKey.setLayout(layout_setKey)

        layout_setKey.addWidget(self.label_setKey)
        layout_setKey.addWidget(self.input_setKey)
        layout_setKey.addWidget(self.btn_setKey)

        layout_main.addWidget(widget_setKey)

        widget_opt = QWidget()
        layout_opt = QHBoxLayout()
        widget_opt.setLayout(layout_opt)

        widget_buttons = QWidget()
        layout_buttons = QVBoxLayout()
        widget_buttons.setLayout(layout_buttons)

        self.btn_setting = QPushButton('设置')
        layout_buttons.addWidget(self.btn_setting)

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

        btn_clear = QPushButton('清空结果')
        btn_clear.clicked.connect(self.clickBtn_clear)
        layout_buttons.addWidget(btn_clear)

        self.btn_output = QPushButton('导出结果')
        self.btn_output.clicked.connect(self.clickBtn_output)
        layout_buttons.addWidget(self.btn_output)

        layout_opt.addWidget(widget_buttons)

        self.output = QTextBrowser(self)
        self.output.append('提示：从【生成位置文件】开始旅程吧~')
        layout_opt.addWidget(self.output)
        
        layout_main.addWidget(widget_opt)

        widget_status = QWidget()
        layout_status = QHBoxLayout()
        widget_status.setLayout(layout_status)

        status_label = QLabel(f"Tips: 不知道要怎么做？看看<a href={APP_HELP_URL} style={A_URL_STYLE}>【帮助文档】</a>吧~")
        status_label.setOpenExternalLinks(True)
        
        author_label = QLabel()
        author_label.setAlignment(Qt.AlignRight)
        author_label.setText(f"<a href={APP_GITHUB_URL} style={A_URL_STYLE}>{APP_NAME}</a> {VERSION} <a href={AUTHOR_GITHUB_URL} style={A_URL_STYLE}>{AUTHOR}</a>")
        author_label.setOpenExternalLinks(True)
        author_label.setToolTip(AUTHOR_TIP)

        layout_status.addWidget(status_label)
        layout_status.addWidget(author_label)

        layout_main.addWidget(widget_status)

    def refreshUI(self):
        confAct.updateSettings()
        self.label_setKey.setText(f"{MAP_NAMES[confAct.MAP_TYPE]}地图API Key：")
        self.input_setKey.setText('')
        self.input_setKey.setEnabled(True)
        self.input_setKey.setPlaceholderText(eval(confAct.MAP_TYPE+'_API_FORMAT'))
        self.btn_setKey.setEnabled(True)

    def clickBtn_setKey(self):
        self.ApiKey = self.input_setKey.text()
        res = testApiKey(self.ApiKey, confAct.MAP_TYPE)
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
        self.excelPath, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", confAct.DATA_DIR, "Excel files(*.xlsx , *.xls)")
        if len(self.excelPath.strip()) != 0:
            testRes = testSpectExcel(self.excelPath)
            if testRes['code'] == 1:
                self.updateOutput(f"<span style='background-color:gray;color:white;'>选择文件路径：{self.excelPath}</span><br>提示：点击【分析位移】或【生成地图】")
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
            self.updateOutput(f"提示：请先设置{MAP_NAMES[confAct.MAP_TYPE]}地图API开发者密钥（Key）~")

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
        self.MapsDir = QFileDialog.getExistingDirectory(self, '选择保存地图的文件夹路径', confAct.DATA_DIR)
        if self.MapsDir != '':
            self.btn_startAnalyse.setEnabled(False)
            self.drawMap()
            self.btn_startAnalyse.setEnabled(True)

    def clickBtn_clear(self):
        self.output.clear()

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
            if confAct.MAP_TYPE == 'AMAP':
                code = str.maketrans(x, 'abcdef')
            elif confAct.MAP_TYPE == 'QQ':
                code = str.maketrans(x.upper(), 'LOVEPY')
            tip = eval(confAct.MAP_TYPE+'_KEYX').translate(code)
            self.input_setKey.setText(tip)

    def showMessageBox(self, msg):
        msgBox = QMessageBox()
        msgBox.setWindowTitle('错误')
        msgBox.setWindowIcon(QIcon(APP_ICON_PATH))
        msgBox.setText(f"友情提示：{msg}\n建议使用【生成位置文件】进行生成")
        return msgBox
