#-*-coding:utf-8-*-
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFileDialog, QCheckBox, QListWidget, QListWidgetItem, QLineEdit, QTextBrowser
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import Qt, QThread

from os import system as os_system
from os.path import split as os_path_split, normpath as os_path_normpath
from time import strftime as time_strftime, localtime as time_localtime

from common.profiles import ProfilesWorker
from configure.config_values import *
from configure import config_action as confAct
from configure.logging_action import Log

log = Log(__name__).getLog()

class ProfileWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.InfoPath = ''
        self.excelsDir = ''
        self.profilesDir = ''

        self.initUI()

    def initUI(self):

        font_Yahei = QFont(FONT_NAME_YAHEI)

        self.setWindowTitle(TITLE_PROFILE)
        self.setWindowIcon(QIcon(APP_ICON_PATH))
        self.setGeometry(450, 200, 600, 400)
        
        layout_main = QVBoxLayout()
        self.setLayout(layout_main)
        
        self.setFont(font_Yahei)

        #-----选择文件路径---------------------------------------

        widget_setInfoPath = QWidget()
        layout_setInfoPath = QHBoxLayout()
        self.label_setInfoPath = QLabel('学生信息文件')
        self.input_setInfoPath = QLineEdit(self)
        self.input_setInfoPath.setReadOnly(True)
        self.btn_setInfoPath = QPushButton('选择文件')
        self.btn_setInfoPath.clicked.connect(self.clickBtn_setInfoPath)
        widget_setInfoPath.setLayout(layout_setInfoPath)
        layout_setInfoPath.addWidget(self.label_setInfoPath)
        layout_setInfoPath.addWidget(self.input_setInfoPath)
        layout_setInfoPath.addWidget(self.btn_setInfoPath)

        widget_setExcelsDir = QWidget()
        layout_setExcelsDir = QHBoxLayout()
        self.label_setExcelsDir = QLabel('每日打卡数据')
        self.input_setExcelsDir = QLineEdit(self)
        self.input_setExcelsDir.setReadOnly(True)
        self.btn_setExcelsDir = QPushButton('选择文件夹')
        self.btn_setExcelsDir.clicked.connect(self.clickBtn_setExcelsDir)
        widget_setExcelsDir.setLayout(layout_setExcelsDir)
        layout_setExcelsDir.addWidget(self.label_setExcelsDir)
        layout_setExcelsDir.addWidget(self.input_setExcelsDir)
        layout_setExcelsDir.addWidget(self.btn_setExcelsDir)

        widget_setProfilesDir = QWidget()
        layout_setProfilesDir = QHBoxLayout()
        self.label_setProfilesDir = QLabel('保存一人一档路径')
        self.input_setProfilesDir = QLineEdit(self)
        self.input_setProfilesDir.setReadOnly(True)
        self.btn_setProfilesDir = QPushButton('选择文件夹')
        self.btn_setProfilesDir.clicked.connect(self.clickBtn_setProfilesDir)
        widget_setProfilesDir.setLayout(layout_setProfilesDir)
        layout_setProfilesDir.addWidget(self.label_setProfilesDir)
        layout_setProfilesDir.addWidget(self.input_setProfilesDir)
        layout_setProfilesDir.addWidget(self.btn_setProfilesDir)

        #-----按钮模块---------------------------------------

        widget_btns = QWidget()
        layout_btns = QHBoxLayout()

        self.btn_clearPath = QPushButton('重置')
        self.btn_clearPath.clicked.connect(self.clickBtn_clearPath)

        self.btn_createProfile = QPushButton('开始生成一人一档')
        self.btn_createProfile.clicked.connect(self.clickBtn_createProfile)

        self.btn_openProfilesDir = QPushButton('打开路径')
        self.btn_openProfilesDir.clicked.connect(self.clickBtn_openDir)
        self.btn_openProfilesDir.setEnabled(False)

        widget_btns.setLayout(layout_btns)
        layout_btns.addWidget(self.btn_clearPath)
        layout_btns.addWidget(self.btn_createProfile)
        layout_btns.addWidget(self.btn_openProfilesDir)

        layout_main.addWidget(widget_setInfoPath)
        layout_main.addWidget(widget_setExcelsDir)
        layout_main.addWidget(widget_setProfilesDir)
        layout_main.addWidget(widget_btns)

        #-----文本框---------------------------------------

        self.output = QTextBrowser(self)
        self.output.append('开始啦~')
        layout_main.addWidget(self.output)
        
        #-----状态显示---------------------------------------

        widget_status = QWidget()
        layout_status = QHBoxLayout()
        widget_status.setLayout(layout_status)

        status_label = QLabel(f"Tips: 不知道要怎么做？看看<a href={APP_HELP_URL} style={A_URL_STYLE}>【帮助文档】</a>吧~")
        status_label.setOpenExternalLinks(True)
        
        author_label = QLabel()
        author_label.setAlignment(Qt.AlignRight)
        author_label.setText(f'<a href={AUTHOR_GITHUB_URL} style="text-decoration:none;color:black">{AUTHOR}</a>')
        author_label.setOpenExternalLinks(True)
        author_label.setToolTip(AUTHOR_TIP)

        layout_status.addWidget(status_label)
        layout_status.addWidget(author_label)

        layout_main.addWidget(widget_status)
    
    def openWin(self):
        self.__init__()
        self.show()
        
    def clickBtn_setInfoPath(self):
        excelPath, _ = QFileDialog.getOpenFileName(self, "选择学生信息Excel文件", confAct.DATA_DIR, "Excel files(*.xlsx , *.xls)")
        if len(excelPath.strip()) != 0:
            self.InfoPath = excelPath
            self.input_setInfoPath.setText(self.InfoPath)
        
    def clickBtn_setExcelsDir(self):
        openDir = QFileDialog.getExistingDirectory(self, '选择每日打卡数据的文件夹路径', confAct.DATA_DIR)
        if openDir != '':
            self.excelsDir = openDir
            self.input_setExcelsDir.setText(self.excelsDir)

    def clickBtn_setProfilesDir(self):
        openDir = QFileDialog.getExistingDirectory(self, '选择保存一人一档的文件夹路径', confAct.DATA_DIR)
        if openDir != '':
            self.profilesDir = openDir
            self.input_setProfilesDir.setText(self.profilesDir)

    def clickBtn_createProfile(self):
        if self.infoPath != '' and self.excelsDir != '' and self.profilesDir != '':
            self.btn_setInfoPath.setEnabled(False)
            self.btn_setExcelsDir.setEnabled(False)
            self.btn_setProfilesDir.setEnabled(False)
            self.btn_clearPath.setEnabled(False)
            self.btn_createProfile.setEnabled(False)
            self.btn_openProfilesDir.setEnabled(False)

            self.createProfiles()

            self.btn_setInfoPath.setEnabled(True)
            self.btn_setExcelsDir.setEnabled(True)
            self.btn_setProfilesDir.setEnabled(True)
            self.btn_clearPath.setEnabled(True)
            self.btn_createProfile.setEnabled(True)
            self.btn_openProfilesDir.setEnabled(True)

    def clickBtn_clearPath(self):
        self.InfoPath = ''
        self.input_setInfoPath.setText('')
        self.excelsDir = ''
        self.input_setExcelsDir.setText('')
        self.profilesDir = ''
        self.input_setProfilesDir.setText('')
        self.output.clear()
    
    def clickBtn_openDir(self):
        profilesDir = os_path_normpath(self.profilesDir)
        os_system("explorer.exe %s" % profilesDir)

    def createProfiles(self):
        self.profilesWorker = ProfilesWorker(self.InfoPath, self.excelsDir, self.profilesDir)
        self.profilesThread = QThread()
        self.profilesWorker._signal.connect(self.appendStatus)
        self.profilesWorker.moveToThread(self.profilesThread)
        self.profilesWorker._finished.connect(self.profilesThread.quit)
        self.profilesThread.started.connect(self.profilesWorker.work)
        self.profilesThread.finished.connect(lambda: self.appendStatus(f'生成一人一档文件结束!【打开路径】吧~ <br>文件夹路径：{self.profilesDir}'))

        self.profilesThread.start()

    def appendStatus(self, msg):
        self.output.append(msg)
