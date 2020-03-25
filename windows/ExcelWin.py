#-*-coding:utf-8-*-
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFileDialog, QCheckBox, QListWidget, QListWidgetItem, QMessageBox
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import Qt, QThread

from os import system as os_system
from os.path import split as os_path_split, normpath as os_path_normpath
from time import strftime as time_strftime, localtime as time_localtime

from common.excel_action import MergeExcelWorker, CreateProfilesWorker, testRawExcel
from configure.config_values import *
from configure import config_action as confAct
from configure.logging_action import Log

log = Log(__name__).getLog()

class ExcelWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.switchType = EXCEL_LOCATION
        self.excelPathList = []
        self.excelOutputPath = ''
        self.excelDir = ''
        self.profilesDir = ''

        self.initUI()

    def initUI(self):

        font_Yahei = QFont(FONT_NAME_YAHEI)

        self.setWindowTitle(EXCEL_LOCATION)
        self.setWindowIcon(QIcon(APP_ICON_PATH))
        self.setGeometry(450, 200, 600, 400)
        
        layout_main = QVBoxLayout()
        self.setLayout(layout_main)
        
        self.setFont(font_Yahei)

        #-----说明模块---------------------------------------
        widget_intro = QWidget()
        layout_intro = QHBoxLayout()

        label_intro = QLabel('说明：添加的文件必须是从钉钉【员工健康】中导出的原始数据表\n  至少添加2个文件哦！')
        self.btn_switch = QPushButton(f'切换为【{EXCEL_PROFILE}】')
        self.btn_switch.clicked.connect(self.clickBtn_switch)

        widget_intro.setLayout(layout_intro)

        layout_intro.addWidget(label_intro)
        layout_intro.addWidget(self.btn_switch)


        layout_main.addWidget(widget_intro)

        widget_opt = QWidget()
        layout_opt = QHBoxLayout()
        widget_opt.setLayout(layout_opt)

        #-----按钮模块---------------------------------------
        widget_buttons = QWidget()
        layout_buttons = QVBoxLayout()
        widget_buttons.setLayout(layout_buttons)

        self.btn_addExcel = QPushButton('添加文件')
        self.btn_addExcel.clicked.connect(self.clickBtn_addExcel)
        layout_buttons.addWidget(self.btn_addExcel)

        self.btn_deleteExcel = QPushButton('删除文件')
        self.btn_deleteExcel.clicked.connect(self.clickbtn_deleteExcel)
        layout_buttons.addWidget(self.btn_deleteExcel)

        self.btn_clearExcel = QPushButton('清空文件')
        self.btn_clearExcel.clicked.connect(self.clickbtn_clearExcel)
        layout_buttons.addWidget(self.btn_clearExcel)

        self.btn_mergeExcel = QPushButton('开始生成')
        self.btn_mergeExcel.clicked.connect(self.clickbtn_mergeExcel)
        self.btn_mergeExcel.setEnabled(False)
        layout_buttons.addWidget(self.btn_mergeExcel)

        self.btn_openMergeExcelDir = QPushButton('打开路径')
        self.btn_openMergeExcelDir.clicked.connect(self.clickbtn_openDir)
        self.btn_openMergeExcelDir.setEnabled(False)
        layout_buttons.addWidget(self.btn_openMergeExcelDir)
        
        layout_opt.addWidget(widget_buttons)

        #-----文件列表---------------------------------------
        widget_path = QWidget()
        self.layout_pathBox = QVBoxLayout()
        widget_path.setLayout(self.layout_pathBox)

        label_pathBox = QLabel('Excel文件列表')
        
        self.widget_pathListBox = QListWidget()

        self.layout_pathBox.addWidget(label_pathBox)
        self.layout_pathBox.addWidget(self.widget_pathListBox)
        layout_opt.addWidget(widget_path)

        layout_main.addWidget(widget_opt)

        #-----状态显示---------------------------------------
        widget_status = QWidget()
        layout_status = QHBoxLayout()
        widget_status.setLayout(layout_status)

        self.status_label = QLabel('请点击【添加文件】开始~')
        
        author_label = QLabel()

        author_label.setAlignment(Qt.AlignRight)
        author_label.setText(f'<a href={AUTHOR_GITHUB_URL} style="text-decoration:none;color:black">{AUTHOR}</a>')
        author_label.setOpenExternalLinks(True)
        author_label.setToolTip(AUTHOR_TIP)

        layout_status.addWidget(self.status_label)
        layout_status.addWidget(author_label)

        layout_main.addWidget(widget_status)
    
    def openWin(self):
        self.__init__()
        self.show()
        
    def clickBtn_switch(self):
        self.btn_switch.setText(f'切换为【{self.switchType}】')
        if self.switchType == EXCEL_LOCATION:
            self.switchType = EXCEL_PROFILE
        else:
            self.switchType = EXCEL_LOCATION
        self.setWindowTitle(self.switchType)
        
    def testBtn_mergeExcel(self):
        if len(self.excelPathList) >= 2:
            return True
        else:
            return False

    def clickBtn_addExcel(self):
        temp_excelPathList, _ = QFileDialog.getOpenFileNames(self, "选择【每日健康打卡】导出数据文件", confAct.DATA_DIR, "Excel files(*.xlsx , *.xls)")
        #FIXME:无法显示
        self.updateStatus('添加中，请稍候...')
        errMsg = []
        for excelPath in temp_excelPathList:
            if excelPath not in self.excelPathList:
                testRes = testRawExcel(excelPath)
                if testRes['code'] == 1:
                    self.widget_pathListBox.addItem(f'{excelPath}')
                    self.excelPathList.append(excelPath)
                else:
                    errMsg.append(testRes['msg'])
        if len(errMsg) != 0:
            self.showMessageBox(errMsg).exec()
        self.updateStatus('添加完毕，请点击【开始生成】')
        self.btn_mergeExcel.setEnabled(self.testBtn_mergeExcel())

    def clickbtn_deleteExcel(self):
        selectedItem = self.widget_pathListBox.currentItem()
        selectedRow = self.widget_pathListBox.row(selectedItem)
        if selectedRow != -1:
            self.widget_pathListBox.takeItem(selectedRow)
            selectedPath = selectedItem.text()
            self.excelPathList.remove(selectedPath)
            self.btn_mergeExcel.setEnabled(self.testBtn_mergeExcel())

    def clickbtn_clearExcel(self):
        count = self.widget_pathListBox.count()
        for i in range(count):
            self.widget_pathListBox.takeItem(0)
        self.excelPathList = []
        self.btn_mergeExcel.setEnabled(False)

    def clickbtn_mergeExcel(self):
        if self.switchType == EXCEL_LOCATION:
            self.mergeExcel_Location()
        else:
            self.mergeExcel_Profile()

    def mergeExcel_Location(self):
        now = time_strftime("%Y-%m-%d-%H-%M-%S", time_localtime())
        self.excelOutputPath, _ = QFileDialog.getSaveFileName(self, "选择Excel保存路径", '每日健康打卡位置汇总（%s）' % now, "Excel files(*.xlsx , *.xls)")
        self.excelDir, self.excelName = os_path_split(self.excelOutputPath)

        if len(self.excelOutputPath.strip()) != 0:
            self.mergeExcelStart()
            self.mergeExcel()
            self.mergeExcelEnd()

    def mergeExcel_Profile(self):
        self.profilesDir = QFileDialog.getExistingDirectory(self, '选择保存一人一档的文件夹路径', confAct.DATA_DIR)
        if self.profilesDir != '':
            self.mergeExcelStart()
            self.createProfiles()
            self.mergeExcelEnd()

    def clickbtn_openDir(self):
        if self.switchType == EXCEL_LOCATION:
            openDir = self.excelDir
        else:
            openDir = self.profilesDir
        if openDir != '':
            openDir = os_path_normpath(openDir)
            os_system("explorer.exe %s" % openDir)

    def mergeExcel(self):

        self.mergeExcelWorker = MergeExcelWorker(self.excelPathList, self.excelOutputPath)
        self.mergeExcelThread = QThread()
        self.mergeExcelWorker._signal.connect(self.updateStatus)
        self.mergeExcelWorker.moveToThread(self.mergeExcelThread)
        self.mergeExcelWorker._finished.connect(self.mergeExcelThread.quit)
        self.mergeExcelThread.started.connect(self.mergeExcelWorker.work)
        self.mergeExcelThread.finished.connect(lambda: self.updateStatus(f'生成位置文件结束!【打开路径】吧~ 文件名：{self.excelName}'))

        self.mergeExcelThread.start()

    def createProfiles(self):
        self.createProfilesWorker = CreateProfilesWorker(self.excelPathList, self.profilesDir)
        self.createProfilesThread = QThread()
        self.createProfilesWorker._signal.connect(self.updateStatus)
        self.createProfilesWorker.moveToThread(self.createProfilesThread)
        self.createProfilesWorker._finished.connect(self.createProfilesThread.quit)
        self.createProfilesThread.started.connect(self.createProfilesWorker.work)
        self.createProfilesThread.finished.connect(lambda: self.updateStatus(f'生成位置文件结束!【打开路径】吧~ 文件夹路径：{self.profilesDir}'))

        self.createProfilesThread.start()

    def mergeExcelStart(self):
        if self.setBtnsUnabled():
            self.updateStatus(f'开始整合{len(self.excelPathList)}个Excel文件...')
        
    def mergeExcelEnd(self):
        self.excelPathList = []
        self.clickbtn_clearExcel()
        self.setBtnsEnabled()

    def updateStatus(self, msg):
        self.status_label.setText(msg)

    def setBtnsUnabled(self):
        try:
            #FIXME:无法将按钮设置成不可用
            self.btn_mergeExcel.setEnabled(False)
            self.btn_addExcel.setEnabled(False)
            self.btn_deleteExcel.setEnabled(False)
            self.btn_clearExcel.setEnabled(False)
            self.btn_openMergeExcelDir.setEnabled(False)
            return True
        except:
            return False

    def setBtnsEnabled(self):
        self.btn_addExcel.setEnabled(True)
        self.btn_deleteExcel.setEnabled(True)
        self.btn_clearExcel.setEnabled(True)
        self.btn_openMergeExcelDir.setEnabled(True)

    def showMessageBox(self, errMsg):
        msgBox = QMessageBox()
        msgBox.setWindowTitle('错误')
        msgBox.setWindowIcon(QIcon(APP_ICON_PATH))
        msgBox.setText(f"友情提示：{len(errMsg)}个文件添加错误了！请重试！请使用钉钉【员工健康】导出的原始文件！")
        msgBox.setDetailedText('\n\n'.join(errMsg))
        return msgBox
