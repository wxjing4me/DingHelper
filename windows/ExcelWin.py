from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFileDialog, QCheckBox, QListWidget, QListWidgetItem
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import Qt, QThread

import os
import time
import sys

from functions.excel_action import MergeExcelWorker

class ExcelWindow(QWidget):
    def __init__(self):
       super().__init__()

       self.excelPathList = []
       self.excelOutputPath = ''

       self.initUI()

    def initUI(self):

        font_Yahei = QFont("Microsoft YaHei")

        self.setWindowTitle('整合文件')
        self.setWindowIcon(QIcon('images/favicon.ico'))
        self.setGeometry(300, 100, 600, 400)
        
        layout_main = QVBoxLayout()
        self.setLayout(layout_main)
        
        self.setFont(font_Yahei)


        #-----说明模块---------------------------------------
        widget_intro = QWidget()
        layout_intro = QVBoxLayout()

        label_intro = QLabel('说明：添加的文件必须是从钉钉【员工健康】中导出的原始数据表')
        
        widget_intro.setLayout(layout_intro)

        layout_intro.addWidget(label_intro)

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

        self.btn_mergeExcel = QPushButton('开始整合')
        self.btn_mergeExcel.clicked.connect(self.clickbtn_mergeExcel)
        self.btn_mergeExcel.setEnabled(False)
        layout_buttons.addWidget(self.btn_mergeExcel)

        self.btn_openMergeExcelDir = QPushButton('打开路径')
        self.btn_openMergeExcelDir.clicked.connect(self.clickbtn_openMergeExcelDir)
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

        self.status_label = QLabel('请点击【添加文件】开始')
        
        author_label = QLabel('@wxjing')

        author_label.setAlignment(Qt.AlignRight)
        author_label.setText('<a href="https://github.com/wxjing4me" style="text-decoration:none;color:black">@wxjing</a>')
        author_label.setOpenExternalLinks(True)
        author_label.setToolTip('你发现了什么？点击有惊喜哦~')

        layout_status.addWidget(self.status_label)
        layout_status.addWidget(author_label)

        layout_main.addWidget(widget_status)
    

    def testBtn_mergeExcel(self):
        if len(self.excelPathList) >= 1:
            return True
        else:
            return False

    def clickBtn_addExcel(self):
        temp_excelPathList, _ = QFileDialog.getOpenFileNames(self, "选择Excel文件", os.getcwd(), "Excel files(*.xlsx , *.xls)")
        for excelPath in temp_excelPathList:
            if excelPath not in self.excelPathList:
                excelDir, excelName = os.path.split(excelPath)
                box_path = QCheckBox(f'{excelName} ({excelDir})')
                self.widget_pathListBox.addItem(f'{excelPath}')
                self.excelPathList.append(excelPath)
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
        # 选择输出文件路径
        now = time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime())
        self.excelOutputPath, _ = QFileDialog.getSaveFileName(self, "选择Excel保存路径", '每日健康打卡位置汇总（%s）' % now, "Excel files(*.xlsx , *.xls)")
        self.excelDir, self.excelName = os.path.split(self.excelOutputPath)

        if len(self.excelOutputPath.strip()) != 0:
            self.mergeExcelStart()
            self.mergeExcel()
            self.mergeExcelEnd()

    def clickbtn_openMergeExcelDir(self):
        excelDir = os.path.normpath(self.excelDir)
        os.system("explorer.exe %s" % excelDir)

    def mergeExcel(self):

        self.mergeExcelWorker = MergeExcelWorker(self.excelPathList, self.excelOutputPath)
        self.mergeExcelThread = QThread()
        self.mergeExcelWorker._signal.connect(self.updateStatus)
        self.mergeExcelWorker.moveToThread(self.mergeExcelThread)
        self.mergeExcelWorker._finished.connect(self.mergeExcelThread.quit)
        self.mergeExcelThread.started.connect(self.mergeExcelWorker.work)
        self.mergeExcelThread.finished.connect(lambda: self.updateStatus(f'整合结束，文件名：{self.excelName}'))

        self.mergeExcelThread.start()

    def mergeExcelStart(self):
        self.updateStatus(f'开始整合{len(self.excelPathList)}个Excel文件...')
        self.setBtnsUnabled()

    def mergeExcelEnd(self):
        self.excelPathList = []
        self.clickbtn_clearExcel()
        self.setBtnsEnabled()

    def updateStatus(self, msg):
        self.status_label.setText(msg)

    def setBtnsUnabled(self):
        self.btn_mergeExcel.setEnabled(False)
        self.btn_addExcel.setEnabled(False)
        self.btn_deleteExcel.setEnabled(False)
        self.btn_clearExcel.setEnabled(False)
        self.btn_openMergeExcelDir.setEnabled(False)

    def setBtnsEnabled(self):
        self.btn_addExcel.setEnabled(True)
        self.btn_deleteExcel.setEnabled(True)
        self.btn_clearExcel.setEnabled(True)
        self.btn_openMergeExcelDir.setEnabled(True)

if __name__ == "__main__":
    
    app = QApplication(sys.argv)
    win = ExcelWindow()
    win.show()
    sys.exit(app.exec_())