#-*-coding:utf-8-*-
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFileDialog, QGroupBox, QRadioButton, QScrollArea, QLineEdit
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import Qt, QThread

from os import getcwd as os_getcwd, system as os_system
from os.path import split as os_path_split, normpath as os_path_normpath
from time import strftime as time_strftime, localtime as time_localtime

import sys
sys.path.append("..")
from configure.logging_setting import Log
from configure.default_setting import *

log = Log(__name__).getLog()

class SettingWindow(QWidget):
    def __init__(self):
       super().__init__()

       self.initUI()

    def initUI(self):

        font_Yahei = QFont(FONT_NAME_YAHEI)

        self.setWindowTitle('设置')
        self.setWindowIcon(QIcon(APP_ICON_PATH))
        self.setGeometry(450, 200, 290, 400)
        
        layout_main = QVBoxLayout()
        self.setLayout(layout_main)
        
        self.setFont(font_Yahei)

        #-----说明模块---------------------------------------
        widget_intro = QWidget()
        layout_intro = QVBoxLayout()

        label_intro = QLabel('设置说明')
        
        widget_intro.setLayout(layout_intro)

        layout_intro.addWidget(label_intro)

        layout_main.addWidget(widget_intro)

        #-----设置选项---------------------------------------
        
        scroll = QScrollArea()
        
        widget_settings = QWidget()
        layout_settings = QVBoxLayout()
        layout_settings.addStretch()

        group_map = QGroupBox('选择地图API')
        layout_map = QVBoxLayout()
        group_map.setLayout(layout_map)
        self.group_map_amap = QRadioButton('高德地图API')
        self.group_map_qq = QRadioButton('腾讯地图API')
        self.group_map_amap.setChecked(True)
        layout_map.addWidget(self.group_map_amap)
        layout_map.addWidget(self.group_map_qq)
        layout_settings.addWidget(group_map)

        group_type = QGroupBox('功能类型')
        layout_type = QVBoxLayout()
        group_type.setLayout(layout_type)
        group_type_clockIn = QRadioButton('每日健康打卡')
        group_type_signOn = QRadioButton('群签到')
        group_type_clockIn.setChecked(True)
        layout_type.addWidget(group_type_clockIn)
        layout_type.addWidget(group_type_signOn)
        layout_settings.addWidget(group_type)

        group_distance = QGroupBox('位移分析是否显示距离')
        layout_distance = QVBoxLayout()
        group_distance.setLayout(layout_distance)
        group_distance_show = QRadioButton('是')
        group_distance_hide = QRadioButton('否')
        group_distance_show.setChecked(True)
        layout_distance.addWidget(group_distance_show)
        layout_distance.addWidget(group_distance_hide)
        layout_settings.addWidget(group_distance)

        group_sheet = QGroupBox('处理Excel时')
        layout_sheet = QVBoxLayout()
        group_sheet.setLayout(layout_sheet)
        group_sheet_first = QRadioButton('只处理第1个工作表')
        group_sheet_all = QRadioButton('处理全部工作表')
        group_sheet_first.setChecked(True)
        layout_sheet.addWidget(group_sheet_first)
        layout_sheet.addWidget(group_sheet_all)
        layout_settings.addWidget(group_sheet)

        group_header = QGroupBox('读取Excel包含表头名称')
        layout_header = QVBoxLayout()
        group_header.setLayout(layout_header)
        group_header_staff = QRadioButton('工号、提交人、当前时间,当前日期')
        group_header_stu = QRadioButton('学号、姓名、当前时间,当前日期')
        group_header_staff.setChecked(True)
        layout_header.addWidget(group_header_staff)
        layout_header.addWidget(group_header_stu)
        layout_settings.addWidget(group_header)

        group_newheader = QGroupBox('生成Excel包含表头名称')
        layout_newheader = QVBoxLayout()
        group_newheader.setLayout(layout_newheader)
        group_newheader_staff = QRadioButton('工号、提交人、x月x日')
        group_newheader_stu = QRadioButton('学号、姓名、x月x日')
        group_newheader_staff.setChecked(True)
        layout_newheader.addWidget(group_newheader_staff)
        layout_newheader.addWidget(group_newheader_stu)
        layout_settings.addWidget(group_newheader)

        group_row = QGroupBox('从Excel的第几行开始读取')
        layout_row = QVBoxLayout()
        group_row.setLayout(layout_row)
        input_row = QLineEdit('1')
        layout_row.addWidget(input_row)
        layout_settings.addWidget(group_row)

        group_dirpath = QGroupBox('文件夹路径')
        layout_dirpath = QVBoxLayout()
        group_dirpath.setLayout(layout_dirpath)
        self.label_setDirpath = QLineEdit('./')
        btn_chooseDirpath = QPushButton('修改文件夹路径')
        btn_chooseDirpath.clicked.connect(self.clickbtn_chooseDirpath)
        layout_dirpath.addWidget(self.label_setDirpath)
        layout_dirpath.addWidget(btn_chooseDirpath)
        layout_settings.addWidget(group_dirpath)

        layout_main.addWidget(scroll)

        widget_settings.setLayout(layout_settings)
        scroll.setWidget(widget_settings)
        
        layout_main.addWidget(scroll)

        #-----按钮---------------------------------------
        widget_buttons = QWidget()
        layout_buttons = QHBoxLayout()
        widget_buttons.setLayout(layout_buttons)

        self.btn_defaultSettings = QPushButton('恢复默认设置')
        self.btn_defaultSettings.clicked.connect(self.clickbtn_defaultSettings)
        layout_buttons.addWidget(self.btn_defaultSettings)

        self.btn_confirmSettings = QPushButton('确定')
        self.btn_confirmSettings.clicked.connect(self.clickbtn_confirmSettings)
        layout_buttons.addWidget(self.btn_confirmSettings)

        self.btn_cancelSettings = QPushButton('取消')
        self.btn_cancelSettings.clicked.connect(self.clickbtn_cancelSettings)
        layout_buttons.addWidget(self.btn_cancelSettings)
        
        layout_main.addWidget(widget_buttons)

        #-----状态显示---------------------------------------
        widget_status = QWidget()
        layout_status = QHBoxLayout()
        widget_status.setLayout(layout_status)

        self.status_label = QLabel('【设置】需谨慎~')
        
        author_label = QLabel()

        author_label.setAlignment(Qt.AlignRight)
        author_label.setText(f'<a href={AUTHOR_GITHUB_URL} style="text-decoration:none;color:black">{AUTHOR}</a>')
        author_label.setOpenExternalLinks(True)
        author_label.setToolTip(AUTHOR_TIP)

        layout_status.addWidget(self.status_label)
        layout_status.addWidget(author_label)

        layout_main.addWidget(widget_status)

    def clickbtn_chooseDirpath(self):
        self.DirPath = QFileDialog.getExistingDirectory(self, '选择默认文件夹路径', './')
        log.debug(f'选择文件夹路径：{self.DirPath}')
        if self.DirPath.strip() != '':
            self.label_setDirpath.setText(self.DirPath)

    def clickbtn_defaultSettings(self):
        log.debug('点击了【恢复默认设置】')
        
    def clickbtn_confirmSettings(self):
        if self.group_map_amap.isChecked():
            log.debug('选择了高德地图')
        elif self.group_map_qq.isChecked():
            log.debug('选择了腾讯地图')
        log.debug('点击了【确定】') 

    def clickbtn_cancelSettings(self):
        log.debug('点击了【取消】')

if __name__ == "__main__":
    
    app = QApplication(sys.argv)
    mainWin = SettingWindow()
    mainWin.show()
    sys.exit(app.exec_())