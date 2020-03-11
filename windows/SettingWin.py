#-*-coding:utf-8-*-
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFileDialog, QGroupBox, QRadioButton, QScrollArea, QLineEdit
from PyQt5.QtGui import QIcon, QFont, QIntValidator
from PyQt5.QtCore import Qt, QThread, pyqtSignal

import sys
from os import getcwd as os_getcwd, system as os_system
from os.path import split as os_path_split, normpath as os_path_normpath, join as os_path_join
from time import strftime as time_strftime, localtime as time_localtime
from json import load as json_load, dump as json_dump

from configure.logging_action import Log
from configure.config_values import *
from configure import config_action as confAct

log = Log(__name__).getLog()

class SettingWindow(QWidget):

    _changedSettings = pyqtSignal()

    def __init__(self):
       super().__init__()

       self.initUI()

    def initUI(self):

        font_Yahei = QFont(FONT_NAME_YAHEI)

        self.setWindowTitle('设置')
        self.setWindowIcon(QIcon(APP_ICON_PATH))
        self.setGeometry(450, 200, 290, 400)
        #TODO:窗口大小不可变
        
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
        
        self.scroll = QScrollArea()
        
        widget_settings = QWidget()
        layout_settings = QVBoxLayout()
        layout_settings.addStretch()

        group_map = QGroupBox('选择地图API')
        layout_map = QVBoxLayout()
        group_map.setLayout(layout_map)
        self.group_map_AMAP = QRadioButton('高德地图API')
        self.group_map_QQ = QRadioButton('腾讯地图API')
        layout_map.addWidget(self.group_map_AMAP)
        layout_map.addWidget(self.group_map_QQ)
        layout_settings.addWidget(group_map)

        group_type = QGroupBox('功能类型')
        layout_type = QVBoxLayout()
        group_type.setLayout(layout_type)
        self.group_type_clockIn = QRadioButton('每日健康打卡')
        self.group_type_signOn = QRadioButton('群签到')
        layout_type.addWidget(self.group_type_clockIn)
        layout_type.addWidget(self.group_type_signOn)
        layout_settings.addWidget(group_type)

        group_distance = QGroupBox('位移分析是否显示距离')
        layout_distance = QVBoxLayout()
        group_distance.setLayout(layout_distance)
        self.group_distance_show = QRadioButton('是')
        self.group_distance_hide = QRadioButton('否')
        layout_distance.addWidget(self.group_distance_show)
        layout_distance.addWidget(self.group_distance_hide)
        layout_settings.addWidget(group_distance)

        group_sheet = QGroupBox('处理Excel时')
        layout_sheet = QVBoxLayout()
        group_sheet.setLayout(layout_sheet)
        self.group_sheet_first = QRadioButton('只处理第1个工作表')
        self.group_sheet_all = QRadioButton('处理全部工作表')
        layout_sheet.addWidget(self.group_sheet_first)
        layout_sheet.addWidget(self.group_sheet_all)
        layout_settings.addWidget(group_sheet)

        group_header = QGroupBox('读取Excel包含表头名称')
        layout_header = QVBoxLayout()
        group_header.setLayout(layout_header)
        self.group_header_staff = QRadioButton('工号、提交人、当前时间,当前日期')
        self.group_header_stu = QRadioButton('学号、姓名、当前时间,当前日期')
        layout_header.addWidget(self.group_header_staff)
        layout_header.addWidget(self.group_header_stu)
        layout_settings.addWidget(group_header)

        group_newheader = QGroupBox('生成Excel包含表头名称')
        layout_newheader = QVBoxLayout()
        group_newheader.setLayout(layout_newheader)
        self.group_newheader_staff = QRadioButton('工号、提交人、x月x日')
        self.group_newheader_stu = QRadioButton('学号、姓名、x月x日')
        layout_newheader.addWidget(self.group_newheader_staff)
        layout_newheader.addWidget(self.group_newheader_stu)
        layout_settings.addWidget(group_newheader)

        group_row = QGroupBox('从Excel的第几行开始读取')
        layout_row = QVBoxLayout()
        group_row.setLayout(layout_row)
        self.input_row = QLineEdit()
        self.input_row.setValidator(QIntValidator(1, 65535))
        layout_row.addWidget(self.input_row)
        layout_settings.addWidget(group_row)

        group_dirpath = QGroupBox('文件夹路径')
        layout_dirpath = QVBoxLayout()
        group_dirpath.setLayout(layout_dirpath)
        self.label_dirpath = QLabel()
        btn_chooseDirpath = QPushButton('修改文件夹路径')
        btn_chooseDirpath.clicked.connect(self.clickbtn_chooseDirpath)
        layout_dirpath.addWidget(self.label_dirpath)
        layout_dirpath.addWidget(btn_chooseDirpath)
        layout_settings.addWidget(group_dirpath)

        layout_main.addWidget(self.scroll)

        widget_settings.setLayout(layout_settings)
        self.scroll.setWidget(widget_settings)
        
        layout_main.addWidget(self.scroll)

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
        self.btn_cancelSettings.clicked.connect(self.close)
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

    def openWin(self):
        self.setSettings()
        self.scroll.verticalScrollBar().setValue(0)
        self.show()

    def setSettings(self):
        config = confAct.updateSettings()
        eval('self.group_map_'+config["MAP_TYPE"]).setChecked(True)
        # self.group_map_AMAP.setChecked(True)
        eval('self.group_type_'+config["FUNC_TYPE"]).setChecked(True)
        # self.group_type_clockIn.setChecked(True)
        eval('self.group_distance_'+config["SHOW_DISTANCE"]).setChecked(True)
        # self.group_distance_show.setChecked(True)
        eval('self.group_sheet_'+config["HANDLE_SHEET"]).setChecked(True)
        # self.group_sheet_first.setChecked(True)
        self.group_header_staff.setChecked(True)
        self.group_newheader_staff.setChecked(True)
        self.input_row.setText(str(config["START_ROW"]))
        self.label_dirpath.setText(config["DATA_DIR"])

    def clickbtn_chooseDirpath(self):
        self.DirPath = QFileDialog.getExistingDirectory(self, '选择默认文件夹路径', './')
        log.debug(f'选择文件夹路径：{self.DirPath}')
        if self.DirPath.strip() != '':
            self.label_dirpath.setText(self.DirPath)

    def clickbtn_defaultSettings(self):
        try:
            os.remove(USER_SETTINGS_JSON)
        except:
            pass
        finally:
            self.openWin()
        
    def clickbtn_confirmSettings(self):
        config = {}
        if self.group_map_AMAP.isChecked():
            config["MAP_TYPE"] = "AMAP"
        elif self.group_map_QQ.isChecked():
            config["MAP_TYPE"] = "QQ"
        if self.group_distance_show.isChecked():
            config["SHOW_DISTANCE"] = 'show'
        else:
            config["SHOW_DISTANCE"] = 'hide'
        if self.group_type_clockIn.isChecked():
            config["FUNC_TYPE"] = 'clockIn'
        elif self.group_type_signOn.isChecked():
            config["FUNC_TYPE"] = 'signOn'
        if self.group_sheet_first.isChecked():
            config["HANDLE_SHEET"] = "first"
        else:
            config["HANDLE_SHEET"] = "all"
        if self.input_row.text().strip() != '':
            config["START_ROW"] = int(self.input_row.text())
        else:
            config["START_ROW"] = 1
        if self.label_dirpath.text().strip() != '':
            config["DATA_DIR"] = self.label_dirpath.text()
        else:
            config["DATA_DIR"] = DATA_DIR
        if confAct.saveConfig(config):
            self._changedSettings.emit()
            self.close()
