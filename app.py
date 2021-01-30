#-*-coding:utf-8-*-
from PyQt5.QtWidgets import QApplication
from windows.MainWin import MainWindow
from windows.ExcelWin import ExcelWindow
from windows.SettingWin import SettingWindow
from windows.ProfileWin import ProfileWindow
import sys

from configure.logging_action import Log

if __name__ == "__main__":
    try:
        log = Log(__name__).getLog()
        app = QApplication(sys.argv)
        mainWin = MainWindow()
        excelWin = ExcelWindow()
        settingWin = SettingWindow()
        profileWin = ProfileWindow()
        mainWin.show()
        log.critical('App启动成功啦！')
        mainWin.btn_profileWin.clicked.connect(profileWin.openWin)
        mainWin.btn_mergeExcel.clicked.connect(excelWin.openWin)
        mainWin.btn_setting.clicked.connect(settingWin.openWin)
        settingWin._changedSettings.connect(mainWin.refreshUI)
        sys.exit(app.exec_())
    except Exception as e:
        log.critical(f'App启动失败 - {e}', exc_info=True)