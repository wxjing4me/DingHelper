#-*-coding:utf-8-*-
from PyQt5.QtWidgets import QApplication
from windows.MainWin import MainWindow
from windows.ExcelWin import ExcelWindow
import sys

from config.logging_setting import Log

if __name__ == "__main__":
    try:
        log = Log(__name__).getLog()
        app = QApplication(sys.argv)
        mainWin = MainWindow()
        excelWin = ExcelWindow()
        mainWin.show()
        log.critical('App启动成功啦！')
        mainWin.btn_mergeExcel.clicked.connect(excelWin.show)
        sys.exit(app.exec_())
    except Exception as e:
        log.critical(f'App启动失败 - {e}', exc_info=True)