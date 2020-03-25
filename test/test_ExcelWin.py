#-*-coding:utf-8-*-
from PyQt5.QtWidgets import QApplication
import sys
sys.path.append("..")
from windows.ExcelWin import ExcelWindow

if __name__ == "__main__":
    app = QApplication(sys.argv)
    excelWin = ExcelWindow()
    excelWin.show()
    sys.exit(app.exec_())