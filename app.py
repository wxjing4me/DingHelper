from PyQt5.QtWidgets import QApplication
from windows.MainWin import MainWindow
from windows.ExcelWin import ExcelWindow
import sys

if __name__ == "__main__":
    
    app = QApplication(sys.argv)
    mainWin = MainWindow()
    excelWin = ExcelWindow()
    mainWin.show()
    mainWin.btn_mergeExcel.clicked.connect(excelWin.show)
    sys.exit(app.exec_())