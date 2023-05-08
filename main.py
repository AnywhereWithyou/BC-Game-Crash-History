# Import required modules
from datetime import datetime
import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication
import mainwindow

from CrashHistory_excel import *


class MainWindowApp(QtWidgets.QMainWindow, mainwindow.Ui_MainWindow):
    def __init__(self, parent=None):
        super(MainWindowApp, self).__init__(parent)
        self.setupUi(self)
        self.runBtn.clicked.connect(self.runBtnAction)

    def runBtnAction(self):
        url = str(self.url_edit.text())
        rowNo = self.rowNo_spin.value()
        rowDelta = self.rowDelta_spin.value()
        CrashHistory(url, rowNo, rowDelta)
def main():   
    app = QApplication(sys.argv)
    form = MainWindowApp()
    form.show()
    app.exec_()
if __name__ == '__main__':
    main()