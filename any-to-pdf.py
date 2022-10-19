from appUI import Ui_mainWindow
from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QMovie
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.uic import loadUiType
import main
import sys


class MainThread(QThread):
    def __init__(self):
        super(MainThread, self).__init__()

startFunction = MainThread()


class GuiStart(QMainWindow):

    def __init__(self):
        super().__init__()
        self.appUI = Ui_mainWindow()
        self.appUI.setupUi(self)

        self.input_path = ""
        self.output_path = ""

        self.appUI.browse1.clicked.connect(self.browseFiles)
        self.appUI.browse2.clicked.connect(self.browseDir)
        startFunction.start()

    def browseFiles(self):
        self.fname = QFileDialog.getOpenFileName(self, "Open File", "C://", "files(*.txt)")
        self.appUI.inputPath.setText(self.fname[0])
        self.appUI.outputPath.setText("/".join(self.fname[0].split("/")[:-1]))
        self.input_path = str(self.fname[0])
        self.output_path = str("/".join(self.fname[0].split("/")[:-1]))

    def browseDir(self):
        self.dir_name = QFileDialog.getExistingDirectory(self, "Open Folder")
        self.appUI.outputPath.setText(str(self.dir_name))
        self.output_path = str(self.dir_name)
        


app = QApplication(sys.argv)
GUI_app = GuiStart()
GUI_app.show()
exit(app.exec_())
