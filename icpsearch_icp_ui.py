from ui_signals import *
from ui_worker import *
from ui_mainwindow import MainWindow
from ui_entry import *
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication
import sys

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 