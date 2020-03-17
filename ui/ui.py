from kiwoom.kiwoom import *

import sys
from PyQt5.QtWidgets import *


class Ui_class():
    def __init__(self):
        self.app = QApplication(sys.argv)

        self.kiwoom = Kiwoom()

        self.app.exec_()