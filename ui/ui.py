from kiwoom.kiwoom import *
from PyQt5.QtWidgets import *
import sys

class Ui_class():
    def __init__(self):
        print("Ui 클래스")

        self.app = QApplication(sys.argv)
        self.kiwoom = Kiwoom()
        self.app.exec_()